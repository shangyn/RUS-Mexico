# -*- coding: utf-8 -*-
"""泰国2欠款发货额度：模块筛选、回款分摊与历史台账。"""
import copy
import json
import os
import pickle
import re
from datetime import datetime

import pandas as pd
import yaml

import quota_core as qc
import quota_shared

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MODULES_CONFIG = os.path.join(BASE_DIR, "quota_modules.yaml")
WORKSPACE_ROOT = os.path.join(BASE_DIR, "workspaces", "quota")

_RE_LADDER_RANGE = re.compile(r"(\d{1,4})\s*#\s*[-~]\s*(\d{1,4})\s*#")
_RE_LADDER = re.compile(r"(\d{1,4})\s*#")


class MiniLogger:
    def __init__(self):
        self.errors = []
        self.notes = []

    def add_error(self, contract_no, message, source=""):
        self.errors.append([str(contract_no), str(message), str(source)])

    def add_note(self, contract_no, message):
        self.notes.append([str(contract_no), str(message)])


def load_modules():
    with open(MODULES_CONFIG, "r", encoding="utf-8") as f:
        return yaml.safe_load(f).get("modules", [])


def get_module(module_id):
    for module in load_modules():
        if module.get("id") == module_id:
            return module
    return None


def module_root(module_id):
    return os.path.join(WORKSPACE_ROOT, module_id)


def module_dirs(module_id):
    base = module_root(module_id)
    return {
        "data": os.path.join(base, "数据源"),
        "out": os.path.join(base, "输出"),
    }


def ledger_path(module_id):
    return os.path.join(module_root(module_id), "欠款台账.json")


def result_path(module_id):
    return os.path.join(module_dirs(module_id)["out"], "_result_" + module_id + ".json")


def _num(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "", regex=False),
                         errors="coerce")


def _ladder_no(bid):
    m = re.search(r"/(\d+)#\s*$", str(bid))
    return int(m.group(1)) if m else None


def _key(contract, ladder):
    return str(contract).strip().upper() + "|" + str(int(ladder))


def _clean_text(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    return "" if text.lower() in ("nan", "none", "nat") else text


def _add_months(date_value, months=10):
    if date_value is None or pd.isna(date_value):
        return ""
    return (pd.Timestamp(date_value) + pd.DateOffset(months=months)).date().isoformat()


def _table_c_cache_path():
    return os.path.join(quota_shared.SHARED_ROOT, "_table_c_raw_v1.pkl")


def _table_c_signature(path):
    stat = os.stat(path)
    return [os.path.abspath(path), int(stat.st_size), int(stat.st_mtime_ns)]


def _read_table_c_frame(path, need):
    """读取表C所需列；同一文件未变化时复用本地缓存，避免每次重读整本台账。"""
    cache_path = _table_c_cache_path()
    signature = _table_c_signature(path)
    try:
        with open(cache_path, "rb") as f:
            payload = pickle.load(f)
        if (payload.get("signature") == signature
                and payload.get("columns") == list(need)
                and isinstance(payload.get("df"), pd.DataFrame)):
            return payload["df"]
    except Exception:
        pass

    df = pd.read_excel(path, sheet_name="sheet1", usecols=need)
    try:
        os.makedirs(os.path.dirname(cache_path), exist_ok=True)
        payload = {"signature": signature, "columns": list(need), "df": df}
        tmp = cache_path + ".tmp"
        with open(tmp, "wb") as f:
            pickle.dump(payload, f, protocol=4)
        os.replace(tmp, cache_path)
    except Exception:
        pass
    return df


def read_module_countries(path_d, module_name):
    df = pd.read_excel(path_d, sheet_name="Sheet3")
    for col in ("国家", "2026年对应模块"):
        if col not in df.columns:
            raise RuntimeError("表D Sheet3 缺少列：" + col)
    names = df["2026年对应模块"].astype(str).str.strip()
    countries = (df.loc[names == str(module_name).strip(), "国家"]
                 .dropna().astype(str).str.strip())
    countries = sorted({c for c in countries if c and c.lower() != "nan"})
    if not countries:
        raise RuntimeError("表D 中没有找到模块：" + str(module_name))
    return countries


def _norm_match(value):
    text = _clean_text(value)
    return re.sub(r"\s+", " ", text).casefold()


def read_scope(path_c, path_d, module):
    match_mode = str(module.get("match_mode") or "module").strip().lower()
    agent_name = str(module.get("agent_name") or "").strip()
    need = ["合同编号", "标的编号", "国家", "签订日期", "组A日期", "产品状态",
            "已回款比例（原币）", "合同额（人民币）", "合同额（原币）", "汇率", "币种", "代理商"]
    df = _read_table_c_frame(path_c, need)
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise RuntimeError("表C缺少列：" + "、".join(missing))

    df["国家"] = df["国家"].astype(str).str.strip()
    df["代理商"] = df["代理商"].astype(str).str.strip()
    if match_mode == "agent":
        if not agent_name:
            raise RuntimeError("模块配置缺少 agent_name")
        target = _norm_match(agent_name)
        matched = df[df["代理商"].map(_norm_match) == target].copy()
        if matched.empty:
            raise RuntimeError("表C 中没有找到代理商：" + agent_name)
        countries = sorted({c for c in matched["国家"].astype(str).str.strip()
                            if c and c.lower() != "nan"})
        df = matched
    else:
        if not path_d:
            raise RuntimeError("模块匹配模式需要表D，请先上传表D")
        countries = read_module_countries(path_d, module.get("module_name") or module.get("name"))
        df = df[df["国家"].isin(countries)].copy()
    df["签订日期"] = pd.to_datetime(df["签订日期"], errors="coerce")
    df["组A日期"] = pd.to_datetime(df["组A日期"], errors="coerce")
    df["产品状态"] = df["产品状态"].astype(str).str.strip()
    for col in ("合同额（人民币）", "合同额（原币）", "汇率", "已回款比例（原币）"):
        df[col] = _num(df[col])
    df["合同额（人民币）"] = df["合同额（人民币）"].fillna(0)
    df["合同额（原币）"] = df["合同额（原币）"].fillna(0)

    date_start = pd.Timestamp(module.get("date_start") or "2024-01-01")
    exclude_status = [str(x).strip() for x in (module.get("exclude_status") or ["已作废"])]
    scope = df[(df["签订日期"] >= date_start) & (~df["产品状态"].isin(exclude_status))].copy()
    scope["_contract"] = scope["合同编号"].astype(str).str.strip()
    scope["_bid"] = scope["标的编号"].astype(str).str.strip().str.upper()
    scope["_ladder"] = scope["_bid"].apply(_ladder_no)
    scope = scope[scope["_ladder"].notna()].copy()
    scope["_ladder"] = scope["_ladder"].astype(int)

    # 同一合同+梯号必须唯一；标的编号偶有空格/大小写差异时仍合并为一条。
    ladder = scope.groupby(["_contract", "_ladder"], as_index=False).agg({
        "国家": "first",
        "_bid": "first",
        "标的编号": "first",
        "合同额（人民币）": "sum",
        "合同额（原币）": "sum",
        "币种": "first",
        "汇率": "first",
        "签订日期": "first",
        "组A日期": "max",
        "产品状态": "first",
        "已回款比例（原币）": "first",
    })
    ladder["ladder_amount_wan"] = ladder["合同额（人民币）"] / 10000.0
    contract_total = ladder.groupby("_contract")["ladder_amount_wan"].sum().to_dict()
    shipped_total = (ladder[ladder["组A日期"].notna()]
                     .groupby("_contract")["ladder_amount_wan"].sum().to_dict())
    ladder["contract_amount_wan"] = ladder["_contract"].map(contract_total).fillna(0.0)
    ladder["shipped_contract_amount_wan"] = ladder["_contract"].map(shipped_total).fillna(0.0)
    return {
        "scope": scope,
        "ladder": ladder,
        "countries": countries,
        "match_mode": match_mode,
        "agent_name": agent_name,
        "counts": {
            "rows": int(len(scope)),
            "contracts": int(scope["_contract"].nunique()),
            "shipped_rows": int(ladder["组A日期"].notna().sum()),
            "shipped_contracts": int(ladder.loc[ladder["组A日期"].notna(), "_contract"].nunique()),
            "unshipped_rows": int(ladder["组A日期"].isna().sum()),
        },
    }


def read_payments(config, data_dir, scope, contract_list):
    paths = config["paths"]
    path_b = qc.find_file(data_dir, paths["table_b"])
    path_e = qc.find_file(data_dir, paths["table_e"])
    path_r = qc.find_file(data_dir, paths["table_r"])
    path_ex = qc.find_file(data_dir, paths["exchange_rate"])
    logger = MiniLogger()
    rate_map = qc.build_rate_map(scope)
    exchange_rates = qc.load_exchange_rates(path_ex)
    df_b = qc.read_table_b(config, path_b)
    df_e = read_table_e_any(config, path_e)
    df_r = qc.read_table_r(config, path_r)
    pay_b = qc.calc_payment_table_b(df_b, rate_map, contract_list, config, logger, exchange_rates)
    pay_e = qc.calc_payment_table_e(df_e, rate_map, contract_list, config, logger)
    pay_r = qc.calc_payment_table_r(df_r, contract_list, config, logger)
    return qc.merge_payment(pay_b, pay_e, pay_r), logger


def _read_table_e_with_engine(config, path, engine):
    cfg = config["table_e"]
    needed = [cfg["contract_col"], cfg["type_col"], cfg["amount_col"], cfg["currency_col"]]
    df = qc._read_sheet(path, 0, needed, header=cfg["header_rows"], engine=engine)
    df = df.rename(columns={
        cfg["contract_col"]: "contract_no",
        cfg["type_col"]: "payment_type",
        cfg["amount_col"]: "amount",
        cfg["currency_col"]: "currency",
    })
    df = df[["contract_no", "payment_type", "amount", "currency"]].copy()
    return df.dropna(subset=["contract_no"])


def read_table_e_any(config, path):
    """表E 兼容 .xls / .xlsx，也兼容扩展名与实际内容不一致的上传文件。"""
    engines = ["xlrd", "openpyxl"] if str(path).lower().endswith(".xls") else ["openpyxl", "xlrd"]
    last_exc = None
    for engine in engines:
        try:
            return _read_table_e_with_engine(config, path, engine)
        except Exception as exc:
            last_exc = exc
    raise last_exc


def parse_ladder_cell(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    if isinstance(value, (int, float)) and not pd.isna(value):
        return [int(value)]
    text = str(value).strip()
    if not text:
        return []
    text = (text.replace("＃", "#").replace("－", "-").replace("—", "-")
            .replace("～", "-").replace("~", "-"))
    if re.fullmatch(r"\d{1,4}", text):
        return [int(text)]
    plain_range = re.fullmatch(r"(\d{1,4})\s*-\s*(\d{1,4})", text)
    if plain_range:
        lo, hi = int(plain_range.group(1)), int(plain_range.group(2))
        return list(range(lo, hi + 1)) if lo <= hi and hi - lo <= 200 else []
    ladders = set()
    for match in _RE_LADDER_RANGE.finditer(text):
        lo, hi = int(match.group(1)), int(match.group(2))
        if lo <= hi and hi - lo <= 200:
            ladders.update(range(lo, hi + 1))
    for match in _RE_LADDER.finditer(text):
        ladders.add(int(match.group(1)))
    return sorted(ladders)


def parse_upload(path):
    """支持两种格式：带表头的 合同号/梯号 列，或 A 列合同号+梯号文本。"""
    errors = []
    warnings = []
    entries = []
    try:
        head = pd.read_excel(path, sheet_name=0, header=0, dtype=object)
    except Exception as exc:
        return {"ok": False, "entries": [], "errors": ["无法读取上传文件：" + str(exc)], "warnings": []}

    columns = [str(c).strip() for c in head.columns]
    ladder_col = next((c for c in head.columns if "梯" in str(c) or "标的" in str(c)), None)
    contract_col = next((c for c in head.columns if "合同" in str(c)), None)
    date_col = next((c for c in head.columns if "发货" in str(c) and "日期" in str(c)), None)
    remark_col = next((c for c in head.columns if "备注" in str(c)), None)

    if ladder_col is not None:
        for pos, (_, row) in enumerate(head.iterrows(), start=2):
            ladder_value = row.get(ladder_col)
            ladders = parse_ladder_cell(ladder_value)
            if not ladders:
                continue
            contract = _clean_text(row.get(contract_col)) if contract_col is not None else ""
            if not contract:
                contract = qc.extract_contract_from_tid(ladder_value) or ""
            contract = contract.upper()
            for ladder in ladders:
                entries.append({
                    "row": pos,
                    "contract": contract,
                    "ladder": int(ladder),
                    "declared_date": _clean_text(row.get(date_col)) if date_col is not None else "",
                    "remark": _clean_text(row.get(remark_col)) if remark_col is not None else "",
                })
    else:
        raw = pd.read_excel(path, sheet_name=0, header=None, dtype=object)
        for pos, value in enumerate(raw.iloc[:, 0].tolist(), start=1):
            text = _clean_text(value)
            if not text:
                continue
            ladders = parse_ladder_cell(text)
            if not ladders:
                continue
            contract = qc.extract_contract_from_tid(text) or ""
            for ladder in ladders:
                entries.append({
                    "row": pos,
                    "contract": contract.upper(),
                    "ladder": int(ladder),
                    "declared_date": "",
                    "remark": "",
                })

    if not entries:
        errors.append("没有识别到梯号，请使用「合同号/梯号」列，或在 A 列填写合同号+梯号")

    deduped = []
    seen = set()
    for entry in entries:
        key = (entry["contract"], entry["ladder"])
        if key in seen:
            warnings.append("第%d行：梯号重复，已忽略重复项 %s#"% (entry["row"], entry["ladder"]))
            continue
        seen.add(key)
        deduped.append(entry)
    return {"ok": not errors, "entries": deduped, "errors": errors, "warnings": warnings}


def resolve_entries(parsed_entries, ladder_df):
    errors = []
    resolved = []
    index = set()
    by_ladder = {}
    for _, row in ladder_df.iterrows():
        key = (str(row["_contract"]).strip().upper(), int(row["_ladder"]))
        index.add(key)
        by_ladder.setdefault(int(row["_ladder"]), []).append(key)
    for entry in parsed_entries:
        contract = (entry.get("contract") or "").strip().upper()
        ladder = int(entry["ladder"])
        if contract:
            key = (contract, ladder)
            if key not in index:
                errors.append("第%d行：合同 %s 下找不到梯号 %d#，或该梯子不在当前代理商范围"% (entry["row"], contract, ladder))
                continue
            resolved.append((key, entry))
            continue
        candidates = by_ladder.get(ladder, [])
        if len(candidates) == 1:
            resolved.append((candidates[0], entry))
        elif not candidates:
            errors.append("第%d行：当前代理商范围中找不到梯号 %d#" % (entry["row"], ladder))
        else:
            names = "、".join(c for c, _ in candidates[:5])
            errors.append("第%d行：梯号 %d# 对应多个合同（%s），请补充合同号"% (entry["row"], ladder, names))
    return {"ok": not errors, "resolved": resolved, "errors": errors}


def load_ledger(module_id):
    path = ledger_path(module_id)
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and isinstance(data.get("items"), dict):
                return data
        except Exception:
            pass
    return {"version": 1, "module_id": module_id, "items": {}, "updated_at": ""}


def save_ledger(module_id, ledger):
    path = ledger_path(module_id)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(ledger, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _apply_item(item, row, contract_paid, now_str):
    amount = round(float(row["ladder_amount_wan"]), 4)
    shipped_total = round(float(row["shipped_contract_amount_wan"]), 4)
    group_a = row.get("组A日期")
    is_shipped = pd.notna(group_a)
    if is_shipped and shipped_total > 0:
        ratio = amount / shipped_total
        allocated = round(float(contract_paid) * ratio, 4)
        outstanding = round(max(amount - allocated, 0.0), 4)
    else:
        allocated = 0.0
        outstanding = amount
    if outstanding <= 0.01:
        outstanding = 0.0
    item.update({
        "contract": str(row["_contract"]).strip(),
        "ladder": int(row["_ladder"]),
        "bid": str(row["_bid"]).strip(),
        "country": str(row["国家"]).strip(),
        "ladder_amount_wan": amount,
        "contract_amount_wan": round(float(row["contract_amount_wan"]), 4),
        "shipped_contract_amount_wan": shipped_total,
        "contract_paid_wan": round(float(contract_paid), 4),
        "allocated_paid_wan": allocated,
        "outstanding_wan": outstanding,
        "group_a_date": "" if pd.isna(group_a) else pd.Timestamp(group_a).date().isoformat(),
        "due_date": _add_months(group_a, 10) if is_shipped else "",
        "status": "待组A" if not is_shipped else ("已补齐" if outstanding <= 0.01 else "欠款中"),
        "last_calc_at": now_str,
    })
    return item


def compute_module(module_id, upload_path=None, output_dir=None):
    module = get_module(module_id)
    if not module:
        raise RuntimeError("unknown module: " + str(module_id))
    required_slots = module.get("shared_required_slots") or None
    if not quota_shared.is_ready(required_slots):
        raise RuntimeError("公共数据尚未上传齐全，请先在「公共数据」页面上传")

    config = qc.load_main_config()
    data_dir = quota_shared.DATA_DIR
    path_c = qc.find_file(data_dir, config["paths"]["table_c"])
    match_mode = str(module.get("match_mode") or "module").strip().lower()
    path_d = None
    if match_mode == "module":
        path_d = qc.find_file(data_dir, config["paths"]["table_d"])
        if not os.path.isfile(path_d):
            raise RuntimeError("按模块匹配需要表D，请先在「公共数据」上传")
    scope_data = read_scope(path_c, path_d, module)
    scope = scope_data["scope"]
    ladder = scope_data["ladder"]
    contract_list = sorted(scope["_contract"].unique())
    payment_wan, pay_logger = read_payments(config, data_dir, scope, contract_list)

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ledger = load_ledger(module_id)
    previous = copy.deepcopy(ledger.get("items") or {})
    warnings = []
    parse_errors = []
    entries = []
    if upload_path:
        parsed = parse_upload(upload_path)
        warnings.extend(parsed.get("warnings") or [])
        parse_errors.extend(parsed.get("errors") or [])
        entries = parsed.get("entries") or []
    elif not previous:
        raise RuntimeError("首次计算需要上传待发货梯号 Excel")

    resolved = []
    if entries:
        resolved_data = resolve_entries(entries, ladder)
        warnings.extend(resolved_data.get("errors") or [])
        resolved = resolved_data.get("resolved") or []
    if parse_errors:
        raise RuntimeError("上传文件存在问题：" + "；".join(parse_errors))

    items = ledger.setdefault("items", {})
    for (contract, ladder_no), entry in resolved:
        key = _key(contract, ladder_no)
        if key not in items:
            items[key] = {
                "contract": contract,
                "ladder": int(ladder_no),
                "first_seen_at": now_str,
                "declared_at": now_str,
                "remark": entry.get("remark") or "",
            }
        elif entry.get("remark"):
            items[key]["remark"] = entry.get("remark")

    ladder_index = {}
    for _, row in ladder.iterrows():
        ladder_index[_key(row["_contract"], row["_ladder"])] = row

    new_count = 0
    new_wan = 0.0
    released_wan = 0.0
    increase_wan = 0.0
    for key, item in list(items.items()):
        row = ladder_index.get(key)
        old = previous.get(key)
        if row is None:
            item["change"] = "沿用"
            warnings.append(str(item.get("contract")) + " / " + str(item.get("ladder")) + "#：当前表C中找不到，沿用上次结果")
            continue
        contract = str(row["_contract"]).strip()
        updated = _apply_item(item, row, payment_wan.get(contract, 0.0), now_str)
        if old is None:
            updated["change"] = "新增"
            new_count += 1
            new_wan += float(updated.get("outstanding_wan") or 0)
        else:
            old_out = float(old.get("outstanding_wan") or 0)
            new_out = float(updated.get("outstanding_wan") or 0)
            if new_out < old_out - 0.01:
                updated["change"] = "已补齐" if new_out <= 0.01 else "减少"
                released_wan += old_out - new_out
            elif new_out > old_out + 0.01:
                updated["change"] = "增加"
                increase_wan += new_out - old_out
            else:
                updated["change"] = "无变化"

    items_list = list(items.values())
    items_list.sort(key=lambda x: (str(x.get("contract") or ""), int(x.get("ladder") or 0)))
    occupied = round(sum(float(x.get("outstanding_wan") or 0)
                         for x in items_list if x.get("status") != "已补齐"), 4)
    limit = round(float(module.get("quota_limit_wan") or 0), 4)
    available = round(limit - occupied, 4)
    status_counts = {}
    for item in items_list:
        status = item.get("status") or ""
        status_counts[status] = status_counts.get(status, 0) + 1

    result = {
        "module_id": module_id,
        "module_name": module.get("name") or module_id,
        "kind": module.get("kind") or "thailand2",
        "match_mode": scope_data.get("match_mode"),
        "agent_name": scope_data.get("agent_name"),
        "generated_at": now_str,
        "quota_limit_wan": limit,
        "occupied_wan": occupied,
        "available_wan": available,
        "over_quota": available < 0,
        "over_quota_wan": round(max(-available, 0.0), 4),
        "item_count": len(items_list),
        "debt_count": status_counts.get("欠款中", 0),
        "pending_group_a_count": status_counts.get("待组A", 0),
        "paid_count": status_counts.get("已补齐", 0),
        "new_count": new_count,
        "new_wan": round(new_wan, 4),
        "released_wan": round(released_wan, 4),
        "increase_wan": round(increase_wan, 4),
        "items": items_list,
        "warnings": warnings,
        "payment_errors": pay_logger.errors[:200],
        "payment_notes": pay_logger.notes[:200],
        "countries": scope_data["countries"],
        "counts": scope_data["counts"],
        "shared_version": quota_shared.version(),
        "data_files": {
            "table_c": os.path.basename(path_c),
            "table_d": os.path.basename(path_d) if path_d else "",
        },
    }

    out_dir = output_dir or module_dirs(module_id)["out"]
    os.makedirs(out_dir, exist_ok=True)
    save_ledger(module_id, ledger)
    _write_result(result, module, out_dir)
    return result


def _write_result(result, module, out_dir):
    prefix = str(module.get("out_prefix") or ("欠款发货额度统计-" + str(module.get("name"))))
    xlsx_path = os.path.join(out_dir, prefix + ".xlsx")
    json_path = os.path.join(out_dir, "_result_" + str(module.get("id")) + ".json")
    _write_excel(result, xlsx_path)
    tmp = json_path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    os.replace(tmp, json_path)


def _write_excel(result, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "汇总"
    ws.append(["项目", "数值"])
    summary = [
        ("统计对象", result["module_name"]),
        ("匹配方式", "代理商" if result.get("match_mode") == "agent" else "模块"),
        ("代理商", result.get("agent_name") or ""),
        ("匹配国家", "、".join(result.get("countries") or [])),
        ("计算时间", result["generated_at"]),
        ("公共数据版本", result.get("shared_version") or ""),
        ("总额度（万元）", result["quota_limit_wan"]),
        ("当前占用（万元）", result["occupied_wan"]),
        ("可用额度（万元）", result["available_wan"]),
        ("是否超额", "是" if result.get("over_quota") else "否"),
        ("超额金额（万元）", result.get("over_quota_wan") or 0),
        ("台账梯数", result["item_count"]),
        ("欠款中", result["debt_count"]),
        ("待组A", result["pending_group_a_count"]),
        ("已补齐", result["paid_count"]),
        ("本次新增梯数", result["new_count"]),
        ("本次新增欠款（万元）", result["new_wan"]),
        ("本次增加金额（万元）", result.get("increase_wan") or 0),
        ("本次释放金额（万元）", result["released_wan"]),
    ]
    for item in summary:
        ws.append(list(item))

    ws2 = wb.create_sheet("欠款发货台账")
    headers = ["合同号", "梯号", "国家", "状态", "梯子合同额（万元）", "合同已回款（万元）",
               "本梯分摊回款（万元）", "当前占用/欠款（万元）", "本次变化",
               "申报日期", "组A日期", "到期日", "备注"]
    ws2.append(headers)
    for item in result.get("items", []):
        ws2.append([
            item.get("contract"), item.get("bid") or str(item.get("ladder")) + "#",
            item.get("country"), item.get("status"), item.get("ladder_amount_wan"),
            item.get("contract_paid_wan"), item.get("allocated_paid_wan"),
            item.get("outstanding_wan"), item.get("change") or "", item.get("declared_at"),
            item.get("group_a_date"), item.get("due_date"), item.get("remark") or "",
        ])

    head_fill = PatternFill("solid", fgColor="D9EAF7")
    bold = Font(bold=True)
    red = Font(bold=True, color="FF0000")
    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for ws_ in (ws, ws2):
        for cell in ws_[1]:
            cell.fill = head_fill
            cell.font = bold
            cell.alignment = Alignment(horizontal="center")
            cell.border = border
        ws_.freeze_panes = "A2"
        for row in ws_.iter_rows(min_row=2):
            for cell in row:
                cell.border = border
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        if isinstance(row[0].value, (int, float)):
            row[0].number_format = "#,##0.00"
    for row in ws2.iter_rows(min_row=2, min_col=5, max_col=8):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.00"
    for row in ws2.iter_rows(min_row=2):
        if row[3].value == "欠款中":
            row[7].font = red
    for ci, width in enumerate([18, 18, 12, 12, 16, 16, 18, 18, 12, 18, 14, 14, 24], start=1):
        ws2.column_dimensions[get_column_letter(ci)].width = width
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 34
    wb.save(path)


def load_result(module_id):
    path = result_path(module_id)
    if not os.path.isfile(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

