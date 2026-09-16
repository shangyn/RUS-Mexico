# -*- coding: utf-8 -*-
"""quota_core —— 发货可用额度公共计算（CLI v1 与上传看板共用，保证口径一致）

compute_region(region, data_dir, output_dir=None)
  region: dict, 见 quota_regions.yaml 中每个大区的配置
  返回 dict: rows / sum2 / sum3 / llc_wan / net_total / grand / 各计数
"""
import glob
import json
import os
import re

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.yaml")

_RE_CONTRACT_NO = re.compile(r'(D[HY]T-\d{5,6}[TAEC]?)')
_RE_VERSION = re.compile(r'\(\d+\)')
_RE_CONTRACT_YEAR_PREFIX = re.compile(r'^D[HY]T-(\d{2})(\d+)')

REMARK_CURRENCIES = ['人民币', '美元', '美金', 'USD', '欧元', '澳元', '澳币',
                     '英镑', '新元', '港元', '日元', '西非法郎元']
_CURRENCY_ALIAS = {'美金': '美元', 'USD': '美元', '澳币': '澳元'}
_RE_REMARK_A = re.compile(r'折合(\d+\.?\d*)(' + '|'.join(REMARK_CURRENCIES) + ')')
_RE_REMARK_B = re.compile(r'折合(' + '|'.join(REMARK_CURRENCIES) + r')(\d+\.?\d*)')


def load_main_config():
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        import yaml
        return yaml.safe_load(f)


def to_num(series):
    return pd.to_numeric(series, errors='coerce')


# ---------------- 文件查找与通用读取 ----------------

def find_file(data_dir, filename):
    """模糊匹配：忽略文件名括号里的 (N) 版本号；找不到再精确匹配"""
    base, ext = os.path.splitext(filename)
    stripped = _RE_VERSION.sub('', base).strip()
    pattern = os.path.join(data_dir, '*' + stripped + '*' + ext)
    hits = glob.glob(pattern)
    if hits:
        return hits[0]
    return os.path.join(data_dir, filename)


def _sheet_actual_cols(filepath, sheet, needed, header=0, engine=None):
    kwargs = {'sheet_name': sheet, 'nrows': 0, 'header': header}
    if engine:
        kwargs['engine'] = engine
    try:
        head = pd.read_excel(filepath, **kwargs)
    except Exception:
        return None
    actual = {str(c).strip(): c for c in head.columns}
    out = []
    for n in needed:
        if n is None:
            continue
        key = str(n).strip()
        if key in actual:
            out.append(actual[key])
    return out if out else None


def _read_sheet(filepath, sheet, needed, header=0, engine=None):
    usecols = _sheet_actual_cols(filepath, sheet, needed, header=header, engine=engine)
    kwargs = {'sheet_name': sheet, 'header': header}
    if engine:
        kwargs['engine'] = engine
    if usecols:
        kwargs['usecols'] = usecols
    return pd.read_excel(filepath, **kwargs)


def extract_contract_from_tid(value):
    if pd.isna(value):
        return None
    m = _RE_CONTRACT_NO.search(str(value).strip())
    return m.group(1) if m else None


# ---------------- 表C ----------------

def read_table_c_scope(path):
    need = ["合同编号", "标的编号", "代理商", "签订日期", "组A日期",
            "产品状态", "已回款比例（原币）", "合同额（人民币）", "汇率"]
    df = pd.read_excel(path, sheet_name="sheet1", usecols=need)
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise RuntimeError("表C缺少列: " + ", ".join(miss))
    df["签订日期"] = pd.to_datetime(df["签订日期"], errors="coerce")
    df["组A日期"] = pd.to_datetime(df["组A日期"], errors="coerce")
    df["已回款比例（原币）"] = to_num(df["已回款比例（原币）"])
    df["合同额（人民币）"] = to_num(df["合同额（人民币）"]).fillna(0)
    df["汇率"] = to_num(df["汇率"])
    df["_agent"] = df["代理商"].astype(str).str.strip()
    df["_contract"] = df["合同编号"].astype(str).str.strip()
    return df


def build_rate_map(scope_df):
    result = {}
    for contract, grp in scope_df.groupby("_contract", sort=False):
        rate = grp.iloc[0]["汇率"]
        result[contract] = float(rate) if pd.notna(rate) and rate != 0 else 0.0
    return result


# ---------------- 表B / 表E / 表R ----------------

def read_table_b(config, path):
    cfg = config["table_b"]
    frames = []
    for year, col_map in cfg["sheets"].items():
        needed = [col_map["amount_col"], col_map["fee_col"], col_map["contract_col"],
                  col_map["currency_col"], col_map["tid_col"]]
        remark_name = col_map.get("remark_col")
        if remark_name:
            needed.append(remark_name)
        df = _read_sheet(path, str(year), needed)
        rename = {
            col_map["amount_col"]: "amount",
            col_map["fee_col"]: "fee",
            col_map["contract_col"]: "contract_no",
            col_map["currency_col"]: "currency",
            col_map["tid_col"]: "tid",
        }
        if remark_name and remark_name in df.columns:
            rename[remark_name] = "remark"
        df = df.rename(columns=rename)
        extracted = df["tid"].apply(extract_contract_from_tid)
        mask = extracted.notna()
        if mask.any():
            df.loc[mask, "contract_no"] = extracted[mask]
        keep = ["contract_no", "amount", "fee", "currency"]
        if "remark" in df.columns:
            keep.append("remark")
        df = df[keep].copy()
        df["year"] = int(year)
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def read_table_e(config, path):
    cfg = config["table_e"]
    needed = [cfg["contract_col"], cfg["type_col"], cfg["amount_col"], cfg["currency_col"]]
    df = _read_sheet(path, 0, needed, header=cfg["header_rows"], engine="xlrd")
    df = df.rename(columns={
        cfg["contract_col"]: "contract_no",
        cfg["type_col"]: "payment_type",
        cfg["amount_col"]: "amount",
        cfg["currency_col"]: "currency",
    })
    df = df[["contract_no", "payment_type", "amount", "currency"]].copy()
    return df.dropna(subset=["contract_no"])


def read_table_r(config, path):
    cfg = config["table_r"]
    frames = []
    nature_name = cfg.get("nature_col")
    for sheet in cfg["sheets"]:
        needed = [cfg["amount_col"], cfg["contract_col"]]
        if nature_name:
            needed.append(nature_name)
        df = _read_sheet(path, sheet, needed)
        df = df.rename(columns={
            cfg["amount_col"]: "amount",
            cfg["contract_col"]: "contract_no",
        })
        if nature_name and nature_name in df.columns:
            extracted = df[nature_name].apply(extract_contract_from_tid)
            mask = extracted.notna()
            if mask.any():
                df.loc[mask, "contract_no"] = extracted[mask]
        df = df[["contract_no", "amount"]].copy()
        df["year"] = int(sheet)
        frames.append(df)
    result = pd.concat(frames, ignore_index=True)
    result = result.dropna(subset=["contract_no"])
    result["amount"] = to_num(result["amount"])
    return result


# ---------------- 汇率 / 备注折合 ----------------

def load_exchange_rates(filepath):
    rates = {"人民币": 1.0, "人民币元": 1.0}
    if not os.path.exists(filepath):
        return rates
    try:
        df = pd.read_csv(filepath, sep="\t", dtype=str)
        for _, row in df.iterrows():
            name = str(row.iloc[0]).strip()
            m = re.search(r"1(.+?)对人民币", name)
            if m:
                rates[m.group(1)] = float(str(row.iloc[1]).strip())
    except Exception:
        pass
    return rates


def parse_remark(remark, exchange_rates):
    if pd.isna(remark):
        return 0.0, False
    text = str(remark).strip()
    for pat in (_RE_REMARK_A, _RE_REMARK_B):
        m = pat.search(text)
        if m:
            if pat is _RE_REMARK_A:
                amount, currency = float(m.group(1)), m.group(2)
            else:
                currency, amount = m.group(1), float(m.group(2))
            currency = _CURRENCY_ALIAS.get(currency, currency)
            rate = exchange_rates.get(currency)
            if rate is None:
                return 0.0, False
            if currency == "人民币":
                return amount, True
            return round(amount * rate, 4), True
    return 0.0, False


def is_rmb(currency, rmb_list):
    if pd.isna(currency):
        return False
    return str(currency).strip() in rmb_list


def get_contract_rate(contract_no, rate_map, logger, context):
    if contract_no not in rate_map:
        return 1.0
    rate = rate_map[contract_no]
    if rate == 0.0:
        logger.add_error(contract_no, f"汇率缺失或为0（{context}），已用默认1.0")
        return 1.0
    return rate


def resolve_contract_no(contract_no, valid_contracts):
    if pd.isna(contract_no) or contract_no is None:
        return None
    s = str(contract_no).strip()
    if not s or s in ("见拆分", "NaN", "nan", "None"):
        return None
    if s in valid_contracts:
        return s
    base = re.sub(r"/\d+#?", "", s).strip()
    if base != s and base in valid_contracts:
        return base
    base2 = re.sub(r"\d+-\d+#?", "", s).strip()
    if base2 != s and base2 in valid_contracts:
        return base2
    base3 = re.sub(r"\d+#$", "", s).strip()
    if base3 != s and base3 in valid_contracts:
        return base3
    for vc in sorted(valid_contracts, key=len, reverse=True):
        if s.startswith(vc):
            return vc
    return None


# ---------------- 三表回款 ----------------

def calc_payment_table_b(df_b, rate_map, contract_list, config, logger, exchange_rates=None):
    rmb_list = config["rmb_currencies"]
    df_b = df_b.copy()
    df_b["amount"] = to_num(df_b["amount"]).fillna(0)
    df_b["fee"] = to_num(df_b["fee"]).fillna(0)
    df_valid = df_b[df_b["contract_no"].isin(contract_list)].copy()

    def calc_row(row):
        total = row["amount"] + row["fee"]
        if is_rmb(row["currency"], rmb_list):
            return total
        remark = row.get("remark", None)
        if exchange_rates and remark is not None and pd.notna(remark):
            remark_amount, parsed = parse_remark(remark, exchange_rates)
            if parsed:
                return remark_amount
        rate = get_contract_rate(row["contract_no"], rate_map, logger, f"表B-{row['year']}")
        return total * rate

    df_valid["payment"] = df_valid.apply(calc_row, axis=1)
    grouped = df_valid.groupby("contract_no")["payment"].sum()
    return {c: round(grouped.get(c, 0.0), 4) for c in contract_list}


def calc_payment_table_e(df_e, rate_map, contract_list, config, logger):
    rmb_list = config["rmb_currencies"]
    df_dy = df_e[df_e["payment_type"] == "抵佣金"].copy()
    df_dy["amount"] = to_num(df_dy["amount"]).fillna(0)
    df_valid = df_dy[df_dy["contract_no"].isin(contract_list)].copy()

    def calc_row(row):
        if is_rmb(row["currency"], rmb_list):
            return row["amount"]
        rate = get_contract_rate(row["contract_no"], rate_map, logger, "表E-抵佣金")
        return row["amount"] * rate

    df_valid["payment"] = df_valid.apply(calc_row, axis=1)
    grouped = df_valid.groupby("contract_no")["payment"].sum()
    return {c: round(grouped.get(c, 0.0), 4) for c in contract_list}


def calc_payment_table_r(df_r, contract_list, config, logger):
    valid = set(contract_list)
    df_r = df_r.copy()
    df_r["amount"] = to_num(df_r["amount"]).fillna(0)
    resolved_map = {}
    for orig in df_r["contract_no"].unique():
        resolved_map[orig] = resolve_contract_no(orig, valid)
    df_r["resolved_no"] = df_r["contract_no"].map(resolved_map)
    df_valid = df_r[df_r["resolved_no"].notna()].copy()
    df_valid["resolved_no"] = df_valid["resolved_no"].astype(str)
    grouped = df_valid.groupby("resolved_no")["amount"].sum()
    return {c: round(grouped.get(c, 0.0), 4) for c in contract_list}


def merge_payment(pay_b, pay_e, pay_r):
    all_c = set(pay_b) | set(pay_e) | set(pay_r)
    return {c: round((pay_b.get(c, 0) + pay_e.get(c, 0) + pay_r.get(c, 0)) / 10000, 4)
            for c in all_c}



# ---------------- 发货申请（订舱）解析与总价 ----------------

_RE_BOOK_RANGE = re.compile(r"^(\d{1,4})#\s*-\s*(\d{1,4})#$")
_RE_BOOK_LADDER = re.compile(r"^(\d{1,4})#$")
_RE_BID_LADDER = re.compile(r"/(\d+)#\s*$")


def _bid_ladder_no(bid):
    """从标的编号中提取梯号，如 DHT-251677T/58# -> 58；无法解析返回 None。"""
    m = _RE_BID_LADDER.search(str(bid))
    return int(m.group(1)) if m else None


def parse_booking_lines(texts):
    """解析订舱 Excel 文本行（A 列，每行一个合同，支持三种写法）：
      1. 仅合同号            -> 该合同全部梯号
      2. 合同号 + 58# 60# 62# -> 只发固定梯号
      3. 合同号 + 1#-23#     -> 从 1# 连续到 23#（含两端）
    同一合同出现在多行、或同一行内梯号重复视为重叠，拒绝计算。
    """
    rows = []
    errors = []
    for idx, raw in enumerate(texts, start=1):
        text = str(raw).strip()
        if not text:
            continue
        upper = re.sub(r"[\s,，、;；]+", " ", text).upper()
        contracts = list(dict.fromkeys(
            m.group(1).upper() for m in _RE_CONTRACT_NO.finditer(upper)))
        if not contracts:
            if "合同" in upper or "梯" in upper:
                continue  # 表头/说明行直接跳过
            errors.append("第{}行：未识别到合同号「{}」".format(idx, text))
            continue
        if len(contracts) > 1:
            errors.append("第{}行：一行含多个合同号，请拆成多行「{}」".format(idx, text))
            continue
        contract = contracts[0]
        rest = upper.replace(contract, " ", 1)
        tokens = rest.split()
        ladders = []
        for tok in tokens:
            m = _RE_BOOK_RANGE.match(tok)
            if m:
                lo, hi = int(m.group(1)), int(m.group(2))
                if lo > hi:
                    errors.append("第{}行：区间倒置「{}」".format(idx, tok))
                else:
                    ladders.extend(range(lo, hi + 1))
                continue
            m = _RE_BOOK_LADDER.match(tok)
            if m:
                ladders.append(int(m.group(1)))
                continue
            errors.append("第{}行：无法识别「{}」，请用 58# 或 1#-23# 写法".format(idx, tok))
        if not errors and ladders and len(ladders) != len(set(ladders)):
            dup = sorted({n for n in ladders if ladders.count(n) > 1})
            dup_txt = "、".join(str(n) + "#" for n in dup)
            errors.append("第{}行：梯号重复「{}」，同一合同请只写一次".format(idx, dup_txt))
        if errors:
            continue
        rows.append({
            "line": idx,
            "text": text,
            "contract": contract,
            "ladders": sorted(ladders) if ladders else None,
            "ladder_text": "全部梯号" if not ladders
                           else "、".join(str(n) + "#" for n in sorted(ladders)),
        })
    by_contract = {}
    for row in rows:
        by_contract.setdefault(row["contract"], []).append(row)
    for contract, sub_rows in by_contract.items():
        if len(sub_rows) > 1:
            line_desc = "、".join("第{}行".format(r["line"]) for r in sub_rows)
            errors.append("合同 {} 在 {} 重复出现（重叠），请合并为一行后重算".format(contract, line_desc))
    return {"ok": not errors, "rows": rows, "errors": errors}


def build_region_scope(region, path_c):
    """区域口径：代理商 + 签订日期 + 非已作废，并识别整单已发货且回款100%的排除合同。"""
    df_c = read_table_c_scope(path_c)
    agents = list(region["agents"])
    date_start = pd.Timestamp(region["date_start"])
    scope = df_c[df_c["_agent"].isin(agents)
                 & (df_c["产品状态"] != "已作废")
                 & (df_c["签订日期"] >= date_start)].copy()
    grp = scope.groupby("_contract")
    all_date = grp["组A日期"].apply(lambda s: s.notna().all())
    all_ratio1 = grp["已回款比例（原币）"].apply(lambda s: (s == 1).all())
    excluded = set(all_date[all_date & all_ratio1].index)
    return scope, excluded


def compute_booking(region, path_c, texts):
    """发货申请总价：按区域口径过滤表C后，取被选梯号的 J 列(合同额人民币)合计。"""
    empty = {"ok": False, "errors": [], "lines": [], "line_totals": [],
             "details": [], "missing": [], "total_rmb": 0.0, "total_wan": 0.0, "count": 0}
    parse = parse_booking_lines(texts)
    if not parse["ok"]:
        empty["errors"] = parse["errors"]
        return empty
    scope, excluded = build_region_scope(region, path_c)
    kept = set(scope["_contract"]) - excluded
    sub = scope[scope["_contract"].isin(kept)].copy()
    sub["_bid"] = sub["标的编号"].astype(str).str.strip().str.upper()
    sub["_ladder"] = sub["_bid"].apply(_bid_ladder_no)
    sub = sub[sub["_ladder"].notna()].copy()
    sub["_ladder"] = sub["_ladder"].astype(int)

    line_totals = []
    details = []
    missing = []
    grand_total = 0.0
    hit_rows = 0
    for row in parse["rows"]:
        rows_c = sub[sub["_contract"] == row["contract"]]
        entry = {"line": row["line"], "text": row["text"], "contract": row["contract"],
                 "spec": row["ladder_text"], "count": 0, "total_rmb": 0.0}
        if rows_c.empty:
            missing.append({"line": row["line"], "contract": row["contract"],
                            "spec": row["ladder_text"],
                            "message": "合同不在当前区域可发货台账中（代理商/签订日期/状态/整单排除过滤后无该合同）"})
            line_totals.append(entry)
            continue
        if row["ladders"] is None:
            sel = rows_c.sort_values("_ladder")
        else:
            want = set(row["ladders"])
            sel = rows_c[rows_c["_ladder"].isin(want)].sort_values("_ladder")
            for n in sorted(want - set(rows_c["_ladder"])):
                missing.append({"line": row["line"], "contract": row["contract"],
                                "spec": str(n) + "#", "message": "该梯号不在当前区域可发货台账中"})
        total = round(float(sel["合同额（人民币）"].sum()), 2)
        entry["count"] = int(len(sel))
        entry["total_rmb"] = total
        grand_total += total
        hit_rows += entry["count"]
        line_totals.append(entry)
        for _, d in sel.iterrows():
            marker = ""
            if pd.notna(d.get("组A日期")):
                marker = "组A已发货 " + str(pd.Timestamp(d["组A日期"]).date())
            details.append({"line": row["line"], "contract": row["contract"],
                            "bid": d["_bid"], "ladder": int(d["_ladder"]),
                            "amount": float(d["合同额（人民币）"]), "marker": marker})
    return {"ok": True, "errors": [], "lines": parse["rows"],
            "line_totals": line_totals, "details": details, "missing": missing,
            "total_rmb": round(grand_total, 2),
            "total_wan": round(grand_total / 10000.0, 4),
            "count": hit_rows}



# ---------------- LLC 无合同号回款 ----------------

def llc_unassigned_sum_wan(region, table_b_path, exchange_path):
    config = load_main_config()
    rmb_list = config["rmb_currencies"]
    exchange_rates = load_exchange_rates(exchange_path)
    with pd.ExcelFile(table_b_path) as xl:
        sheets = [s for s in xl.sheet_names
                  if re.fullmatch(r"\d{4}", str(s)) and int(s) >= int(region["llc_year_from"])]
    total_yuan = 0.0
    hit = 0
    for sheet in sheets:
        df = pd.read_excel(table_b_path, sheet_name=sheet)
        col_cmp = next((c for c in df.columns if "公司" in str(c)), None)
        col_contract = next((c for c in df.columns if str(c).strip() == "合同号"), None)
        col_currency = next((c for c in df.columns if str(c).strip() == "币种"), None)
        col_amount = next((c for c in df.columns
                           if str(c).strip() in ("收款金额", "金额（小）")), None)
        col_fee = next((c for c in df.columns if "手续费" in str(c)), None)
        col_remark = next((c for c in df.columns if "备注" in str(c)), None)
        if not all([col_cmp, col_contract, col_currency, col_amount]):
            continue
        company = df[col_cmp].astype(str).str.strip()
        no_contract = df[col_contract].isna() | (df[col_contract].astype(str).str.strip() == "")
        mask = (company == str(region["llc_company"])) & no_contract
        sub = df[mask]
        if not len(sub):
            continue
        amount = to_num(sub[col_amount]).fillna(0)
        fee = to_num(sub[col_fee]).fillna(0) if col_fee else pd.Series(0.0, index=sub.index)
        for idx in sub.index:
            currency = sub.at[idx, col_currency]
            row_total = float(amount.at[idx]) + float(fee.at[idx])
            if is_rmb(currency, rmb_list):
                total_yuan += row_total
            else:
                remark = sub.at[idx, col_remark] if col_remark else None
                parsed = None
                if pd.notna(remark):
                    val, ok = parse_remark(str(remark), exchange_rates)
                    if ok:
                        parsed = val
                if parsed is None:
                    name = str(currency).strip() if not pd.isna(currency) else ""
                    parsed = row_total * exchange_rates.get(name, 1.0)
                total_yuan += parsed
            hit += 1
    if hit:
        print(f"  [LLC] 无合同号命中 {hit} 行")
    return round(total_yuan / 10000.0, 4)


# ---------------- 主流程 ----------------

def _contract_year_seq(contract):
    m = _RE_CONTRACT_YEAR_PREFIX.match(contract)
    if not m:
        return (-1, -1)
    return int(m.group(1)), int(m.group(2))


def build_rows(scope_df, contract_list, payment_wan):
    info = {}
    for contract, grp in scope_df.groupby("_contract"):
        if contract not in contract_list:
            continue
        date_mask = grp["组A日期"].notna()
        info[contract] = {
            "has_a": bool(date_mask.any()),
            "shipped": round(grp.loc[date_mask, "合同额（人民币）"].sum() / 10000.0, 4),
        }

    def sort_key(c):
        block = 0 if info[c]["has_a"] else 1
        y, s = _contract_year_seq(c)
        return (block, -y, -s)

    rows = []
    for contract in sorted(info, key=sort_key):
        col2 = payment_wan.get(contract, 0.0)
        col3 = info[contract]["shipped"]
        rows.append({"contract": contract, "col2": col2, "col3": col3,
                     "col4": round(col2 - col3, 4), "has_a": info[contract]["has_a"]})
    return rows


def compute_region(region, data_dir, output_dir=None):
    """核心计算。region 配置见 quota_regions.yaml；data_dir 内含表C/B/E/R。"""
    cfg = load_main_config()
    print("=" * 60)
    print(f"发货额度统计 — {region.get('name', region.get('id', '区域'))}")
    print(f"数据目录: {data_dir}")

    path_c = find_file(data_dir, cfg["paths"]["table_c"])
    path_b = find_file(data_dir, cfg["paths"]["table_b"])
    path_e = find_file(data_dir, cfg["paths"]["table_e"])
    path_r = find_file(data_dir, cfg["paths"]["table_r"])
    exchange_path = find_file(data_dir, cfg["paths"]["exchange_rate"])

    # 1) 表C筛选 + 整单排除（公共口径，与发货申请共用）
    scope, excluded = build_region_scope(region, path_c)
    count_scope = int(scope["_contract"].nunique())
    print(f"命中 {len(scope)} 行 / {count_scope} 合同")
    contract_list = sorted(set(scope["_contract"]) - excluded)
    print(f"整单排除 {len(excluded)} 个；保留合同 {len(contract_list)} 个")

    # 2) 三表回款
    class MiniLogger:
        def __init__(self):
            self.errors = []

        def add_error(self, contract_no, message, source=""):
            self.errors.append((contract_no, message, source))

    logger = MiniLogger()
    rate_map = build_rate_map(scope)
    exchange_rates = load_exchange_rates(exchange_path)
    df_b = read_table_b(cfg, path_b)
    df_e = read_table_e(cfg, path_e)
    df_r = read_table_r(cfg, path_r)
    pay_b = calc_payment_table_b(df_b, rate_map, contract_list, cfg, logger, exchange_rates)
    pay_e = calc_payment_table_e(df_e, rate_map, contract_list, cfg, logger)
    pay_r = calc_payment_table_r(df_r, contract_list, cfg, logger)
    payment_wan = merge_payment(pay_b, pay_e, pay_r)

    rows = build_rows(scope, set(contract_list), payment_wan)
    llc_wan = llc_unassigned_sum_wan(region, path_b, exchange_path)
    sum2 = round(sum(r["col2"] for r in rows), 4)
    sum3 = round(sum(r["col3"] for r in rows), 4)
    net_total = round(sum2 - sum3, 4)
    fixed = float(region.get("fixed_quota", 0))
    grand = round(net_total + fixed + llc_wan, 4)

    print(f"合同行数: {len(rows)}")
    print(f"已回款金额合计: {sum2:,.2f} 万元")
    print(f"已发货合同金额合计: {sum3:,.2f} 万元")
    print(f"固定额度: {fixed:,.2f} 万元 | LLC(无合同号): {llc_wan:,.2f} 万元")
    print(f"发货可用额度 = {grand:,.2f} 万元")

    result = {
        "region_id": region.get("id"),
        "region_name": (region.get("name") or "").replace("大区", ""),
        "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        "rows": rows,
        "count_scope": count_scope,
        "count_excluded": len(excluded),
        "count_kept": len(contract_list),
        "sum2": sum2, "sum3": sum3, "llc": llc_wan,
        "fixed": fixed, "net_total": net_total, "grand": grand,
    }
    if output_dir:
        _write_outputs(result, region, output_dir)
    return result


def _write_outputs(result, region, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    prefix = str(region.get("out_prefix", "发货额度统计"))
    xlsx_path = os.path.join(output_dir, prefix + ".xlsx")
    html_path = os.path.join(output_dir, prefix + ".html")
    json_path = os.path.join(output_dir, f"_result_{region.get('id', 'region')}.json")

    _write_excel(result, xlsx_path)
    _write_html(result, html_path)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False)
    print(f"[输出] Excel: {xlsx_path}")
    print(f"[输出] HTML : {html_path}")


def _write_excel(result, xlsx_path, loan=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    rows = result["rows"]
    headers = ["合同号", "已回款金额", "已发货合同金额", "可用额度"]
    wb = Workbook()
    ws = wb.active
    ws.title = "发货额度统计"
    head_fill = PatternFill("solid", fgColor="D9D9D9")
    top_fill = PatternFill("solid", fgColor="BDD7EE")
    total_fill = PatternFill("solid", fgColor="FFF2CC")
    bold = Font(bold=True)
    red = Font(bold=True, color="FF0000")
    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.append(headers)
    for cell in ws[1]:
        cell.fill = head_fill
        cell.font = bold
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    for r in rows:
        ws.append([r["contract"], r["col2"], r["col3"], r["col4"]])
        n = ws.max_row
        for ci in range(1, 5):
            cell = ws.cell(row=n, column=ci)
            cell.border = border
            if ci in (2, 3, 4):
                cell.number_format = "#,##0.00"
        if r["has_a"]:
            ws.cell(row=n, column=1).fill = top_fill

    footer = [
        ("合计", result["sum2"], result["sum3"], result["net_total"], bold, None),
        (f"固定额度({result['fixed']:g}万)", None, None, result["fixed"], None, None),
        (f"{'LLC无合同号回款'}", None, None, result["llc"], None, None),
        ("发货可用额度", None, None, result["grand"], bold, total_fill),
    ]
    loan = float(loan or 0)
    if loan > 0:
        final_balance = round(float(result.get("grand") or 0) - loan, 4)
        footer.append(("借款额", None, None, loan, None, None))
        footer.append(("最终余额", None, None, final_balance,
                       red if final_balance < 0 else bold, total_fill))
    ws.cell(row=ws.max_row + 1, column=1).border = Border(top=Side(style="double"))
    for label, v2, v3, v4, font, fill in footer:
        ws.append([label, v2 if v2 is not None else None,
                   v3 if v3 is not None else None,
                   v4 if v4 is not None else None])
        n = ws.max_row
        for ci in range(1, 5):
            cell = ws.cell(row=n, column=ci)
            cell.border = border
            if font:
                cell.font = font
            if fill:
                cell.fill = fill
            if ci in (2, 3, 4) and ws.cell(row=n, column=ci).value is not None:
                ws.cell(row=n, column=ci).number_format = "#,##0.00"

    for ci, name in enumerate(headers, start=1):
        max_len = max([len(name)] +
                      [len(str(ws.cell(row=r, column=ci).value or ""))
                       for r in range(2, ws.max_row + 1)])
        ws.column_dimensions[get_column_letter(ci)].width = max(10, max_len * 1.8 + 2)
    ws.freeze_panes = "A2"
    wb.save(xlsx_path)


def _write_html(result, html_path):
    rows = result["rows"]

    def esc(x):
        return "" if x is None else str(x)

    body = []
    for r in rows:
        color = ' style="background:#BDD7EE"' if r["has_a"] else ""
        body.append(
            f'<tr><td{color}>{esc(r["contract"])}</td>'
            f'<td class="num">{r["col2"]:,.2f}</td>'
            f'<td class="num">{r["col3"]:,.2f}</td>'
            f'<td class="num">{r["col4"]:,.2f}</td></tr>'
        )
    footer = (
        f'<tr class="sum"><td>合计</td><td class="num">{result["sum2"]:,.2f}</td>'
        f'<td class="num">{result["sum3"]:,.2f}</td>'
        f'<td class="num">{result["net_total"]:,.2f}</td></tr>'
        f'<tr><td>固定额度({result["fixed"]:g}万)</td><td></td><td></td>'
        f'<td class="num">{result["fixed"]:,.2f}</td></tr>'
        f'<tr><td>LLC无合同号回款</td><td></td><td></td>'
        f'<td class="num">{result["llc"]:,.2f}</td></tr>'
        f'<tr class="grand"><td>发货可用额度</td><td></td><td></td>'
        f'<td class="num">{result["grand"]:,.2f}</td></tr>'
    )
    html = f"""<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8"><title>发货额度统计</title>
<style>
 body {{ font-family: "Microsoft YaHei", sans-serif; margin: 24px; }}
 table {{ border-collapse: collapse; }}
 th, td {{ border: 1px solid #aaa; padding: 4px 12px; font-size: 13px; }}
 th {{ background: #e0e0e0; }}
 .num {{ text-align: right; }}
 tr.sum td {{ border-top: 3px double #000; font-weight: bold; }}
 tr.grand td {{ background: #fff2cc; font-weight: bold; }}
</style></head><body><h2>发货额度统计（{esc(result['region_name'])}）</h2>
<table><thead><tr><th>合同号</th><th>已回款金额</th><th>已发货合同金额</th><th>可用额度</th></tr></thead>
<tbody>{''.join(body)}{footer}</tbody></table>
<p>发货可用额度 = Σ可用额度 + 固定额度 + LLC无合同号回款（浅蓝首列=存在已发货梯）</p>
</body></html>"""
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)


# ---------------- 区域明细 x 发货申请联动（预发货金额 + 余额） ----------------

def merge_booking_into_rows(result, booking):
    """按合同号把最近一次发货申请金额(万元)并入区域明细行。

    返回 (rows, book_sum, bal_sum)：
      rows 每项 dict(contract, col2, col3, has_a, pre, bal)
        pre = 本次申请预发货金额(万元)；bal = 余额 = 已回款(col2) - 已发货(col3) - 预发货(pre)
    """
    book_wan = {}
    for lt in booking.get("line_totals") or []:
        c = str(lt.get("contract") or "").strip().upper()
        if not c:
            continue
        try:
            amt = float(lt.get("total_rmb") or 0)
        except (TypeError, ValueError):
            amt = 0.0
        book_wan[c] = round(book_wan.get(c, 0.0) + amt / 10000.0, 4)
    rows = []
    book_sum = 0.0
    bal_sum = 0.0
    for r in result.get("rows") or []:
        contract = str(r.get("contract") or "").strip()
        pre = book_wan.get(contract.upper(), 0.0)
        col2 = float(r.get("col2") or 0)
        col3 = float(r.get("col3") or 0)
        bal = round(col2 - col3 - pre, 4)
        rows.append({
            "contract": contract,
            "col2": col2, "col3": col3,
            "has_a": bool(r.get("has_a")),
            "pre": round(pre, 4), "bal": bal,
        })
        book_sum += pre
        bal_sum += bal
    return rows, round(book_sum, 4), round(bal_sum, 4)


def _write_excel_linked(result, booking, xlsx_path, loan=None):
    """有发货申请时的区域明细 Excel：5 列（...预发货金额 | 余额），负数红字。"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    rows, book_sum, bal_sum = merge_booking_into_rows(result, booking)
    headers = ["合同号", "已回款金额", "已发货合同金额", "预发货金额", "余额"]
    wb = Workbook()
    ws = wb.active
    ws.title = "发货额度统计"
    head_fill = PatternFill("solid", fgColor="D9D9D9")
    top_fill = PatternFill("solid", fgColor="BDD7EE")
    total_fill = PatternFill("solid", fgColor="FFF2CC")
    bold = Font(bold=True)
    red = Font(bold=True, color="FF0000")
    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.append(headers)
    for cell in ws[1]:
        cell.fill = head_fill
        cell.font = bold
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    for r in rows:
        ws.append([r["contract"], r["col2"], r["col3"],
                   r["pre"] if r["pre"] else None, r["bal"]])
        n = ws.max_row
        for ci in range(1, 6):
            cell = ws.cell(row=n, column=ci)
            cell.border = border
            if ci in (2, 3, 4, 5):
                cell.number_format = "#,##0.00"
        if r["has_a"]:
            ws.cell(row=n, column=1).fill = top_fill
        if r["bal"] < 0:
            ws.cell(row=n, column=5).font = Font(color="FF0000")

    grand = float(result.get("grand") or 0)
    balance_after = round(grand - book_sum, 4)
    ws.cell(row=ws.max_row + 1, column=1).border = Border(top=Side(style="double"))
    footer = [
        ("合计", result.get("sum2"), result.get("sum3"), book_sum, bal_sum, bold, None),
        (f"固定额度({result.get('fixed') or 0:g}万)", None, None, None, result.get("fixed"), None, None),
        ("未确认回款", None, None, None, result.get("llc"), None, None),
        ("发货可用额度", None, None, None, grand, bold, total_fill),
        ("余额（本次发货后）", None, None, book_sum, balance_after,
         red if balance_after < 0 else bold, total_fill),
    ]
    loan = float(loan or 0)
    if loan > 0:
        final_balance = round(balance_after - loan, 4)
        footer.append(("借款额", None, None, None, loan, None, None))
        footer.append(("最终余额", None, None, None, final_balance,
                       red if final_balance < 0 else bold, total_fill))
    for label, v2, v3, v4, v5, font, fill in footer:
        ws.append([label, v2 if v2 is not None else None,
                   v3 if v3 is not None else None,
                   v4 if v4 is not None else None,
                   v5 if v5 is not None else None])
        n = ws.max_row
        for ci in range(1, 6):
            cell = ws.cell(row=n, column=ci)
            cell.border = border
            if font:
                cell.font = font
            if fill:
                cell.fill = fill
            if ci in (2, 3, 4, 5) and ws.cell(row=n, column=ci).value is not None:
                ws.cell(row=n, column=ci).number_format = "#,##0.00"

    for ci in range(1, 6):
        max_len = max([len(headers[ci - 1])] +
                      [len(str(ws.cell(row=r, column=ci).value or ""))
                       for r in range(2, ws.max_row + 1)])
        ws.column_dimensions[get_column_letter(ci)].width = max(10, max_len * 1.8 + 2)
    ws.freeze_panes = "A2"
    wb.save(xlsx_path)


def write_region_excel(region, result, booking, out_dir, loan=None):
    """按是否有成功的发货申请，重写大区「下载 Excel」：
    无申请 -> 原 4 列回款明细；有申请 -> 联动 5 列（预发货金额 + 余额），负数红字。
    """
    os.makedirs(out_dir, exist_ok=True)
    prefix = str(region.get("out_prefix") or "发货额度统计")
    xlsx_path = os.path.join(out_dir, prefix + ".xlsx")
    if result and booking and booking.get("ok"):
        _write_excel_linked(result, booking, xlsx_path, loan=loan)
    elif result:
        _write_excel(result, xlsx_path, loan=loan)
    else:
        return None
    return xlsx_path


# ---------------- 发货申请（订舱）结果落盘：Excel + JSON ----------------

def save_booking_result(result, region, out_dir):
    """把一次成功的发货申请计算结果保存为 Excel + JSON，供下载与后续缓存使用。
    返回生成的 xlsx 路径。result 需含 total_rmb/total_wan/grand_wan/balance_wan 等。
    """
    import os
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    os.makedirs(out_dir, exist_ok=True)
    xlsx_path = os.path.join(out_dir, "发货申请计算明细.xlsx")
    json_path = os.path.join(out_dir, "_booking_last.json")

    def num_or_dash(v):
        return v if isinstance(v, (int, float)) else "—"

    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="D9D9D9")
    bold = Font(bold=True)
    num_fmt = "#,##0.00"

    # Sheet1 汇总
    ws = wb.active
    ws.title = "汇总"
    summary = [
        ("大区", result.get("region_name", "")),
        ("计算时间", result.get("calculated_at", "")),
        ("上传文件", result.get("uploaded_name", "")),
        ("申请行数", len(result.get("lines", []))),
        ("命中梯数", result.get("count", 0)),
        ("本次发货总价（元）", num_or_dash(result.get("total_rmb"))),
        ("本次发货总价（万元）", num_or_dash(result.get("total_wan"))),
        ("发货可用额度（万元）", num_or_dash(result.get("grand_wan"))),
        ("余额（万元）", num_or_dash(result.get("balance_wan"))),
        ("未命中提示条数", len(result.get("missing", []))),
    ]
    for label, value in summary:
        ws.append([label, value])
    for row in ws.iter_rows(min_row=1, max_row=len(summary)):
        row[0].font = bold
    for r in range(2, len(summary) + 1):
        v = ws.cell(row=r, column=2).value
        if isinstance(v, (int, float)):
            ws.cell(row=r, column=2).number_format = num_fmt
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 30
    ws.freeze_panes = "A2"

    # Sheet2 按申请汇总
    ws2 = wb.create_sheet("按申请汇总")
    headers2 = ["行", "合同号", "梯号规格", "梯数", "金额（元）"]
    ws2.append(headers2)
    for cell in ws2[1]:
        cell.fill = head_fill
        cell.font = bold
    total_rows = 0
    total_rmb = 0.0
    for lt in result.get("line_totals", []):
        ws2.append([lt["line"], lt["contract"], lt["spec"], lt["count"], lt["total_rmb"]])
        total_rows += lt["count"]
        total_rmb += lt["total_rmb"]
        ws2.cell(row=ws2.max_row, column=5).number_format = num_fmt
    ws2.append(["合计", "", "", total_rows, round(total_rmb, 2)])
    for cell in ws2[ws2.max_row]:
        cell.font = bold
    ws2.cell(row=ws2.max_row, column=5).number_format = num_fmt

    # Sheet3 逐梯明细（用户要看的“153 行”）
    ws3 = wb.create_sheet("逐梯明细")
    headers3 = ["行", "合同号", "标的编号", "金额（元）", "备注"]
    ws3.append(headers3)
    for cell in ws3[1]:
        cell.fill = head_fill
        cell.font = bold
    for d in result.get("details", []):
        ws3.append([d["line"], d["contract"], d["bid"], d["amount"], d.get("marker", "")])
        ws3.cell(row=ws3.max_row, column=4).number_format = num_fmt
    if result.get("missing"):
        ws3.append(["", "", "", "", ""])
        ws3.append(["未命中/异常", "", "", "", ""])
        n = ws3.max_row
        for cell in ws3[n]:
            cell.font = bold
        for m in result.get("missing", []):
            ws3.append(["第%s行" % m["line"], m["contract"], m.get("spec", ""), "",
                        m.get("message", "")])
    for ws_ in (ws2, ws3):
        for ci in range(1, 6):
            width = max([len(str(ws_.cell(row=r, column=ci).value or ""))
                         for r in range(1, ws_.max_row + 1)] + [6])
            ws_.column_dimensions[get_column_letter(ci)].width = min(max(width * 1.6 + 2, 8), 46)
        ws_.freeze_panes = "A2"
    wb.save(xlsx_path)

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    return xlsx_path


# ==================== 墨西哥：欠款发货额度（新增区块） ====================
# 与俄罗斯的差异：外币计量 + 按梯号拆分（已组A / 已订舱未组A）+ 预算成本帐海运费。
# 本区块全部为新增代码，不修改俄罗斯既有函数的行为。

_MEX_C_COLS = ["合同编号", "标的编号", "签订日期", "组A日期", "产品状态", "国家",
               "已回款比例（原币）", "合同额（原币）", "汇率"]
_MEX_COLS = ["合同号", "合同额（万美元）", "已回款金额（万美元）",
             "已组A金额（万美元）", "已组A海运费金额（万美元）", "已组A设备金额（万美元）",
             "已订舱未组A金额（万美元）", "已订舱未组A海运费金额（万美元）", "已订舱未组A设备金额（万美元）",
             "设备欠款金额（万美元）", "海运费欠款金额（万美元）", "汇率",
             "设备欠款金额人民币（万元）", "海运费欠款金额人民币（万元）"]
# 列名 -> 结果字典的键（顺序即输出列顺序）
_MEX_KEY_MAP = [("合同号", "contract"), ("合同额（万美元）", "B"), ("已回款金额（万美元）", "C"),
                ("已组A金额（万美元）", "D"), ("已组A海运费金额（万美元）", "E"),
                ("已组A设备金额（万美元）", "F"),
                ("已订舱未组A金额（万美元）", "G"), ("已订舱未组A海运费金额（万美元）", "H"),
                ("已订舱未组A设备金额（万美元）", "I"), ("设备欠款金额（万美元）", "J"),
                ("海运费欠款金额（万美元）", "K"), ("汇率", "L"),
                ("设备欠款金额人民币（万元）", "M"), ("海运费欠款金额人民币（万元）", "N")]
_MEX_NUM_KEYS = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "M", "N"]


class _MxLogger:
    """收集口径提示（汇率缺失等），不影响计算流程。"""

    def __init__(self):
        self.errors = []

    def add_error(self, contract_no, message, source=""):
        self.errors.append((contract_no, message, source))


def read_table_c_mexico(path):
    """读取表C（墨西哥口径）：需要 合同额（原币）E列 与 汇率 I列。"""
    df = pd.read_excel(path, sheet_name="sheet1", usecols=_MEX_C_COLS)
    miss = [c for c in _MEX_C_COLS if c not in df.columns]
    if miss:
        raise RuntimeError("表C缺少列: " + ", ".join(miss))
    df["签订日期"] = pd.to_datetime(df["签订日期"], errors="coerce")
    df["组A日期"] = pd.to_datetime(df["组A日期"], errors="coerce")
    df["已回款比例（原币）"] = to_num(df["已回款比例（原币）"])
    df["合同额（原币）"] = to_num(df["合同额（原币）"]).fillna(0)
    df["汇率"] = to_num(df["汇率"])
    df["_contract"] = df["合同编号"].astype(str).str.strip()
    df["_key"] = df["_contract"].str.upper()
    df["_bid"] = df["标的编号"].astype(str).str.strip().str.upper()
    df["_ladder"] = df["_bid"].apply(_bid_ladder_no)
    return df


def build_mexico_scope(df_c, region):
    """墨西哥有效行：国家含关键词 + 非作废 + 签订日期 >= 起始；并标记“已收满”。"""
    pattern = str(region.get("country_pattern") or "墨西哥")
    date_start = pd.Timestamp(region.get("date_start") or "2024-01-01")
    threshold = float(region.get("ratio_paid_threshold") or 1.0)
    scope = df_c[df_c["国家"].astype(str).str.contains(pattern, na=False)
                 & (df_c["产品状态"] != "已作废")
                 & (df_c["签订日期"] >= date_start)].copy()
    # 已收满 = 已回款比例四舍五入到 2 位后 >= 阈值（Q1 确认口径）
    scope["_paid"] = scope["已回款比例（原币）"].notna() & (
        scope["已回款比例（原币）"].round(2) >= threshold)
    return scope


def read_budget_cost(path):
    """预算成本帐：返回 ({标的号大写: 预计海运费（外币）}, 重复标的号清单)。"""
    need = ["标的号", "预计海运费（外币）"]
    df = pd.read_excel(path, sheet_name=0, usecols=need)
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise RuntimeError("预算成本帐缺少列: " + ", ".join(miss))
    df["_bid"] = df["标的号"].astype(str).str.strip().str.upper()
    df["_fee"] = to_num(df["预计海运费（外币）"]).fillna(0)
    dup = sorted({b for b, n in df["_bid"].value_counts().items() if n > 1})
    return df.groupby("_bid")["_fee"].sum().to_dict(), dup


def read_mexico_booking_list(path):
    """订舱清单：A 列一行一个合同（合同号 + 可选梯号写法）。"""
    df = pd.read_excel(path, header=None, dtype=object)
    texts = []
    for value in df.iloc[:, 0].tolist():
        if value is None:
            continue
        text = str(value).strip()
        if text and text.lower() != "nan":
            texts.append(text)
    return texts


def calc_payment_table_b_fx(df_b, rate_map, contract_list, config, logger, exchange_rates=None):
    """表B 回款（外币口径）：人民币 ÷ 汇率；备注“折合”按人民币 ÷ 汇率；外币取原币。"""
    rmb_list = config["rmb_currencies"]
    df_b = df_b.copy()
    df_b["amount"] = to_num(df_b["amount"]).fillna(0)
    df_b["fee"] = to_num(df_b["fee"]).fillna(0)
    df_valid = df_b[df_b["contract_no"].isin(contract_list)].copy()

    def calc_row(row):
        year = row.get("year", "")
        rate = get_contract_rate(row["contract_no"], rate_map, logger, "表B-" + str(year))
        total = float(row["amount"]) + float(row["fee"])
        if is_rmb(row["currency"], rmb_list):
            return total / rate if rate else total
        remark = row.get("remark", None)
        if exchange_rates and remark is not None and pd.notna(remark):
            remark_amount, parsed = parse_remark(remark, exchange_rates)
            if parsed:
                return remark_amount / rate if rate else remark_amount
        return total

    df_valid["payment"] = df_valid.apply(calc_row, axis=1)
    grouped = df_valid.groupby("contract_no")["payment"].sum()
    return {c: round(grouped.get(c, 0.0), 4) for c in contract_list}


def calc_payment_table_e_fx(df_e, rate_map, contract_list, config, logger):
    """表E 回款（外币口径，仅“抵佣金”）：人民币 ÷ 汇率；外币取原币。"""
    rmb_list = config["rmb_currencies"]
    df_dy = df_e[df_e["payment_type"] == "抵佣金"].copy()
    df_dy["amount"] = to_num(df_dy["amount"]).fillna(0)
    df_valid = df_dy[df_dy["contract_no"].isin(contract_list)].copy()

    def calc_row(row):
        amount = float(row["amount"])
        if is_rmb(row["currency"], rmb_list):
            rate = get_contract_rate(row["contract_no"], rate_map, logger, "表E-抵佣金")
            return amount / rate if rate else amount
        return amount

    df_valid["payment"] = df_valid.apply(calc_row, axis=1)
    grouped = df_valid.groupby("contract_no")["payment"].sum()
    return {c: round(grouped.get(c, 0.0), 4) for c in contract_list}


def calc_payment_table_r_fx(df_r, rate_map, contract_list, config, logger):
    """表R 回款（人民币台账，外币口径需 ÷ 汇率）。"""
    valid = set(contract_list)
    df_r = df_r.copy()
    df_r["amount"] = to_num(df_r["amount"]).fillna(0)
    resolved_map = {}
    for orig in df_r["contract_no"].unique():
        resolved_map[orig] = resolve_contract_no(orig, valid)
    df_r["resolved_no"] = df_r["contract_no"].map(resolved_map)
    df_valid = df_r[df_r["resolved_no"].notna()].copy()
    df_valid["resolved_no"] = df_valid["resolved_no"].astype(str)

    def calc_row(row):
        amount = float(row["amount"])
        rate = get_contract_rate(row["resolved_no"], rate_map, logger, "表R")
        return amount / rate if rate else amount

    df_valid["payment"] = df_valid.apply(calc_row, axis=1)
    grouped = df_valid.groupby("resolved_no")["payment"].sum()
    return {c: round(grouped.get(c, 0.0), 4) for c in contract_list}


def merge_payment_fx(*dicts):
    """外币口径合并：三个 dict 直接相加（不再 ÷10000）。"""
    keys = set()
    for item in dicts:
        keys |= set(item)
    return {c: round(sum(d.get(c, 0.0) for d in dicts), 4) for c in keys}


def _mex_round(value, digits=2):
    return round(float(value) + 0.0, digits)


def _mex_wan(value, digits=2):
    """金额统一按「万」计（÷10000，保留 2 位小数）：美元列 → 万美元，人民币列 → 万元。"""
    return _mex_round(float(value) / 10000.0, digits)


def _mex_sort_key(contract):
    """合同号排序：年份降序 -> 同年编号降序。"""
    m = _RE_CONTRACT_YEAR_PREFIX.match(str(contract))
    if not m:
        return (1, 0, 0)
    return (0, -int(m.group(1)), -int(m.group(2)))


def compute_mexico(region, data_dir, output_dir=None):
    """墨西哥欠款发货额度：每合同一块（合同行 + 梯号明细行）的 14 列结果。"""
    cfg = load_main_config()
    # 界面/导出显示名不带「大区」（与俄罗斯的命名习惯保持一致）
    name = (region.get("name") or region.get("id") or "墨西哥").replace("大区", "")
    paths = cfg["paths"]
    print("=" * 60)
    print("欠款发货额度统计 — " + str(name))
    print("数据目录: " + str(data_dir))

    path_c = find_file(data_dir, paths["table_c"])
    path_b = find_file(data_dir, paths["table_b"])
    path_e = find_file(data_dir, paths["table_e"])
    path_r = find_file(data_dir, paths["table_r"])
    path_bk = find_file(data_dir, paths["mexico_booking"])
    path_bud = find_file(data_dir, paths["budget_cost"])
    exchange_path = find_file(data_dir, paths["exchange_rate"])

    # 1) 表C 有效行
    df_c = read_table_c_mexico(path_c)
    scope = build_mexico_scope(df_c, region)
    scope_keys = {}
    for contract in scope["_contract"]:
        text = str(contract).strip()
        scope_keys.setdefault(text.upper(), text)
    print("表C 墨西哥有效行 %d / %d 合同" % (len(scope), len(scope_keys)))

    # 2) 订舱清单（异常：直接报错终止，不产出结果）
    if not os.path.isfile(path_bk):
        raise RuntimeError("缺少订舱清单文件（" + str(paths["mexico_booking"]) + "）")
    parsed = parse_booking_lines(read_mexico_booking_list(path_bk))
    if not parsed["ok"]:
        raise RuntimeError("订舱清单存在问题，已终止计算：" + "；".join(parsed["errors"]))
    booking_by_contract = {row["contract"]: row for row in parsed["rows"]}
    print("订舱清单 %d 个合同" % len(booking_by_contract))

    # 3) 预算成本帐（海运费，按标的号精确匹配）
    fee_map, dup_bids = read_budget_cost(path_bud)
    warnings = []
    if dup_bids:
        warnings.append("预算成本帐存在重复标的号（已按求和处理）：" +
                        "、".join(dup_bids[:10]) + ("…" if len(dup_bids) > 10 else ""))

    # 4) A 列合同号 = 组A又欠款 ∪ 已订舱未组A
    group_a_rows = scope[scope["组A日期"].notna() & (~scope["_paid"])]
    keys_a = {str(k).strip().upper() for k in group_a_rows["_contract"]}
    keys_all = sorted(set(keys_a) | set(booking_by_contract), key=_mex_sort_key)
    print("组A又欠款 %d 合同；A 列合计 %d 合同" % (len(keys_a), len(keys_all)))

    # 5) 三表回款（外币口径）
    logger = _MxLogger()
    rate_map = build_rate_map(scope)
    exchange_rates = load_exchange_rates(exchange_path)
    df_b = read_table_b(cfg, path_b)
    df_e = read_table_e(cfg, path_e)
    df_r = read_table_r(cfg, path_r)
    contract_list = [scope_keys.get(k, k) for k in keys_all]
    pay_b = calc_payment_table_b_fx(df_b, rate_map, contract_list, cfg, logger, exchange_rates)
    pay_e = calc_payment_table_e_fx(df_e, rate_map, contract_list, cfg, logger)
    pay_r = calc_payment_table_r_fx(df_r, rate_map, contract_list, cfg, logger)
    payment_fx = merge_payment_fx(pay_b, pay_e, pay_r)

    # 6) 逐合同组装（块 = 合同行 + 梯号明细行）
    blocks = []
    count_details = 0
    for key in keys_all:
        contract = scope_keys.get(key, key)
        rows = scope[scope["_key"] == key]
        if rows.empty:
            raise RuntimeError("合同 " + contract +
                               " 不在墨西哥台账口径内（国家/签订日期/产品状态过滤后无该合同）")
        rates = sorted({float(r) for r in rows["汇率"].dropna().tolist() if float(r) != 0})
        if not rates:
            rate = 0.0
            warnings.append(contract + "：表C 汇率缺失或为 0，人民币等值按 0 处理")
        else:
            rate = rates[0]
            if len(rates) > 1:
                warnings.append(contract + "：表C 存在多个汇率 " + str(rates) + "，取第一个")
        amount = float(rows["合同额（原币）"].sum())

        details = []
        # 已组A部分：组A日期非空 且 未收满（Q3 确认口径）
        a_rows = rows[rows["组A日期"].notna() & (~rows["_paid"])].sort_values("_ladder")
        d_amount = 0.0
        e_fee = 0.0
        for _, row in a_rows.iterrows():
            ladder = row["_ladder"]
            bid = str(row["_bid"]).strip().upper()
            if pd.isna(ladder):
                warnings.append(contract + "：标的编号无法解析梯号（" + str(row["标的编号"]) +
                                "），已跳过该明细行")
                continue
            fee = float(fee_map.get(bid, 0.0))
            if bid not in fee_map:
                warnings.append(contract + " " + str(int(ladder)) + "#：预算成本帐无该标的号，海运费按 0")
            row_amount = float(row["合同额（原币）"])
            d_amount += row_amount
            e_fee += fee
            details.append({"ladder": int(ladder), "group": "ga",
                            "amount": _mex_round(row_amount),
                            "freight": _mex_round(fee),
                            "equip": _mex_round(row_amount - fee)})

        # 已订舱未组A部分：订舱清单指定的梯号
        g_amount = 0.0
        h_fee = 0.0
        booking = booking_by_contract.get(key)
        if booking is not None:
            if booking["ladders"] is None:
                selected = rows.sort_values("_ladder")
            else:
                want = set(booking["ladders"])
                have = {int(v) for v in rows["_ladder"].dropna().tolist()}
                missing = sorted(want - have)
                if missing:
                    raise RuntimeError(
                        "订舱清单存在问题，已终止计算：" + contract + " 的梯号 " +
                        "、".join(str(n) + "#" for n in missing) + " 在表C台账中不存在")
                selected = rows[rows["_ladder"].isin(want)].sort_values("_ladder")
            for _, row in selected.iterrows():
                ladder = row["_ladder"]
                bid = str(row["_bid"]).strip().upper()
                if pd.isna(ladder):
                    warnings.append(contract + "：标的编号无法解析梯号（" + str(row["标的编号"]) +
                                    "），已跳过该明细行")
                    continue
                fee = float(fee_map.get(bid, 0.0))
                if bid not in fee_map:
                    warnings.append(contract + " " + str(int(ladder)) +
                                    "#：预算成本帐无该标的号，海运费按 0")
                row_amount = float(row["合同额（原币）"])
                g_amount += row_amount
                h_fee += fee
                details.append({"ladder": int(ladder), "group": "bk",
                                "amount": _mex_round(row_amount),
                                "freight": _mex_round(fee),
                                "equip": _mex_round(row_amount - fee)})

        if key in keys_a and booking is not None:
            source = "组A+订舱"
        elif key in keys_a:
            source = "组A欠款"
        else:
            source = "已订舱未组A"

        f_amount = _mex_round(d_amount - e_fee)
        i_amount = _mex_round(g_amount - h_fee)
        c_amount = _mex_round(payment_fx.get(contract, 0.0))
        j_amount = _mex_round(max(0.0, f_amount + i_amount - c_amount))
        k_amount = _mex_round(amount - c_amount) if j_amount == 0 else \
            _mex_round(amount - c_amount - j_amount)
        # 展示与导出一律按「万美元 / 万元」（÷10000，两位小数），与俄罗斯的「万元」口径对齐
        blocks.append({
            "contract": contract, "source": source,
            "B": _mex_wan(amount), "C": _mex_wan(c_amount),
            "D": _mex_wan(d_amount), "E": _mex_wan(e_fee), "F": _mex_wan(f_amount),
            "G": _mex_wan(g_amount), "H": _mex_wan(h_fee), "I": _mex_wan(i_amount),
            "J": _mex_wan(j_amount), "K": _mex_wan(k_amount), "L": round(rate, 4),
            "M": _mex_wan(j_amount * rate), "N": _mex_wan(k_amount * rate),
            "details": [{"ladder": d["ladder"], "group": d["group"],
                         "amount": _mex_wan(d["amount"]),
                         "freight": _mex_wan(d["freight"]),
                         "equip": _mex_wan(d["equip"])} for d in details],
        })
        count_details += len(details)

    blocks.sort(key=lambda item: _mex_sort_key(item["contract"]))
    totals = {key: _mex_round(sum(item[key] for item in blocks)) for key in _MEX_NUM_KEYS}
    for item in logger.errors:
        warnings.append("汇率提示：" + " | ".join(str(part) for part in item))

    print("合同块 %d 个 / 明细行 %d 行 / 数据行 %d 行" %
          (len(blocks), count_details, len(blocks) + count_details))
    print("设备欠款合计 %s 万美元；海运费欠款合计 %s 万美元" %
          (format(totals["J"], ",.2f"), format(totals["K"], ",.2f")))
    print("折人民币 %.2f 万元 / %.2f 万元" % (totals["M"], totals["N"]))

    result = {
        "kind": "mexico",
        "region_id": region.get("id"),
        "region_name": name,
        "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        "blocks": blocks,
        "totals": totals,
        "count_contracts": len(blocks),
        "count_details": count_details,
        "count_group_a": len(keys_a),
        "count_booking": len(booking_by_contract),
        "scope_rows": int(len(scope)),
        "scope_contracts": len(scope_keys),
        "warnings": warnings,
    }
    if output_dir:
        _write_mexico_outputs(result, region, output_dir)
    return result


def _write_mexico_outputs(result, region, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    prefix = str(region.get("out_prefix") or "欠款发货额度统计-墨西哥")
    xlsx_path = os.path.join(output_dir, prefix + ".xlsx")
    html_path = os.path.join(output_dir, prefix + ".html")
    json_path = os.path.join(output_dir, "_result_" + str(region.get("id", "mexico")) + ".json")
    _write_mexico_excel(result, xlsx_path)
    _write_mexico_html(result, html_path)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False)
    print("[输出] Excel: " + xlsx_path)
    print("[输出] HTML : " + html_path)
    return xlsx_path


def _write_mexico_excel(result, xlsx_path):
    """块状输出：表头 + （合同行 + 梯号明细行）× N + 合计行。"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "欠款发货额度统计"
    head_fill = PatternFill("solid", fgColor="D9D9D9")
    contract_fill = PatternFill("solid", fgColor="E7E6E6")
    ga_fill = PatternFill("solid", fgColor="DDEBF7")      # 已组A明细：浅蓝
    bk_fill = PatternFill("solid", fgColor="FFF2CC")      # 已订舱明细：浅黄
    total_fill = PatternFill("solid", fgColor="FCE4D6")
    bold = Font(bold=True)
    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    num_fmt = "#,##0.00"
    n_cols = len(_MEX_COLS)

    ws.append(_MEX_COLS)
    for cell in ws[1]:
        cell.fill = head_fill
        cell.font = bold
        cell.border = border
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    def put_row(values, fill=None, font=None, fill_cols=None, font_cols=None):
        ws.append(values)
        row_idx = ws.max_row
        for ci in range(1, n_cols + 1):
            cell = ws.cell(row=row_idx, column=ci)
            cell.border = border
            if ci >= 2:
                cell.number_format = num_fmt
            if fill and (fill_cols is None or ci in fill_cols):
                cell.fill = fill
            if font and (font_cols is None or ci in font_cols):
                cell.font = font
        return row_idx

    for block in result["blocks"]:
        put_row([block[key] for _, key in _MEX_KEY_MAP],
                fill=contract_fill, font=bold)
        for detail in block["details"]:
            values = [None] * n_cols
            values[0] = "    " + str(detail["ladder"]) + "#"
            if detail["group"] == "ga":
                values[3] = detail["amount"]
                values[4] = detail["freight"]
                values[5] = detail["equip"]
                put_row(values, fill=ga_fill, fill_cols=(4, 5, 6))
            else:
                values[6] = detail["amount"]
                values[7] = detail["freight"]
                values[8] = detail["equip"]
                put_row(values, fill=bk_fill, fill_cols=(7, 8, 9))

    # 合计行：汇率为比率，不参与求和（留空）
    put_row(["合计"] + [result["totals"].get(key) for _, key in _MEX_KEY_MAP[1:]],
            fill=total_fill, font=bold)

    ws.column_dimensions["A"].width = 18
    for ci in range(2, n_cols + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 16
    ws.row_dimensions[1].height = 34
    ws.freeze_panes = "A2"
    wb.save(xlsx_path)


def _write_mexico_html(result, html_path):
    """本地自包含 HTML 看板（与 Flask 看板同结构、同配色）。"""
    def num(value):
        return format(float(value or 0), ",.2f")

    n_cols = len(_MEX_KEY_MAP) - 1
    body = []
    for block in result["blocks"]:
        cells = "".join('<td class="num">' + num(block.get(key)) + "</td>"
                        for _, key in _MEX_KEY_MAP[1:])
        body.append('<tr class="contract"><td>' + str(block["contract"]) + "</td>" +
                    cells + "</tr>")
        for detail in block["details"]:
            row = ['<td>' + "    " + str(detail["ladder"]) + "#</td>"] + [""] * n_cols
            css = "ga" if detail["group"] == "ga" else "bk"
            start = 3 if detail["group"] == "ga" else 6
            for offset, key in enumerate(("amount", "freight", "equip")):
                row[start + offset] = ('<td class="num ' + css + '">' +
                                       num(detail[key]) + "</td>")
            body.append("<tr>" + "".join(row) + "</tr>")
    total_cells = "".join(
        ('<td class="num">' + num(result["totals"][key]) + "</td>")
        if key in result["totals"] else '<td class="num"></td>'
        for _, key in _MEX_KEY_MAP[1:])
    head = "".join("<th>" + name + "</th>" for name in _MEX_COLS)
    html = (
        "<!DOCTYPE html>\n<html lang=\"zh\"><head><meta charset=\"utf-8\">"
        "<title>欠款发货额度统计</title>\n<style>\n"
        "body { font-family: \"Microsoft YaHei\", sans-serif; margin: 20px; }\n"
        "table { border-collapse: collapse; }\n"
        "th, td { border: 1px solid #aaa; padding: 3px 8px; font-size: 12px; white-space: nowrap; }\n"
        "th { background: #d9d9d9; }\n"
        ".num { text-align: right; }\n"
        "tr.contract td { background: #e7e6e6; font-weight: 700; }\n"
        "td.ga { background: #ddebf7; }\n"
        "td.bk { background: #fff2cc; }\n"
        "tr.total td { background: #fce4d6; font-weight: 700; border-top: 3px double #000; }\n"
        ".legend span { display: inline-block; margin-right: 14px; font-size: 12px; padding: 2px 8px; }\n"
        "</style></head><body>\n"
        "<h2>欠款发货额度统计（" + str(result["region_name"]) + "）</h2>\n"
        "<p class=\"legend\"><span style=\"background:#ddebf7\">浅蓝 = 已组A梯号明细</span>"
        "<span style=\"background:#fff2cc\">浅黄 = 已订舱未组A梯号明细</span>"
        "合同行 = 该合同合计　生成时间 " + str(result["generated_at"]) + "</p>\n"
        "<table><thead><tr>" + head + "</tr></thead>\n<tbody>" + "".join(body) +
        '<tr class="total"><td>合计</td>' + total_cells + "</tr>\n</tbody></table>\n"
        "</body></html>"
    )
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
