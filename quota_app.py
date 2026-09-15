# -*- coding: utf-8 -*-
"""发货额度统计 —— 本地上传看板（Flask）

运行：python quota_app.py  →  http://localhost:5050
上传表C/表B/表E/表R（汇率文件可选）→ 点击运行 → 查看各区域可用额度。

多区域 + 上传缓存（需求规格 v5 定稿）：
- 单页 Tab：俄罗斯 / 美洲 各占一个视图（美洲暂为预留占位）；
- 上传缓存：有缓存时可省略文件直接运行；只替换需要更新的文件，未选的沿用缓存；
- 仅成功运行才更新缓存，失败保留上一次数据与上传日期。
"""
import json
import os
import shutil
import sys
import tempfile
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import yaml
import pandas as pd
from flask import Flask, render_template, request, send_file, abort, redirect, url_for

from quota_core import (compute_region, compute_booking, load_main_config, find_file,
                       save_booking_result, merge_booking_into_rows, write_region_excel)
from quota_core import compute_mexico

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WORKSPACE_ROOT = os.path.join(BASE_DIR, "workspaces", "quota")
UPLOAD_PASSWORD = "888888"  # 部署时建议改为环境变量/config
CACHE_META_NAME = "_cache_meta.json"
LOAN_META_NAME = "_loan.json"

app = Flask(__name__)
app.jinja_env.filters['f2'] = lambda x: f'{x:,.2f}'
app.config["MAX_CONTENT_LENGTH"] = 512 * 1024 * 1024  # 512MB 上限

UPLOAD_KEYS = ["table_c", "table_b", "table_e", "table_r"]
# 墨西哥：同样四张台账 + 订舱清单 + 预算成本帐（海运费）
MEXICO_KEYS = UPLOAD_KEYS + ["mexico_booking", "budget_cost"]
ALL_SLOTS = UPLOAD_KEYS + ["exchange_rate"]
SLOT_ATTRS = {
    "table_c":       {"label": "表C · 国贸合同标的台账 (.xlsx)", "accept": ".xlsx", "optional": False},
    "table_b":       {"label": "表B · 外币回款台账 (.xlsx)",     "accept": ".xlsx", "optional": False},
    "table_e":       {"label": "表E · gm_ht_hthkmx (.xls/.xlsx)", "accept": ".xls,.xlsx", "optional": False},
    "table_r":       {"label": "表R · 人民币回款台账 (.xlsx)",   "accept": ".xlsx", "optional": False},
    "mexico_booking": {"label": "订舱清单 · 墨西哥已订舱未组A合同号 (.xlsx)", "accept": ".xlsx", "optional": False},
    "budget_cost":   {"label": "预算成本帐 · 预计海运费 (.xlsx)",  "accept": ".xlsx", "optional": False},
    "exchange_rate": {"label": "汇率文件（可选）",               "accept": ".txt", "optional": True},
}


def region_slots(region):
    """该大区需要上传/缓存的文件槽位（顺序即页面展示顺序）。"""
    keys = MEXICO_KEYS if region.get("kind") == "mexico" else UPLOAD_KEYS
    return list(keys) + ["exchange_rate"]


def load_regions():
    with open(os.path.join(BASE_DIR, "quota_regions.yaml"), "r", encoding="utf-8") as f:
        return yaml.safe_load(f)["regions"]


def display_name(region):
    """界面显示名：统一去掉「大区」后缀（tab / 卡片标题 / 提示语保持一致）"""
    return (region.get("name") or region.get("id") or "").replace("大区", "")


def region_dirs(region_id):
    base = os.path.join(WORKSPACE_ROOT, region_id)
    return {"data": os.path.join(base, "数据源"), "out": os.path.join(base, "输出")}


def region_base(region_id):
    return os.path.join(WORKSPACE_ROOT, region_id)


def cache_meta_path(region_id):
    return os.path.join(region_base(region_id), CACHE_META_NAME)


def load_cache_meta(region_id):
    path = cache_meta_path(region_id)
    if os.path.isfile(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache_meta(region_id, meta):
    path = cache_meta_path(region_id)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)


def loan_meta_path(region_id):
    return os.path.join(region_base(region_id), LOAN_META_NAME)


def load_loan(region_id):
    """读取该大区缓存的借款额（万元）；没有记录时返回 0 和空时间。"""
    path = loan_meta_path(region_id)
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return {"amount": round(float(data.get("amount") or 0), 2),
                    "updated_at": str(data.get("updated_at") or "")}
        except Exception:
            pass
    return {"amount": 0.0, "updated_at": ""}


def save_loan(region_id, amount):
    """写入借款额缓存（万元），返回更新时间字符串。"""
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    path = loan_meta_path(region_id)
    parent = os.path.dirname(path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"amount": round(float(amount), 2), "updated_at": now_str},
                  f, ensure_ascii=False, indent=2)
    return now_str


def _file_mtime_str(path):
    try:
        return datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M:%S")
    except OSError:
        return ""


def cached_path(region_id, slot):
    """返回该大区正式数据目录中某表文件的真实路径；不存在返回 None"""
    cfg = load_main_config()
    data_dir = region_dirs(region_id)["data"]
    path = find_file(data_dir, cfg["paths"][slot])
    return path if os.path.isfile(path) else None


def build_cache_fields(region):
    meta = load_cache_meta(region["id"])
    fields = []
    for slot in region_slots(region):
        attrs = SLOT_ATTRS[slot]
        path = cached_path(region["id"], slot)
        present = path is not None
        entry = meta.get(slot) or {}
        fields.append({
            "slot": slot,
            "label": attrs["label"],
            "accept": attrs["accept"],
            "optional": attrs["optional"],
            "required": (not present) and (not attrs["optional"]),
            "present": present,
            "original": entry.get("original") or (os.path.basename(path) if path else ""),
            "uploaded_at": entry.get("uploaded_at") or (_file_mtime_str(path) if path else ""),
        })
    return fields


def load_result(region):
    out = region_dirs(region["id"])["out"]
    json_path = os.path.join(out, f"_result_{region['id']}.json")
    xlsx = os.path.join(out, f"{region['out_prefix']}.xlsx")
    if not os.path.isfile(json_path) or not os.path.isfile(xlsx):
        return None
    with open(json_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    data["xlsx_name"] = os.path.basename(xlsx)
    return data


def load_booking(region_id):
    """读取最近一次成功计算的发货申请结果；没有则返回 None。"""
    out = region_dirs(region_id)["out"]
    json_path = os.path.join(out, "_booking_last.json")
    if not os.path.isfile(json_path):
        return None
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) and data.get("ok") else None
    except Exception:
        return None


def sync_booking_view(item):
    """依据 item['result'] 与 item['booking'] 准备联动明细视图数据。"""
    result = item.get("result")
    booking = item.get("booking")
    item["has_booking"] = bool(result and booking and booking.get("ok"))
    item["view_rows"] = None
    item["book_sum"] = 0.0
    item["bal_sum"] = 0.0
    item["balance_after"] = None
    if item["has_booking"]:
        rows, book_sum, bal_sum = merge_booking_into_rows(result, booking)
        item["view_rows"] = rows
        item["book_sum"] = book_sum
        item["bal_sum"] = bal_sum
        if isinstance(result.get("grand"), (int, float)):
            item["balance_after"] = round(float(result["grand"]) - book_sum, 4)


def region_excel_is_linked(out_dir, prefix):
    """判断大区下载 Excel 是否已是含预发货金额的 5 列联动版。"""
    path = os.path.join(out_dir, (prefix or "发货额度统计") + ".xlsx")
    if not os.path.isfile(path):
        return False
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True)
        ws = wb.active
        head = [str(c.value) for c in next(ws.iter_rows(min_row=1, max_row=1))]
        wb.close()
        return head == ["合同号", "已回款金额", "已发货合同金额", "预发货金额", "余额"]
    except Exception:
        return False


def ensure_region_linked(item):
    """启动/刷新时对账：有发货申请但大区 Excel 还没联动 -> 自动补写，保证页面与下载一致。"""
    if not (item.get("has_booking") and item.get("result")):
        return
    try:
        out = region_dirs(item["id"])["out"]
        if not region_excel_is_linked(out, item.get("out_prefix")):
            write_region_excel(item, item["result"], item["booking"], out,
                               loan=(item.get("loan") or {}).get("amount"))
    except Exception:
        pass  # 文件被占用等情况忽略，页面照常可用

def fmt(x):
    return f"{x:,.2f}" if isinstance(x, (int, float)) else ""


def _read_booking_excel(path):
    """读取订舱 Excel：A 列一行一个合同（合同号 + 可选梯号写法）。"""
    df = pd.read_excel(path, header=None, dtype=object)
    texts = []
    for v in df.iloc[:, 0].tolist():
        if v is None:
            continue
        s = str(v).strip()
        if s and s.lower() != "nan":
            texts.append(s)
    return texts


def build_view():
    regions = load_regions()
    view = []
    for region in regions:
        item = dict(region)
        item["result"] = load_result(region)
        item["loan"] = load_loan(region["id"])
        if region.get("enabled"):
            item["fields"] = build_cache_fields(region)
            item["has_cache"] = all(f["present"] for f in item["fields"] if not f["optional"])
            item["booking"] = load_booking(region["id"])
            sync_booking_view(item)
            ensure_region_linked(item)
        view.append(item)
    return view


def render_page(view, msg=""):
    return render_template("quota_dashboard.html", regions=view, today=date.today(),
                           msg=msg, fmt=fmt)


@app.route("/")
def index():
    return render_page(build_view(), request.args.get("msg", ""))


@app.route("/booking", methods=["POST"])
def booking():
    view = build_view()
    region_id = request.form.get("region", "")
    target = next((v for v in view if v["id"] == region_id and v.get("enabled")), None)

    def set_error(msg):
        if target is not None:
            target["booking"] = {"ok": False, "errors": [msg]}
        return render_page(view)

    if target is None:
        return render_page(view, "无效或未启用的区域")
    if target.get("kind") == "mexico":
        return render_page(view, "墨西哥大区不使用「发货申请」功能（订舱清单在「更新数据」里上传）")
    fs = request.files.get("booking_file")
    if fs is None or not fs.filename:
        return set_error("请先选择订舱发货 Excel 文件")
    path_c = cached_path(region_id, "table_c")
    if path_c is None:
        return set_error("该大区还没有表C台账数据，请先在上方「更新数据」上传表C并运行一次额度计算")
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", prefix="booking_",
                                         dir=WORKSPACE_ROOT, delete=False) as f:
            fs.save(f.name)
            tmp_path = f.name
        texts = _read_booking_excel(tmp_path)
        res = compute_booking(target, path_c, texts)
        res["uploaded_name"] = fs.filename
        res["calculated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if res.get("ok"):
            grand = None
            if target["result"] and isinstance(target["result"].get("grand"), (int, float)):
                grand = float(target["result"]["grand"])
            res["grand_wan"] = grand
            res["balance_wan"] = round(grand - res["total_wan"], 4) if grand is not None else None
            res["region_name"] = display_name(target) or region_id
            out_dir = region_dirs(region_id)["out"]
            try:
                # 联动：把本次申请并入大区「下载 Excel」（预发货金额 + 余额）
                write_region_excel(target, target.get("result"), res, out_dir)
            except Exception as e3:
                res["save_error"] = "区域明细联动失败（若Excel正打开下载文件，请关闭后重新计算一次）：" + str(e3)
            try:
                xlsx_path = save_booking_result(res, target, out_dir)
                res["excel_name"] = os.path.basename(xlsx_path)
            except Exception as e2:
                prev = res.get("save_error")
                res["save_error"] = ((prev + "；") if prev else "") + "明细Excel保存失败：" + str(e2)
        target["booking"] = res
        sync_booking_view(target)
        return render_page(view)
    except Exception as e:
        return set_error("计算失败：" + str(e))
    finally:
        if tmp_path:
            try:
                os.remove(tmp_path)
            except OSError:
                pass


@app.route("/booking/download/<region_id>")
def booking_download(region_id):
    """下载最近一次成功计算的发货申请明细 Excel。"""
    out = region_dirs(region_id)["out"]
    xlsx = os.path.join(out, "发货申请计算明细.xlsx")
    if not os.path.isfile(xlsx):
        abort(404)
    return send_file(xlsx, as_attachment=True, download_name=os.path.basename(xlsx))


@app.route("/run", methods=["POST"])
def run():
    regions = load_regions()
    region_id = request.form.get("region", "")
    region = next((r for r in regions if r["id"] == region_id and r.get("enabled")), None)

    def back(text):
        return redirect(url_for("index", msg=text, _anchor=region_id))

    if region is None:
        return back("无效或未启用的区域")
    if request.form.get("password", "") != UPLOAD_PASSWORD:
        return back("密码错误")

    cfg = load_main_config()
    base = region_base(region_id)
    final_data = region_dirs(region_id)["data"]
    final_out = region_dirs(region_id)["out"]
    stage = os.path.join(base, "_staging")
    stage_data = os.path.join(stage, "数据源")
    stage_out = os.path.join(stage, "输出")
    if os.path.isdir(stage):
        shutil.rmtree(stage)
    os.makedirs(stage_data, exist_ok=True)
    os.makedirs(stage_out, exist_ok=True)

    # 备份最近一次发货申请（数据重算提交会整体替换输出目录）
    prev_booking_raw = None
    if os.path.isdir(final_out):
        _prev_path = os.path.join(final_out, "_booking_last.json")
        try:
            if os.path.isfile(_prev_path):
                with open(_prev_path, "rb") as f:
                    prev_booking_raw = f.read()
        except OSError:
            prev_booking_raw = None

    meta_old = load_cache_meta(region_id)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    meta_new = {}
    missing = []
    uploaded_slots = []          # 本次新上传的槽位（运行失败时也要保留，避免重试重新上传）

    def prepare_slot(slot, file_storage):
        """把本次要用的文件放入暂存目录：新上传优先，其次沿用缓存。
        返回 True=已就绪；False=该槽位无文件（仅可选文件允许）。"""
        target = os.path.join(stage_data, cfg["paths"][slot])
        present_cache = cached_path(region_id, slot)
        if file_storage is not None:
            file_storage.save(target)
            meta_new[slot] = {"original": file_storage.filename, "uploaded_at": now_str}
            uploaded_slots.append(slot)
            return True
        if present_cache:
            shutil.copy2(present_cache, target)
            old = meta_old.get(slot) or {}
            meta_new[slot] = {
                "original": old.get("original") or os.path.basename(present_cache),
                "uploaded_at": old.get("uploaded_at") or _file_mtime_str(present_cache),
            }
            return True
        return False

    def keep_uploaded_files():
        """运行失败时保留本次上传的文件：写入正式数据目录并更新缓存日期。
        只覆盖本次新上传的槽位，输出目录与其余缓存一律不动，
        这样修正后重跑无需重新上传，也不用重新等待表C读取。"""
        if not uploaded_slots:
            return 0
        os.makedirs(final_data, exist_ok=True)
        kept = 0
        for slot in uploaded_slots:
            name = cfg["paths"].get(slot)
            if not name:
                continue
            src = os.path.join(stage_data, name)
            if not os.path.isfile(src):
                continue
            try:
                shutil.copy2(src, os.path.join(final_data, name))
            except OSError:
                continue      # 单个文件被占用也不影响其他文件
            kept += 1
        meta = dict(meta_old)
        meta.update({s: meta_new[s] for s in uploaded_slots if s in meta_new})
        save_cache_meta(region_id, meta)
        return kept

    for slot in region_slots(region):
        if slot == "exchange_rate":
            continue
        fs = request.files.get(slot)
        file_obj = fs if fs and fs.filename else None
        if not prepare_slot(slot, file_obj):
            missing.append(SLOT_ATTRS[slot]["label"])

    ex = request.files.get("exchange_rate")
    ex_file = ex if ex and ex.filename else None
    if not prepare_slot("exchange_rate", ex_file):
        default_ex = os.path.join(BASE_DIR, "数据源", cfg["paths"]["exchange_rate"])
        if os.path.isfile(default_ex):
            target = os.path.join(stage_data, cfg["paths"]["exchange_rate"])
            shutil.copy2(default_ex, target)
            meta_new["exchange_rate"] = {
                "original": cfg["paths"]["exchange_rate"] + "（默认）",
                "uploaded_at": _file_mtime_str(default_ex),
            }

    if missing:
        kept = keep_uploaded_files()
        shutil.rmtree(stage, ignore_errors=True)
        kept_note = ("；本次已上传的 %d 个文件已保存为缓存，下次只需补上传缺失的文件" % kept) if kept else ""
        return back("缺少文件：" + "、".join(missing) + "（无缓存，必须上传）" + kept_note)

    try:
        if region.get("kind") == "mexico":
            compute_mexico(region, stage_data, stage_out)
        else:
            compute_region(region, stage_data, stage_out)
    except Exception as e:
        kept = keep_uploaded_files()
        shutil.rmtree(stage, ignore_errors=True)
        kept_note = ("；本次上传的 %d 个文件已保存为缓存，修正后可直接重跑，无需重新上传"
                     "（上次的计算结果与下载文件不受影响）") % kept if kept else ""
        return back("计算失败：" + str(e) + kept_note)

    # 计算成功后才提交：整目录原子替换。
    # 正式目录内任一文件被 Excel 等占用（未以共享删除方式打开）时目录无法整体改名，
    # 提交整体失败——此时尚未删除/移动任何文件，原缓存与输出完整保留。
    def _probe_replaceable(dirpath):
        """探测目录能否被整体替换：改名再改回，不修改内容。"""
        if not os.path.isdir(dirpath):
            return
        probe = dirpath + ".__probe"
        if os.path.isdir(probe):
            shutil.rmtree(probe, ignore_errors=True)
        os.rename(dirpath, probe)
        try:
            os.rename(probe, dirpath)
        finally:
            if not os.path.isdir(dirpath) and os.path.isdir(probe):
                os.rename(probe, dirpath)

    def commit_dir(stage_dir, final_dir):
        """暂存目录原子替换到正式目录：旧目录先改名备份，新目录移入后再删备份。"""
        parent = os.path.dirname(final_dir)
        if parent and not os.path.isdir(parent):
            os.makedirs(parent)
        backup = final_dir + ".__old"
        if os.path.isdir(backup):
            shutil.rmtree(backup, ignore_errors=True)
        had_final = os.path.isdir(final_dir)
        if had_final:
            os.rename(final_dir, backup)
        try:
            os.rename(stage_dir, final_dir)
        except OSError:
            if had_final and not os.path.isdir(final_dir) and os.path.isdir(backup):
                os.rename(backup, final_dir)  # 还原旧目录
            raise
        if had_final:
            shutil.rmtree(backup, ignore_errors=True)

    try:
        _probe_replaceable(final_data)
        _probe_replaceable(final_out)
    except OSError as e:
        kept = keep_uploaded_files()
        shutil.rmtree(stage, ignore_errors=True)
        kept_note = ("；本次上传的 %d 个文件已保存为缓存，关闭占用后可直接重跑" % kept) if kept else ""
        return back("保存失败：" + str(e) +
                    "（目录内文件正被其他程序占用，通常是 Excel 打开了旧数据/输出文件；"
                    "请关闭后重试。本次运行未改动原数据）" + kept_note)

    try:
        commit_dir(stage_data, final_data)
        commit_dir(stage_out, final_out)
    except Exception as e:
        kept = keep_uploaded_files()
        shutil.rmtree(stage, ignore_errors=True)
        kept_note = ("；本次上传的 %d 个文件已保存为缓存，关闭占用后可直接重跑" % kept) if kept else ""
        return back("保存失败：" + str(e) + "（文件正被占用，请关闭相关程序后重试）" + kept_note)

    if os.path.isdir(stage):
        shutil.rmtree(stage, ignore_errors=True)
    save_cache_meta(region_id, meta_new)

    # 数据更新后：沿用最近一次发货申请并重新联动（总额度/余额按新结果刷新）
    if prev_booking_raw:
        try:
            booking = json.loads(prev_booking_raw.decode("utf-8"))
            result_json = os.path.join(final_out, f"_result_{region_id}.json")
            if booking and booking.get("ok") and os.path.isfile(result_json):
                with open(result_json, "r", encoding="utf-8") as f:
                    result = yaml.safe_load(f)
                grand = result.get("grand")
                if isinstance(grand, (int, float)) and isinstance(booking.get("total_wan"), (int, float)):
                    booking["grand_wan"] = float(grand)
                    booking["balance_wan"] = round(float(grand) - float(booking["total_wan"]), 4)
                booking["region_name"] = display_name(region)
                save_booking_result(booking, region, final_out)
                write_region_excel(region, result, booking, final_out,
                                   loan=load_loan(region_id)["amount"])
        except Exception as e:
            print("重算后联动发货申请失败（可忽略）：", e)

    return redirect(url_for("index",
                            msg="运行完成，已更新「" + display_name(region) + "」",
                            _anchor=region_id))


@app.route("/loan", methods=["POST"])
def loan():
    """更新某大区借款额（万元）：密码校验 -> 写缓存 -> 立即重写该大区下载 Excel。"""
    regions = load_regions()
    region_id = request.form.get("region", "")
    region = next((r for r in regions if r["id"] == region_id and r.get("enabled")), None)

    def back(text):
        return redirect(url_for("index", msg=text, _anchor=region_id))

    if region is None:
        return back("无效或未启用的区域")
    if region.get("kind") == "mexico":
        return back("墨西哥大区不使用「借款额」功能")
    if request.form.get("password", "") != UPLOAD_PASSWORD:
        return back("密码错误")

    raw = (request.form.get("amount") or "").strip()
    try:
        amount = round(float(raw), 2)
    except ValueError:
        return back("借款额请填写数字（万元）")
    if amount < 0:
        return back("借款额不能为负数")

    save_loan(region_id, amount)

    # 立刻重写该大区下载 Excel，保证下载到的表格与页面一致
    warn = ""
    try:
        result = load_result(region)
        if result:
            write_region_excel(region, result, load_booking(region_id),
                               region_dirs(region_id)["out"], loan=amount)
    except PermissionError:
        warn = "；但 Excel 正被占用，表格未刷新，请关闭 Excel 后重新提交"
    except Exception as e:
        warn = "；Excel 刷新失败：" + str(e)

    return redirect(url_for("index",
                            msg="借款额已更新为 " + f"{amount:,.2f}" + " 万元" + warn,
                            _anchor=region_id))


@app.route("/download/<region_id>")
def download(region_id):
    regions = load_regions()
    region = next((r for r in regions if r["id"] == region_id), None)
    if region is None:
        abort(404)
    out = region_dirs(region_id)["out"]
    xlsx = os.path.join(out, f"{region['out_prefix']}.xlsx")
    if not os.path.isfile(xlsx):
        abort(404)
    return send_file(xlsx, as_attachment=True, download_name=os.path.basename(xlsx))


if __name__ == "__main__":
    print("发货额度统计看板: http://localhost:5050")
    app.run(host="127.0.0.1", port=5050, debug=False, threaded=True)
