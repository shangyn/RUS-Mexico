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
                       save_booking_result)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WORKSPACE_ROOT = os.path.join(BASE_DIR, "workspaces", "quota")
UPLOAD_PASSWORD = "888888"  # 部署时建议改为环境变量/config
CACHE_META_NAME = "_cache_meta.json"

app = Flask(__name__)
app.jinja_env.filters['f2'] = lambda x: f'{x:,.2f}'
app.config["MAX_CONTENT_LENGTH"] = 512 * 1024 * 1024  # 512MB 上限

UPLOAD_KEYS = ["table_c", "table_b", "table_e", "table_r"]
ALL_SLOTS = UPLOAD_KEYS + ["exchange_rate"]
SLOT_ATTRS = {
    "table_c":       {"label": "表C · 国贸合同标的台账 (.xlsx)", "accept": ".xlsx", "optional": False},
    "table_b":       {"label": "表B · 外币回款台账 (.xlsx)",     "accept": ".xlsx", "optional": False},
    "table_e":       {"label": "表E · gm_ht_hthkmx (.xls/.xlsx)", "accept": ".xls,.xlsx", "optional": False},
    "table_r":       {"label": "表R · 人民币回款台账 (.xlsx)",   "accept": ".xlsx", "optional": False},
    "exchange_rate": {"label": "汇率文件（可选）",               "accept": ".txt", "optional": True},
}


def load_regions():
    with open(os.path.join(BASE_DIR, "quota_regions.yaml"), "r", encoding="utf-8") as f:
        return yaml.safe_load(f)["regions"]


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
    for slot in ALL_SLOTS:
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
        if region.get("enabled"):
            item["fields"] = build_cache_fields(region)
            item["has_cache"] = all(f["present"] for f in item["fields"] if not f["optional"])
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
            res["region_name"] = target.get("name") or region_id
            try:
                out_dir = region_dirs(region_id)["out"]
                xlsx_path = save_booking_result(res, target, out_dir)
                res["excel_name"] = os.path.basename(xlsx_path)
            except Exception as e2:
                res["save_error"] = str(e2)
        target["booking"] = res
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

    meta_old = load_cache_meta(region_id)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    meta_new = {}
    missing = []

    def prepare_slot(slot, file_storage):
        """把本次要用的文件放入暂存目录：新上传优先，其次沿用缓存。
        返回 True=已就绪；False=该槽位无文件（仅可选文件允许）。"""
        target = os.path.join(stage_data, cfg["paths"][slot])
        present_cache = cached_path(region_id, slot)
        if file_storage is not None:
            file_storage.save(target)
            meta_new[slot] = {"original": file_storage.filename, "uploaded_at": now_str}
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

    for slot in UPLOAD_KEYS:
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
        shutil.rmtree(stage, ignore_errors=True)
        return back("缺少文件：" + "、".join(missing) + "（无缓存，必须上传）")

    try:
        compute_region(region, stage_data, stage_out)
    except Exception as e:
        shutil.rmtree(stage, ignore_errors=True)
        return back("计算失败：" + str(e))

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
        shutil.rmtree(stage, ignore_errors=True)
        return back("保存失败：" + str(e) +
                    "（目录内文件正被其他程序占用，通常是 Excel 打开了旧数据/输出文件；"
                    "请关闭后重试。本次运行未改动原数据）")

    try:
        commit_dir(stage_data, final_data)
        commit_dir(stage_out, final_out)
    except Exception as e:
        shutil.rmtree(stage, ignore_errors=True)
        return back("保存失败：" + str(e) + "（文件正被占用，请关闭相关程序后重试）")

    if os.path.isdir(stage):
        shutil.rmtree(stage, ignore_errors=True)
    save_cache_meta(region_id, meta_new)
    return redirect(url_for("index", msg="运行完成，已更新「" + region["name"] + "」", _anchor=region_id))


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
