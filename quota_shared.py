# -*- coding: utf-8 -*-
"""公共数据（表C/B/E/R、汇率、表D）的缓存与上传处理。"""
import hashlib
import json
import os
import shutil
from datetime import datetime

import pandas as pd

from quota_core import find_file, load_main_config

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SHARED_ROOT = os.path.join(BASE_DIR, "workspaces", "quota", "_shared")
DATA_DIR = os.path.join(SHARED_ROOT, "数据源")
STAGING_DIR = os.path.join(SHARED_ROOT, "_staging")
BACKUP_DIR = os.path.join(SHARED_ROOT, "_prev_data")
META_NAME = "_cache_meta.json"

SHARED_SLOTS = ["table_c", "table_b", "table_e", "table_r", "exchange_rate", "table_d"]
SLOT_ATTRS = {
    "table_c": {"label": "表C · 国贸合同标的台账", "accept": ".xlsx", "required": True},
    "table_b": {"label": "表B · 外币回款台账", "accept": ".xlsx", "required": True},
    "table_e": {"label": "表E · gm_ht_hthkmx", "accept": ".xls,.xlsx", "required": True},
    "table_r": {"label": "表R · 人民币回款台账", "accept": ".xlsx", "required": True},
    "exchange_rate": {"label": "汇率文件", "accept": ".txt", "required": True},
    "table_d": {"label": "表D · 国家-市场-业务员", "accept": ".xlsx", "required": False},
}


def meta_path():
    return os.path.join(SHARED_ROOT, META_NAME)


def load_meta():
    path = meta_path()
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except Exception:
            pass
    return {}


def save_meta(meta):
    os.makedirs(SHARED_ROOT, exist_ok=True)
    path = meta_path()
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def cfg():
    return load_main_config()


def canonical_name(slot):
    return cfg()["paths"][slot]


def cached_path(slot):
    path = find_file(DATA_DIR, canonical_name(slot))
    return path if os.path.isfile(path) else None


def is_ready(required_slots=None):
    if required_slots is None:
        required_slots = [slot for slot in SHARED_SLOTS if SLOT_ATTRS[slot]["required"]]
    return all(cached_path(slot) for slot in required_slots)


def version():
    return str(load_meta().get("version") or "")


def _file_hash(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _validate_file(slot, path):
    """只做轻量校验，避免上传时重复完整读取大表。"""
    if not os.path.isfile(path) or os.path.getsize(path) == 0:
        return "文件为空或不存在"
    if slot == "exchange_rate":
        return None
    if slot == "table_e":
        last_exc = None
        for engine in ("xlrd", "openpyxl"):
            try:
                pd.read_excel(path, sheet_name=0, nrows=0, engine=engine)
                return None
            except Exception as exc:
                last_exc = exc
        return "无法读取：" + str(last_exc)
    engine = None
    try:
        if slot == "table_d":
            pd.read_excel(path, sheet_name="Sheet3", nrows=0, engine=engine)
        else:
            pd.read_excel(path, sheet_name=0, nrows=0, engine=engine)
    except Exception as exc:
        return "无法读取：" + str(exc)
    return None


def build_fields():
    meta = load_meta()
    files = meta.get("files") or {}
    fields = []
    for slot in SHARED_SLOTS:
        path = cached_path(slot)
        entry = files.get(slot) or {}
        attrs = SLOT_ATTRS[slot]
        fields.append({
            "slot": slot,
            "label": attrs["label"],
            "accept": attrs["accept"],
            "required": attrs["required"],
            "present": path is not None,
            "original": entry.get("original") or (os.path.basename(path) if path else ""),
            "uploaded_at": entry.get("uploaded_at") or "",
        })
    return fields


def save_upload(files):
    """保存公共数据：新上传优先，未上传沿用缓存；校验通过后整体替换。"""
    old_meta = load_meta()
    old_files = old_meta.get("files") or {}
    if os.path.isdir(STAGING_DIR):
        shutil.rmtree(STAGING_DIR)
    stage_data = os.path.join(STAGING_DIR, "数据源")
    os.makedirs(stage_data, exist_ok=True)

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_files = {}
    errors = []
    uploaded_slots = []

    for slot in SHARED_SLOTS:
        name = canonical_name(slot)
        target = os.path.join(stage_data, name)
        f = files.get(slot) if files else None
        if f is not None and getattr(f, "filename", ""):
            f.save(target)
            new_files[slot] = {"original": f.filename, "uploaded_at": now_str}
            uploaded_slots.append(slot)
        else:
            old_path = cached_path(slot)
            if old_path:
                shutil.copy2(old_path, target)
                entry = old_files.get(slot) or {}
                new_files[slot] = {
                    "original": entry.get("original") or os.path.basename(old_path),
                    "uploaded_at": entry.get("uploaded_at") or "",
                }
            elif SLOT_ATTRS[slot]["required"]:
                errors.append(SLOT_ATTRS[slot]["label"] + "：尚未上传")

    for slot in SHARED_SLOTS:
        target = os.path.join(stage_data, canonical_name(slot))
        if not os.path.isfile(target):
            continue
        message = _validate_file(slot, target)
        if message:
            errors.append(SLOT_ATTRS[slot]["label"] + "：" + message)

    if errors:
        shutil.rmtree(STAGING_DIR, ignore_errors=True)
        return {"ok": False, "errors": errors, "uploaded": uploaded_slots}

    os.makedirs(SHARED_ROOT, exist_ok=True)
    if os.path.isdir(BACKUP_DIR):
        shutil.rmtree(BACKUP_DIR)
    try:
        if os.path.isdir(DATA_DIR):
            os.rename(DATA_DIR, BACKUP_DIR)
        os.rename(stage_data, DATA_DIR)
    except Exception:
        if not os.path.isdir(DATA_DIR) and os.path.isdir(BACKUP_DIR):
            os.rename(BACKUP_DIR, DATA_DIR)
        shutil.rmtree(STAGING_DIR, ignore_errors=True)
        raise
    shutil.rmtree(STAGING_DIR, ignore_errors=True)

    for slot in SHARED_SLOTS:
        path = os.path.join(DATA_DIR, canonical_name(slot))
        if os.path.isfile(path):
            new_files.setdefault(slot, {})["sha256"] = _file_hash(path)
            new_files[slot]["size"] = os.path.getsize(path)

    meta = {
        "version": now_str,
        "updated_at": now_str,
        "files": new_files,
    }
    save_meta(meta)
    return {"ok": True, "version": now_str, "uploaded": uploaded_slots}

