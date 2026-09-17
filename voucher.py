# -*- coding: utf-8 -*-
"""凭证 PDF 生成 + 台账归档（发货额度 / 订舱申请）。

凭证 = 编号 + 台账 + 归档 PDF：
  · 编号   <大区前缀>-<YYYYMMDD>-<当日序号>，首次生成时固定；
  · 台账   workspaces/quota/<region>/凭证/凭证台账.json（+ 同名 .xlsx 供查阅）；
  · 归档   workspaces/quota/<region>/凭证/<编号>_<标题>.pdf。

台账目录与「数据源 / 输出」同级，重算时整体替换那两个目录不会动到已归档凭证。
同一份申请（申请时间相同）重复下载复用同一张凭证，不会重复建号。

对外只暴露 generate_voucher()。
"""
import json
import os
from datetime import datetime

from fpdf import FPDF
from fpdf.enums import AccessPermission, EncryptionMethod, XPos, YPos
from fpdf.fonts import FontFace

VOUCHER_DIR_NAME = "凭证"
LEDGER_JSON = "凭证台账.json"
LEDGER_XLSX = "凭证台账.xlsx"
OWNER_PASSWORD = "quota-voucher-lock"   # 仅用于锁定编辑权限，打开 PDF 无需密码

BLUE = (29, 78, 216)
GREEN = (21, 128, 61)
DARK = (31, 45, 61)
GRAY = (100, 116, 139)

# 台账表结构（按大区类型区分）
LEDGER_COLUMNS = {
    "russia": ["凭证编号", "生成时间", "申请时间", "数据日期", "申请文件",
               "本次发货总价(万元)", "发货可用额度(万元)", "发货后余额(万元)", "梯数", "PDF文件"],
    "mexico": ["凭证编号", "生成时间", "订舱申请时间", "订舱清单文件",
               "本次订舱合同数", "设备欠款(万美元)", "海运费欠款(万美元)", "欠款合计(万美元)", "PDF文件"],
}
LEDGER_KEYS = {
    "russia": ["no", "created_at", "apply_time", "data_time", "file",
               "total_wan", "grand_wan", "balance_wan", "ladders", "pdf"],
    "mexico": ["no", "created_at", "apply_time", "file",
               "booking_contracts", "arrears_equip", "arrears_freight", "arrears_total", "pdf"],
}

FONT_CANDIDATES = (
    os.path.join("fonts", "simhei.ttf"),
    os.path.join("fonts", "msyh.ttc"),
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/truetype/arphic/uming.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
)
BOLD_CANDIDATES = (
    os.path.join("fonts", "simhei.ttf"),
    os.path.join("fonts", "simsunb.ttf"),
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
)


# ------------------------------------------------------------------ 基础工具

def _resolve(app_dir, candidates):
    for name in candidates:
        path = name if os.path.isabs(name) else os.path.join(app_dir, name)
        if os.path.isfile(path):
            return path
    return None


def _fonts(app_dir):
    regular = _resolve(app_dir, FONT_CANDIDATES)
    if not regular:
        raise RuntimeError("找不到中文字体，请在项目 fonts/ 目录放入 simhei.ttf")
    return regular, (_resolve(app_dir, BOLD_CANDIDATES) or regular)


def _money(value):
    try:
        return f"{float(value):,.2f}"
    except (TypeError, ValueError):
        return "-"


def _new_pdf(app_dir):
    regular, bold = _fonts(app_dir)
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_font("cn", "", regular)
    pdf.add_font("cn", "B", bold)
    pdf.set_margins(15, 15, 15)
    pdf.set_auto_page_break(auto=True, margin=16)
    pdf.set_text_color(*DARK)
    pdf.add_page()
    return pdf


def _head(pdf, title, subtitle, no, created_at, meta_rows):
    pdf.set_font("cn", size=19)
    pdf.cell(0, 11, title, align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("cn", size=12)
    pdf.set_text_color(*GRAY)
    pdf.cell(0, 7, subtitle, align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_text_color(*DARK)
    pdf.ln(1)
    pdf.set_draw_color(148, 163, 184)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(3)

    pdf.set_font("cn", size=10)
    with pdf.table(col_widths=(30, 60, 30, 60), line_height=6.5,
                   borders_layout="NONE", padding=0.6) as t:
        row = t.row()
        row.cell("凭证编号", style=FontFace(emphasis="BOLD"))
        row.cell(no)
        row.cell("生成时间", style=FontFace(emphasis="BOLD"))
        row.cell(created_at)
        for left, right in meta_rows:
            row = t.row()
            row.cell(left, style=FontFace(emphasis="BOLD"))
            row.cell(right)
            row.cell("")
            row.cell("")
    pdf.ln(2)


def _section(pdf, text):
    pdf.set_font("cn", size=12)
    pdf.set_text_color(*BLUE)
    pdf.cell(0, 8, text, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_text_color(*DARK)


def _banner(pdf, text, color=GREEN, fill=(240, 253, 244)):
    pdf.set_font("cn", size=12)
    pdf.set_text_color(*color)
    pdf.set_fill_color(*fill)
    pdf.cell(0, 10, text, align="C", fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_text_color(*DARK)


def _lines(pdf, items, size=10, color=None):
    pdf.set_font("cn", size=size)
    if color:
        pdf.set_text_color(*color)
    for text in items:
        pdf.cell(0, 6.5, text, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    if color:
        pdf.set_text_color(*DARK)


def _table(pdf, headings, rows, widths, total_row=None, aligns=None, size=9.5):
    pdf.set_font("cn", size=size)
    headings_style = FontFace(emphasis="BOLD", fill_color=(224, 239, 255))
    with pdf.table(col_widths=widths, line_height=6.2, headings_style=headings_style,
                   text_align=aligns or ("LEFT",), padding=1.2) as t:
        header = t.row()
        for h in headings:
            header.cell(h)
        for r in rows:
            row = t.row()
            for v in r:
                row.cell(str(v))
        if total_row:
            row = t.row()
            for v in total_row:
                row.cell(str(v), style=FontFace(emphasis="BOLD", fill_color=(254, 243, 199)))
    pdf.ln(2)


def _notes(pdf, items):
    pdf.set_font("cn", size=9)
    pdf.set_text_color(*GRAY)
    for text in items:
        pdf.multi_cell(0, 5.2, "· " + text, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_text_color(*DARK)


def _lock(pdf):
    perms = (AccessPermission.PRINT_HIGH_RES | AccessPermission.PRINT_LOW_RES
             | AccessPermission.COPY_FOR_ACCESSIBILITY)
    pdf.set_encryption(owner_password=OWNER_PASSWORD, user_password="",
                       encryption_method=EncryptionMethod.AES_128, permissions=perms)


# ------------------------------------------------------------------ 台账读写

def voucher_dir(region_base):
    return os.path.join(region_base, VOUCHER_DIR_NAME)


def load_ledger(region_base):
    path = os.path.join(voucher_dir(region_base), LEDGER_JSON)
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and isinstance(data.get("records"), list):
                return data["records"]
        except (OSError, ValueError):
            pass
    return []


def save_ledger(region_base, records):
    target = voucher_dir(region_base)
    os.makedirs(target, exist_ok=True)
    payload = {"updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "records": records}
    with open(os.path.join(target, LEDGER_JSON), "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def _export_ledger(region_kind, records, region_base):
    """台账 Excel：每次新增凭证后整体重写，方便直接查阅/外发。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    keys = LEDGER_KEYS[region_kind]
    wb = Workbook()
    ws = wb.active
    ws.title = "凭证台账"
    ws.append(LEDGER_COLUMNS[region_kind])
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E8FB")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for record in records:
        ws.append([record.get(k, "") for k in keys])

    widths = [18, 20, 20, 20, 34, 20, 20, 20, 18, 40]
    for idx in range(1, len(LEDGER_COLUMNS[region_kind]) + 1):
        ws.column_dimensions[get_column_letter(idx)].width = widths[idx - 1] if idx <= len(widths) else 18
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.00"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(os.path.join(voucher_dir(region_base), LEDGER_XLSX))


def _next_no(records, prefix):
    day = datetime.now().strftime("%Y%m%d")
    head = prefix + "-" + day + "-"
    used = [r.get("no", "") for r in records if str(r.get("no", "")).startswith(head)]
    return head + "%03d" % (len(used) + 1)


def _safe(name):
    for ch in '\\/:*?"<>|':
        name = name.replace(ch, "")
    return name.strip()


# ------------------------------------------------------------------ 俄罗斯

def _build_russia(pdf, region, result, booking, no, created_at):
    grand = float(result["grand"])
    total_wan = float(booking["total_wan"])
    balance = float(booking.get("balance_wan") or 0)
    ok = balance >= 0

    _head(pdf, "发货额度申请凭证",
          (region.get("name") or "").replace("大区", "") + "大区 · " + (region.get("note") or "发货额度"),
          no, created_at,
          [("申请时间", booking.get("calculated_at") or ""),
           ("数据日期", result.get("generated_at") or "")])

    _section(pdf, "一、额度核定结论")
    _banner(pdf,
            ("额度充足  " if ok else "额度不足  ") +
            "本次发货总价 " + _money(total_wan) + " 万元 ≤ 发货可用额度 " + _money(grand) + " 万元",
            GREEN if ok else (185, 28, 28),
            (240, 253, 244) if ok else (254, 242, 242))
    _lines(pdf, [
        "发货后余额：" + _money(balance) + " 万元（＝ 发货可用额度 " + _money(grand) +
        " - 本次发货总价 " + _money(total_wan) + "）",
        "本次发货后余额 ≥ 0 即视为额度充足；本次申请 " + str(booking.get("count", 0)) +
        " 个梯、" + str(len(booking.get("line_totals") or [])) + " 条申请行。",
    ])

    _section(pdf, "二、发货可用额度构成（单位：万元）")
    _table(pdf, ["项目", "金额", "说明"],
           [["已回款金额合计", _money(result["sum2"]), "表B + 表E抵佣金 + 表R"],
            ["已发货合同金额合计", _money(result["sum3"]), "组A有日期梯的合同额合计"],
            ["净额（已回款 - 已发货）", _money(result["net_total"]), ""],
            ["固定额度", _money(result["fixed"]), "区域固定追加"],
            ["未确认回款（LLC 无合同号）", _money(result["llc"]), ""],
            ["发货可用额度", _money(grand), "＝ 净额 + 固定额度 + 未确认回款"],
            ["本次发货总价", _money(total_wan), "本次申请梯号合同额合计"],
            ["发货后余额", _money(balance), "＝ 发货可用额度 - 本次发货总价"]],
           (62, 28, 90), aligns=("LEFT", "RIGHT", "LEFT"))

    _notes(pdf, [
        "金额口径：合同额取表C「合同额（人民币）」，回款为表B / 表E抵佣金 / 表R 三表按合同汇总，单位万元。",
        "已回款金额合计 " + _money(result["sum2"]) + " 万元，已发货合同金额合计 " +
        _money(result["sum3"]) + " 万元，数据日期 " + str(result.get("generated_at") or "") + "。",
        "本凭证由系统按台账自动生成并归档，PDF 已锁定编辑权限；如与系统记录不符，以系统记录为准。",
    ])

    pdf.add_page()
    _section(pdf, "三、本次发货申请明细（申请文件：" + str(booking.get("uploaded_name") or "") + "）")
    rows = [[lt.get("line", ""), lt.get("contract", ""), lt.get("spec", ""),
             lt.get("count", ""), _money(lt.get("total_rmb"))]
            for lt in (booking.get("line_totals") or [])]
    _table(pdf, ["行", "合同号", "梯号规格", "梯数", "金额小计（元）"], rows,
           (12, 38, 78, 14, 38),
           total_row=["合计", "", "", booking.get("count", 0), _money(booking.get("total_rmb"))],
           aligns=("CENTER", "LEFT", "LEFT", "CENTER", "RIGHT"))
    pdf.set_font("cn", size=10)
    pdf.cell(0, 7, "合计： " + _money(booking.get("total_rmb")) + " 元 ＝ " + _money(total_wan) + " 万元",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(2)


def _record_russia(region, result, booking, no, created_at, pdf_name):
    return {
        "no": no,
        "created_at": created_at,
        "source_time": str(booking.get("calculated_at") or ""),
        "apply_time": str(booking.get("calculated_at") or ""),
        "data_time": str(result.get("generated_at") or ""),
        "file": str(booking.get("uploaded_name") or ""),
        "total_wan": round(float(booking.get("total_wan") or 0), 2),
        "grand_wan": round(float(result.get("grand") or 0), 2),
        "balance_wan": round(float(booking.get("balance_wan") or 0), 2),
        "ladders": int(booking.get("count") or 0),
        "pdf": pdf_name,
    }


# ------------------------------------------------------------------ 墨西哥

def _build_mexico(pdf, region, result, booking_file, no, created_at):
    totals = result["totals"]
    equip = float(totals["J"])
    freight = float(totals["K"])

    _head(pdf, "订舱申请凭证",
          (region.get("name") or "").replace("大区", "") + "大区 · 欠款发货额度（万美元）",
          no, created_at,
          [("订舱申请时间", result.get("generated_at") or ""),
           ("数据日期", result.get("generated_at") or "")])

    _section(pdf, "一、订舱申请登记")
    _banner(pdf,
            "本次订舱申请 " + str(result.get("count_booking", 0)) + " 个合同 ／ 共 " +
            str(result.get("count_contracts", 0)) + " 个合同纳入欠款发货额度",
            BLUE, (240, 247, 255))
    _lines(pdf, [
        "订舱清单文件：" + (booking_file or "（缓存）"),
        "订舱清单中指定的梯号计入「已订舱未组A」，其余为已组A又欠款 " +
        str(result.get("count_group_a", 0)) + " 个合同。",
    ])

    _section(pdf, "二、欠款发货额度快照（单位：万美元）")
    _table(pdf, ["项目", "金额（万美元）", "说明"],
           [["合同额合计", _money(totals["B"]), "纳入口径合同额合计"],
            ["已回款合计", _money(totals["C"]), "表B + 表E抵佣金 + 表R"],
            ["已组A金额合计", _money(totals["D"]), "含设备与海运费"],
            ["已组A设备", _money(totals["F"]), "回款优先抵设备款"],
            ["已订舱未组A金额合计", _money(totals["G"]), "含本次订舱"],
            ["已订舱未组A设备", _money(totals["I"]), ""],
            ["设备欠款", _money(equip), "＝ 已组A设备 + 已订舱未组A设备 - 已回款"],
            ["海运费欠款", _money(freight), ""],
            ["欠款合计", _money(equip + freight), "＝ 设备欠款 + 海运费欠款"]],
           (58, 34, 88), aligns=("LEFT", "RIGHT", "LEFT"),
           )
    _banner(pdf,
            "设备欠款 " + _money(equip) + " 万美元　+　海运费欠款 " + _money(freight) +
            " 万美元　=　合计 " + _money(equip + freight) + " 万美元",
            BLUE, (240, 247, 255))

    notes = [
        "金额口径：均为万美元（原币），回款为表B / 表E抵佣金 / 表R 三表按合同汇总。",
        "数据日期 " + str(result.get("generated_at") or "") + "。",
        "本凭证由系统按台账与订舱清单自动生成并归档，PDF 已锁定编辑权限；如与系统记录不符，以系统记录为准。",
    ]
    if result.get("warnings"):
        notes.append("计算提示：" + "；".join(result["warnings"]))
    _notes(pdf, notes)


def _record_mexico(region, result, booking_file, no, created_at, pdf_name):
    totals = result["totals"]
    equip = float(totals["J"])
    freight = float(totals["K"])
    return {
        "no": no,
        "created_at": created_at,
        "source_time": str(result.get("generated_at") or ""),
        "apply_time": str(result.get("generated_at") or ""),
        "file": booking_file or "",
        "booking_contracts": int(result.get("count_booking") or 0),
        "arrears_equip": round(equip, 2),
        "arrears_freight": round(freight, 2),
        "arrears_total": round(equip + freight, 2),
        "pdf": pdf_name,
    }


# ------------------------------------------------------------------ 对外入口

def generate_voucher(region, region_base, app_dir, result, booking=None, booking_file=""):
    """生成（或复用）该大区当前这份申请的凭证 PDF。

    返回 (pdf_path, record)。同一申请时间已归档过则直接复用，不重复建号。
    """
    kind = "mexico" if region.get("kind") == "mexico" else "russia"
    source_time = str((result or {}).get("generated_at") or "")
    if kind == "russia":
        source_time = str((booking or {}).get("calculated_at") or "")

    records = load_ledger(region_base)
    for record in records:
        if record.get("source_time") == source_time and record.get("kind") == kind:
            path = os.path.join(voucher_dir(region_base), record.get("pdf", ""))
            if os.path.isfile(path):
                return path, record

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    prefix = str(region.get("voucher_prefix") or region.get("id", "XX")[:2].upper())
    no = _next_no(records, prefix)

    pdf = _new_pdf(app_dir)
    if kind == "mexico":
        _build_mexico(pdf, region, result, booking_file, no, now)
        title = "订舱申请凭证"
        record = _record_mexico(region, result, booking_file, no, now, "")
    else:
        _build_russia(pdf, region, result, booking or {}, no, now)
        title = "发货额度申请凭证"
        record = _record_russia(region, result, booking or {}, no, now, "")

    pdf_name = _safe(no + "_" + title) + ".pdf"
    target = voucher_dir(region_base)
    os.makedirs(target, exist_ok=True)
    path = os.path.join(target, pdf_name)
    _lock(pdf)
    pdf.output(path)

    record["kind"] = kind
    record["region_id"] = region.get("id")
    record["region_name"] = (region.get("name") or "").replace("大区", "")
    record["pdf"] = pdf_name
    records.append(record)
    save_ledger(region_base, records)
    try:
        _export_ledger(kind, records, region_base)
    except Exception as exc:      # 台账 Excel 被占用等情况不影响凭证下载
        print("凭证台账 Excel 写入失败（可忽略）：", exc)
    return path, record
