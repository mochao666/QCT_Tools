# -*- coding: utf-8 -*-
"""从 PDT 构建 QCT 内存数据（按 Sheet 的行列表），并支持写入 Excel。"""

from datetime import datetime

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from config import (
    SHEET_SDTM,
    SHEET_ADAM_TFL,
    OUTPUT_TYPE_SDTM,
    QCT_HEADERS_SDTM,
    QCT_HEADERS_ADAM_TFL,
)


def _normalize_cell_value(val):
    if pd.isna(val):
        return ""
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime("%Y-%m-%d")
    return val


def _row_to_qct_values(headers_config, row: pd.Series) -> list:
    """根据 (QCT列名, PDT列名) 配置和 PDT 的一行，生成 QCT 该行各列的值（列表）。"""
    values = []
    for qct_col, pdt_col in headers_config:
        if pdt_col is None:
            values.append("")
        else:
            val = _normalize_cell_value(row.get(pdt_col))
            values.append(val)
    return values


# 行数据：前 9 列为 QCT 表内容，第 10 列为 Event，第 11 列为 Category（仅用于 Comments 过滤）。按行区分 Event 以支持「在已有 QCT 上增加行」。
QCT_NUM_COLUMNS = 9
EVENT_COL_INDEX = 9
CATEGORY_COL_INDEX = 10


def build_qct_rows_from_pdt(pdt_df: pd.DataFrame, event_value: str = ""):
    """
    从已清洗的 PDT DataFrame 生成两个 Sheet 的行数据。
    每行末尾带 Event（索引 9）、Category（索引 10）；event_value 用于新生成或增加行时标记该批行的 Event。
    返回 (sdtm_rows, adam_tfl_rows)，每个为 list of list，顺序与 QCT_HEADERS_* 一致 + Event + Category。
    """
    sdtm_rows = []
    adam_tfl_rows = []
    for _, row in pdt_df.iterrows():
        output_type = str(row.get("Output Type", "")).strip().upper()
        raw_cat = row.get("Category", "")
        if raw_cat is None or (isinstance(raw_cat, float) and pd.isna(raw_cat)):
            raw_cat = ""
        category = str(raw_cat).strip()
        # 将 "nan"、"None" 等视为空；仅统计并加入 Category 非空的行
        if not category or category.upper() in ("NAN", "NONE", "NAT"):
            continue
        if output_type == OUTPUT_TYPE_SDTM:
            sdtm_rows.append(_row_to_qct_values(QCT_HEADERS_SDTM, row) + [event_value, category])
        else:
            adam_tfl_rows.append(_row_to_qct_values(QCT_HEADERS_ADAM_TFL, row) + [event_value, category])
    return sdtm_rows, adam_tfl_rows


def get_sdtm_headers():
    return [h[0] for h in QCT_HEADERS_SDTM]


def get_adam_tfl_headers():
    return [h[0] for h in QCT_HEADERS_ADAM_TFL]


# 可编辑列在表头中的索引（0-based），供 GUI 使用
EDITABLE_COL_QC_DESC = 2   # QC results description (e.g. details of findings or passed)
EDITABLE_COL_DATE_OF_QC = 3   # Date of QC
EDITABLE_COL_FOLLOWUP_NOTES = 8   # Specify Notes (If Final Status="Followup")


def _openpyxl_cell_value(cell):
    """从 openpyxl 单元格取值，空为 ''，日期格式化为 Y-m-d。"""
    v = cell.value
    if v is None:
        return ""
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.strftime("%Y-%m-%d")
    return "" if pd.isna(v) else str(v)


def read_qct_workbook(path: str):
    """
    从 QCT Excel 文件读取两个 Sheet 的数据，返回 (sdtm_rows, adam_tfl_rows)。
    每行为 9 列 QCT 数据 + Category（导入时默认 "Output"）。
    会过滤掉「QC results description (e.g. details of findings or passed)」为空的行。
    """
    wb = load_workbook(path, data_only=True)
    sdtm_rows = []
    adam_tfl_rows = []

    first_event = ""

    def sheet_to_rows(sheet_name, out_list):
        nonlocal first_event
        if sheet_name not in wb.sheetnames:
            return
        ws = wb[sheet_name]
        for row_cells in ws.iter_rows(min_row=2, max_col=10):  # A=Event, B-J=9 列 QCT
            values = [_openpyxl_cell_value(c) for c in row_cells]
            if len(values) < 10:
                values.extend([""] * (10 - len(values)))
            if not first_event and values[0]:
                first_event = str(values[0]).strip()
            qct_vals = values[1:10]
            if len(qct_vals) < 9:
                qct_vals.extend([""] * (9 - len(qct_vals)))
            qct_vals = qct_vals[:9]
            qc_desc = (qct_vals[EDITABLE_COL_QC_DESC] or "").strip()
            if not qc_desc:
                continue
            event_val = str(values[0]).strip() if values[0] else ""
            out_list.append(qct_vals + [event_val, "Output"])

    sheet_to_rows(SHEET_SDTM, sdtm_rows)
    sheet_to_rows(SHEET_ADAM_TFL, adam_tfl_rows)
    return sdtm_rows, adam_tfl_rows, first_event


# 导出 Comments 表头：第一列 Event（单行），前 5 列「中文\n英文」两行，最后两列 Developers、Validators 单行
COMMENTS_HEADERS_ZH = ["表格编号", "标题", "问题描述", "记录日期", "记录人"]
COMMENTS_HEADERS_EN = ["TFL No.", "TFL Title", "Issue Description", "Date Issued", "Issued by"]
COMMENTS_HEADERS_COMBINED = ["Event"] + [f"{zh}\n{en}" for zh, en in zip(COMMENTS_HEADERS_ZH, COMMENTS_HEADERS_EN)] + ["Developers", "Validators"]
# QCT 行中 Developers / Validators 对应列（Person Responsible for Resolution、QC programmer Name）
DEVELOPERS_COL_INDEX = 5
VALIDATORS_COL_INDEX = 4


def write_comments_workbook(sdtm_rows: list, adam_tfl_rows: list, path: str, event_value: str = ""):
    """将当前 QCT 中 Category=Output 的审阅意见导出为单表 Excel。第一列 Event 按 event_value 自动填充（与导入 PDT 时选择的 Event 一致）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "审阅意见"
    # 单行表头：每格内为「中文\n英文」，灰色背景、自动换行
    ws.append(COMMENTS_HEADERS_COMBINED)
    grey_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    wrap_align = Alignment(wrap_text=True)
    for c in range(1, len(COMMENTS_HEADERS_COMBINED) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = grey_fill
        cell.alignment = wrap_align
    # 列宽：Event、表格编号、标题、问题描述、记录日期、记录人、Developers、Validators
    col_widths = (12, 18, 60, 60, 16, 14, 28, 18)
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    # 隐藏 G/H 列（Developers、Validators），数据仍保留
    ws.column_dimensions[get_column_letter(7)].hidden = True
    ws.column_dimensions[get_column_letter(8)].hidden = True
    # 只保留 Category=Output 的记录；A=Event 使用 event_value（导入 PDT 时选择的 Event）自动填充
    def emit_output_rows(rows):
        for row in rows:
            if len(row) <= CATEGORY_COL_INDEX:
                continue
            cat = str(row[CATEGORY_COL_INDEX]).strip().upper()
            if cat != "OUTPUT":
                continue
            developers = row[DEVELOPERS_COL_INDEX] if len(row) > DEVELOPERS_COL_INDEX else ""
            validators = row[VALIDATORS_COL_INDEX] if len(row) > VALIDATORS_COL_INDEX else ""
            ws.append([event_value or "", row[0], row[1], "", "", "", developers, validators])
    emit_output_rows(sdtm_rows)
    emit_output_rows(adam_tfl_rows)
    wb.save(path)


# QCT 两个 Sheet：第一列为 Event（整份文件同一值），后 9 列为 QCT 内容；列宽
QCT_COL_WIDTHS = (12, 15, 60, 45, 15, 15, 28, 15, 15, 45)
# 表头第一行（仅 9 列 QCT；导出时在首列前插入 Event）
QCT_HEADER_ROW_SDTM = [
    "SDTM Datasets",
    "QC checklist-index",
    "QC results description\n(e.g. details of findings or passed)",
    "Date of QC",
    "QC programmer\nName",
    "Person Responsible\nfor Resolution if\nFindings",
    "Date of Resolved\nif Findings",
    "Final Status if\nFindings",
    'Specify Notes (If\nFinal Status="Followup")',
]
QCT_HEADER_ROW_ADAM = [
    "ADaM Dataset/\nTFL Number",
    "QC checklist-index",
    "QC results description\n(e.g. details of findings or passed)",
    "Date of QC",
    "QC programmer\nName",
    "Person Responsible\nfor Resolution if\nFindings",
    "Date of Resolved\nif Findings",
    "Final Status if\nFindings",
    'Specify Notes (If\nFinal Status="Followup")',
]
FILL_LIGHT_BLUE = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
FILL_LIGHT_GREEN = PatternFill(start_color="D8E4BC", end_color="D8E4BC", fill_type="solid")
HEADER_FONT = Font(bold=True)
HEADER_ALIGN = Alignment(wrap_text=True)
# List Sheet 表头（Final Status for Findings 单元格内换行）
SHEET_LIST_NAME = "List"
LIST_HEADERS = ["Event", "Final Status for\nFindings", "User List"]
LIST_COL_WIDTHS = (15, 22, 25)
# Event 选项，展示在 List 表 A 列
EVENT_OPTIONS_LIST = ["CSR", "Dryrun", "IA", "DMC", "EOP2"]
# Final Status for Findings 三行选项，展示在 List 表 B 列，并供 SDTM/ADaM 的 Final Status if Findings 下拉引用
FINAL_STATUS_OPTIONS_LIST = ["Open", "Closed", "Follow up"]


def _read_users_from_pdt(pdt_path: str):
    """从 PDT 文件的 Users 表第一列读取内容，返回非空字符串列表。"""
    if not pdt_path or not isinstance(pdt_path, str):
        return []
    try:
        pdt_wb = load_workbook(pdt_path, data_only=True)
        ws = None
        for name in pdt_wb.sheetnames:
            if name.strip().upper() == "USERS":
                ws = pdt_wb[name]
                break
        if ws is None:
            return []
        users = []
        for row in ws.iter_rows(min_col=1, max_col=1, min_row=1):
            val = row[0].value
            if val is not None and str(val).strip():
                users.append(str(val).strip())
        return users
    except Exception:
        return []


def _set_qct_header_row(ws, headers: list):
    """设置 QCT 表头第一行：文案、换行、加粗、浅蓝/浅绿背景（C 列即 QC checklist-index 浅绿，其余浅蓝）。"""
    for c, text in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c)
        cell.value = text
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.fill = FILL_LIGHT_GREEN if c == 3 else FILL_LIGHT_BLUE


def write_qct_workbook(sdtm_rows: list, adam_tfl_rows: list, path: str, pdt_path: str = None, event_value: str = "", export_mode: str = "initial"):
    """将内存中的两个 Sheet 行数据写入 Excel。两 Sheet 第一列（A）为 Event，后 9 列为 QCT 内容；含 List 表。
    export_mode:
      - "initial"（初版QCT）: 所有行 A 列统一使用 event_value
      - "append"（新增Event）: 若行本身带 Event（row[EVENT_COL_INDEX] 非空）则使用该值；否则回退使用 event_value。"""
    use_single_event = export_mode == "initial"
    header_sdtm = ["Event"] + QCT_HEADER_ROW_SDTM
    header_adam = ["Event"] + QCT_HEADER_ROW_ADAM
    wb = Workbook()
    ws_sdtm = wb.active
    ws_sdtm.title = SHEET_SDTM
    _set_qct_header_row(ws_sdtm, header_sdtm)
    for row in sdtm_rows:
        if use_single_event:
            ev = event_value or ""
        else:
            row_event = ""
            if len(row) > EVENT_COL_INDEX:
                row_event = (row[EVENT_COL_INDEX] or "").strip()
            ev = row_event or (event_value or "")
        ws_sdtm.append([ev] + row[:QCT_NUM_COLUMNS])
    for i, w in enumerate(QCT_COL_WIDTHS, start=1):
        ws_sdtm.column_dimensions[get_column_letter(i)].width = w

    ws_adam = wb.create_sheet(title=SHEET_ADAM_TFL)
    _set_qct_header_row(ws_adam, header_adam)
    for row in adam_tfl_rows:
        if use_single_event:
            ev = event_value or ""
        else:
            row_event = ""
            if len(row) > EVENT_COL_INDEX:
                row_event = (row[EVENT_COL_INDEX] or "").strip()
            ev = row_event or (event_value or "")
        ws_adam.append([ev] + row[:QCT_NUM_COLUMNS])
    for i, w in enumerate(QCT_COL_WIDTHS, start=1):
        ws_adam.column_dimensions[get_column_letter(i)].width = w

    # List 表：第一行浅蓝；A 列 Event 选项；B 列 Final Status 选项；C 列从 PDT Users 表第一列读取
    ws_list = wb.create_sheet(title=SHEET_LIST_NAME)
    for c, text in enumerate(LIST_HEADERS, start=1):
        cell = ws_list.cell(row=1, column=c)
        cell.value = text
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.fill = FILL_LIGHT_BLUE
    for i, w in enumerate(LIST_COL_WIDTHS, start=1):
        ws_list.column_dimensions[get_column_letter(i)].width = w
    # A 列（Event）展示选项：CSR, Dryrun, IA, DMC, EOP2
    for r, opt in enumerate(EVENT_OPTIONS_LIST, start=2):
        ws_list.cell(row=r, column=1, value=opt)
    # B 列（Final Status for Findings）展示三行内容：Open, Closed, Follow up
    for r, opt in enumerate(FINAL_STATUS_OPTIONS_LIST, start=2):
        ws_list.cell(row=r, column=2, value=opt)
    # C 列（User List）从 PDT Users 表第一列填入
    user_list = _read_users_from_pdt(pdt_path) if pdt_path else []
    for r, user in enumerate(user_list, start=2):
        ws_list.cell(row=r, column=3, value=user)

    # 两 Sheet 的 A 列（Event）设为下拉，选项引用 List 表 A2:A6
    event_ref = f"{SHEET_LIST_NAME}!$A$2:$A${1 + len(EVENT_OPTIONS_LIST)}"
    dv_event = DataValidation(type="list", formula1=event_ref, allow_blank=True)
    dv_event.add("A2:A1000")
    ws_sdtm.add_data_validation(dv_event)
    ws_adam.add_data_validation(dv_event)
    # 「Final Status if Findings」列（I 列，因首列插入了 Event）
    list_ref = f"{SHEET_LIST_NAME}!$B$2:$B$4"
    dv_final = DataValidation(type="list", formula1=list_ref, allow_blank=True)
    dv_final.add("I2:I1000")
    ws_sdtm.add_data_validation(dv_final)
    ws_adam.add_data_validation(dv_final)
    # F、G 列（User List 下拉，原 E、F 列顺延）
    if user_list:
        last_row = 1 + len(user_list)
        user_ref = f"{SHEET_LIST_NAME}!$C$2:$C${last_row}"
        dv_user_sdtm = DataValidation(type="list", formula1=user_ref, allow_blank=True)
        dv_user_sdtm.add("F2:F1000")
        dv_user_sdtm.add("G2:G1000")
        ws_sdtm.add_data_validation(dv_user_sdtm)
        dv_user_adam = DataValidation(type="list", formula1=user_ref, allow_blank=True)
        dv_user_adam.add("F2:F1000")
        dv_user_adam.add("G2:G1000")
        ws_adam.add_data_validation(dv_user_adam)

    wb.save(path)
