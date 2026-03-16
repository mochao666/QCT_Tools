# -*- coding: utf-8 -*-
"""
PDT 与 QCT 列名、Sheet 名等配置。
修改本文件可适配不同模板的 PDT/QCT 表头。
"""

# ---------------------------------------------------------------------------
# PDT 列配置
# ---------------------------------------------------------------------------
# 必选列（逻辑名）：程序从 PDT 中按这些名字或 PDT_COLUMN_MAPPING 匹配列
PDT_COLUMNS = [
    "Output Type",
    "Output Reference",
    "Title",
    "Developers",
    "Validators",
    "Date Checked by Trial Statistician",
    "Category",
]

# 逻辑名 -> 实际表头名（匹配时忽略大小写、换行视为空格）
# 若 PDT 表头与逻辑名不同（如旧模板 "OUTTYPE"），在此配置映射
# RTF Combine 为可选列：存在时导出 Comments 仅保留该列为 'Y' 的行
PDT_COLUMN_MAPPING = {
    "Output Type": "Output Type",
    "Output Reference": "Output Reference",
    "Title": "Title",
    "Developers": "Developers",
    "Validators": "Validators",
    "Date Checked by Trial Statistician": "Validation Date",
    "Category": "Category",
    "RTF Combine": "RTF Combine",
}

# ---------------------------------------------------------------------------
# QCT Sheet 名称与输出类型
# ---------------------------------------------------------------------------
SHEET_SDTM = "SDTM(aCRF, SPEC and Coding)"
SHEET_ADAM_TFL = "ADaM(SPEC and Coding) and TFL"

# Output Type 为 SDTM 时写入第一个 Sheet，否则写入第二个 Sheet
OUTPUT_TYPE_SDTM = "SDTM"

# ---------------------------------------------------------------------------
# QCT 表头与 PDT 列对应关系（10 列）
# 每项为 (QCT 列名, PDT 列名)，PDT 列名为 None 表示该列留空由用户填写
# ---------------------------------------------------------------------------
QCT_HEADERS_SDTM = [
    ("SDTM Datasets", "Output Reference"),
    ("QC checklist-index", "Title"),
    ("QC results description (e.g. details of findings or passed)", None),
    ("Date of QC", "Date Checked by Trial Statistician"),
    ("QC programmer Name", "Validators"),
    ("Person Responsible for Resolution if Findings", "Developers"),
    ("Date of Resolved if Findings", None),
    ("Resolution Details", None),
    ("Final Status if Findings", None),
    ("Special Notes", None),
]

QCT_HEADERS_ADAM_TFL = [
    ("ADaM Dataset / TFL Number", "Output Reference"),
    ("QC checklist-index", "Title"),
    ("QC results description (e.g. details of findings or passed)", None),
    ("Date of QC", "Date Checked by Trial Statistician"),
    ("QC programmer Name", "Validators"),
    ("Person Responsible for Resolution if Findings", "Developers"),
    ("Date of Resolved if Findings", None),
    ("Resolution Details", None),
    ("Final Status if Findings", None),
    ("Special Notes", None),
]
