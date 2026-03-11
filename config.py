# -*- coding: utf-8 -*-
"""PDT 与 QCT 列名、Sheet 名等配置。"""

# PDT 文件中需要提取的列（逻辑名，用于程序内部）
PDT_COLUMNS = [
    "Output Type",
    "Output Reference",
    "Title",
    "Developers",
    "Validators",
    "Date Checked by Trial Statistician",
    "Category",
]

# PDT 实际列名映射（若 Excel 表头与上不同，在此配置：逻辑名 -> 实际列名）
# 当前按 PDT_Template.xlsx 表头：OUTTYPE, OUTREF, OUTTITLE, USERDEV, USERQC, STATDATE, CATEGORY
PDT_COLUMN_MAPPING = {
    "Output Type": "OUTTYPE",
    "Output Reference": "OUTREF",
    "Title": "OUTTITLE",
    "Developers": "USERDEV",
    "Validators": "USERQC",
    "Date Checked by Trial Statistician": "STATDATE",
    "Category": "CATEGORY",
}

# QCT 两个 Sheet 的名称
SHEET_SDTM = "SDTM(aCRF, SPEC and Coding)"
SHEET_ADAM_TFL = "ADaM(SPEC and Coding) and TFL"

# Output Type 取值：用于判断写入哪个 Sheet
OUTPUT_TYPE_SDTM = "SDTM"

# 各 Sheet 的表头及与 PDT 的对应关系（9 列）
# 键为 QCT 列名，值为 PDT 列名（None 表示该列留空）
QCT_HEADERS_SDTM = [
    ("SDTM Datasets", "Output Reference"),
    ("QC checklist-index", "Title"),
    ("QC results description (e.g. details of findings or passed)", None),
    ("Date of QC", "Date Checked by Trial Statistician"),
    ("QC programmer Name", "Validators"),
    ("Person Responsible for Resolution if Findings", "Developers"),
    ("Date of Resolved if Findings", None),
    ("Final Status if Findings", None),
    ("Specify Notes (If Final Status=\"Followup\")", None),
]

QCT_HEADERS_ADAM_TFL = [
    ("ADaM Dataset / TFL Number", "Output Reference"),
    ("QC checklist-index", "Title"),
    ("QC results description (e.g. details of findings or passed)", None),
    ("Date of QC", "Date Checked by Trial Statistician"),
    ("QC programmer Name", "Validators"),
    ("Person Responsible for Resolution if Findings", "Developers"),
    ("Date of Resolved if Findings", None),
    ("Final Status if Findings", None),
    ("Specify Notes (If Final Status=\"Followup\")", None),
]
