# -*- coding: utf-8 -*-
"""
创建仅含表头的 QCT 模板工作簿。
供命令行 pdt_to_qct 在未指定模板时使用，生成空表头的 SDTM + ADaM 两个 Sheet。
"""

from openpyxl import Workbook

from config import (
    SHEET_SDTM,
    SHEET_ADAM_TFL,
    QCT_HEADERS_SDTM,
    QCT_HEADERS_ADAM_TFL,
)


def get_sdtm_headers():
    """SDTM Sheet 表头列名列表（与 config 中 QCT_HEADERS_SDTM 一致）。"""
    return [h[0] for h in QCT_HEADERS_SDTM]


def get_adam_tfl_headers():
    """ADaM/TFL Sheet 表头列名列表（与 config 中 QCT_HEADERS_ADAM_TFL 一致）。"""
    return [h[0] for h in QCT_HEADERS_ADAM_TFL]


def create_empty_qct_workbook():
    """在内存中创建仅含两 Sheet 表头、无数据行的工作簿。返回 openpyxl.Workbook。"""
    wb = Workbook()
    ws_sdtm = wb.active
    ws_sdtm.title = SHEET_SDTM
    ws_sdtm.append(get_sdtm_headers())
    ws_adam = wb.create_sheet(title=SHEET_ADAM_TFL)
    ws_adam.append(get_adam_tfl_headers())
    return wb


def save_qct_template(path: str):
    """将仅含表头的 QCT 模板保存到指定路径。"""
    wb = create_empty_qct_workbook()
    wb.save(path)
    print(f"QCT 模板已保存: {path}")
