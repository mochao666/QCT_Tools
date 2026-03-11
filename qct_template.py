# -*- coding: utf-8 -*-
"""创建仅含表头的 QCT 模板工作簿。"""

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from config import (
    SHEET_SDTM,
    SHEET_ADAM_TFL,
    QCT_HEADERS_SDTM,
    QCT_HEADERS_ADAM_TFL,
)


def get_sdtm_headers():
    """Sheet1 表头列名列表。"""
    return [h[0] for h in QCT_HEADERS_SDTM]


def get_adam_tfl_headers():
    """Sheet2 表头列名列表。"""
    return [h[0] for h in QCT_HEADERS_ADAM_TFL]


def create_empty_qct_workbook():
    """
    在内存中创建一个只有表头、无数据的 QCT 工作簿。
    返回 openpyxl.Workbook 实例。
    """
    wb = Workbook()
    # 默认会有一个 Sheet，用作 SDTM
    ws_sdtm = wb.active
    ws_sdtm.title = SHEET_SDTM
    ws_sdtm.append(get_sdtm_headers())

    ws_adam = wb.create_sheet(title=SHEET_ADAM_TFL)
    ws_adam.append(get_adam_tfl_headers())

    return wb


def save_qct_template(path: str):
    """将仅含表头的 QCT 模板保存到 path。"""
    wb = create_empty_qct_workbook()
    wb.save(path)
    print(f"QCT 模板已保存: {path}")
