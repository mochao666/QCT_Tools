# -*- coding: utf-8 -*-
"""
命令行：从 PDT 文件生成 QCT 文件（仅两 Sheet 数据，无 List/数据验证）。
用法示例：
  python pdt_to_qct.py "path/to/PDT_Study123.xlsx"
  python pdt_to_qct.py "path/to/PDT.xlsx" -o "output_QCT.xlsx" -t "template.xlsx"
"""

import argparse
import os
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook

from config import (
    SHEET_SDTM,
    SHEET_ADAM_TFL,
    OUTPUT_TYPE_SDTM,
    QCT_HEADERS_SDTM,
    QCT_HEADERS_ADAM_TFL,
)
from pdt_reader import read_and_clean_pdt
from qct_template import create_empty_qct_workbook


# ---------------------------------------------------------------------------
# 单元格与行转换
# ---------------------------------------------------------------------------

def _normalize_cell_value(val):
    """将单元格值转为字符串；空/NaN 为 ''，日期格式化为 YYYY-MM-DD。"""
    if pd.isna(val):
        return ""
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime("%Y-%m-%d")
    return val


def _row_to_qct_values(headers_config, row: pd.Series) -> list:
    """根据 (QCT列名, PDT列名) 配置，将 PDT 一行转为 QCT 一行 10 列值。"""
    values = []
    for qct_col, pdt_col in headers_config:
        if pdt_col is None:
            values.append("")
        else:
            values.append(_normalize_cell_value(row.get(pdt_col)))
    return values


def build_output_path(pdt_path: str, output_path: str = None) -> str:
    """未指定输出路径时，在 PDT 同目录下生成「原文件名_QCT_Template.xlsx」。"""
    if output_path:
        return output_path
    base = os.path.splitext(os.path.basename(pdt_path))[0]
    return os.path.join(os.path.dirname(pdt_path), f"{base}_QCT_Template.xlsx")


# ---------------------------------------------------------------------------
# 主逻辑：PDT -> QCT 写入
# ---------------------------------------------------------------------------

def generate_qct(
    pdt_path: str,
    output_path: str = None,
    qct_template_path: str = None,
) -> str:
    """
    从 PDT 生成 QCT：读取 PDT、按 Output Type 分 Sheet 写入行数据。
    qct_template_path 若提供则用该工作簿的 SDTM/ADaM 表头，否则用 create_empty_qct_workbook。
    返回生成的 QCT 文件路径。
    """
    pdt_clean = read_and_clean_pdt(pdt_path)
    out_path = build_output_path(pdt_path, output_path)

    if qct_template_path and os.path.isfile(qct_template_path):
        wb = load_workbook(qct_template_path)
        ws_sdtm = wb[SHEET_SDTM]
        ws_adam = wb[SHEET_ADAM_TFL]
    else:
        wb = create_empty_qct_workbook()
        ws_sdtm = wb[SHEET_SDTM]
        ws_adam = wb[SHEET_ADAM_TFL]

    for _, row in pdt_clean.iterrows():
        output_type = str(row.get("Output Type", "")).strip().upper()
        if output_type == OUTPUT_TYPE_SDTM:
            ws_sdtm.append(_row_to_qct_values(QCT_HEADERS_SDTM, row))
        else:
            ws_adam.append(_row_to_qct_values(QCT_HEADERS_ADAM_TFL, row))
    wb.save(out_path)
    return out_path


# ---------------------------------------------------------------------------
# 命令行入口
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="从 PDT Excel 生成 QCT 文件，实现 PDT 与 QCT 联动。"
    )
    parser.add_argument(
        "pdt",
        help="PDT 文件路径（Excel）",
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="输出的 QCT 文件路径（默认：与 PDT 同目录，文件名加 _QCT_Template）",
    )
    parser.add_argument(
        "-t", "--template",
        default=None,
        help="可选：QCT 模板文件路径（仅表头）。不指定则程序自动生成表头。",
    )
    args = parser.parse_args()

    pdt_path = os.path.abspath(args.pdt)
    if not os.path.isfile(pdt_path):
        print(f"错误：找不到 PDT 文件 {pdt_path}")
        return 1

    try:
        out_path = generate_qct(
            pdt_path,
            output_path=args.output,
            qct_template_path=args.template,
        )
        print(f"QCT 文件已生成: {out_path}")
        return 0
    except Exception as e:
        print(f"错误: {e}")
        return 1


if __name__ == "__main__":
    exit(main())
