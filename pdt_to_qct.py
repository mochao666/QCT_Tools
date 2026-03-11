# -*- coding: utf-8 -*-
"""
从 PDT 文件生成 QCT 文件，实现 PDT 与 QCT 联动。
用法示例：
  python pdt_to_qct.py "path/to/PDT_Study123.xlsx" --output "HRxxxxx_xxx_CSR_01_QCT_v1.0_Template.xlsx"
  python pdt_to_qct.py "path/to/PDT.xlsx"  # 输出到同目录下自动命名的 QCT 文件
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


def _normalize_cell_value(val):
    """将单元格值转为适合写入 QCT 的形式（日期格式化为 YYYY-MM-DD）。"""
    if pd.isna(val):
        return ""
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime("%Y-%m-%d")
    return val


def _row_to_qct_values(headers_config, row: pd.Series) -> list:
    """根据 (QCT列名, PDT列名) 配置和 PDT 的一行，生成 QCT 该行各列的值。"""
    values = []
    for qct_col, pdt_col in headers_config:
        if pdt_col is None:
            values.append("")
        else:
            val = _normalize_cell_value(row.get(pdt_col))
            values.append(val)
    return values


def build_output_path(pdt_path: str, output_path: str = None) -> str:
    """若未指定输出路径，则根据 PDT 路径生成默认 QCT 文件名。"""
    if output_path:
        return output_path
    base = os.path.splitext(os.path.basename(pdt_path))[0]
    # 可选：从 base 解析 study 等生成 HRxxxxx_xxx_CSR_01_QCT_v1.0_Template.xlsx
    dirname = os.path.dirname(pdt_path)
    return os.path.join(dirname, f"{base}_QCT_Template.xlsx")


def generate_qct(
    pdt_path: str,
    output_path: str = None,
    qct_template_path: str = None,
) -> str:
    """
    从 PDT 生成 QCT 文件。
    - pdt_path: PDT Excel 路径
    - output_path: 输出 QCT 路径，为空则自动生成
    - qct_template_path: 可选；若提供则从该模板加载（仅用表头），否则内存创建表头
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
            values = _row_to_qct_values(QCT_HEADERS_SDTM, row)
            ws_sdtm.append(values)
        else:
            # ADaM 或 TFL 等均写入第二个 Sheet
            values = _row_to_qct_values(QCT_HEADERS_ADAM_TFL, row)
            ws_adam.append(values)

    wb.save(out_path)
    return out_path


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
