# -*- coding: utf-8 -*-
"""生成仅含表头的 QCT 模板文件，便于自定义后配合 pdt_to_qct.py 的 -t 参数使用。"""

import argparse
import os

from qct_template import save_qct_template


def main():
    parser = argparse.ArgumentParser(description="生成仅含表头的 QCT 模板 Excel 文件。")
    parser.add_argument(
        "-o", "--output",
        default="QCT_template.xlsx",
        help="输出模板文件路径（默认: QCT_template.xlsx）",
    )
    args = parser.parse_args()
    path = os.path.abspath(args.output)
    save_qct_template(path)


if __name__ == "__main__":
    main()
