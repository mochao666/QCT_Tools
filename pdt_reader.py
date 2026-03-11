# -*- coding: utf-8 -*-
"""PDT 文件读取与清洗，供命令行和 GUI 共用。"""

import pandas as pd

from config import PDT_COLUMNS, PDT_COLUMN_MAPPING


def _normalize_col_key(name):
    """表头匹配用：任意空白（含换行）压成单个空格，再转大写。"""
    return " ".join(str(name).split()).upper()


def _find_column_ignore_case(df, *candidates):
    """按「忽略大小写、空白/换行视为空格」匹配列名，返回实际列名，未找到返回 None。"""
    col_map = {_normalize_col_key(c): c for c in df.columns}
    for name in candidates:
        key = _normalize_col_key(name)
        if key in col_map:
            return col_map[key]
    return None


def read_and_clean_pdt(pdt_path: str) -> pd.DataFrame:
    """
    读取 PDT 文件并提取所需列，形成主词典。
    若存在 PDT_COLUMN_MAPPING，则按映射从实际列名读取并重命名为逻辑列名。
    Category 列按「去除首尾空格、忽略大小写」匹配，避免表头不一致导致整列被误填为 Output。
    """
    # PDT 表头从第 3 行开始（Excel 行号 3，pandas header=2）
    pdt_df = pd.read_excel(pdt_path, header=2)
    # 统一把表头首尾空格去掉，便于后续匹配（不改原表，只用于找列）
    required_cols = [c for c in PDT_COLUMNS if c != "Category"]
    required_ok = all(_find_column_ignore_case(pdt_df, c) is not None for c in required_cols)
    if not required_ok:
        missing = [c for c in required_cols if _find_column_ignore_case(pdt_df, c) is None]
        if not PDT_COLUMN_MAPPING:
            raise ValueError(
                f"PDT 文件中缺少以下列: {missing}。当前列名: {list(pdt_df.columns)}"
            )
        physical_cols = list(PDT_COLUMN_MAPPING.values())
        category_physical = PDT_COLUMN_MAPPING.get("Category")
        required_physical = [c for c in physical_cols if c != category_physical] if category_physical else physical_cols
        missing = [c for c in required_physical if _find_column_ignore_case(pdt_df, c) is None]
        if missing:
            raise ValueError(
                f"PDT 文件中缺少映射列: {missing}。当前列名: {list(pdt_df.columns)}"
            )
        physical_to_logical = {v: k for k, v in PDT_COLUMN_MAPPING.items()}
        cols_to_read = []
        rename_to_logical = {}
        for phys in physical_cols:
            actual = _find_column_ignore_case(pdt_df, phys)
            if actual:
                cols_to_read.append(actual)
                if phys in physical_to_logical and actual != physical_to_logical[phys]:
                    rename_to_logical[actual] = physical_to_logical[phys]
        pdt_clean = pdt_df[cols_to_read].copy()
        pdt_clean = pdt_clean.rename(columns=rename_to_logical)
        category_col = _find_column_ignore_case(pdt_clean, "Category", PDT_COLUMN_MAPPING.get("Category") or "CATEGORY")
        if not category_col and len(pdt_df.columns) > 0:
            category_col = pdt_df.columns[0]  # PDT 中 Category 常为第一列
        if category_col and category_col not in pdt_clean.columns and category_col in pdt_df.columns:
            pdt_clean[category_col] = pdt_df[category_col].values
        if category_col and category_col in pdt_clean.columns and category_col != "Category":
            pdt_clean = pdt_clean.rename(columns={category_col: "Category"})
        if "Category" not in pdt_clean.columns:
            pdt_clean["Category"] = ""  # 无 Category 列时视为全空，只统计列存在且非空的行
        return pdt_clean

    # 所需列都存在（按忽略大小写+去空格匹配）
    cols = []
    for c in PDT_COLUMNS:
        if c == "Category":
            continue
        actual = _find_column_ignore_case(pdt_df, c)
        if actual:
            cols.append(actual)
    category_col = _find_column_ignore_case(
        pdt_df, "Category", PDT_COLUMN_MAPPING.get("Category") if PDT_COLUMN_MAPPING else None, "CATEGORY"
    )
    if not category_col and len(pdt_df.columns) > 0:
        category_col = pdt_df.columns[0]  # PDT 中 Category 是第一列
    if category_col:
        if category_col not in cols:
            cols.append(category_col)
    pdt_clean = pdt_df[cols].copy()
    if category_col and category_col != "Category":
        pdt_clean = pdt_clean.rename(columns={category_col: "Category"})
    if "Category" not in pdt_clean.columns:
        pdt_clean["Category"] = ""  # 无 Category 列时视为全空，只统计列存在且非空的行
    # 统一成逻辑列名：对非 Category 的列，若实际列名与逻辑名不同则重命名（按 config 映射或忽略大小写）
    rename_map = {}
    for logical in PDT_COLUMNS:
        if logical == "Category":
            continue
        actual = _find_column_ignore_case(pdt_clean, logical)
        if actual and actual != logical:
            rename_map[actual] = logical
    if rename_map:
        pdt_clean = pdt_clean.rename(columns=rename_map)
    return pdt_clean
