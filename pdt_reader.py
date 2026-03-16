# -*- coding: utf-8 -*-
"""
PDT 文件读取与清洗，供命令行（pdt_to_qct）和 GUI（app_gui）共用。
- 表头从 Excel 第 3 行读取（header=2）
- 按 config 中的列名或映射匹配列，缺失必选列会抛错
- 可选列 RTF Combine：若存在则参与导出 Comments 过滤，不存在则默认整列 'Y'
"""

import pandas as pd

from config import PDT_COLUMNS, PDT_COLUMN_MAPPING


# ---------------------------------------------------------------------------
# 列名匹配与可选列补全
# ---------------------------------------------------------------------------

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


def _ensure_rtf_combine_column(pdt_df: pd.DataFrame, pdt_clean: pd.DataFrame) -> None:
    """若 PDT 有 RTF Combine 列则加入 pdt_clean，否则添加一列默认 'Y'（兼容无该列的旧 PDT）。"""
    if "RTF Combine" in pdt_clean.columns:
        return
    rtf_col = _find_column_ignore_case(
        pdt_df, "RTF Combine", (PDT_COLUMN_MAPPING or {}).get("RTF Combine") or "RTF Combine"
    )
    if rtf_col and rtf_col in pdt_df.columns:
        pdt_clean["RTF Combine"] = pdt_df[rtf_col].astype(str).str.strip()
    else:
        pdt_clean["RTF Combine"] = "Y"


# ---------------------------------------------------------------------------
# 主入口：读取并清洗 PDT
# ---------------------------------------------------------------------------

def read_and_clean_pdt(pdt_path: str) -> pd.DataFrame:
    """
    读取 PDT 文件并提取所需列，形成统一逻辑列名的 DataFrame。
    - 若必选列直接用逻辑名匹配不到，则用 PDT_COLUMN_MAPPING 按物理列名匹配并重命名
    - Category 列特殊处理：常为第一列或名称为 CATEGORY，避免表头不一致导致整列被误填
    - 最后补全可选列 RTF Combine
    """
    # PDT 表头从第 3 行开始（Excel 行号 3）
    pdt_df = pd.read_excel(pdt_path, header=2)

    required_cols = [c for c in PDT_COLUMNS if c != "Category"]
    required_ok = all(_find_column_ignore_case(pdt_df, c) is not None for c in required_cols)

    if not required_ok:
        # 走映射分支：按 PDT_COLUMN_MAPPING 的物理列名找列，RTF Combine 不参与必选校验
        missing = [c for c in required_cols if _find_column_ignore_case(pdt_df, c) is None]
        if not PDT_COLUMN_MAPPING:
            raise ValueError(
                f"PDT 文件中缺少以下列: {missing}。当前列名: {list(pdt_df.columns)}"
            )
        physical_cols = list(PDT_COLUMN_MAPPING.values())
        category_physical = PDT_COLUMN_MAPPING.get("Category")
        rtf_physical = PDT_COLUMN_MAPPING.get("RTF Combine")
        required_physical = [
            c for c in physical_cols
            if c != category_physical and c != rtf_physical
        ]
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
        # Category 可能不在 cols_to_read 中，从原表补一列
        category_col = _find_column_ignore_case(pdt_clean, "Category", PDT_COLUMN_MAPPING.get("Category") or "CATEGORY")
        if not category_col and len(pdt_df.columns) > 0:
            category_col = pdt_df.columns[0]
        if category_col and category_col not in pdt_clean.columns and category_col in pdt_df.columns:
            pdt_clean[category_col] = pdt_df[category_col].values
        if category_col and category_col in pdt_clean.columns and category_col != "Category":
            pdt_clean = pdt_clean.rename(columns={category_col: "Category"})
        if "Category" not in pdt_clean.columns:
            pdt_clean["Category"] = ""
        _ensure_rtf_combine_column(pdt_df, pdt_clean)
        return pdt_clean

    # 所需列都存在（按逻辑名忽略大小写匹配）
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
        category_col = pdt_df.columns[0]
    if category_col:
        if category_col not in cols:
            cols.append(category_col)
    pdt_clean = pdt_df[cols].copy()
    if category_col and category_col != "Category":
        pdt_clean = pdt_clean.rename(columns={category_col: "Category"})
    if "Category" not in pdt_clean.columns:
        pdt_clean["Category"] = ""
    # 统一成逻辑列名（与 config 中名称一致）
    rename_map = {}
    for logical in PDT_COLUMNS:
        if logical == "Category":
            continue
        actual = _find_column_ignore_case(pdt_clean, logical)
        if actual and actual != logical:
            rename_map[actual] = logical
    if rename_map:
        pdt_clean = pdt_clean.rename(columns=rename_map)
    _ensure_rtf_combine_column(pdt_df, pdt_clean)
    return pdt_clean
