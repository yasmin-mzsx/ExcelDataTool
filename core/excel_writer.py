# -*- coding: utf-8 -*-
# ============================================================
# 📌【BACKEND IMPORTS START】
# 后端：DataFrame -> Excel模板填充，纯逻辑不碰Qt
# ============================================================
import os
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
# ============================================================
# 📌【BACKEND IMPORTS END】
# ============================================================


# ============================================================
# 📌【BACKEND START】
# 负责把处理好的 DataFrame 填充到模板对应sheet，并另存为新文件
# ============================================================

# 折旧明细表是两行表头，数据从第4行开始；其余sheet从第3行开始
_SPECIAL_START_ROW = {
    "折旧明细表": 4,
}
_DEFAULT_START_ROW = 3


def _cell_value(v):
    """把df里的值转成Excel能写的值，NaN/NaT -> None（空单元格）"""
    if v is None:
        return None
    # pandas 的 NA / NaT / nan 都判一下
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    # numpy 类型转 python 原生，避免 openpyxl 报错
    if hasattr(v, "item"):
        try:
            return v.item()
        except Exception:
            pass
    return v


def write_df_to_template(template_path: str, output_dir: str,
                         source_filename: str, doc_key: str,
                         df: pd.DataFrame) -> str:
    """
    把 df 写入模板对应sheet，另存为新文件。

    参数：
        template_path: 模板文件路径
        output_dir:    输出目录
        source_filename: 源文件名（输出文件用这个名字）
        doc_key:       单据名（对应模板里的sheet名）
        df:            处理好的数据（不带列名索引）

    返回：
        输出文件的完整路径
    """
    # 1. 加载模板
    wb = load_workbook(template_path)

    # 2. 找sheet，并删掉其他所有sheet（输出文件只保留这一个）
    if doc_key not in wb.sheetnames:
        raise ValueError(f"模板中找不到sheet：{doc_key}")
    for name in list(wb.sheetnames):
        if name != doc_key:
            del wb[name]
    ws = wb[doc_key]

    # 3. 确定起始行
    start_row = _SPECIAL_START_ROW.get(doc_key, _DEFAULT_START_ROW)

    # 4. 逐格写入数据（不写列名、不写索引）
    n_rows = len(df)
    n_cols = len(df.columns)
    for r in range(n_rows):
        for c in range(n_cols):
            val = _cell_value(df.iat[r, c])
            ws.cell(row=start_row + r, column=1 + c, value=val)

    # 5. 给实际写入的数据区域加框线 + 缩小字体填充
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    align = Alignment(shrink_to_fit=True)
    for r in range(start_row, start_row + n_rows):
        for c in range(1, 1 + n_cols):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            cell.alignment = align

    # 6. 另存为（保留源文件名，扩展名统一xlsx）
    out_name = Path(source_filename).stem + ".xlsx"
    out_path = os.path.join(output_dir, out_name)
    wb.save(out_path)
    wb.close()
    return out_path

# ============================================================
# 📌【BACKEND END】
# ============================================================
