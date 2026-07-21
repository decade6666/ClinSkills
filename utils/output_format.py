"""utils/output_format.py — 报表输出聚合入口

历史上本模块同时承载 docx 三线表与 xlsx 清单两类输出（约 650 行、双职责）；
现已按介质拆分为 output_docx.py 与 output_xlsx.py。本文件保留为聚合入口，
re-export 两模块的全部公开函数，使既有 `from utils.output_format import ...` 继续可用。

新代码可直接从对应介质模块导入（output_docx / output_xlsx）；沿用本入口亦可。
"""
from utils.output_docx import (
    set_cell_border,
    set_cell_background,
    set_run_font,
    set_cell_font,
    set_row_height,
    set_table_width,
    calculate_column_widths,
    merge_cells_vertical,
    save_table_to_docx_threeline,
)
from utils.output_xlsx import (
    export_to_excel_with_format,
    export_to_one_excel_with_format,
    export_to_excel_twoheader,
)

__all__ = [
    # docx（output_docx.py）
    "set_cell_border",
    "set_cell_background",
    "set_run_font",
    "set_cell_font",
    "set_row_height",
    "set_table_width",
    "calculate_column_widths",
    "merge_cells_vertical",
    "save_table_to_docx_threeline",
    # xlsx（output_xlsx.py）
    "export_to_excel_with_format",
    "export_to_one_excel_with_format",
    "export_to_excel_twoheader",
]
