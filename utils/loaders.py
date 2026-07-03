"""
utils/loaders.py — 数据读取层

统一使用 load_sheet / load_rand 读取 EDC 导出的 Excel 数据。
"""

import json
import pandas as pd
from config import raw_path
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent

# 读取 EDC 类型
_meta_path = _PROJECT_ROOT / "02 metadata" / "FormField.json"
with open(_meta_path, encoding="utf-8") as _f:
    _meta = json.load(_f)
EDC_TYPE = _meta.get("_meta", {}).get("edcType", "")


# ── 系统列注册表 ──
# 同一 EDC 系统的 rawdata 系统列跨研究固定，故作确定性知识登记于此。
# 模板/脚本通过 system_cols() 取值，函数体不硬编码系统列名。
# 新增 EDC 类型时在此补一行即可。
#
# 6 个角色可完全定位 EDC 中的每一个数据点：
#   center      中心编号
#   subject     筛选号
#   visit_name  访视名称
#   visit_seq   访视序号（标记访视重复）
#   form_name   表单名称
#   row         字段行号（标记表单内重复记录）
SYSTEM_COLUMNS: dict[str, dict[str, str]] = {
    "clinflash": {
        "center":     "试验中心编号",
        "subject":    "受试者编号",
        "visit_name": "数据节",
        "visit_seq":  "Instance顺序号",
        "form_name":  "数据页",
        "row":        "行号",
    },
    "taimei5": {
        "center":     "SITEID",
        "subject":    "SUBJID",
        "visit_name": "VISIT",
        "visit_seq":  "VISTREP",
        "form_name":  "FORMNM",
        "row":        "RECREP",
    },
    "taimei6": {
        "center":     "SITEID",
        "subject":    "SUBJID",
        "visit_name": "VISIT",
        "visit_seq":  "VISTREP",
        "form_name":  "FORMNM",
        "row":        "RECREP",
    },
    "cmis": {
        "center":     "SITEID",
        "subject":    "SUBJID",
        "visit_name": "VISIT",
        "visit_seq":  "VISITNUM",
        "form_name":  "FORMNAME",
        "row":        "TOPICSEQ",
    },
}


def system_cols(role: str | None = None) -> dict | str:
    """返回当前 EDC 类型的系统列名。

    Args:
        role: 指定角色（center/subject/visit_name/visit_seq/form_name/row）
              则返回对应列名；省略则返回整 dict。

    Raises:
        ValueError: 当前 EDC 类型未登记，或已登记但缺少请求的角色——
                    提示在 SYSTEM_COLUMNS 中补全。
    """
    cols = SYSTEM_COLUMNS.get(EDC_TYPE)
    if cols is None:
        raise ValueError(
            f"EDC 类型 '{EDC_TYPE}' 未在 utils/loaders.py 的 SYSTEM_COLUMNS 登记，"
            f"请补全该 EDC 的系统列名（center/subject/visit_name/visit_seq/form_name/row）。"
        )
    if role is None:
        return cols
    if role not in cols:
        raise ValueError(
            f"EDC '{EDC_TYPE}' 未登记系统列角色 '{role}'，"
            f"请在 utils/loaders.py 的 SYSTEM_COLUMNS['{EDC_TYPE}'] 补全。"
        )
    return cols[role]


def _resolve_sheet_name(form_oid: str, form_name: str | None = None) -> str:
    """根据 EDC 类型构造实际 sheet 名称。

    cmis 的 sheet 名为 '{formOID}--{formName}'，
    taimei5/taimei6 的 sheet 名直接是 formOID。
    """
    if EDC_TYPE == "cmis":
        if form_name is None:
            # 从元数据自动查找 formName
            for rec in _meta["variables"]:
                if rec["formOID"] == form_oid:
                    form_name = rec["formName"]
                    break
        if form_name:
            return f"{form_oid}--{form_name}"
    return form_oid


def load_sheet(
    form_oid: str,
    usecols: list[str] | None = None,
    form_name: str | None = None,
    dtype: dict | type | None = None,
) -> pd.DataFrame:
    """读取 EDC 导出 Excel 的指定 sheet。

    Args:
        form_oid: 表单 OID（即 sheet 名的核心部分）
        usecols: 只读取指定列（中文列名列表）
        form_name: 表单中文名（cmis 用于拼接 sheet 名；省略则自动从元数据查找）
        dtype: 列类型覆盖（传给 pd.read_excel 的 dtype 参数，用于保留前导零等）

    Returns:
        DataFrame
    """
    sheet = _resolve_sheet_name(form_oid, form_name)
    kwargs = {"header": 0, "usecols": usecols, "dtype": dtype}
    if EDC_TYPE != "clinflash":
        kwargs["skiprows"] = [1]
    return pd.read_excel(raw_path, sheet_name=sheet, **kwargs)
