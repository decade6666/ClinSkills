"""
loaders.py — 薄加载层（写死本试验 schema，不做通用化）

每个函数对应一个原始 Excel sheet 的标准化读取。
调用方按需取列、转 dtype、填空值。
"""

from pathlib import Path
import yaml
import pandas as pd

# 从 config.yaml 加载 raw_path（相对于 utils/ 所在目录的上一级）
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_CONFIG = _PROJECT_ROOT / "config.yaml"
with open(_CONFIG, "r", encoding="utf-8") as _f:
    _cfg = yaml.safe_load(_f)
raw_path = str(_PROJECT_ROOT / _cfg["path"]["raw_path"])


def load_sheet(sheet_name, cols, path=None):
    """读取原始 Excel 的单个 sheet，统一 header=0 / skiprows=[1] / dtype=str。

    Parameters
    ----------
    sheet_name : str
        Sheet 名称。
    cols : list[str]
        需要的列名。
    path : str, optional
        Excel 文件路径，默认使用 raw_path。

    Returns
    -------
    pd.DataFrame
    """
    if path is None:
        path = raw_path
    return pd.read_excel(
        path, sheet_name=sheet_name,
        header=0, skiprows=[1],
        usecols=cols, dtype=str,
    )


def load_rand(cols, path=None):
    """读取 DS_RAND（受试者随机化信息）。"""
    return load_sheet("DS_RAND", cols, path)


def load_completion(path=None):
    """读取 DS_END 并过滤为"试验完成情况总结"行。

    返回 DataFrame，含 ["受试者", "是否完成试验"] 两列。
    """
    df = load_sheet("DS_END", ["受试者", "页面名称", "是否完成试验_TXT"], path).fillna("")
    df = df[df["页面名称"] == "试验完成情况总结"].drop(columns="页面名称")
    if "受试者是否完成试验_TXT" in df.columns:
        df = df.rename(columns={"受试者是否完成试验_TXT": "是否完成试验"})
    else:
        df = df.rename(columns={"是否完成试验_TXT": "是否完成试验"})
    return df


def load_first_dose(path=None):
    """读取 EC sheet，返回每受试者最早服药日期。"""
    ec = load_sheet("EC", ["受试者", "服药日期"], path).fillna("")
    ec["服药日期"] = pd.to_datetime(ec["服药日期"], errors="coerce")
    ec = ec[ec["服药日期"].notna()]
    return (
        ec.groupby("受试者", dropna=False)["服药日期"]
        .min()
        .reset_index()
        .rename(columns={"服药日期": "首次用药日期"})
    )
