"""
loaders.py — 薄加载层（写死本试验 schema，不做通用化）

每个函数对应一个原始 Excel sheet 的标准化读取。
调用方按需取列、转 dtype、填空值。
"""

import pandas as pd


def load_rand(cols, path=None):
    """读取 DS_RAND（受试者随机化信息）。

    Parameters
    ----------
    cols : list[str]
        需要的列名，如 ["受试者", "随机号"]。
    path : str, optional
        Excel 文件路径，默认使用 env.py 中的 raw_path。

    Returns
    -------
    pd.DataFrame
    """
    if path is None:
        path = raw_path  # noqa: F821 — 由 env.py 注入全局
    return pd.read_excel(
        path, sheet_name="DS_RAND",
        header=0, skiprows=[1],
        usecols=cols,
    )


def load_completion(path=None):
    """读取 DS_END 并过滤为"试验完成情况总结"行。

    返回 DataFrame，含 ["受试者", "是否完成试验"] 两列。
    原始列名 "受试者是否完成试验_TXT" 或 "是否完成试验_TXT" 统一映射为 "是否完成试验"。
    NaN 填充为 ""。
    """
    if path is None:
        path = raw_path  # noqa: F821
    df = pd.read_excel(
        path, sheet_name="DS_END",
        header=0, skiprows=[1],
        usecols=["受试者", "页面名称", "是否完成试验_TXT"],
        dtype=str,
    ).fillna("")
    df = df[df["页面名称"] == "试验完成情况总结"].drop(columns="页面名称")
    # 兼容两种原始列名
    if "受试者是否完成试验_TXT" in df.columns:
        df = df.rename(columns={"受试者是否完成试验_TXT": "是否完成试验"})
    else:
        df = df.rename(columns={"是否完成试验_TXT": "是否完成试验"})
    return df


def load_first_dose(path=None):
    """读取 EC sheet，返回每受试者最早服药日期。

    Returns DataFrame with ["受试者", "首次用药日期"]。
    """
    if path is None:
        path = raw_path  # noqa: F821
    ec = pd.read_excel(
        path, sheet_name="EC",
        header=0, skiprows=[1],
        usecols=["受试者", "服药日期"],
        dtype=str,
    ).fillna("")
    ec["服药日期"] = pd.to_datetime(ec["服药日期"], errors="coerce")
    ec = ec[ec["服药日期"].notna()]
    return (
        ec.groupby("受试者", dropna=False)["服药日期"]
        .min()
        .reset_index()
        .rename(columns={"服药日期": "首次用药日期"})
    )
