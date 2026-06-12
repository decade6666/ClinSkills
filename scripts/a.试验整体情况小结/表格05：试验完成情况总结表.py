# %run ../../env.py
from utils.loaders import load_rand

# =====================================================================
# 表2 试验完成情况总结表
# =====================================================================

# --- 步骤 1·读取 ---

df_end = pd.read_excel(
    raw_path, sheet_name="DS_END",
    header=0, skiprows=[1],
    usecols=["受试者", "页面名称", "受试者是否完成试验_TXT"],
    dtype=str,
)

df_rand = load_rand(cols=["中心编号", "研究中心", "受试者", "受试者是否随机入组_TXT"]).fillna("")

# --- 步骤 6·连接 ---

df = df_rand.merge(df_end, on="受试者", how="left")
df["研究中心"] = df["中心编号"] + "-" + df["研究中心"]

# --- 步骤 5·派生 — 按入组与完成状态分类 ---

def classify_subject_status(row):
    """根据是否随机入组、是否完成试验，判定受试者类别。"""
    if row["受试者是否随机入组_TXT"] == "否":
        return "筛选失败"
    elif row["受试者是否随机入组_TXT"] == "是":
        return "完成试验" if row["受试者是否完成试验_TXT"] == "是" else "退出试验"
    return "未知"

df["类别"] = df.apply(classify_subject_status, axis=1)
df = df.drop(columns=["中心编号", "受试者", "受试者是否随机入组_TXT", "受试者是否完成试验_TXT"])

# --- 步骤 4·变形 — 按中心 × 类别交叉计数，汇总为宽表 ---

ct = pd.crosstab(df["研究中心"], df["类别"])

n_fail     = ct.get("筛选失败", pd.Series(0, index=ct.index))
n_complete = ct.get("完成试验", pd.Series(0, index=ct.index))
n_dropout  = ct.get("退出试验", pd.Series(0, index=ct.index))

summary = pd.DataFrame({
    "筛选总人数":   ct.sum(axis=1),
    "筛选失败人数": n_fail,
    "随机入组人数": n_complete + n_dropout,
    "退出试验人数": n_dropout,
    "完成试验人数": n_complete,
}, index=ct.index)

summary["筛败率"] = summary["筛选失败人数"] / summary["筛选总人数"] * 100
summary["入组率"] = summary["随机入组人数"] / summary["筛选总人数"] * 100
summary["脱落率"] = summary["退出试验人数"] / summary["随机入组人数"] * 100
summary["脱落率"] = summary["脱落率"].replace([np.inf, -np.inf], np.nan)

# 追加合计行
total = summary[["筛选总人数", "筛选失败人数", "随机入组人数",
                 "退出试验人数", "完成试验人数"]].sum()

total["筛败率"] = total["筛选失败人数"] / total["筛选总人数"] * 100
total["入组率"] = total["随机入组人数"] / total["筛选总人数"] * 100
total["脱落率"] = (
    total["退出试验人数"] / total["随机入组人数"] * 100
    if total["随机入组人数"] != 0 else np.nan
)

summary.loc["合计"] = total

# --- 步骤 7·格式化 ---

count_cols = ["筛选总人数", "筛选失败人数", "随机入组人数", "退出试验人数", "完成试验人数"]
summary[count_cols] = summary[count_cols].fillna(0).astype(int)

for col in ["筛败率", "入组率", "脱落率"]:
    summary[col] = summary[col].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "")

summary = summary.reset_index().rename(columns={"index": "研究中心"})
summary = summary[["研究中心", "筛选总人数", "筛选失败人数", "筛败率",
                    "随机入组人数", "入组率", "退出试验人数", "脱落率", "完成试验人数"]]

# --- 步骤 8·输出 ---

notes = [
    "筛败率%=筛选失败人数/筛选总人数*100%",
    "入组率%=随机入组人数/筛选总人数*100%",
    "脱落率%=退出试验人数/随机入组人数*100%",
]

save_table_to_docx_threeline(
    summary,
    f'{output_path}/table/表2 试验完成情况总结表.docx',
    '表2 试验完成情况总结表',
    notes,
    row_height_cm=0.6,
    auto_width=True,
)
summary
