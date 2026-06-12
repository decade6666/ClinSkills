# %%
# %run ../../env.py
from utils.loaders import load_rand

# %% [markdown]
# ## 清单： 完成试验受试者

# %%
# 随机号 随机日期
subj = load_rand(cols=['受试者', '受试者状态', '随机时间', '随机号'])
subj = subj[subj["受试者状态"] == "完成试验"]

# 最晚一次访视完成日期
cols = ["受试者", "试验完成日期","提前退出日期"]
END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = cols, dtype=str)
# 注：原脚本此处误用了上一张表残留的全局变量 end/col1，本块的 END 为同样的读取，改为引用 END（行为一致）
END["研究完成日期"] = np.where(END["试验完成日期"].notna(), END["试验完成日期"], END["提前退出日期"])

# 查找最晚访视日期来计算试验时长
cols = ["受试者", "访视OID", "访视日期"]
SV = pd.read_excel(raw_path, sheet_name = "SV", header = 0, skiprows = [1], usecols = cols, dtype=str)
SV = SV[(SV["访视OID"] != "V90") & (SV["访视OID"] != "V80")]

SV["访视日期"] = pd.to_datetime(SV["访视日期"], errors="coerce")
idx = SV.groupby("受试者")["访视日期"].idxmax()
SV = SV.loc[idx, ["受试者", "访视日期"]]

SV["访视日期"] = SV["访视日期"].dt.strftime("%Y-%m-%d")

cols = ["受试者", "知情同意书签署日期"]
DS = pd.read_excel(raw_path, sheet_name = "DS_ICF", header = 0, skiprows = [1], usecols = cols, dtype=str)

# SV = SV.merge(DS, on = "受试者", how = "left")
# SV["访视日期"] = pd.to_datetime(SV["访视日期"], errors="coerce")
# SV["知情同意书签署日期"] = pd.to_datetime(SV["知情同意书签署日期"], errors="coerce")
# SV["试验时长（天）"] = (SV["访视日期"] - SV["知情同意书签署日期"]).dt.days + 1

# 通过计算首末次服药日期来计算治疗天数
cols = ["受试者", "开始日期", "结束日期"]
EC = pd.read_excel(raw_path, sheet_name = "EC_ED", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
EC["开始日期"] = pd.to_datetime(EC["开始日期"], errors="coerce")
EC["结束日期"] = pd.to_datetime(EC["结束日期"], errors="coerce")

EC1 = (
    EC.groupby("受试者", dropna=False)["开始日期"]
      .agg(["min"])
      .rename(columns={"min": "首次用药日期"})
      )

EC2 = (
    EC.groupby("受试者", dropna=False)["结束日期"]
      .agg(["max"])
      .rename(columns={"max": "末次用药日期"})
      )

EC = EC1.merge(EC2, on = "受试者", how = "inner")

# 计算治疗天数（末次 - 首次 + 1 天）
EC["治疗天数（天）"] = (
    (EC["末次用药日期"] - EC["首次用药日期"]).dt.days + 1
)

# 如果某些受试者只有1条记录或日期缺失，避免 NaN
EC["治疗天数（天）"] = EC["治疗天数（天）"].where(
    EC["治疗天数（天）"] > 0, np.nan
)

EC = EC.reset_index()

df = (subj.merge(EC, on = "受试者", how = "left")
          .merge(DS, on = "受试者", how = "left")
          .merge(END, on = "受试者", how = "left")
     )

df["研究完成日期"] = pd.to_datetime(df["研究完成日期"], errors="coerce")
df["知情同意书签署日期"] = pd.to_datetime(df["知情同意书签署日期"], errors="coerce")
df["试验时长（天）"] = (df["研究完成日期"] - df["知情同意书签署日期"]).dt.days + 1

df = df.rename(columns = {
    "提前退出原因_TXT":"提前退出原因",
    "受试者":"筛选号",
    "知情同意书签署日期":"首次知情同意书签署日期"
})

df["首次用药日期"] = df["首次用药日期"].dt.strftime("%Y-%m-%d")
df["末次用药日期"] = df["末次用药日期"].dt.strftime("%Y-%m-%d")
df["研究完成日期"] = df["研究完成日期"].dt.strftime("%Y-%m-%d")
df["首次知情同意书签署日期"] = df["首次知情同意书签署日期"].dt.strftime("%Y-%m-%d")

stand_cols = [
    '筛选号',
    '随机号',
    '首次知情同意书签署日期',
    '随机时间',
    '首次用药日期',
    '末次用药日期',
    '治疗天数（天）',
    '试验完成日期',
    '试验时长（天）']

df = df[stand_cols]


n = len(df)
df.insert(0, "No.", range(1, len(df) + 1))

export_to_excel_with_format(
    df,
    f"{output_path}/listing/表34 完成试验受试者清单.xlsx",
    "表34 完成试验受试者清单",
    f"表34 完成试验受试者清单（{n}例）"
)
df
