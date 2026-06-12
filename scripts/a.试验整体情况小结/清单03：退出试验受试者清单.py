# %%
# %run ../../env.py
from utils.loaders import load_rand

# %%
# 随机号 随机日期
RAND = load_rand(cols=['受试者', '受试者状态', '随机时间', '随机号'])

# 是否完成试验
cols = ["受试者", "页面名称","受试者是否完成试验_TXT"]
DS_END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = cols, dtype = str)

# %% [markdown]
# ## 清单： 退出试验受试者清单
# ####

# %%
# 退出日期，使用备注文件的手动记录的退出日期
# cols = ["受试者编号", "页面", "备注内容"]
# EXIT = pd.read_excel(remark_path, sheet_name = "备注清单", header = 4, usecols = cols)
# EXIT = EXIT[EXIT["页面"] == "治疗结束页"].drop(columns = "页面")
# EXIT[["日期类型", "日期"]] = EXIT["备注内容"].str.split("是", expand=True)
# EXIT["日期"] = pd.to_datetime(EXIT["日期"], format="%Y年%m月%d日")
# EXIT = EXIT.pivot(index = "受试者编号", columns = "日期类型", values = "日期").reset_index()
# EXIT = EXIT.rename(columns = {"受试者编号":"受试者"})
# EXIT['退出日期'] = pd.to_datetime(EXIT['退出日期'], errors='coerce')

# 研究开始日期：首次知情同意日期
cols = ["受试者", "知情同意书签署日期"]
DS = pd.read_excel(raw_path, sheet_name = "DS_ICF", header = 0, skiprows = [1], usecols = cols, dtype = str)
DS = DS.rename(columns={
    "知情同意书签署日期": "研究开始日期"
})

# 随机号 随机日期
subj = RAND[RAND["受试者状态"] == "中止退出"]

# 末次计划访视
cols = ["受试者", "访视OID", "访视名称", "访视日期"]
SV = pd.read_excel(raw_path, sheet_name = "SV", header = 0, skiprows = [1], usecols = cols, dtype = str)
SV = SV[(SV["访视OID"] != "V90") & (SV["访视OID"] != "V80")]
SV["访视日期"] = pd.to_datetime(SV["访视日期"], errors="coerce")

idx = SV.groupby("受试者")["访视日期"].idxmax()
SV = SV.loc[idx, ["受试者", "访视名称", "访视日期"]]

SV = SV.rename(columns={
    "访视名称": "末次已完成的计划内访视",
})

# 是否进行提前退出访视
cols = ["受试者", "访视OID", "访视日期"]
EXITYN = pd.read_excel(raw_path, sheet_name = "SV", header = 0, skiprows = [1], usecols = cols, dtype = str)
EXITYN = EXITYN[(EXITYN["访视OID"] == "V80")]
EXITYN["是否进行提前退出访视"] = '是'

# 试验完成日期
cols = ["受试者", "页面名称", "受试者退出试验原因_TXT", "试验完成日期", "提前退出日期"]
END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols =cols, dtype=str)
END["研究结束日期"] = np.where(END["试验完成日期"].notna(), END["试验完成日期"], END["提前退出日期"])

# 是否提前终止治疗
cols = ["受试者", "页面名称", "受试者是否永久终止试验干预_TXT"]
END1 = pd.read_excel(raw_path, sheet_name = "DS_INTED", header = 0, skiprows = [1], usecols = cols, dtype = str)

# 首末次用药日期
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
          .merge(SV, on = "受试者", how = "left")
          .merge(DS_END, on = "受试者", how = "left")
          .merge(END, on = "受试者", how = "left")
          .merge(END1, on = "受试者", how = "left")
          .merge(DS, on = "受试者", how = "left")
          .merge(EXITYN, on = "受试者", how = "left")
     )

df = df.rename(columns = {
    "受试者退出试验原因_TXT":"提前退出原因",
    "受试者":"筛选号",
    "受试者是否永久终止试验干预_TXT": "是否提前终止治疗"
})

# 计算试验时长（结束 - 开始 + 1 天）
df["研究结束日期"] = pd.to_datetime(df["研究结束日期"], errors="coerce").dt.normalize()
df["研究开始日期"] = pd.to_datetime(df["研究开始日期"], errors="coerce").dt.normalize()
df["试验时长（天）"] = (
    (df["研究结束日期"] - df["研究开始日期"]).dt.days + 1
)
df["研究结束日期"] = df["研究结束日期"].dt.strftime("%Y-%m-%d")
df["研究开始日期"] = df["研究开始日期"].dt.strftime("%Y-%m-%d")
# df["退出日期"] = df["退出日期"].dt.strftime("%Y-%m-%d").fillna('')

df["用药后安全性指标评估情况"] = "有/无"
df["用药后疗效性指标评估情况"] = "有/无"

stand_cols = [
    '筛选号',
    '随机号',
    '研究开始日期',
    '随机时间',
    '首次用药日期',
    '末次用药日期',
    '治疗天数（天）',
    '是否提前终止治疗',
    '研究结束日期',
    '试验时长（天）',
    '末次已完成的计划内访视',
    '是否进行提前退出访视',
    "用药后安全性指标评估情况",
    "用药后疗效性指标评估情况",
    '提前退出原因'
]
df = df[stand_cols]

n = len(df)
df.insert(0, "No.", range(1, len(df) + 1))

notes = [
    "提前退出：受试者未进行访视6（V6，D71±3）；",
    "治疗天数（天）=末次用药日期-首次用药日期+1；",
    "研究开始日期：最早一次知情同意书签署日期；",
    "研究结束日期：最晚一次访视完成日期；",
    "试验时长（天）=研究结束日期-研究开始日期+1。"
]

save_table_to_docx_threeline(
        df,
        f'{output_path}/table/表5 退出试验受试者清单（{n}例）.docx',
        '表5 退出试验受试者清单',
        notes,
        row_height_cm=0.6,
        auto_width=True
    )
df
