# %%
# %run env.py

subj = "受试者"

# %%
# 随机号 随机日期
cols = ["受试者", "受试者状态", "随机时间", "随机号"]
RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = cols, dtype = str)

# 是否完成试验
cols = ["受试者", "页面名称","受试者是否完成试验_TXT"]
DS_END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = cols, dtype = str)

# %% [markdown]
# ## 表格： 首末例受试者情况
# #### 1.计算最晚访视日期的时候是否考虑计划外访视？
# ####

# %%
# 研究开始日期：最早一次知情同意书签署日期
subj = "受试者"
col1 = "知情同意书签署日期"
col2 = "知情同意书签署时间"

start = pd.read_excel(
    raw_path,
    sheet_name="DS_ICF",
    header=0,
    dtype=str,
    skiprows=[1],
    usecols=[subj, col1, col2]
)

start["签署日期时间"] = pd.to_datetime(
    start[col1].astype(str).str.strip() + " " +
    start[col2].fillna("00:00").astype(str).str.strip(),
    errors="coerce"
)

m = start["签署日期时间"].min()

first_case = start.loc[
    start["签署日期时间"] == m,
    [subj]
].copy()

first_case["首末例"] = "首例"
first_case

# 研究结束日期：最晚一次访视完成日期；
col1 = "试验完成日期"
col2 = "提前退出日期"

end = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = [subj, col1, col2], dtype=str)
end["研究结束日期"] = np.where(end[col1].notna(), end[col1], end[col2])
dt = pd.to_datetime(end["研究结束日期"], errors="coerce")
m = dt.max()
last_case = end.loc[dt == m, ["受试者"]]
last_case["首末例"] = "末例"

df = pd.concat([first_case, last_case])
df = (df.merge(start, on = ["受试者"], how = "left")
        .merge(end, on = ["受试者"], how = "left")
        .merge(RAND, on = ["受试者"], how = "left")
        .merge(DS_END, on = ["受试者"], how = "left")
     )

df = df.rename(columns={
    "受试者":"筛选号",
    "知情同意书签署日期":"研究开始日期",
    "受试者是否完成试验_TXT":"是否完成试验",
    "首末例":"受试者",
})


df_first = df[df['受试者'] == '首例']
df_first = df_first.loc[df_first['研究结束日期'].idxmax(),:].to_frame().T
df_last = df[df['受试者'] == '末例']
df_last = df_last.loc[df_last['研究开始日期'].idxmax(),:].to_frame().T

df = pd.concat([df_first, df_last])

df["研究结束日期"] = pd.to_datetime(df["研究结束日期"], errors="coerce")
df["研究开始日期"] = pd.to_datetime(df["研究开始日期"], errors="coerce")
df["试验时长（天）"] = (df["研究结束日期"] - df["研究开始日期"]).dt.days + 1
df["研究开始日期"] = df["研究开始日期"].dt.strftime("%Y-%m-%d")
df["研究结束日期"] = df["研究结束日期"].dt.strftime("%Y-%m-%d")

stand_cols = ["受试者" ,"筛选号" ,"随机号" ,"研究开始日期" ,"随机时间" ,"研究结束日期" ,"试验时长（天）" ,"是否完成试验"]
df = df[stand_cols]
print(df)
notes = [
    "首例病例为入组受试者中第一例签署知情同意书的受试者；末例病例为入组受试者中最后结束研究的受试者；",
    "研究开始日期：最早一次知情同意书签署日期；",
    "研究结束日期：最晚一次访视完成日期；",
    "试验时长（天）=研究结束日期-研究开始日期+1。"
]

save_table_to_docx_threeline(
        df, 
        f'{output_path}/table/表1 首末例受试者情况.docx', 
        '表1 首末例受试者情况', 
        notes,
        row_height_cm=0.6,
        auto_width=True
    )

# %% [markdown]
# ## 表格： 试验完成情况总结表
# ####

# %%
cols = ["中心编号", "研究中心", "受试者","受试者是否随机入组_TXT"]
subj = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")

df = subj.merge(DS_END, on = "受试者", how = "left")
df["研究中心"] = df["中心编号"] + "-" + df["研究中心"]

# 定义类别列
def assign_category(row):
    if row["受试者是否随机入组_TXT"] == "否":
        return "筛选失败"
    elif row["受试者是否随机入组_TXT"] == "是":
        if row["受试者是否完成试验_TXT"] == "是":
            return "完成试验"
        else:
            return "退出试验"
    return "未知"

# 使用 apply 来应用该逻辑
df["类别"] = df.apply(assign_category, axis=1)
df = df.drop(columns=["中心编号", "受试者", "受试者是否随机入组_TXT", "受试者是否完成试验_TXT"])
ct = pd.crosstab(df["研究中心"], df["类别"])

# 防止某些类别在某中心不存在，先安全取一下
fail = ct.get("筛选失败", pd.Series(0, index=ct.index))
complete = ct.get("完成试验", pd.Series(0, index=ct.index))
dropout = ct.get("退出试验", pd.Series(0, index=ct.index))

# 按规则生成汇总表
summary = pd.DataFrame(index=ct.index)
summary["筛选总人数"]   = ct.sum(axis=1)
summary["筛选失败人数"] = fail
summary["随机入组人数"] = complete + dropout
summary["退出试验人数"] = dropout
summary["完成试验人数"] = complete

# 计算比例（先用数值，方便你后面有需要再计算）
summary["筛败率"] = summary["筛选失败人数"] / summary["筛选总人数"] * 100
summary["入组率"] = summary["随机入组人数"] / summary["筛选总人数"] * 100
summary["脱落率"] = summary["退出试验人数"] / summary["随机入组人数"] * 100

# 避免除以 0 的情况
summary["脱落率"] = summary["脱落率"].replace([np.inf, -np.inf], np.nan)

# 加一行 “合计”
tot_counts = summary[["筛选总人数","筛选失败人数","随机入组人数",
                      "退出试验人数","完成试验人数"]].sum()

total_row = pd.Series({
    "筛选总人数":   tot_counts["筛选总人数"],
    "筛选失败人数": tot_counts["筛选失败人数"],
    "随机入组人数": tot_counts["随机入组人数"],
    "退出试验人数": tot_counts["退出试验人数"],
    "完成试验人数": tot_counts["完成试验人数"],
})

total_row["筛败率"] = total_row["筛选失败人数"] / total_row["筛选总人数"] * 100
total_row["入组率"] = total_row["随机入组人数"] / total_row["筛选总人数"] * 100
total_row["脱落率"] = (
    total_row["退出试验人数"] / total_row["随机入组人数"] * 100
    if total_row["随机入组人数"] != 0 else np.nan
)

summary.loc["合计"] = total_row

int_cols = [
    "筛选总人数",
    "筛选失败人数",
    "随机入组人数",
    "退出试验人数",
    "完成试验人数",
]
summary[int_cols] = summary[int_cols].fillna(0).astype(int)

# 格式化成XX.XX% 的样式
for col in ["筛败率", "入组率", "脱落率"]:
    summary[col] = summary[col].apply(
        lambda x: f"{x:.2f}%" if pd.notna(x) else ""
    )

# 把研究中心变成一列，列顺序按图里来
summary = summary.reset_index().rename(columns={"index": "研究中心"})
summary = summary[[
    "研究中心",
    "筛选总人数",
    "筛选失败人数",
    "筛败率",
    "随机入组人数",
    "入组率",
    "退出试验人数",
    "脱落率",
    "完成试验人数",
]]

notes = [
    "筛败率%=筛选失败人数/筛选总人数*100%",
    "入组率%=随机入组人数/筛选总人数*100%",
    "脱落率%=退出试验人数/随机入组人数*100%"
]

save_table_to_docx_threeline(
        summary, 
        f'{output_path}/table/表2 试验完成情况总结表.docx', 
        '表2 试验完成情况总结表', 
        notes,
        row_height_cm=0.6,
        auto_width=True
    )
summary

# %% [markdown]
# ## 清单： 完成试验受试者

# %%
# 随机号 随机日期
cols = ["受试者", "受试者状态","随机时间", "随机号"]
subj = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
subj = subj[subj["受试者状态"] == "完成试验"]

# 最晚一次访视完成日期
cols = ["受试者", "试验完成日期","提前退出日期"]
END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = cols, dtype=str)
END["研究完成日期"] = np.where(end[col1].notna(), end["试验完成日期"], end["提前退出日期"])

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

# %% [markdown]
# ## 表格： 筛选失败原因分类

# %%
usecols = ["受试者是否随机入组_TXT", "不符合入选标准", "符合排除标准", "撤回知情同意", "失访，尝试联系≥3次均未成功", "其他"]
subj = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = usecols, dtype = str)
subj = subj[subj["受试者是否随机入组_TXT"] == "否"]
subj = subj.drop(columns = ["受试者是否随机入组_TXT"])

# 需要统计的筛选失败原因
reasons = ["不符合入选标准", "符合排除标准", "撤回知情同意", "失访，尝试联系≥3次均未成功", "其他"]

# 按原因统计“例次”：统计该列非空的行数即可
df = pd.DataFrame({
    "筛选失败原因": reasons,
    "例次": [(subj[col] == "1").sum() for col in reasons]
})

# 加一行“合计”
total = df["例次"].sum()
df.loc[len(df)] = ["合计", total]

# 人数/例次确保为整数
df["例次"] = df["例次"].astype(int)

notes = [
    "根据筛选失败原因，拆分信息按例次计算。"
]
save_table_to_docx_threeline(
        df, 
        f'{output_path}/table/表3 筛选失败原因分类.docx', 
        '表3 筛选失败原因分类', 
        notes,
        row_height_cm=0.6,
        auto_width=True
    )
df

# %% [markdown]
# ## 清单： 筛选失败受试者

# %%
cols1 = [ "不符合入选标准", "符合排除标准", "撤回知情同意", "失访，尝试联系≥3次均未成功", "其他筛选失败原因"]
index = ["受试者", "受试者状态", "受试者是否随机入组_TXT",]
subj = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = index + cols1, dtype = str)
subj = subj.melt(id_vars=index, value_vars=cols1 , var_name='筛选失败原因', value_name='结果')
subj= subj[(subj["受试者状态"] == "筛选失败") & (subj["结果"] == '1')]

cols2 = ["受试者", "知情同意书签署日期"]
DS = pd.read_excel(raw_path, sheet_name = "DS_ICF", header = 0, skiprows = [1], usecols = cols2, dtype = str)

df = subj.merge(DS, on = "受试者", how = "left")
df = df.rename(columns = {
    "受试者":"筛选号"
})
df = df[["筛选号", "知情同意书签署日期", "筛选失败原因"]]
df.insert(0, "No.", range(1, len(df) + 1))
n = len(df)

export_to_excel_with_format(
    df, 
    f"{output_path}/listing/表32 筛选失败受试者清单.xlsx", 
    "表32 筛选失败受试者清单", 
    f"表32 筛选失败受试者清单（{n}例）"
)
df

# %% [markdown]
# ## 清单： 入组未用药的受试者
# #### 1. 当前数据集中入组未用药的受试者为0例，后续有数据后再补充代码？
# ####

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

# %% [markdown]
# ## 表格： 因与试验用药品相关的不良事件而发生的提前终止治疗

# %%
cols = ["受试者", "页面名称", "永久终止试验干预原因_TXT"]
DS_END1 = pd.read_excel(raw_path, sheet_name = "DS_INTED", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
DS_END1 = DS_END1[(DS_END1["永久终止试验干预原因_TXT"] == "试验期间受试者发生不良事件，研究者认为受试者需永久停止服用试验用药品")]
DS_END1 = DS_END1.drop(columns = ["页面名称"], axis = 1)

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

EC["首次用药日期"] = pd.to_datetime(EC["首次用药日期"], errors="coerce")
EC["末次用药日期"] = pd.to_datetime(EC["末次用药日期"], errors="coerce")

EC["治疗天数（天）"] = (EC["末次用药日期"] - EC["首次用药日期"]).dt.days + 1
EC["治疗天数（天）"] = (EC["治疗天数（天）"].astype("Int64").astype("string").fillna(""))

EC["首次用药日期"] = EC["首次用药日期"].dt.strftime("%Y-%m-%d")
EC["末次用药日期"] = EC["末次用药日期"].dt.strftime("%Y-%m-%d")

cols = ["受试者", "页面名称", "受试者是否完成试验_TXT"]
DS_END2 = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
DS_END2 = DS_END2.drop(columns = ["页面名称"], axis = 1)

cols = ["受试者", "不良事件名称", "对试验药物采取的措施-1_TXT", "对试验药物采取的措施-2_TXT", "对试验药物采取的措施-3_TXT", 
                                  "对试验药物采取的措施-4_TXT", "对试验药物采取的措施-5_TXT", "对试验药物采取的措施-6_TXT", "与试验药物的关系_TXT", "导致死亡"]
AE = pd.read_excel(raw_path, sheet_name = "AE", header = 0, skiprows = [1], usecols = cols, dtype = str)
AE["对试验药物采取的措施"] = AE[measure_cols].apply(
    lambda row: ",".join(
        [
            x.strip()
            for x in row
            if pd.notna(x) and str(x).strip() not in ["", "nan", "NaN"]
        ]
    ),
    axis=1
)
AE = AE[(AE["对试验药物采取的措施"] == "永久停药") | (AE["导致死亡"] == "Y")]

df = (DS_END1.merge(RAND, on = "受试者", how = "left")
        .merge(EC, on = "受试者", how = "left")
        .merge(DS_END2, on = "受试者", how = "left")
        .merge(AE, on = "受试者", how = "left")
     )

df = df.rename(columns = {
    "受试者":"筛选号",
    "研究者决定原因_TXT":"提前终止治疗的原因",
    "是否完成试验_TXT":"是否完成试验",
    "与试验药物的关系_TXT":"与试验用药品的关系",
})

df = df[[
 '筛选号',
 '随机号',
 '首次用药日期',
 '末次用药日期',
 '治疗天数（天）',
 '不良事件名称',
 '与试验用药品的关系',
 '提前终止治疗的原因',
 '是否完成试验'
]]

df.insert(0, "No.", range(1, len(df) + 1))
df

notes = [
    "治疗天数（天）= 试验药物末次用药日期 - 试验药物首次用药日期 + 1；"
]

save_table_to_docx_threeline(
        df, 
        f'{output_path}/table/表9 因与试验用药品相关的不良事件而发生的提前终止治疗.docx', 
        '表9 因与试验用药品相关的不良事件而发生的提前终止治疗', 
        notes,
        row_height_cm=0.6,
        auto_width=True
    )

# %%
cols = ["受试者", "不良事件名称", "对试验药物采取的措施-1_TXT", "对试验药物采取的措施-2_TXT", "对试验药物采取的措施-3_TXT", 
                                  "对试验药物采取的措施-4_TXT", "对试验药物采取的措施-5_TXT", "对试验药物采取的措施-6_TXT", "与试验药物的关系_TXT"]
measure_cols = [
    "对试验药物采取的措施-1_TXT",
    "对试验药物采取的措施-2_TXT",
    "对试验药物采取的措施-3_TXT",
    "对试验药物采取的措施-4_TXT",
    "对试验药物采取的措施-5_TXT",
    "对试验药物采取的措施-6_TXT"
]
AE = pd.read_excel(raw_path, sheet_name = "AE", header = 0, skiprows = [1], usecols = cols, dtype = str)
AE["对试验药物采取的措施"] = AE[measure_cols].apply(
    lambda row: ",".join(
        [
            x.strip()
            for x in row
            if pd.notna(x) and str(x).strip() not in ["", "nan", "NaN"]
        ]
    ),
    axis=1
)
AE = AE[AE["对试验药物采取的措施"] != ""]
AE

# %% [markdown]
# ## 表格： 因其他原因而发生的提前终止治疗

# %%
cols = ["受试者", "页面名称", "其他原因"]
DS_END1 = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
DS_END1 = DS_END1[(DS_END1["页面名称"] == "治疗结束页") & (DS_END1["其他原因"] != "")]  #如果需要考虑研究者判断，这里需要补充条件
DS_END1 = DS_END1.drop(columns = ["页面名称"], axis = 1)

cols = ["受试者", "服药日期"]
EC = pd.read_excel(raw_path, sheet_name = "EC", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
EC["服药日期"] = pd.to_datetime(EC["服药日期"], errors="coerce")
EC = EC[EC["服药日期"].notna()]
EC = (
    EC.groupby("受试者", dropna=False)["服药日期"]
      .agg(["min", "max"])
      .rename(columns={"min": "首次用药日期", "max": "末次用药日期"})
)

EC["首次用药日期"] = pd.to_datetime(EC["首次用药日期"], errors="coerce")
EC["末次用药日期"] = pd.to_datetime(EC["末次用药日期"], errors="coerce")

EC["治疗天数（天）"] = (EC["末次用药日期"] - EC["首次用药日期"]).dt.days + 1
EC["治疗天数（天）"] = (EC["治疗天数（天）"].astype("Int64").astype("string").fillna(""))

EC["首次用药日期"] = EC["首次用药日期"].dt.strftime("%Y-%m-%d")
EC["末次用药日期"] = EC["末次用药日期"].dt.strftime("%Y-%m-%d")

cols = ["受试者", "页面名称", "是否完成试验_TXT"]
DS_END2 = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
DS_END2 = DS_END2[DS_END2["页面名称"] == "试验完成情况总结"]
DS_END2 = DS_END2.drop(columns = ["页面名称"], axis = 1)

cols = ["受试者", "不良事件名称", "对试验药物采取的措施_TXT", "与试验药物的关系_TXT", "导致死亡"]
AE = pd.read_excel(raw_path, sheet_name = "AE", header = 0, skiprows = [1], usecols = cols, dtype = str)
AE = AE[(AE["对试验药物采取的措施_TXT"] == "永久停药") | (AE["导致死亡"] == "Y")]
AE
df = (DS_END1.merge(RAND, on = "受试者", how = "left")
             .merge(EC, on = "受试者", how = "left")
             .merge(DS_END2, on = "受试者", how = "left")
             .merge(AE, on = "受试者", how = "left")
     )

df = df.rename(columns = {
    "受试者":"筛选号",
    "其他原因":"提前终止治疗的原因",
    "是否完成试验_TXT":"是否完成试验",
    "与试验药物的关系_TXT":"与试验用药品的关系",
})

df = df[[
 '筛选号',
 '随机号',
 '首次用药日期',
 '末次用药日期',
 '治疗天数（天）',
 '不良事件名称',
 '与试验用药品的关系',
 '提前终止治疗的原因',
 '是否完成试验'
]]

df.insert(0, "No.", range(1, len(df) + 1))
df

notes = [
    "治疗天数（天）= 试验药物末次用药日期 - 试验药物首次用药日期 + 1；"
]

save_table_to_docx_threeline(
        df, 
        f'{output_path}/table/表10 因其他原因而发生的提前终止治疗.docx', 
        '表10 因其他原因而发生的提前终止治疗', 
        notes,
        row_height_cm=0.6,
        auto_width=True
    )
