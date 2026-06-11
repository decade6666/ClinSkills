# %%
# %run env.py

# %%
RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = ["受试者", "随机号"])

DS_END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = ["受试者", "页面名称", "是否完成试验_TXT"], dtype = str).fillna("")
DS_END = DS_END[DS_END["页面名称"] == "试验完成情况总结"].drop(columns = "页面名称")

# %%
index = ["受试者", "记录号"]

# %% [markdown]
# # 清单：不良事件

# %% [markdown]
# ## 不良事件预处理
# - 这里只列出了3个AE严重程度变化，实际设计了6组收集字段，但是后面均无数据

# %%
cols1 = [
    "受试者", 
    "不良事件名称", 
    "发生日期",
    "转归日期",
    "试验结束时，转归_TXT",
    "初始严重程度_TXT", 
    "严重程度是否有变化_TXT", 
    "严重程度变化日期-1", 
    "严重程度-1_TXT",
    "严重程度变化日期-2", 
    "严重程度-2_TXT",
    "严重程度变化日期-3", 
    "严重程度-3_TXT",
    "对试验药物采取的措施_TXT",
    "与试验药物的关系_TXT",
    "是否符合严重不良事件定义_TXT",
    "严重不良事件开始日期",
    "是否因此不良事件退出试验_TXT",
    "是否为特别关注不良事件_TXT",
    "PT术语",
    "SOC术语",
]

cols2 = ["未采取措施", "药物治疗", "非药物治疗"]
cols4 = ["其他，请描述"]
cols3 = ["导致死亡", "危及生命", "导致住院或延长住院时间", "永久或严重的残疾或者功能丧失", "先天性异常或者出生缺陷", "其他重要的医学事件"]

AE = pd.read_excel(raw_path, sheet_name = "AE", header = 0, skiprows = [1], usecols = cols1 + cols2 + cols3 + cols4)
AE = AE[AE["不良事件名称"].notna()]
m = AE["其他，请描述"].notna() & AE["其他，请描述"].astype(str).str.strip().ne("")
AE.loc[m, "其他，请描述"] = "其他:" + AE.loc[m, "其他，请描述"].astype(str)

# 将字段下的“√”变成具体的值
for index, row in AE.iterrows():
    for col in cols2 + cols3:
        if pd.notna(row[col]):
            AE.at[index, col] = col

# 拼接对试验药物采取的措施、严重不良事件定义等多选字段
AE["是否采取治疗措施"] = AE[cols2 + cols4].apply(lambda row: ";".join(row.dropna()), axis=1)
AE["严重不良事件定义"] = AE[cols3].apply(lambda row: ";".join(row.dropna()), axis=1)
AE = AE.drop(columns = cols2 + cols3 + cols4)

AE = (AE.merge(RAND, on = "受试者", how = "left").merge(DS_END, on = "受试者", how = "left"))

AE.columns = [col.replace("_TXT", "") for col in AE.columns]
AE = AE.rename(columns = {
    "受试者":"筛选号",
    "PT术语":"PT",
    "SOC术语":"SOC",
})

stand_cols = [
    "筛选号",
    "随机号",
    "不良事件名称", 
    "SOC",
    "PT",
    "发生日期",
    "转归日期",
    "试验结束时，转归",
    "初始严重程度", 
    "严重程度是否有变化", 
    "严重程度变化日期-1", 
    "严重程度-1",
    "严重程度变化日期-2", 
    "严重程度-2",
    "严重程度变化日期-3", 
    "严重程度-3",
    "是否采取治疗措施",
    "对试验药物采取的措施",
    "与试验药物的关系",
    "是否符合严重不良事件定义",
    "严重不良事件定义",
    "严重不良事件开始日期",
    "是否因此不良事件退出试验",
    "是否为特别关注不良事件",
    "是否完成试验"
     ]
AE = AE[stand_cols]
AE.insert(0, "No.", range(1, len(AE) + 1))

# %% [markdown]
# ## 清单：与试验药物有关的不良事件（XXX例次 XXX例）

# %%
AE_related = AE[
(AE["与试验药物的关系"] == "肯定有关") | 
(AE["与试验药物的关系"] == "很可能有关") | 
(AE["与试验药物的关系"] == "可能有关")
]

lc = len(AE_related)
ls = len(AE_related.drop_duplicates(subset = "筛选号"))

export_to_excel_with_format(
    AE_related, 
    f"{output_path}/listing/表40 与试验药物有关的不良事件清单.xlsx", 
    "表40 与试验药物有关的不良事件清单", 
    f"表40 与试验药物有关的不良事件清单（{lc}例次{ls}例）"
)

# %% [markdown]
# ## 清单：与试验药物无关的不良事件（XXX例次 XXX例）

# %%
AE_unrelated = AE[
(AE["与试验药物的关系"] == "可能无关") | 
(AE["与试验药物的关系"] == "无关")
]

lc = len(AE_unrelated)
ls = len(AE_unrelated.drop_duplicates(subset = "筛选号"))

export_to_excel_with_format(
    AE_unrelated, 
    f"{output_path}/listing/表41 与试验药物无关的不良事件清单.xlsx", 
    "表41 与试验药物无关的不良事件清单", 
    f"表41 与试验药物无关的不良事件清单（{lc}例次{ls}例）"
)

# %% [markdown]
# ##  清单：特别关注不良事件（XXX例次 XXX例）

# %%
AE_special = AE[(AE["是否为特别关注不良事件"] == "是")]

lc = len(AE_special)
ls = len(AE_special.drop_duplicates(subset = "筛选号"))

export_to_excel_with_format(
    AE_special, 
    f"{output_path}/listing/表42 特别关注不良事件清单.xlsx", 
    "表42 特别关注不良事件清单", 
    f"表42 特别关注不良事件清单（{lc}例次{ls}例）"
)

# %% [markdown]
# # 清单：人口学资料

# %%
cols = ["受试者", "性别_TXT", "出生日期", "年龄", "民族_TXT", "其他民族"]
DM = pd.read_excel(raw_path, sheet_name = "DM", header = 0, skiprows = [1], usecols = cols)

cols = ["受试者", "是否为育龄期女性_TXT", "页面名称"]
RP = pd.read_excel(raw_path, sheet_name = "RP", header = 0, skiprows = [1], usecols = cols)
RP = RP[RP["页面名称"] == "人口学资料"]

cols = ["受试者", "生育情况_TXT", "婚姻情况_TXT", "页面名称"]
SC = pd.read_excel(raw_path, sheet_name = "SC", header = 0, skiprows = [1], usecols = cols)
SC = SC[SC["页面名称"] == "人口学资料"]

DM = (DM.merge(RP, on = "受试者", how = "left").merge(SC, on = "受试者", how = "left"))
DM = (DM.merge(RAND, on = "受试者", how = "left"))

DM.columns = [col.replace("_TXT", "") for col in DM.columns]

DM["民族"] = np.where(
    DM["其他民族"].notna(),
    DM["其他民族"],
    DM["民族"],
)
DM = DM.rename(columns = {
    "受试者":"筛选号",
    "年龄":"年龄（岁）",
})

stand_cols = ["筛选号", "随机号", "性别", "出生日期", "年龄（岁）", "民族", "婚姻情况", "生育情况", "是否为育龄期女性"]
DM = DM[stand_cols]
DM.insert(0, "No.", range(1, len(DM) + 1))

lc = len(DM)

export_to_excel_with_format(
    DM, 
    f"{output_path}/listing/表43 人口学资料清单.xlsx", 
    "表43 人口学资料清单", 
    f"表43 人口学资料清单（{lc}例）"
)

# %% [markdown]
# # 清单：个人生活史

# %%
cols = ["受试者", "既往月经是否规律_TXT", "末次月经开始时间", "页面名称"]
RP = pd.read_excel(raw_path, sheet_name = "RP", header = 0, skiprows = [1], usecols = cols)
RP = RP[RP["页面名称"] == "个人生活史"]

cols = ["受试者", "是否吸烟_TXT", "是否饮酒_TXT"]
SU = pd.read_excel(raw_path, sheet_name = "SU", header = 0, skiprows = [1], usecols = cols)

cols = ["受试者", "性别_TXT"]
DM = pd.read_excel(raw_path, sheet_name = "DM", header = 0, skiprows = [1], usecols = cols)

SU = (SU.merge(RP, on = "受试者", how = "left")
        .merge(RAND, on = "受试者", how = "left")
        .merge(DM, on = "受试者", how = "left")
     )

SU.columns = [col.replace("_TXT", "") for col in SU.columns]

SU = SU.rename(columns = {
    "受试者":"筛选号",
})

stand_cols = ["筛选号", "随机号", "性别", "是否吸烟", "是否饮酒", "既往月经是否规律", "末次月经开始时间"]
SU = SU[stand_cols]
SU.insert(0, "No.", range(1, len(SU) + 1))

lc = len(SU)

export_to_excel_with_format(
    SU, 
    f"{output_path}/listing/表44 个人生活史清单.xlsx", 
    "表44 个人生活史清单", 
    f"表44 个人生活史清单（{lc}例）"
)

# %% [markdown]
# # 清单：过敏史

# %%
cols = ["受试者", "过敏原_TXT", "其他过敏原", "严重程度_TXT"]
cols1 = [
 '疼痛',
 '脱皮',
 '发红',
 '发痒',
 '长疹子',
 '长痘',
 '红肿']
cols2 = ['其他过敏症状']

MH_ALLE = pd.read_excel(raw_path, sheet_name = "MH_ALLE", header = 0, skiprows = [1], usecols = cols + cols1 + cols2)

MH_ALLE["过敏原_TXT"] = np.where(
    MH_ALLE["其他过敏原"].notna(),
    MH_ALLE["其他过敏原"],
    MH_ALLE["过敏原_TXT"],
)
MH_ALLE = MH_ALLE[MH_ALLE["过敏原_TXT"].notna()]
MH_ALLE = MH_ALLE.merge(RAND, on = "受试者", how = "left")

# 将字段下的“√”变成具体的值
for index, row in MH_ALLE.iterrows():
    for col in cols1:
        if pd.notna(row[col]):
            MH_ALLE.at[index, col] = col

# 拼接对试验药物采取的措施、严重不良事件定义等多选字段
MH_ALLE["过敏症状"] = MH_ALLE[cols1 + cols2].apply(lambda row: ";".join(row.dropna()), axis=1)

MH_ALLE.columns = [col.replace("_TXT", "") for col in MH_ALLE.columns]
MH_ALLE = MH_ALLE.rename(columns = {
    "受试者":"筛选号",
})

stand_cols = ["筛选号", "随机号", "过敏原", "过敏症状", "严重程度"]
MH_ALLE = MH_ALLE[stand_cols]

MH_ALLE.insert(0, "No.", range(1, len(MH_ALLE) + 1))

lc = len(MH_ALLE)

export_to_excel_with_format(
    MH_ALLE, 
    f"{output_path}/listing/表45 过敏史清单.xlsx", 
    "表45 过敏史清单", 
    f"表45 过敏史清单（{lc}例）"
)

# %% [markdown]
# # 清单：既往史及现病史

# %%
cols = ["受试者", "疾病名称", "开始日期", "筛选期问询时，是否持续_TXT", "结束日期"]
MH = pd.read_excel(raw_path, sheet_name = "MH", header = 0, skiprows = [1], usecols = cols)
MH = MH[MH["疾病名称"].notna()]

MH = MH.merge(RAND, on = "受试者", how = "left")
MH.columns = [col.replace("_TXT", "") for col in MH.columns]

MH = MH.rename(columns = {
    "受试者":"筛选号",
})

stand_cols = ["筛选号", "随机号", "疾病名称", "开始日期", "筛选期问询时，是否持续", "结束日期"]
MH = MH[stand_cols]
MH.insert(0, "No.", range(1, len(MH) + 1))

lc = len(MH)

export_to_excel_with_format(
    MH, 
    f"{output_path}/listing/表46 既往病史清单.xlsx", 
    "表46 既往病史清单", 
    f"表46 既往病史清单（{lc}例）"
)

# %% [markdown]
# # 清单：帕金森病史

# %%
index = ["受试者"]
cols1 = [
    "是否明确诊断帕金森综合征_TXT",
    "最早诊断帕金森病时间",
    "是否诊断为帕金森病_TXT",
]

cols2 = ["运动迟缓", "静止性震颤", "肌强直"]
cols3 = ["确诊不存在绝对排除标准", "至少存在2条支持标准", "没有警示征象"]
cols4 = ["是否发生精神症状_TXT", "精神症状_TXT", "首次出现日期", "筛选前一个月每周是否出现_TXT"]

MH_PD = pd.read_excel(raw_path, sheet_name = "MH_PD", header = 0, skiprows = [1], usecols = index + cols1 + cols2 + cols3+ cols4)
# 多选拼接
for index, row in MH_PD.iterrows():
    for col in cols2:
        if pd.notna(row[col]):
            MH_PD.at[index, col] = col
            
for index, row in MH_PD.iterrows():
    for col in cols3:
        if pd.notna(row[col]):
            MH_PD.at[index, col] = col
            
MH_PD["帕金森综合征诊断条件"] = MH_PD[cols2].apply(lambda row: ";".join(row.dropna()), axis=1)
MH_PD["帕金森病具备的诊断条件"] = MH_PD[cols3].apply(lambda row: ";".join(row.dropna()), axis=1)

MH_PD = MH_PD.merge(RAND, on = "受试者", how = "left")
MH_PD.columns = [col.replace("_TXT", "") for col in MH_PD.columns]

MH_PD = MH_PD.rename(columns = {
    "受试者":"筛选号",
})

stand_cols = [
    "筛选号", 
    "随机号", 
    "是否明确诊断帕金森综合征", 
    "最早诊断帕金森病时间", 
    "帕金森综合征诊断条件", 
    "是否诊断为帕金森病",
    "帕金森病具备的诊断条件",
    "是否发生精神症状",
    "精神症状",
    "首次出现日期",
    "筛选前一个月每周是否出现",
]
MH_PD = MH_PD[stand_cols]
MH_PD.insert(0, "No.", range(1, len(MH_PD) + 1))

lc = len(MH_PD)
ls = len(MH_PD.drop_duplicates(subset = "筛选号"))

export_to_excel_with_format(
    MH_PD, 
    f"{output_path}/listing/表47 帕金森病史清单.xlsx", 
    "表47 帕金森病史清单", 
    f"表47 帕金森病史清单（{lc}例次{ls}例）"
)

# %% [markdown]
# # 清单：既往/合并用药

# %% editable=true slideshow={"slide_type": ""}
index = ["受试者"]
cols1 = [
        "药物名称（通用名）",
        "单次给药剂量",
        "单位_TXT",
        "其他单位，请注明",
        "给药频率_TXT",
        "其他频率，请注明",
        "给药途径_TXT",
        "其他途径，请注明",
        "该用药开始日期",
        "试验结束时，是否持续_TXT",
        "该用药结束日期",
       ]
cols2 = ["帕金森病"]
cols3 = ["病史名称", "不良事件名称", "预防用药，请说明", "其他，请说明"]        

CM = pd.read_excel(raw_path, sheet_name = "CM", header = 0, skiprows = [1], usecols = index + cols1 + cols2 + cols3)
CM = CM[CM["药物名称（通用名）"].notna()]

for index, row in CM.iterrows():
    for col in cols2:
        if pd.notna(row[col]):
            CM.at[index, col] = col
            
CM["给药原因"] = CM[cols2 + cols3].apply(lambda row: ";".join(row.dropna()), axis=1)
CM["单位_TXT"] = np.where(CM["其他单位，请注明"].notna(), CM["其他单位，请注明"], CM["单位_TXT"])
CM["给药频率_TXT"] = np.where(CM["其他频率，请注明"].notna(), CM["其他频率，请注明"], CM["给药频率_TXT"])
CM["给药途径_TXT"] = np.where(CM["其他途径，请注明"].notna(), CM["其他途径，请注明"], CM["给药途径_TXT"])

CM = CM.merge(RAND, on = "受试者", how = "left")
CM.columns = [col.replace("_TXT", "") for col in CM.columns]

CM = CM.rename(columns = {
    "受试者":"筛选号",
})

stand_cols = [
    "筛选号", 
    "随机号", 
    "药物名称（通用名）", 
    "给药原因", 
    "单次给药剂量", 
    "单位",
    "给药频率",
    "给药途径",
    "该用药开始日期",
    "试验结束时，是否持续",
    "该用药结束日期",
]
CM = CM[stand_cols]
CM.insert(0, "No.", range(1, len(CM) + 1))

lc = len(CM)
ls = len(CM.drop_duplicates(subset = "筛选号"))

export_to_excel_with_format(
    CM, 
    f"{output_path}/listing/表48 既往合并用药清单.xlsx", 
    "表48 既往合并用药清单", 
    f"表48 既往合并用药清单（{lc}例次{ls}例）"
)

# %% [markdown]
# # 清单：既往/合并非药物治疗

# %%
index = ["受试者"]
cols1 = [
        "治疗名称",
        "治疗频率_TXT",
        "其他频率，请注明",
        "该治疗开始日期",
        "试验结束时，是否持续_TXT",
        "该治疗结束日期",
       ]
cols2 = ["帕金森病"]
cols3 = ["病史名称", "不良事件名称", "预防治疗，请说明", "其他，请说明"]        

PR = pd.read_excel(raw_path, sheet_name = "PR", header = 0, skiprows = [1], usecols = index + cols1 + cols2 + cols3)
PR = PR[PR["治疗名称"].notna()]

for index, row in PR.iterrows():
    for col in cols2:
        if pd.notna(row[col]):
            PR.at[index, col] = col
            
PR["治疗原因"] = PR[cols2 + cols3].apply(lambda row: ";".join(row.dropna()), axis=1)
PR["治疗频率_TXT"] = np.where(PR["其他频率，请注明"].notna(), PR["其他频率，请注明"], PR["治疗频率_TXT"])

PR = PR.merge(RAND, on = "受试者", how = "left")
PR.columns = [col.replace("_TXT", "") for col in PR.columns]

PR = PR.rename(columns = {
    "受试者":"筛选号",
})

stand_cols = [
    "筛选号", 
    "随机号", 
    "治疗名称", 
    "治疗原因", 
    "治疗频率",
    "该治疗开始日期",
    "试验结束时，是否持续",
    "该治疗结束日期",
]
PR = PR[stand_cols]
PR.insert(0, "No.", range(1, len(PR) + 1))

lc = len(PR)
ls = len(PR.drop_duplicates(subset = "筛选号"))

export_to_excel_with_format(
    PR, 
    f"{output_path}/listing/表49 既往合并非药物治疗清单.xlsx", 
    "表49 既往合并非药物治疗清单", 
    f"表49 既往合并非药物治疗清单（{lc}例次{ls}例）"
)

