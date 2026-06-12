# %%
# %run ../../env.py
from utils.loaders import load_rand

# %%
# 随机号 随机日期
RAND = load_rand(cols=['受试者', '受试者状态', '随机时间', '随机号'])

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

# 注：原脚本此处用到的 measure_cols 在本块之后才定义，拆分后在此提前定义
measure_cols = [
    "对试验药物采取的措施-1_TXT",
    "对试验药物采取的措施-2_TXT",
    "对试验药物采取的措施-3_TXT",
    "对试验药物采取的措施-4_TXT",
    "对试验药物采取的措施-5_TXT",
    "对试验药物采取的措施-6_TXT"
]

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
