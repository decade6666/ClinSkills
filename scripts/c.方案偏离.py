# %%
# %run ../env.py

# %%
# 随机号 随机日期
RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = ["受试者", "随机日期", "随机号"])

# 是否完成试验
DS_END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = ["受试者", "页面名称","是否完成试验_TXT"])
DS_END = DS_END[DS_END["页面名称"] == "试验完成情况总结"]

# %%
index = ["受试者", "受试者状态", "访视名称","页面名称"]

# %% [markdown]
# ## 表12 中心层面方案偏离情况

# %%
PD = pd.read_excel(pd_path, sheet_name = "方案偏离", usecols = ["中心编号", "筛选号", "严重程度", "分类"], header = 4, dtype = str)
PD = PD[PD["筛选号"].isna()]
df = PD.drop_duplicates()

df = df.rename(columns={"中心编号":"中心"}).fillna("")
stand_cols = ["中心", "严重程度", "分类"]

notes = []
save_table_to_docx_threeline(
        df, 
        f'{output_path}/table/表12 中心层面方案偏离情况.docx', 
        '表12 中心层面方案偏离情况', 
        notes,
        row_height_cm=0.6,
        auto_width=True
    )

# %% [markdown]
# ## 表13 受试者层面方案偏离情况

# %%
PD = pd.read_excel(pd_path, sheet_name = "方案偏离", usecols = ["筛选号", "严重程度", "分类"], header = 4, dtype = str)
# 按"严重程度"和"分类"分组，计算例次和例数
df = PD.groupby(["严重程度", "分类"]).agg(
    例次=("筛选号", "count"),      # 每个分组内的总记录数
    例数=("筛选号", "nunique")     # 每个分组内去重后的筛选号数量
).reset_index()

# 调整列的顺序
df = df[["严重程度", "分类", "例次", "例数"]]

save_table_to_docx_threeline(
        df, 
        f'{output_path}/table/表13 受试者层面方案偏离情况.docx', 
        '表13 受试者层面方案偏离情况', 
        notes,
        row_height_cm=0.6,
        auto_width=True,
        include_notes=False,
        merge_columns=['严重程度']
    )

# %% [markdown]
# ## 表14 严重方案偏离情况 （XXX例次XXX例）

# %%
PD = pd.read_excel(pd_path, sheet_name = "方案偏离", usecols = ["中心编号", "筛选号", "严重程度", "分类", "详述"], header = 4, dtype = str)
PD = PD[PD["严重程度"] == "严重方案偏离(Major PD)"]
RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = ["受试者","随机号", "随机日期"], dtype = str).fillna("")
DS_END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = ["受试者", "页面名称", "是否完成试验_TXT"], dtype = str).fillna("")
DS_END = DS_END[DS_END["页面名称"] == "试验完成情况总结"]

df = (PD.merge(RAND, left_on = "筛选号", right_on = "受试者", how = "left")
       .merge(DS_END, left_on = "筛选号", right_on = "受试者", how = "left"))

df = df.rename(columns = {
    "是否完成试验_TXT":"是否完成试验"
})

df =df[["中心编号", "筛选号", "随机号", "随机日期", "是否完成试验", "分类", "详述"]]
df.insert(0, "No.", range(1, len(df) + 1))

lc = len(df)
ls = len(df.drop_duplicates(subset = ["筛选号"]))

save_table_to_docx_threeline(
        df, 
        f'{output_path}/table/表14 严重方案偏离情况.docx', 
        f'表14 严重方案偏离情况 （{lc}例次{ls}例）', 
        notes,
        row_height_cm=0.6,
        auto_width=True,
        include_notes=False
    )
