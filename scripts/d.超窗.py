# %%
# %run ../env.py

# %%
TW = pd.read_excel(timewin_path, sheet_name = "时间窗", usecols = ["类别", "访视名称", "时间窗下限", "时间窗上限"])
TW["时间窗下限"] = TW["时间窗下限"].astype("Int32")
TW["时间窗上限"] = TW["时间窗上限"].astype("Int32")

# %%
index = ["受试者", "受试者状态", "访视名称","页面名称"]

# %% [markdown]
# # 超窗情况汇总

# %% [markdown]
# # 访视

# %% [markdown]
# ## 表格：访视超窗

# %%
SV = pd.read_excel(raw_path, sheet_name = "SV", header = 0, skiprows = [1], usecols = index + ["访视日期"]).rename(columns={"访视日期":"评估日期"})
df = SV.sort_values(by = ["受试者", "访视名称", "页面名称", "评估日期"])
df = df.drop_duplicates()

df = df[(df["受试者状态"] != "筛选失败")]
df["类别"] = "其他指标超窗"

RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = ["受试者", "随机日期"])

df = (df.merge(TW, left_on = ["类别", "访视名称"], right_on = ["类别", "访视名称"], how = "left")
        .merge(RAND, on = "受试者", how = "left")
     )

# 确保日期列是 datetime 格式
df["随机日期"] = pd.to_datetime(df["随机日期"], errors='coerce')
df["评估日期"] = pd.to_datetime(df["评估日期"], errors='coerce')

# 确保时间窗列是数值型
df["时间窗上限"] = pd.to_numeric(df["时间窗上限"], errors='coerce')
df["时间窗下限"] = pd.to_numeric(df["时间窗下限"], errors='coerce')

# 定义条件：访视名称为"筛选期（V1，D-15~-13）"
condition = df["访视名称"] == "筛选期（V1，D-15~-13）"

df.loc[~condition, "上限"] = df.loc[~condition, "随机日期"] + pd.to_timedelta(df.loc[~condition, "时间窗上限"], unit='D')
df.loc[~condition, "下限"] = df.loc[~condition, "随机日期"] + pd.to_timedelta(df.loc[~condition, "时间窗下限"], unit='D')

df.loc[condition, "上限"] = df.loc[condition, "随机日期"] - pd.to_timedelta(df.loc[condition, "时间窗上限"], unit='D')
df.loc[condition, "下限"] = df.loc[condition, "随机日期"] - pd.to_timedelta(df.loc[condition, "时间窗下限"], unit='D')

# 比较评估日期与上下限，判断超窗与否
df["超窗"] = np.where((df["评估日期"] > df["上限"]) | (df["评估日期"] < df["下限"]), "超窗", "未超窗")
df = df[df["超窗"] == "超窗"]

# %%
visit = df.copy()
visit["超窗时间（天）"] = np.where(
    visit["评估日期"] > visit["上限"],
    (visit["评估日期"] - visit["上限"]).dt.days,
    (visit["评估日期"] - visit["下限"]).dt.days
)

visit["计划时间窗"] = visit["下限"].astype(str) + "-" + visit["上限"].astype(str)

cols = ["受试者", "服药日期"]
EC = pd.read_excel(raw_path, sheet_name = "EC", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
EC["服药日期"] = pd.to_datetime(EC["服药日期"], errors="coerce")
EC = EC[EC["服药日期"].notna()]
EC = (EC.groupby("受试者", dropna=False)["服药日期"].agg(["min"]).rename(columns={"min": "首次用药日期"}))

cols = ["受试者", "页面名称", "是否完成试验_TXT"]
DS_END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
DS_END = DS_END[DS_END["页面名称"] == "试验完成情况总结"].drop(columns = "页面名称")

cols = ["受试者", "受试者状态", "随机号"]
RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
RAND = RAND[RAND["受试者状态"] != "筛选失败"]

visit = (visit.merge(DS_END, on = "受试者", how = "left")
              .merge(EC, on = "受试者", how = "left")
              .merge(RAND, on = "受试者", how = "left")
     )

cols = ["受试者", "随机号", "访视名称", "页面名称", "评估日期", "首次用药日期", "计划时间窗", "超窗时间（天）", "是否完成试验_TXT"]
visit = visit[cols]
visit.insert(0, "No.", range(1, len(visit) + 1))

visit["评估日期"] = visit["评估日期"].dt.strftime('%Y-%m-%d')
visit["首次用药日期"] = visit["首次用药日期"].dt.strftime('%Y-%m-%d')
visit = visit.reindex(visit["超窗时间（天）"].abs().sort_values(ascending=False).index)

visit = visit.rename(columns = {
        "受试者":"筛选号",
        "页面名称":"表单名称",
        "是否完成试验_TXT":"是否完成试验",
        "评估日期":"发生日期",
                         })

lc = len(visit)
ls = len(visit.drop_duplicates(subset = ["筛选号"]))

export_to_excel_with_format(
    visit, 
    f"{output_path}/listing/访视超窗清单.xlsx", 
    "访视超窗清单", 
    f"访视超窗清单（{lc}例次{ls}例）"
)

# %% [markdown]
# ## 汇总部分：访视超窗

# %%
visit_table = df.copy()

# %% [markdown]
# # 疗效指标

# %% [markdown]
# ## 表格：疗效评价指标超窗

# %% editable=true slideshow={"slide_type": ""}
cols = ["评估日期"]
QS_SAPS = pd.read_excel(raw_path, sheet_name = "QS_SAPS", header = 0, skiprows = [1], usecols = index + cols)
QS_SAPS = QS_SAPS.drop_duplicates()
QS_CGIS = pd.read_excel(raw_path, sheet_name = "QS_CGIS", header = 0, skiprows = [1], usecols = index + cols)
QS_CGII = pd.read_excel(raw_path, sheet_name = "QS_CGII", header = 0, skiprows = [1], usecols = index + cols)
QS_MDS = pd.read_excel(raw_path, sheet_name = "QS_MDS", header = 0, skiprows = [1], usecols = index + cols)
df = pd.concat([QS_SAPS, QS_CGIS, QS_CGII, QS_MDS]).sort_values(by = index + cols)
df = df[(df["受试者状态"] != "筛选失败")]
df["类别"] = "疗效评价指标超窗"

RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = ["受试者", "随机日期"])

df = (df.merge(TW, left_on = ["类别", "访视名称"], right_on = ["类别", "访视名称"], how = "left")
        .merge(RAND, on = "受试者", how = "left")
     )

# 确保日期列是 datetime 格式
df["随机日期"] = pd.to_datetime(df["随机日期"], errors='coerce')
df["评估日期"] = pd.to_datetime(df["评估日期"], errors='coerce')

# 确保时间窗列是数值型
df["时间窗上限"] = pd.to_numeric(df["时间窗上限"], errors='coerce')
df["时间窗下限"] = pd.to_numeric(df["时间窗下限"], errors='coerce')

# 定义条件：访视名称为"筛选期（V1，D-15~-13）"
condition = df["访视名称"] == "筛选期（V1，D-15~-13）"

df.loc[~condition, "上限"] = df.loc[~condition, "随机日期"] + pd.to_timedelta(df.loc[~condition, "时间窗上限"], unit='D')
df.loc[~condition, "下限"] = df.loc[~condition, "随机日期"] + pd.to_timedelta(df.loc[~condition, "时间窗下限"], unit='D')

df.loc[condition, "上限"] = df.loc[condition, "随机日期"] - pd.to_timedelta(df.loc[condition, "时间窗上限"], unit='D')
df.loc[condition, "下限"] = df.loc[condition, "随机日期"] - pd.to_timedelta(df.loc[condition, "时间窗下限"], unit='D')

# 比较评估日期与上下限，判断超窗与否
df["超窗"] = np.where((df["评估日期"] > df["上限"]) | (df["评估日期"] < df["下限"]), "超窗", "未超窗")
df = df[df["超窗"] == "超窗"]

# 删除已经被视为访视超窗的数据
# efficacy_drop = visit_table[["受试者", "访视名称", "超窗", "评估日期"]].rename(columns = {"超窗":"访视超窗", "评估日期":"访视日期"})
# df = df.merge(efficacy_drop, on = ["受试者", "访视名称"], how = "left")
# df = df[(df["访视超窗"].isna()) | (df["访视日期"] != df["评估日期"])]

# %% editable=true slideshow={"slide_type": ""}
efficacy = df.copy()
efficacy["超窗时间（天）"] = np.where(
    efficacy["评估日期"] > efficacy["上限"],
    (efficacy["评估日期"] - efficacy["上限"]).dt.days,
    (efficacy["评估日期"] - efficacy["下限"]).dt.days
)

efficacy["计划时间窗"] = efficacy["下限"].astype(str) + "-" + efficacy["上限"].astype(str)

EC = pd.read_excel(raw_path, sheet_name = "EC", header = 0, skiprows = [1], usecols = ["受试者", "服药日期"], dtype = str).fillna("")
EC["服药日期"] = pd.to_datetime(EC["服药日期"], errors="coerce")
EC = EC[EC["服药日期"].notna()]
EC = (EC.groupby("受试者", dropna=False)["服药日期"].agg(["min"]).rename(columns={"min": "首次用药日期"}))

cols = ["受试者", "页面名称", "是否完成试验_TXT"]
DS_END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
DS_END = DS_END[DS_END["页面名称"] == "试验完成情况总结"].drop(columns = "页面名称")

cols = ["受试者", "受试者状态", "随机号"]
RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
RAND = RAND[RAND["受试者状态"] != "筛选失败"].drop(columns = "受试者状态")

#是否完成访视5
cols = ["受试者", "访视OID", "是否进行本次访视_TXT"]
SV = pd.read_excel(raw_path, sheet_name = "SV", header = 0, skiprows = [1], usecols = cols)
SV= SV[(SV["访视OID"] == "V50")]

efficacy = (efficacy.merge(DS_END, on = "受试者", how = "left")
         .merge(EC, on = "受试者", how = "left")
         .merge(RAND, on = "受试者", how = "left")
         .merge(SV, on = "受试者", how = "left")
     )

efficacy = efficacy.rename(columns = {
    "受试者":"筛选号",
    "页面名称":"表单名称",
    "是否完成试验_TXT":"是否完成试验",
    "是否进行本次访视_TXT":"是否完成访视5"
                         })

stand_cols = [
    "筛选号", 
    "随机号", 
    "访视名称",
    "表单名称", 
    "评估日期", 
    "首次用药日期", 
    "计划时间窗", 
    "超窗时间（天）", 
    "是否完成访视5", 
    "是否完成试验"]

efficacy = efficacy[stand_cols].fillna("")

efficacy["评估日期"] = efficacy["评估日期"].dt.strftime('%Y-%m-%d')
efficacy["首次用药日期"] = efficacy["首次用药日期"].dt.strftime('%Y-%m-%d')
efficacy = efficacy.reindex( efficacy["超窗时间（天）"].abs().sort_values(ascending=False).index)

notes = []

table_no = 16
for visit_name, sub_df in efficacy.groupby("访视名称", sort=False, dropna=False):
    visit_disp = "未知访视" if pd.isna(visit_name) else str(visit_name)
    
    lc = len(sub_df)
    ls = len(sub_df.drop_duplicates(subset = ["筛选号"]))
    sub_df.insert(0, "No.", range(1, len(sub_df) + 1))
    
    file_path = f"{output_path}/table/表{table_no} 疗效评价指标超窗清单（{visit_disp}）.docx"
    title = f"表{table_no} 疗效评价指标超窗清单（{visit_disp}）（{lc}例次{ls}例）"
    
    save_table_to_docx_threeline(
        sub_df,
        str(file_path),
        title,
        notes,
        row_height_cm=0.6,
        auto_width=True,
        include_notes=False,
    )

    table_no += 1

# %% [markdown]
# ## 汇总部分：疗效评价指标超窗

# %%
efficacy_table = df.copy()

# %% [markdown]
# # 安全性指标

# %% [markdown]
# ## 表格：安全性评价指标超窗

# %%
VS = pd.read_excel(raw_path, sheet_name = "VS", header = 0, skiprows = [1], usecols = index + ["检查日期"]).rename(columns={"检查日期":"评估日期"})
PE = pd.read_excel(raw_path, sheet_name = "PE", header = 0, skiprows = [1], usecols = index + ["检查日期"]).rename(columns={"检查日期":"评估日期"})
EG = pd.read_excel(raw_path, sheet_name = "EG", header = 0, skiprows = [1], usecols = index + ["检查日期"]).rename(columns={"检查日期":"评估日期"})

LB_HEM = pd.read_excel(raw_path, sheet_name = "LB_HEM", header = 0, skiprows = [1], usecols = index + ["采样日期"]).rename(columns={"采样日期":"评估日期"})
LB_LFT = pd.read_excel(raw_path, sheet_name = "LB_LFT", header = 0, skiprows = [1], usecols = index + ["采样日期"]).rename(columns={"采样日期":"评估日期"})
LB_RFT = pd.read_excel(raw_path, sheet_name = "LB_RFT", header = 0, skiprows = [1], usecols = index + ["采样日期"]).rename(columns={"采样日期":"评估日期"})
LB_ELECT = pd.read_excel(raw_path, sheet_name = "LB_ELECT", header = 0, skiprows = [1], usecols = index + ["采样日期"]).rename(columns={"采样日期":"评估日期"})
LB_FBG = pd.read_excel(raw_path, sheet_name = "LB_FBG", header = 0, skiprows = [1], usecols = index + ["采样日期"]).rename(columns={"采样日期":"评估日期"})
LB_URI = pd.read_excel(raw_path, sheet_name = "LB_URI", header = 0, skiprows = [1], usecols = index + ["采样日期"]).rename(columns={"采样日期":"评估日期"})
LB_HCG1 = pd.read_excel(raw_path, sheet_name = "LB_HCG1", header = 0, skiprows = [1], usecols = index + ["采样日期"]).rename(columns={"采样日期":"评估日期"})
LB_HCG2 = pd.read_excel(raw_path, sheet_name = "LB_HCG2", header = 0, skiprows = [1], usecols = index + ["采样日期"]).rename(columns={"采样日期":"评估日期"})

df = pd.concat([VS, PE, EG, LB_HEM, LB_LFT, LB_RFT, LB_ELECT, LB_FBG, LB_URI, LB_HCG1, LB_HCG2]).sort_values(by = ["受试者", "访视名称", "页面名称", "评估日期"])
df = df.drop_duplicates()

df = df[(df["受试者状态"] != "筛选失败")]
df["类别"] = "安全性评价指标超窗"

RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = ["受试者", "随机日期"])

df = (df.merge(TW, left_on = ["类别", "访视名称"], right_on = ["类别", "访视名称"], how = "left")
        .merge(RAND, on = "受试者", how = "left")
     )

# 确保日期列是 datetime 格式
df["随机日期"] = pd.to_datetime(df["随机日期"], errors='coerce')
df["评估日期"] = pd.to_datetime(df["评估日期"], errors='coerce')

# 确保时间窗列是数值型
df["时间窗上限"] = pd.to_numeric(df["时间窗上限"], errors='coerce')
df["时间窗下限"] = pd.to_numeric(df["时间窗下限"], errors='coerce')

# 定义条件：访视名称为"筛选期（V1，D-15~-13）"
condition = df["访视名称"] == "筛选期（V1，D-15~-13）"

df.loc[~condition, "上限"] = df.loc[~condition, "随机日期"] + pd.to_timedelta(df.loc[~condition, "时间窗上限"], unit='D')
df.loc[~condition, "下限"] = df.loc[~condition, "随机日期"] + pd.to_timedelta(df.loc[~condition, "时间窗下限"], unit='D')

df.loc[condition, "上限"] = df.loc[condition, "随机日期"] - pd.to_timedelta(df.loc[condition, "时间窗上限"], unit='D')
df.loc[condition, "下限"] = df.loc[condition, "随机日期"] - pd.to_timedelta(df.loc[condition, "时间窗下限"], unit='D')

# 比较评估日期与上下限，判断超窗与否
df["超窗"] = np.where((df["评估日期"] > df["上限"]) | (df["评估日期"] < df["下限"]), "超窗", "未超窗")
df = df[df["超窗"] == "超窗"]

# 移除访视超窗的情况
visit_drop = visit_table[["受试者", "访视名称", "超窗", "评估日期"]].rename(columns = {"超窗":"访视超窗", "评估日期":"访视日期"})
df = df.merge(visit_drop, on = ["受试者", "访视名称"], how = "left")
df = df[(df["访视超窗"].isna()) | (df["访视日期"] != df["评估日期"])]

# %%
safe = df.copy()
safe["超窗时间（天）"] = np.where(
    safe["评估日期"] > safe["上限"],
    (safe["评估日期"] - safe["上限"]).dt.days,
    (safe["评估日期"] - safe["下限"]).dt.days
)

safe["计划时间窗"] = safe["下限"].astype(str) + "-" + safe["上限"].astype(str)

cols = ["受试者", "服药日期"]
EC = pd.read_excel(raw_path, sheet_name = "EC", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
EC["服药日期"] = pd.to_datetime(EC["服药日期"], errors="coerce")
EC = EC[EC["服药日期"].notna()]
EC = (EC.groupby("受试者", dropna=False)["服药日期"].agg(["min"]).rename(columns={"min": "首次用药日期"}))

cols = ["受试者", "页面名称", "是否完成试验_TXT"]
DS_END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
DS_END = DS_END[DS_END["页面名称"] == "试验完成情况总结"].drop(columns = "页面名称")

cols = ["受试者", "受试者状态", "随机号"]
RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
RAND = RAND[RAND["受试者状态"] != "筛选失败"]

safe = (safe.merge(DS_END, on = "受试者", how = "left")
            .merge(EC, on = "受试者", how = "left")
            .merge(RAND, on = "受试者", how = "left")
     )


safe["评估日期"] = safe["评估日期"].dt.strftime('%Y-%m-%d')
safe["首次用药日期"] = safe["首次用药日期"].dt.strftime('%Y-%m-%d')
safe = safe.reindex(safe["超窗时间（天）"].abs().sort_values(ascending=False).index)

safe = safe.rename(columns = {
    "受试者":"筛选号",
    "页面名称":"表单名称",
    "是否完成试验_TXT":"是否完成试验",
    "评估日期":"采样/检查/测量日期",
                         })

lc = len(safe)
ls = len(safe.drop_duplicates(subset = ["筛选号"]))

stand_cols = [
    "筛选号", 
    "随机号", 
    "访视名称", 
    "表单名称", 
    "采样/检查/测量日期", 
    "首次用药日期", 
    "计划时间窗", 
    "超窗时间（天）", 
    "是否完成试验"]

safe = safe[stand_cols]
safe.insert(0, "No.", range(1, len(safe) + 1))

export_to_excel_with_format(
    safe, 
    f"{output_path}/listing/表34 安全性评价指标超窗清单.xlsx", 
    "表34 安全性评价指标超窗清单", 
    f"表34 安全性评价指标超窗清单（{lc}例次{ls}例）"
)

# %% [markdown]
# ## 汇总部分：安全性指标超窗

# %%
safe_table = df.copy()

# %% [markdown]
# # 其他指标

# %% [markdown]
# ## 表格：其他指标

# %%
VS_HW = pd.read_excel(raw_path, sheet_name = "VS_HW", header = 0, skiprows = [1], usecols = index + ["测量日期"]).rename(columns={"测量日期":"评估日期"})
VS_W_1 = pd.read_excel(raw_path, sheet_name = "VS_W_1", header = 0, skiprows = [1], usecols = index + ["测量日期"]).rename(columns={"测量日期":"评估日期"})
QS_NPI = pd.read_excel(raw_path, sheet_name = "QS_NPI", header = 0, skiprows = [1], usecols = index + ["评估日期"])
QS_SSRS = pd.read_excel(raw_path, sheet_name = "QS_SSRS", header = 0, skiprows = [1], usecols = index + ["评估日期"])
QS_MMSE = pd.read_excel(raw_path, sheet_name = "QS_MMSE", header = 0, skiprows = [1], usecols = index + ["评估日期"])
DA_DD = pd.read_excel(raw_path, sheet_name = "DA_DD", header = 0, skiprows = [1], usecols = index + ["发药日期"]).rename(columns={"发药日期":"评估日期"})
DA_DR = pd.read_excel(raw_path, sheet_name = "DA_DR", header = 0, skiprows = [1], usecols = index + ["回收日期"]).rename(columns={"回收日期":"评估日期"})

df = pd.concat([VS_HW, VS_W_1, QS_NPI, QS_SSRS, QS_MMSE, DA_DD, DA_DR]).sort_values(by = ["受试者", "访视名称", "页面名称", "评估日期"])
df = df.drop_duplicates()

df = df[(df["受试者状态"] != "筛选失败")]
df["类别"] = "其他指标超窗"

RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = ["受试者", "随机日期"])

df = (df.merge(TW, left_on = ["类别", "访视名称"], right_on = ["类别", "访视名称"], how = "left")
        .merge(RAND, on = "受试者", how = "left")
     )

# 确保日期列是 datetime 格式
df["随机日期"] = pd.to_datetime(df["随机日期"], errors='coerce')
df["评估日期"] = pd.to_datetime(df["评估日期"], errors='coerce')

# 确保时间窗列是数值型
df["时间窗上限"] = pd.to_numeric(df["时间窗上限"], errors='coerce')
df["时间窗下限"] = pd.to_numeric(df["时间窗下限"], errors='coerce')

# 定义条件：访视名称为"筛选期（V1，D-15~-13）"
condition = df["访视名称"] == "筛选期（V1，D-15~-13）"

df.loc[~condition, "上限"] = df.loc[~condition, "随机日期"] + pd.to_timedelta(df.loc[~condition, "时间窗上限"], unit='D')
df.loc[~condition, "下限"] = df.loc[~condition, "随机日期"] + pd.to_timedelta(df.loc[~condition, "时间窗下限"], unit='D')

df.loc[condition, "上限"] = df.loc[condition, "随机日期"] - pd.to_timedelta(df.loc[condition, "时间窗上限"], unit='D')
df.loc[condition, "下限"] = df.loc[condition, "随机日期"] - pd.to_timedelta(df.loc[condition, "时间窗下限"], unit='D')

# 比较评估日期与上下限，判断超窗与否
df["超窗"] = np.where((df["评估日期"] > df["上限"]) | (df["评估日期"] < df["下限"]), "超窗", "未超窗")
df = df[df["超窗"] == "超窗"]

# 移除访视超窗的情况
visit_drop = visit_table[["受试者", "访视名称", "超窗", "评估日期"]].rename(columns = {"超窗":"访视超窗", "评估日期":"访视日期"})
df = df.merge(visit_drop, on = ["受试者", "访视名称"], how = "left")
df = df[(df["访视超窗"].isna()) | (df["访视日期"] != df["评估日期"])]

# %%
oth = df.copy()
oth["超窗时间（天）"] = np.where(
    oth["评估日期"] > oth["上限"],
    (oth["评估日期"] - oth["上限"]).dt.days,
    (oth["评估日期"] - oth["下限"]).dt.days
)

oth["计划时间窗"] = oth["下限"].astype(str) + "-" + oth["上限"].astype(str)

cols = ["受试者", "服药日期"]
EC = pd.read_excel(raw_path, sheet_name = "EC", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
EC["服药日期"] = pd.to_datetime(EC["服药日期"], errors="coerce")
EC = EC[EC["服药日期"].notna()]
EC = (EC.groupby("受试者", dropna=False)["服药日期"].agg(["min"]).rename(columns={"min": "首次用药日期"}))

cols = ["受试者", "页面名称", "是否完成试验_TXT"]
DS_END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
DS_END = DS_END[DS_END["页面名称"] == "试验完成情况总结"].drop(columns = "页面名称")

cols = ["受试者", "受试者状态", "随机号"]
RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
RAND = RAND[RAND["受试者状态"] != "筛选失败"]

oth = (oth.merge(DS_END, on = "受试者", how = "left")
            .merge(EC, on = "受试者", how = "left")
            .merge(RAND, on = "受试者", how = "left")
     )

oth["评估日期"] = oth["评估日期"].dt.strftime('%Y-%m-%d')
oth["首次用药日期"] = oth["首次用药日期"].dt.strftime('%Y-%m-%d')
oth = oth.reindex(oth["超窗时间（天）"].abs().sort_values(ascending=False).index)

oth = oth.rename(columns = {
    "受试者":"筛选号",
    "页面名称":"表单名称",
    "是否完成试验_TXT":"是否完成试验",
    "评估日期":"发生日期",
                         })
stand_cols = [
    "筛选号", 
    "随机号", 
    "访视名称", 
    "表单名称", 
    "发生日期", 
    "首次用药日期", 
    "计划时间窗", 
    "超窗时间（天）", 
    "是否完成试验"]

oth = oth[stand_cols]
oth.insert(0, "No.", range(1, len(oth) + 1))

lc = len(oth)
ls = len(oth.drop_duplicates(subset = ["筛选号"]))

export_to_excel_with_format(
    oth, 
    f"{output_path}/listing/表35 其他指标超窗清单.xlsx", 
    "表35 其他指标超窗清单", 
    f"表35 其他指标超窗清单（{lc}例次{ls}例）"
)

# %% [markdown]
# ## 汇总部分：其他指标超窗

# %%
oth_table = df.copy()

# %% [markdown]
# # 汇总

# %% [markdown]
# ## 表格：超窗汇总结果

# %%
df = pd.concat([efficacy_table, safe_table, oth_table, visit_table])

df["超窗时间（天）"] = np.where(
    df["评估日期"] > df["上限"],
    (df["评估日期"] - df["上限"]).dt.days,
    (df["下限"] - df["评估日期"]).dt.days
)

page_stats = df.groupby(["类别", "页面名称"]).agg(
    例次=("受试者", "count"),           # 总记录数
    例数=("受试者", "nunique"),         # 去重后的受试者数
    最小超窗时间=("超窗时间（天）", "min"),  # 最小超窗时间
    最大超窗时间=("超窗时间（天）", "max")   # 最大超窗时间
).reset_index()

# 步骤2：计算每个类别的汇总数据
category_stats = df.groupby("类别").agg(
    例次=("受试者", "count"),
    例数=("受试者", "nunique"),
    最小超窗时间=("超窗时间（天）", "min"),
    最大超窗时间=("超窗时间（天）", "max")
).reset_index()

# 为类别汇总行添加空的页面名称列
category_stats["页面名称"] = ""
category_page_order = {
    "疗效评价指标超窗": ["SAPS量表", "疾病严重程度量表（CGI-S）", "总体进步量表（CGI-I）", "MDS统一帕金森病评定量表（MDS-UPDRS）"],
    "安全性评价指标超窗": ["生命体征", "体格检查", "血常规", "肝功能", "肾功能", "电解质", "空腹血糖", "尿常规", "血妊娠", "尿妊娠", "12导联心电图"],
    "其他指标超窗": ["体重", "访视日期", "身高体重", "神经精神量表（NPI）", "哥伦比亚-自杀严重程度评定量表（C-SSRS）", "简易精神状态检查量表（MMSE）", "试验药物发放记录", "试验药物回收记录"]
}
result_rows = []

for category, page_list in category_page_order.items():
    # 检查该类别是否存在
    if category not in category_stats["类别"].values:
        continue
    
    # 添加类别汇总行
    category_row = category_stats[category_stats["类别"] == category].iloc[0]
    result_rows.append({
        "未遵循研究方案时间窗子类别": category_row["类别"],
        "例次": category_row["例次"],
        "例数": category_row["例数"],
        "最小超窗时间（天）": category_row["最小超窗时间"],
        "最大超窗时间（天）": category_row["最大超窗时间"],
        "层级": "类别"
    })
    
    # 按照指定顺序添加页面行
    pages = page_stats[page_stats["类别"] == category]
    for page_name in page_list:
        page_row = pages[pages["页面名称"] == page_name]
        if len(page_row) > 0:
            page_row = page_row.iloc[0]
            result_rows.append({
                "未遵循研究方案时间窗子类别": f"        {page_row['页面名称']}",
                "例次": page_row["例次"],
                "例数": page_row["例数"],
                "最小超窗时间（天）": page_row["最小超窗时间"],
                "最大超窗时间（天）": page_row["最大超窗时间"],
                "层级": "页面"
            })

# %% editable=true slideshow={"slide_type": ""}
df = pd.DataFrame(result_rows)
df = df[["未遵循研究方案时间窗子类别", "例次", "例数", "最小超窗时间（天）", "最大超窗时间（天）"]]
df["最小超窗时间（天）"] = df["最小超窗时间（天）"].astype("Int32")
df["最大超窗时间（天）"] = df["最大超窗时间（天）"].astype("Int32")
notes = []
save_table_to_docx_threeline(
        df, 
        f'{output_path}/table/表18 超窗情况汇总.docx', 
        f'表18 超窗情况汇总', 
        notes,
        row_height_cm=0.6,
        auto_width=True,
        include_notes=False,
        alignment = WD_TABLE_ALIGNMENT.LEFT
    )
