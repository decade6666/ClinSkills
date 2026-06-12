# %%
# %run ../../env.py
from utils.loaders import load_first_dose
from utils.loaders import load_completion
from utils.loaders import load_rand

# %%
index = ["受试者", "受试者状态", "访视名称", "页面名称"]
sig =["病史名称", "不良事件名称", "帕金森病_TXT", "其他，请说明"]

prefix_map = { "病史名称": "MH:", "不良事件名称": "AE:", "帕金森病_TXT": "研究疾病:", "其他，请说明": "其他:" }
gcols = ["受试者", "页面名称", "项目"]

def pick_rows(g):
    base = g[g["访视名称"].eq("基线期（V2，D1）")]
    if not base.empty:
        if (base["临床意义"] == "异常有临床意义").any():
            return g.iloc[0:0]
        return base

    scr = g[g["访视名称"].eq("筛选期（V1，D-15~-13）")]
    scr = scr[scr["临床意义"].ne("异常有临床意义")]
    if not scr.empty:
        return scr

    return g.iloc[0:0]

# %% [markdown]
# ## 重算：生命体征 -> VS

# %%
cols = ["检查日期", "项目_TXT", "结果", "临床意义_TXT", "单位_TXT"]
VS = pd.read_excel(raw_path, sheet_name = "VS", header = 0, skiprows = [1], usecols = index + cols + sig)
VS = VS.rename(columns={"检查日期":"评估日期", "项目_TXT":"项目", "临床意义_TXT":"临床意义", "单位_TXT":"单位"})

df = pd.concat([VS])
df = df[(df["受试者状态"] != "筛选失败") & (df["临床意义"].notna())]

EC = load_first_dose().rename(columns={"首次用药日期": "服药日期"})
DS_END = load_completion()
RAND = load_rand(cols=['受试者', '随机号'])

df = (df.merge(EC, on = ["受试者"], how = "left").merge(DS_END, on = ["受试者"], how = "left").merge(RAND, on = "受试者", how = "left"))
df["分组"] = df.apply( lambda row: "给药前检查" if row["评估日期"] <= row["服药日期"] else "给药后检查", axis=1 )

pre = df.drop(columns=sig)
pre = pre[pre["分组"] == "给药前检查"]
pre = pre.sort_values(by=["受试者", "页面名称", "项目", "服药日期"])
pre = (pre.groupby(gcols, group_keys=False).apply(pick_rows).reset_index(drop=True))

post = df[(df["分组"] == "给药后检查") & (df["临床意义"] == "异常有临床意义")].copy()
post["异常有临床意义，请描述"] = post[sig].apply( lambda row: ";".join( f"{prefix_map[col]}{str(val).replace('√', '帕金森病')}" for col, val in row.items() if pd.notna(val) and str(val).strip() != "" ), axis=1 )
post = post[["受试者", "访视名称", "页面名称", "评估日期", "项目", "结果", "临床意义", "异常有临床意义，请描述"]]

merge = pre.merge(post, on = ["受试者", "页面名称", "项目"], how = "left")
merge = merge[~((merge["结果_x"].isna()) | (merge["结果_y"].isna()))]
merge = merge.rename(columns = {
    "访视名称_y":"访视名称_首次用药后", "结果_y":"检查结果_首次用药后",
    "页面名称":"表单名称", "项目":"检查项", "受试者":"筛选号", "是否完成试验_TXT":"是否完成试验",
})
merge.insert(0, "No.", range(1, len(merge) + 1))
merge['temp_id_visit'] = merge['筛选号'].astype(str) + merge['表单名称'].astype(str) + merge['检查项'].astype(str) + "_" + merge['访视名称_首次用药后'].astype(str) + merge['检查结果_首次用药后'].astype(str)
VS = merge.copy()

# %% [markdown]
# ## 重算：体格检查 -> PE

# %%
cols = ["检查日期", "检查部位_TXT", "临床意义_TXT", "如有异常予以描述："]
PE = pd.read_excel(raw_path, sheet_name = "PE", header = 0, skiprows = [1], usecols = index + cols + sig)
PE = PE.rename(columns={"检查日期":"评估日期", "检查部位_TXT":"项目", "临床意义_TXT":"临床意义", "如有异常予以描述：":"异常描述"})
PE["结果"] = PE["临床意义"]

df = pd.concat([PE])
df = df[(df["受试者状态"] != "筛选失败") & (df["临床意义"].notna())]

EC = load_first_dose().rename(columns={"首次用药日期": "服药日期"})
df = (df.merge(EC, on = ["受试者"], how = "left").merge(DS_END, on = ["受试者"], how = "left").merge(RAND, on = "受试者", how = "left"))
df["分组"] = df.apply( lambda row: "给药前检查" if row["评估日期"] <= row["服药日期"] else "给药后检查", axis=1 )

pre = df.drop(columns=sig)
pre = pre[pre["分组"] == "给药前检查"]
pre = pre.sort_values(by=["受试者", "页面名称", "项目", "服药日期"])
pre = (pre.groupby(gcols, group_keys=False).apply(pick_rows).reset_index(drop=True))

post = df[(df["分组"] == "给药后检查") & (df["临床意义"] == "异常有临床意义")].copy()
post["异常有临床意义，请描述"] = post[sig].apply( lambda row: ";".join( f"{prefix_map[col]}{str(val).replace('√', '帕金森病')}" for col, val in row.items() if pd.notna(val) and str(val).strip() != "" ), axis=1 )
post = post[["受试者", "访视名称", "页面名称", "评估日期", "项目", "结果", "临床意义", "异常描述", "异常有临床意义，请描述"]]

merge = pre.merge(post, on = ["受试者", "页面名称", "项目"], how = "left")
merge = merge[~((merge["结果_x"].isna()) | (merge["结果_y"].isna()))]
merge = merge.rename(columns = {
    "访视名称_y":"访视名称_首次用药后", "结果_y":"检查结果_首次用药后",
    "页面名称":"表单名称", "项目":"检查项", "受试者":"筛选号", "是否完成试验_TXT":"是否完成试验",
})
merge.insert(0, "No.", range(1, len(merge) + 1))
merge['temp_id_visit'] = merge['筛选号'].astype(str) + merge['表单名称'].astype(str) + merge['检查项'].astype(str) + "_" + merge['访视名称_首次用药后'].astype(str)
PE = merge.copy()

# %% [markdown]
# ## 重算：12导联心电图 -> EG

# %%
cols1 = ["检查日期", "临床意义_TXT", "异常描述"]
EG = pd.read_excel(raw_path, sheet_name = "EG", header = 0, skiprows = [1], usecols = index + cols1 + sig)
EG = EG.rename(columns={"检查日期":"评估日期", "临床意义_TXT":"临床意义"})
EG["项目"] = EG["页面名称"]
EG["结果"] = EG["临床意义"]

df = pd.concat([EG])
df = df[(df["受试者状态"] != "筛选失败") & (df["临床意义"].notna())]

EC = load_first_dose().rename(columns={"首次用药日期": "服药日期"})
df = (df.merge(EC, on = ["受试者"], how = "left").merge(DS_END, on = ["受试者"], how = "left").merge(RAND, on = "受试者", how = "left"))
df["分组"] = df.apply( lambda row: "给药前检查" if row["评估日期"] <= row["服药日期"] else "给药后检查", axis=1 )

pre = df.drop(columns=sig)
pre = pre[pre["分组"] == "给药前检查"]
pre = pre.sort_values(by=["受试者", "页面名称", "项目", "服药日期"])
pre = (pre.groupby(gcols, group_keys=False).apply(pick_rows).reset_index(drop=True))

post = df[(df["分组"] == "给药后检查") & (df["临床意义"] == "异常有临床意义")].copy()
post["异常有临床意义，请描述"] = post[sig].apply( lambda row: ";".join( f"{prefix_map[col]}{str(val).replace('√', '帕金森病')}" for col, val in row.items() if pd.notna(val) and str(val).strip() != "" ), axis=1 )
post = post[["受试者", "访视名称", "页面名称", "评估日期", "项目", "结果", "临床意义", "异常描述", "异常有临床意义，请描述"]]

merge = pre.merge(post, on = ["受试者", "页面名称", "项目"], how = "left")
merge = merge[~((merge["结果_x"].isna()) | (merge["结果_y"].isna()))]
merge = merge.rename(columns = {
    "访视名称_y":"访视名称_首次用药后", "结果_y":"检查结果_首次用药后",
    "页面名称":"表单名称", "项目":"检查项", "受试者":"筛选号", "是否完成试验_TXT":"是否完成试验",
})
merge.insert(0, "No.", range(1, len(merge) + 1))
merge['temp_id_visit'] = merge['筛选号'].astype(str) + merge['表单名称'].astype(str) + merge['检查项'].astype(str) + "_" + merge['访视名称_首次用药后'].astype(str)
EG = merge.copy()

# %% [markdown]
# ## 重算：实验室检查 -> LB

# %%
cols = ["项目.1", "测定值", "临床意义_TXT", "采样日期", "下限", "上限", "单位"]
LB_HEM = pd.read_excel(raw_path, sheet_name = "LB_HEM", header = 0, skiprows = [1], usecols = index + cols + sig)
LB_LFT = pd.read_excel(raw_path, sheet_name = "LB_LFT", header = 0, skiprows = [1], usecols = index + cols + sig)
LB_RFT = pd.read_excel(raw_path, sheet_name = "LB_RFT", header = 0, skiprows = [1], usecols = index + cols + sig)
LB_ELECT = pd.read_excel(raw_path, sheet_name = "LB_ELECT", header = 0, skiprows = [1], usecols = index + cols + sig)
LB_FBG = pd.read_excel(raw_path, sheet_name = "LB_FBG", header = 0, skiprows = [1], usecols = index + cols + sig)
LB_URI = pd.read_excel(raw_path, sheet_name = "LB_URI", header = 0, skiprows = [1], usecols = index + cols + sig)
LB_HCG1 = pd.read_excel(raw_path, sheet_name = "LB_HCG1", header = 0, skiprows = [1], usecols = index + cols + sig)

LB = pd.concat([LB_HEM, LB_LFT, LB_RFT, LB_ELECT, LB_FBG, LB_URI, LB_HCG1])
LB = LB.rename(columns={"临床意义_TXT":"临床意义", "采样日期":"评估日期", "项目.1":"项目", "测定值":"结果", "上限":"正常值范围上限", "下限":"正常值范围下限"})

cols = ["尿妊娠_TXT", "临床意义_TXT", "采样日期"]
LB_HCG2 = pd.read_excel(raw_path, sheet_name = "LB_HCG2", header = 0, skiprows = [1], usecols = index + cols + sig)
LB_HCG2 = LB_HCG2.rename(columns={"尿妊娠_TXT":"结果", "临床意义_TXT":"临床意义", "采样日期":"评估日期"})
LB_HCG2["项目"] = LB_HCG2["页面名称"]

df = pd.concat([LB, LB_HCG2])
df = df[(df["受试者状态"] != "筛选失败") & (df["临床意义"].notna())]

EC = load_first_dose().rename(columns={"首次用药日期": "服药日期"})
df = (df.merge(EC, on = ["受试者"], how = "left").merge(DS_END, on = ["受试者"], how = "left").merge(RAND, on = "受试者", how = "left"))
df["分组"] = df.apply( lambda row: "给药前检查" if row["评估日期"] <= row["服药日期"] else "给药后检查", axis=1 )

pre = df.drop(columns=sig)
pre = pre[pre["分组"] == "给药前检查"]
pre = pre.sort_values(by=["受试者", "页面名称", "项目", "服药日期"])
pre = (pre.groupby(gcols, group_keys=False).apply(pick_rows).reset_index(drop=True))

post = df[(df["分组"] == "给药后检查") & (df["临床意义"] == "异常有临床意义")].copy()
post["异常有临床意义，请描述"] = post[sig].apply( lambda row: ";".join( f"{prefix_map[col]}{str(val).replace('√', '帕金森病')}" for col, val in row.items() if pd.notna(val) and str(val).strip() != "" ), axis=1 )
post = post[["受试者", "访视名称", "页面名称", "评估日期", "项目", "结果", "临床意义", "异常有临床意义，请描述"]]

merge = pre.merge(post, on = ["受试者", "页面名称", "项目"], how = "left")
merge = merge[~((merge["结果_x"].isna()) | (merge["结果_y"].isna()))]
merge = merge.rename(columns = {
    "访视名称_y":"访视名称_首次用药后", "结果_y":"检查结果_首次用药后",
    "页面名称":"表单名称", "项目":"检查项", "受试者":"筛选号", "是否完成试验_TXT":"是否完成试验",
})
merge.insert(0, "No.", range(1, len(merge) + 1))
merge['temp_id_visit'] = merge['筛选号'].astype(str) + merge['表单名称'].astype(str) + merge['检查项'].astype(str) + "_" + merge['访视名称_首次用药后'].astype(str)
LB = merge.copy()

# %% [markdown]
# ## 表格：用药后检查异常有临床意义整体情况

# %%
merge = pd.concat([VS, PE, EG, LB])
summary = merge.groupby("表单名称").agg(
    例数=("筛选号", "nunique"),
    例次=("temp_id_visit", "nunique")
).reset_index()

merge.drop(columns=['temp_id_visit'], inplace=True)

ls = merge["筛选号"].nunique()
lc = summary["例次"].sum()

row = pd.DataFrame({
    "表单名称": ["合计"],
    "例数": [ls],
    "例次": [lc]
})

summary = pd.concat([summary, row], ignore_index=True)

order = ["生命体征", "体格检查", "血常规", "肝功能", "肾功能", "电解质", "空腹血糖", "尿常规", "血妊娠", "12导联心电图", "合计"]
summary['表单名称'] = pd.Categorical(summary['表单名称'], categories=order, ordered=True)
summary = summary.sort_values("表单名称").reset_index(drop=True)

notes = ["注：用药后检查异常有临床意义详细清单见附件：“用药后检查异常有临床意义清单”。"]
save_table_to_docx_threeline(
        summary,
        f'{output_path}/table/表27 用药后检查异常有临床意义整体情况.docx',
        f'表27 用药后检查异常有临床意义整体情况',
        notes,
        row_height_cm=0.6,
        auto_width=True,
        include_notes=True,
    )
