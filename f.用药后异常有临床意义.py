# %%
# %run env.py

# %% editable=true slideshow={"slide_type": ""}
index = ["受试者", "受试者状态", "访视名称", "页面名称"]
sig =["病史名称", "不良事件名称", "帕金森病_TXT", "其他，请说明"]

# %% [markdown]
# # 用药后检查异常有临床意义整体情况、用药后检查异常有临床意义清单

# %% [markdown]
# ## 生命体征

# %% editable=true slideshow={"slide_type": ""}
cols = ["检查日期", "项目_TXT", "结果", "临床意义_TXT", "单位_TXT"]
VS = pd.read_excel(raw_path, sheet_name = "VS", header = 0, skiprows = [1], usecols = index + cols + sig)
VS = VS.rename(columns={
    "检查日期":"评估日期",
    "项目_TXT":"项目",
    "临床意义_TXT":"临床意义",
    "单位_TXT":"单位",
})

# %%
df = pd.concat([VS])
df = df[(df["受试者状态"] != "筛选失败") & (df["临床意义"].notna())]

# 单个受试者最早服药日期，用来判断当前检查是给药前分组，还是给药后的分组

cols1 = ["服药日期"]
cols2 = ["受试者"]

EC = pd.read_excel(raw_path, sheet_name = "EC", header = 0, skiprows = [1], usecols = cols1 + cols2)
EC = (EC.sort_values(["受试者", "服药日期"]).drop_duplicates(subset=["受试者"], keep="first"))

DS_END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = ["受试者", "页面名称", "是否完成试验_TXT"], dtype = str).fillna("")
DS_END = DS_END[DS_END["页面名称"] == "试验完成情况总结"].drop(columns = "页面名称")

RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = ["受试者", "随机号"])

df = (df.merge(EC, on = cols2, how = "left")
        .merge(DS_END, on = cols2, how = "left")
        .merge(RAND, on = "受试者", how = "left")
     )

df["分组"] = df.apply( lambda row: "给药前检查" if row["评估日期"] <= row["服药日期"] else "给药后检查", axis=1 )

# 给药前的检查数据
# 离首次给药日期最近
# 检查结果正常或异常无临床意义
pre = df.drop(columns=sig)
pre = pre[pre["分组"] == "给药前检查"]
pre = pre.sort_values(by=["受试者", "页面名称", "项目", "服药日期"])

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

pre = (
    pre.groupby(gcols, group_keys=False)
       .apply(pick_rows)
       .reset_index(drop=True)
)


prefix_map = { "病史名称": "MH:", "不良事件名称": "AE:", "帕金森病_TXT": "研究疾病:", "其他，请说明": "其他:" }
post = df[(df["分组"] == "给药后检查") & (df["临床意义"] == "异常有临床意义")].copy()
post["异常有临床意义，请描述"] = post[sig].apply( 
    lambda row: ";".join( 
        f"{prefix_map[col]}{str(val).replace('√', '帕金森病')}" 
        for col, val in row.items() 
        if pd.notna(val) and str(val).strip() != "" ), axis=1 )

post = post[["受试者", "访视名称", "页面名称", "评估日期", "项目", "结果", "临床意义", "异常有临床意义，请描述"]]

merge = pre.merge(post, on = ["受试者", "页面名称", "项目"], how = "left")
merge = merge[~((merge["结果_x"].isna()) | (merge["结果_y"].isna()))]

merge = merge.rename(columns = {
    "访视名称_x":"访视名称_首次用药前",
    "访视名称_y":"访视名称_首次用药后",
    "结果_x":"检查结果_首次用药前",
    "结果_y":"检查结果_首次用药后",
    "评估日期_x":"检查日期_首次用药前",
    "评估日期_y":"检查日期_首次用药后",
    "临床意义_x":"临床意义_首次用药前",
    "临床意义_y":"临床意义_首次用药后",
    "异常有临床意义，请描述":"异常有临床意义，请描述_首次用药后",
    "页面名称":"表单名称",
    "项目":"检查项",
    "受试者":"筛选号",
    "是否完成试验_TXT":"是否完成试验",
})

merge = merge[[
    "筛选号", 
    "随机号", 
    "表单名称", 
    "检查项", 
    "单位",
    "访视名称_首次用药前",
    "检查日期_首次用药前",
    "检查结果_首次用药前",
    "临床意义_首次用药前",
    "访视名称_首次用药后",
    "检查日期_首次用药后",
    "检查结果_首次用药后",
    "临床意义_首次用药后",
    "异常有临床意义，请描述_首次用药后",
    "是否完成试验",
]]
merge.insert(0, "No.", range(1, len(merge) + 1))
merge['temp_id_visit'] = merge['筛选号'].astype(str) + merge['表单名称'].astype(str) + merge['检查项'].astype(str) + "_" + merge['访视名称_首次用药后'].astype(str) + merge['检查结果_首次用药后'].astype(str)

# %%
VS = merge.copy()

# %%
vs_merge = merge.copy().drop(columns = ["temp_id_visit"])
file_name = f"{output_path}/listing/表39-2 生命体征用药后检查异常有临床意义清单.xlsx"
sheet_name = "表39-2 用药后检查异常有临床意义清单"
with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    # 数据从第 4 行开始写 (索引为 3)
    vs_merge.to_excel(writer, sheet_name=sheet_name, startrow=3, index=False, header=False)
    
    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # --- 格式定义 ---
    header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D3D3D3'})
    title_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
    data_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter'})

    # --- 1. 大标题 ---
    lc = len(vs_merge)
    ls = len(vs_merge.drop_duplicates(subset = "筛选号"))
    worksheet.merge_range(0, 0, 0, 15, f'表 39-2 用药后检查异常有临床意义清单 ({lc}例次{ls}例)', title_fmt)

    # --- 2. 第一层合并表头 (Row 1 & 2) ---
    # 固定列：跨两行合并 (1, col, 2, col)
    fixed_cols = ['No.', '筛选号', '随机号', '表单名称', '检查项', '单位']
    for i, col_name in enumerate(fixed_cols):
        worksheet.merge_range(1, i, 2, i, col_name, header_fmt)
    
    # 首次用药前
    worksheet.merge_range(1, 6, 1, 9, '首次用药前', header_fmt)
    worksheet.write(2, 6,  '访视名称', header_fmt)
    worksheet.write(2, 7,  '检查日期', header_fmt)
    worksheet.write(2, 8, '检查结果', header_fmt)
    worksheet.write(2, 9, '临床意义', header_fmt)
    
    
    # 首次用药后：
    worksheet.merge_range(1, 10, 1, 14, '首次用药后', header_fmt)
    worksheet.write(2, 10,  '访视名称', header_fmt)
    worksheet.write(2, 11,  '检查日期', header_fmt)
    worksheet.write(2, 12,  '检查结果', header_fmt)
    worksheet.write(2, 13,  '临床意义', header_fmt)
    worksheet.write(2, 14,  '异常有临床意义，请描述', header_fmt)
    
    # 是否完成试验 (最后一列)
    worksheet.merge_range(1, 15, 2, 15, '是否完成试验', header_fmt)

    # --- 3. 设置列宽 (根据需要调整数值) ---
    worksheet.set_column(0, 0, 5)
    worksheet.set_column(1, 2, 8)
    worksheet.set_column(3, 4, 12)
    worksheet.set_column(5, 5, 5)
    worksheet.set_column(6, 13, 18)
    worksheet.set_column(14, 14, 30)
    
    worksheet.set_column(15, 15, 14)
    
    # --- 4. 给数据区域补上边框 ---
    # 从第 4 行到数据末尾，所有列画上边框
    rows, cols = vs_merge.shape
    for r in range(rows):
        for c in range(cols):
            # 获取当前单元格的值
            val = vs_merge.iloc[r, c]
            
            # 处理空值 (NaN) 转换为 Excel 的空字符串
            if pd.isna(val):
                val = ""
            
            # 写入 Excel
            # r + 3 是因为数据从第 4 行开始写（索引 3）
            worksheet.write(r + 3, c, val, data_fmt)
print(f"处理完成，文件：{file_name}")

# %% [markdown]
# ## 体格检查

# %% editable=true slideshow={"slide_type": ""}
# 没有检查结果字段，则直接将临床意义结果视为检查结果

cols = ["检查日期", "检查部位_TXT", "临床意义_TXT", "如有异常予以描述："]
PE = pd.read_excel(raw_path, sheet_name = "PE", header = 0, skiprows = [1], usecols = index + cols + sig)
PE = PE.rename(columns={
    "检查日期":"评估日期",
    "检查部位_TXT":"项目",
    "临床意义_TXT":"临床意义",
    "如有异常予以描述：":"异常描述",
})
PE["结果"] = PE["临床意义"]

# %%
df = pd.concat([PE])
df = df[(df["受试者状态"] != "筛选失败") & (df["临床意义"].notna())]

# 单个受试者最早服药日期，用来判断当前检查是给药前分组，还是给药后的分组

cols1 = ["服药日期"]
cols2 = ["受试者"]

EC = pd.read_excel(raw_path, sheet_name = "EC", header = 0, skiprows = [1], usecols = cols1 + cols2)
EC = (EC.sort_values(["受试者", "服药日期"]).drop_duplicates(subset=["受试者"], keep="first"))

DS_END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = ["受试者", "页面名称", "是否完成试验_TXT"], dtype = str).fillna("")
DS_END = DS_END[DS_END["页面名称"] == "试验完成情况总结"].drop(columns = "页面名称")

RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = ["受试者", "随机号"])

df = (df.merge(EC, on = cols2, how = "left")
        .merge(DS_END, on = cols2, how = "left")
        .merge(RAND, on = "受试者", how = "left")
     )

df["分组"] = df.apply( lambda row: "给药前检查" if row["评估日期"] <= row["服药日期"] else "给药后检查", axis=1 )

# 给药前的检查数据
# 离首次给药日期最近
# 检查结果正常或异常无临床意义
pre = df.drop(columns=sig)
pre = pre[pre["分组"] == "给药前检查"]
pre = pre.sort_values(by=["受试者", "页面名称", "项目", "服药日期"])

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

pre = (
    pre.groupby(gcols, group_keys=False)
       .apply(pick_rows)
       .reset_index(drop=True)
)


prefix_map = { "病史名称": "MH:", "不良事件名称": "AE:", "帕金森病_TXT": "研究疾病:", "其他，请说明": "其他:" }
post = df[(df["分组"] == "给药后检查") & (df["临床意义"] == "异常有临床意义")].copy()
post["异常有临床意义，请描述"] = post[sig].apply( 
    lambda row: ";".join( 
        f"{prefix_map[col]}{str(val).replace('√', '帕金森病')}" 
        for col, val in row.items() 
        if pd.notna(val) and str(val).strip() != "" ), axis=1 )

post = post[["受试者", "访视名称", "页面名称", "评估日期", "项目", "结果", "临床意义", "异常描述", "异常有临床意义，请描述"]]

merge = pre.merge(post, on = ["受试者", "页面名称", "项目"], how = "left")
merge = merge[~((merge["结果_x"].isna()) | (merge["结果_y"].isna()))]

merge = merge.rename(columns = {
    "访视名称_x":"访视名称_首次用药前",
    "访视名称_y":"访视名称_首次用药后",
    "结果_x":"检查结果_首次用药前",
    "结果_y":"检查结果_首次用药后",
    "评估日期_x":"检查日期_首次用药前",
    "评估日期_y":"检查日期_首次用药后",
    "临床意义_x":"临床意义_首次用药前",
    "临床意义_y":"临床意义_首次用药后",
    "异常描述_x":"异常描述_首次用药前",
    "异常描述_y":"异常描述_首次用药后",
    "异常有临床意义，请描述":"异常有临床意义，请描述_首次用药后",
    "页面名称":"表单名称",
    "项目":"检查项",
    "受试者":"筛选号",
    "是否完成试验_TXT":"是否完成试验",
})

merge = merge[[
    "筛选号", 
    "随机号", 
    "表单名称", 
    "检查项",
    "访视名称_首次用药前",
    "检查日期_首次用药前",
    "检查结果_首次用药前",
    "临床意义_首次用药前",
    "异常描述_首次用药前",
    "访视名称_首次用药后",
    "检查日期_首次用药后",
    "检查结果_首次用药后",
    "临床意义_首次用药后",
    "异常描述_首次用药后",
    "异常有临床意义，请描述_首次用药后",
    "是否完成试验",
]]
merge.insert(0, "No.", range(1, len(merge) + 1))
merge['temp_id_visit'] = merge['筛选号'].astype(str) + merge['表单名称'].astype(str) + merge['检查项'].astype(str) + "_" + merge['访视名称_首次用药后'].astype(str)

# %%
PE = merge.copy()

# %%
pe_merge = merge.copy().drop(columns = ["temp_id_visit"])

file_name = f"{output_path}/listing/表39-3 体格检查用药后检查异常有临床意义清单.xlsx"
sheet_name = "表39-3 用药后检查异常有临床意义清单"
with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    # 数据从第 4 行开始写 (索引为 3)
    pe_merge.to_excel(writer, sheet_name=sheet_name, startrow=3, index=False, header=False)
    
    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # --- 格式定义 ---
    header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D3D3D3'})
    title_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
    data_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter'})

    # --- 1. 大标题 ---
    lc = len(pe_merge)
    ls = len(pe_merge.drop_duplicates(subset = "筛选号"))
    worksheet.merge_range(0, 0, 0, 17, f'表 39-3 用药后检查异常有临床意义清单 ({lc}例次{ls}例)', title_fmt)

    # --- 2. 第一层合并表头 (Row 1 & 2) ---
    # 固定列：跨两行合并 (1, col, 2, col)
    fixed_cols = ['No.', '筛选号', '随机号', '表单名称', '检查项']
    for i, col_name in enumerate(fixed_cols):
        worksheet.merge_range(1, i, 2, i, col_name, header_fmt)
    
    # 首次用药前
    worksheet.merge_range(1, 5, 1, 9, '首次用药前', header_fmt)
    worksheet.write(2, 5,  '访视名称', header_fmt)
    worksheet.write(2, 6,  '检查日期', header_fmt)
    worksheet.write(2, 7, '检查结果', header_fmt)
    worksheet.write(2, 8, '临床意义', header_fmt)
    worksheet.write(2, 9, '异常描述', header_fmt)
    
    
    # 首次用药后：
    worksheet.merge_range(1, 10, 1, 15, '首次用药后', header_fmt)
    worksheet.write(2, 10,  '访视名称', header_fmt)
    worksheet.write(2, 11,  '检查日期', header_fmt)
    worksheet.write(2, 12,  '检查结果', header_fmt)
    worksheet.write(2, 13,  '临床意义', header_fmt)
    worksheet.write(2, 14,  '异常描述', header_fmt)
    worksheet.write(2, 15,  '异常有临床意义，请描述', header_fmt)
    
    # 是否完成试验 (最后一列)
    worksheet.merge_range(1, 16, 2, 16, '是否完成试验', header_fmt)

    # --- 3. 设置列宽 (根据需要调整数值) ---
    worksheet.set_column(0, 0, 5)
    worksheet.set_column(1, 2, 8)
    worksheet.set_column(3, 4, 12)
    worksheet.set_column(5, 6, 16)
    worksheet.set_column(7, 7, 5)
    worksheet.set_column(7, 14, 18)
    worksheet.set_column(15, 15, 30)
    worksheet.set_column(16, 16, 14)
    
    # --- 4. 给数据区域补上边框 ---
    # 从第 4 行到数据末尾，所有列画上边框
    rows, cols = pe_merge.shape
    for r in range(rows):
        for c in range(cols):
            # 获取当前单元格的值
            val = pe_merge.iloc[r, c]
            
            # 处理空值 (NaN) 转换为 Excel 的空字符串
            if pd.isna(val):
                val = ""
            
            # 写入 Excel
            # r + 3 是因为数据从第 4 行开始写（索引 3）
            worksheet.write(r + 3, c, val, data_fmt)
print(f"处理完成，文件：{file_name}")

# %% [markdown]
# ## 12导联心电图

# %% editable=true slideshow={"slide_type": ""}

# 没有检查结果字段，则直接将临床意义结果视为检查结果

cols1 = ["检查日期", "临床意义_TXT", "异常描述"]
# cols2 = ["心率", "QT", "QTcF", "PR间期"]
# cols3 = ["心率_UNIT", "QT_UNIT", "QTcF_UNIT", "PR间期_UNIT"]

EG = pd.read_excel(raw_path, sheet_name = "EG", header = 0, skiprows = [1], usecols = index + cols1 + sig)

EG = EG.rename(columns={
    "检查日期":"评估日期",
    "临床意义_TXT":"临床意义",
})

EG["项目"] = EG["页面名称"]
EG["结果"] = EG["临床意义"]

# %%
df = pd.concat([EG])
df = df[(df["受试者状态"] != "筛选失败") & (df["临床意义"].notna())]

# 单个受试者最早服药日期，用来判断当前检查是给药前分组，还是给药后的分组

cols1 = ["服药日期"]
cols2 = ["受试者"]

EC = pd.read_excel(raw_path, sheet_name = "EC", header = 0, skiprows = [1], usecols = cols1 + cols2)
EC = (EC.sort_values(["受试者", "服药日期"]).drop_duplicates(subset=["受试者"], keep="first"))

DS_END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = ["受试者", "页面名称", "是否完成试验_TXT"], dtype = str).fillna("")
DS_END = DS_END[DS_END["页面名称"] == "试验完成情况总结"].drop(columns = "页面名称")

RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = ["受试者", "随机号"])

df = (df.merge(EC, on = cols2, how = "left")
        .merge(DS_END, on = cols2, how = "left")
        .merge(RAND, on = "受试者", how = "left")
     )

df["分组"] = df.apply( lambda row: "给药前检查" if row["评估日期"] <= row["服药日期"] else "给药后检查", axis=1 )

# 给药前的检查数据
# 离首次给药日期最近
# 检查结果正常或异常无临床意义
pre = df.drop(columns=sig)
pre = pre[pre["分组"] == "给药前检查"]
pre = pre.sort_values(by=["受试者", "页面名称", "项目", "服药日期"])

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

pre = (
    pre.groupby(gcols, group_keys=False)
       .apply(pick_rows)
       .reset_index(drop=True)
)


prefix_map = { "病史名称": "MH:", "不良事件名称": "AE:", "帕金森病_TXT": "研究疾病:", "其他，请说明": "其他:" }
post = df[(df["分组"] == "给药后检查") & (df["临床意义"] == "异常有临床意义")].copy()
post["异常有临床意义，请描述"] = post[sig].apply( 
    lambda row: ";".join( 
        f"{prefix_map[col]}{str(val).replace('√', '帕金森病')}" 
        for col, val in row.items() 
        if pd.notna(val) and str(val).strip() != "" ), axis=1 )

post = post[["受试者", "访视名称", "页面名称", "评估日期", "项目", "结果", "临床意义", "异常描述", "异常有临床意义，请描述"]]

merge = pre.merge(post, on = ["受试者", "页面名称", "项目"], how = "left")
merge = merge[~((merge["结果_x"].isna()) | (merge["结果_y"].isna()))]

merge = merge.rename(columns = {
    "访视名称_x":"访视名称_首次用药前",
    "访视名称_y":"访视名称_首次用药后",
    "结果_x":"检查结果_首次用药前",
    "结果_y":"检查结果_首次用药后",
    "评估日期_x":"检查日期_首次用药前",
    "评估日期_y":"检查日期_首次用药后",
    "临床意义_x":"临床意义_首次用药前",
    "临床意义_y":"临床意义_首次用药后",
    "异常描述_x":"异常描述_首次用药前",
    "异常描述_y":"异常描述_首次用药后",
    "异常有临床意义，请描述":"异常有临床意义，请描述_首次用药后",
    "页面名称":"表单名称",
    "项目":"检查项",
    "受试者":"筛选号",
    "是否完成试验_TXT":"是否完成试验",
})

merge = merge[[
    "筛选号", 
    "随机号", 
    "表单名称", 
    "检查项", 
    # "正常值范围下限", 
    # "正常值范围上限", 
    # "单位",
    "访视名称_首次用药前",
    "检查日期_首次用药前",
    "检查结果_首次用药前",
    "临床意义_首次用药前",
    "异常描述_首次用药前",
    "访视名称_首次用药后",
    "检查日期_首次用药后",
    "检查结果_首次用药后",
    "临床意义_首次用药后",
    "异常描述_首次用药后",
    "异常有临床意义，请描述_首次用药后",
    "是否完成试验",
]]
merge.insert(0, "No.", range(1, len(merge) + 1))
merge['temp_id_visit'] = merge['筛选号'].astype(str) + merge['表单名称'].astype(str) + merge['检查项'].astype(str) + "_" + merge['访视名称_首次用药后'].astype(str)

# %%
EG = merge.copy()

# %%
eg_merge = merge.copy().drop(columns = ["temp_id_visit"])
file_name = f"{output_path}/listing/表39-4 12导联心电图用药后检查异常有临床意义清单.xlsx"
sheet_name = "表39-4 用药后检查异常有临床意义清单"
with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    # 数据从第 4 行开始写 (索引为 3)
    eg_merge.to_excel(writer, sheet_name=sheet_name, startrow=3, index=False, header=False)
    
    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # --- 格式定义 ---
    header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D3D3D3'})
    title_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
    data_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter'})

    # --- 1. 大标题 ---
    lc = len(eg_merge)
    ls = len(eg_merge.drop_duplicates(subset = "筛选号"))
    worksheet.merge_range(0, 0, 0, 17, f'表 39-4 用药后检查异常有临床意义清单 ({lc}例次{ls}例)', title_fmt)

    # --- 2. 第一层合并表头 (Row 1 & 2) ---
    # 固定列：跨两行合并 (1, col, 2, col)
    fixed_cols = ['No.', '筛选号', '随机号', '表单名称', '检查项']
    for i, col_name in enumerate(fixed_cols):
        worksheet.merge_range(1, i, 2, i, col_name, header_fmt)
    
    # 首次用药前
    worksheet.merge_range(1, 5, 1, 9, '首次用药前', header_fmt)
    worksheet.write(2, 5,  '访视名称', header_fmt)
    worksheet.write(2, 6,  '检查日期', header_fmt)
    worksheet.write(2, 7, '检查结果', header_fmt)
    worksheet.write(2, 8, '临床意义', header_fmt)
    worksheet.write(2, 9, '异常描述', header_fmt)
    
    
    # 首次用药后：
    worksheet.merge_range(1, 10, 1, 15, '首次用药后', header_fmt)
    worksheet.write(2, 10,  '访视名称', header_fmt)
    worksheet.write(2, 11,  '检查日期', header_fmt)
    worksheet.write(2, 12,  '检查结果', header_fmt)
    worksheet.write(2, 13,  '临床意义', header_fmt)
    worksheet.write(2, 14,  '异常描述', header_fmt)
    worksheet.write(2, 15,  '异常有临床意义，请描述', header_fmt)
    
    # 是否完成试验 (最后一列)
    worksheet.merge_range(1, 16, 2, 16, '是否完成试验', header_fmt)

    # --- 3. 设置列宽 (根据需要调整数值) ---
    worksheet.set_column(0, 0, 5)
    worksheet.set_column(1, 2, 8)
    worksheet.set_column(3, 4, 12)
    worksheet.set_column(5, 6, 16)
    worksheet.set_column(7, 7, 5)
    worksheet.set_column(7, 14, 18)
    worksheet.set_column(15, 15, 30)
    worksheet.set_column(16, 16, 14)
    
    # --- 4. 给数据区域补上边框 ---
    # 从第 4 行到数据末尾，所有列画上边框
    rows, cols = eg_merge.shape
    for r in range(rows):
        for c in range(cols):
            # 获取当前单元格的值
            val = eg_merge.iloc[r, c]
            
            # 处理空值 (NaN) 转换为 Excel 的空字符串
            if pd.isna(val):
                val = ""
            
            # 写入 Excel
            # r + 3 是因为数据从第 4 行开始写（索引 3）
            worksheet.write(r + 3, c, val, data_fmt)
print(f"处理完成，文件：{file_name}")

# %% [markdown]
# ## 实验室检查

# %% editable=true slideshow={"slide_type": ""}
cols = ["项目.1", "测定值", "临床意义_TXT", "采样日期", "下限", "上限", "单位"]
LB_HEM = pd.read_excel(raw_path, sheet_name = "LB_HEM", header = 0, skiprows = [1], usecols = index + cols + sig)
LB_LFT = pd.read_excel(raw_path, sheet_name = "LB_LFT", header = 0, skiprows = [1], usecols = index + cols + sig)
LB_RFT = pd.read_excel(raw_path, sheet_name = "LB_RFT", header = 0, skiprows = [1], usecols = index + cols + sig)
LB_ELECT = pd.read_excel(raw_path, sheet_name = "LB_ELECT", header = 0, skiprows = [1], usecols = index + cols + sig)
LB_FBG = pd.read_excel(raw_path, sheet_name = "LB_FBG", header = 0, skiprows = [1], usecols = index + cols + sig)
LB_URI = pd.read_excel(raw_path, sheet_name = "LB_URI", header = 0, skiprows = [1], usecols = index + cols + sig)
LB_HCG1 = pd.read_excel(raw_path, sheet_name = "LB_HCG1", header = 0, skiprows = [1], usecols = index + cols + sig)

LB = pd.concat([LB_HEM, LB_LFT, LB_RFT, LB_ELECT, LB_FBG, LB_URI, LB_HCG1])
LB = LB.rename(columns={
    "临床意义_TXT":"临床意义",
    "采样日期":"评估日期",
    "项目.1":"项目",
    "测定值":"结果",
    "上限":"正常值范围上限",
    "下限":"正常值范围下限",
})

# %% editable=true slideshow={"slide_type": ""}
cols = ["尿妊娠_TXT", "临床意义_TXT", "采样日期"]
LB_HCG2 = pd.read_excel(raw_path, sheet_name = "LB_HCG2", header = 0, skiprows = [1], usecols = index + cols + sig)
LB_HCG2 = LB_HCG2.rename(columns={
    "尿妊娠_TXT":"结果",
    "临床意义_TXT":"临床意义",
    "采样日期":"评估日期",
})
LB_HCG2["项目"] = LB_HCG2["页面名称"]

# %% editable=true slideshow={"slide_type": ""}
df = pd.concat([LB, LB_HCG2])
df = df[(df["受试者状态"] != "筛选失败") & (df["临床意义"].notna())]

# 单个受试者最早服药日期，用来判断当前检查是给药前分组，还是给药后的分组

cols1 = ["服药日期"]
cols2 = ["受试者"]

EC = pd.read_excel(raw_path, sheet_name = "EC", header = 0, skiprows = [1], usecols = cols1 + cols2)
EC = (EC.sort_values(["受试者", "服药日期"]).drop_duplicates(subset=["受试者"], keep="first"))

DS_END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = ["受试者", "页面名称", "是否完成试验_TXT"], dtype = str).fillna("")
DS_END = DS_END[DS_END["页面名称"] == "试验完成情况总结"].drop(columns = "页面名称")

RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = ["受试者", "随机号"])

df = (df.merge(EC, on = cols2, how = "left")
        .merge(DS_END, on = cols2, how = "left")
        .merge(RAND, on = "受试者", how = "left")
     )

df["分组"] = df.apply( lambda row: "给药前检查" if row["评估日期"] <= row["服药日期"] else "给药后检查", axis=1 )

# 给药前的检查数据
# 离首次给药日期最近
# 检查结果正常或异常无临床意义
pre = df.drop(columns=sig)
pre = pre[pre["分组"] == "给药前检查"]
pre = pre.sort_values(by=["受试者", "页面名称", "项目", "服药日期"])

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

pre = (
    pre.groupby(gcols, group_keys=False)
       .apply(pick_rows)
       .reset_index(drop=True)
)


prefix_map = { "病史名称": "MH:", "不良事件名称": "AE:", "帕金森病_TXT": "研究疾病:", "其他，请说明": "其他:" }
post = df[(df["分组"] == "给药后检查") & (df["临床意义"] == "异常有临床意义")].copy()
post["异常有临床意义，请描述"] = post[sig].apply( 
    lambda row: ";".join( 
        f"{prefix_map[col]}{str(val).replace('√', '帕金森病')}" 
        for col, val in row.items() 
        if pd.notna(val) and str(val).strip() != "" ), axis=1 )

post = post[["受试者", "访视名称", "页面名称", "评估日期", "项目", "结果", "临床意义", "异常有临床意义，请描述"]]

merge = pre.merge(post, on = ["受试者", "页面名称", "项目"], how = "left")
merge = merge[~((merge["结果_x"].isna()) | (merge["结果_y"].isna()))]

merge = merge.rename(columns = {
    "访视名称_x":"访视名称_首次用药前",
    "访视名称_y":"访视名称_首次用药后",
    "结果_x":"检查结果_首次用药前",
    "结果_y":"检查结果_首次用药后",
    "评估日期_x":"检查日期_首次用药前",
    "评估日期_y":"检查日期_首次用药后",
    "临床意义_x":"临床意义_首次用药前",
    "临床意义_y":"临床意义_首次用药后",
    "异常有临床意义，请描述":"异常有临床意义，请描述_首次用药后",
    "页面名称":"表单名称",
    "项目":"检查项",
    "受试者":"筛选号",
    "是否完成试验_TXT":"是否完成试验",
})

merge = merge[[
    "筛选号", 
    "随机号", 
    "表单名称", 
    "检查项", 
    "正常值范围下限", 
    "正常值范围上限", 
    "单位",
    "访视名称_首次用药前",
    "检查日期_首次用药前",
    "检查结果_首次用药前",
    "临床意义_首次用药前",
    "访视名称_首次用药后",
    "检查日期_首次用药后",
    "检查结果_首次用药后",
    "临床意义_首次用药后",
    "异常有临床意义，请描述_首次用药后",
    "是否完成试验",
]]
merge.insert(0, "No.", range(1, len(merge) + 1))
merge['temp_id_visit'] = merge['筛选号'].astype(str) + merge['表单名称'].astype(str) + merge['检查项'].astype(str) + "_" + merge['访视名称_首次用药后'].astype(str)

# %%
LB = merge.copy()

# %% editable=true slideshow={"slide_type": ""}
lb_merge = merge.copy().drop(columns = ["temp_id_visit"])
file_name = f"{output_path}/listing/表39-1 实验室检查用药后检查异常有临床意义清单.xlsx"
sheet_name = "表39-1 用药后检查异常有临床意义清单"
with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    # 数据从第 4 行开始写 (索引为 3)
    lb_merge.to_excel(writer, sheet_name=sheet_name, startrow=3, index=False, header=False)
    
    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # --- 格式定义 ---
    header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D3D3D3'})
    title_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
    data_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter'})

    # --- 1. 大标题 ---
    lc = len(lb_merge)
    ls = len(lb_merge.drop_duplicates(subset = "筛选号"))
    worksheet.merge_range(0, 0, 0, 17, f'表 39-1 用药后检查异常有临床意义清单 ({lc}例次{ls}例)', title_fmt)

    # --- 2. 第一层合并表头 (Row 1 & 2) ---
    # 固定列：跨两行合并 (1, col, 2, col)
    fixed_cols = ['No.', '筛选号', '随机号', '表单名称', '检查项', '正常值范围下限', '正常值范围上限', '单位']
    for i, col_name in enumerate(fixed_cols):
        worksheet.merge_range(1, i, 2, i, col_name, header_fmt)
    
    # 首次用药前
    worksheet.merge_range(1, 8, 1, 11, '首次用药前', header_fmt)
    worksheet.write(2, 8,  '访视名称', header_fmt)
    worksheet.write(2, 9,  '检查日期', header_fmt)
    worksheet.write(2, 10, '检查结果', header_fmt)
    worksheet.write(2, 11, '临床意义', header_fmt)
    
    
    # 首次用药后：
    worksheet.merge_range(1, 12, 1, 16, '首次用药后', header_fmt)
    worksheet.write(2, 12,  '访视名称', header_fmt)
    worksheet.write(2, 13,  '检查日期', header_fmt)
    worksheet.write(2, 14,  '检查结果', header_fmt)
    worksheet.write(2, 15,  '临床意义', header_fmt)
    worksheet.write(2, 16,  '异常有临床意义，请描述', header_fmt)
    
    # 是否完成试验 (最后一列)
    worksheet.merge_range(1, 17, 2, 17, '是否完成试验', header_fmt)

    # --- 3. 设置列宽 (根据需要调整数值) ---
    worksheet.set_column(0, 0, 5)
    worksheet.set_column(1, 2, 8)
    worksheet.set_column(3, 4, 12)
    worksheet.set_column(5, 6, 16)
    worksheet.set_column(7, 7, 5)
    worksheet.set_column(7, 15, 18)
    worksheet.set_column(16, 16, 30)
    worksheet.set_column(17, 17, 14)
    
    # --- 4. 给数据区域补上边框 ---
    # 从第 4 行到数据末尾，所有列画上边框
    rows, cols = lb_merge.shape
    for r in range(rows):
        for c in range(cols):
            # 获取当前单元格的值
            val = lb_merge.iloc[r, c]
            
            # 处理空值 (NaN) 转换为 Excel 的空字符串
            if pd.isna(val):
                val = ""
            
            # 写入 Excel
            # r + 3 是因为数据从第 4 行开始写（索引 3）
            worksheet.write(r + 3, c, val, data_fmt)
print(f"处理完成，文件：{file_name}")

# %% editable=true slideshow={"slide_type": ""}
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
