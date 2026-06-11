# %%
# %run env.py

# %%
RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = ["受试者", "随机号"], dtype = str)

DS_END = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = ["受试者", "页面名称", "是否完成试验_TXT"])
DS_END = DS_END[DS_END["页面名称"] == "试验完成情况总结"]

# %%
index = ["受试者", "记录号"]

# %% [markdown]
# ## 表 7 方案禁止使用的合并药物或治疗
# #### 1. 是否为伴发事件，CRF中没有相应字段，需要进一步判断逻辑
# #### 2. 治疗天数（天）表中没有列出，但是脚注中列出了，需进一步确认？

# %%
cols = [
    "受试者",
    "药物名称（通用名）",
    "该用药开始日期",
    "该用药结束日期",
    "单次给药剂量",
    "单位_TXT",
    "其他单位，请注明",
    "给药频率_TXT",
    "其他频率，请注明",
    "给药途径_TXT",
    "其他途径，请注明",
    "病史名称",
    "不良事件名称",
    "帕金森病",
    "预防用药，请说明",
    "其他，请说明"
]
CM = pd.read_excel(raw_path, sheet_name = "CM", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
CM["用药途径"] = np.where(
    CM["给药途径_TXT"] != "其他",
    CM["给药途径_TXT"],
    CM["其他途径，请注明"],
)

CM["给药频率"] = np.where(
    CM["给药频率_TXT"] != "其他",
    CM["给药频率_TXT"],
    CM["其他频率，请注明"],
)

CM["单位"] = np.where(
    CM["单位_TXT"] != "其他",
    CM["单位_TXT"],
    CM["其他单位，请注明"],
)

CM = CM.drop(columns = [
    "单位_TXT",
    "其他单位，请注明",
    "给药频率_TXT",
    "其他频率，请注明",
    "给药途径_TXT",
    "其他途径，请注明",])


# 正则模式：只捕获“名称”那一段，日期部分用非捕获组
pattern = r"\d+(.+?)\s(?:\d{4}|UK|uk)-(?:\d{2}|UK|uk)-(?:\d{2}|UK|uk)"

# 病史名称：提取所有匹配的第 1 捕获组，用 "，" 连接，并在前面加 MH:
def extract_mh(text: str) -> str:
    matches = re.findall(pattern, text)
    if not matches:
        return ""
    # 去掉两端空白再拼接
    names = ["".join(m).strip() if isinstance(m, tuple) else str(m).strip() for m in matches]
    return "MH:" + "，".join(names)

# 不良事件名称：同理，前面加 AE:
def extract_ae(text: str) -> str:
    matches = re.findall(pattern, text)
    if not matches:
        return ""
    names = ["".join(m).strip() if isinstance(m, tuple) else str(m).strip() for m in matches]
    return "AE:" + "，".join(names)

CM["病史名称"] = CM["病史名称"].apply(extract_mh)
CM["不良事件名称"] = CM["不良事件名称"].apply(extract_ae)

# 按规则拼接为“用药原因”，分隔符为 "；"
def build_reason(row):
    parts = []

    # MH 部分
    if row["病史名称"]:
        parts.append(row["病史名称"])

    # AE 部分
    if row["不良事件名称"]:
        parts.append(row["不良事件名称"])

    # 帕金森病列不为空，则拼接固定文案“帕金森病”
    if row["帕金森病"].strip() != "":
        parts.append("帕金森病")

    # 预防用药，请说明
    if row["预防用药，请说明"].strip() != "":
        parts.append(row["预防用药，请说明"].strip())

    # 其他，请说明
    if row["其他，请说明"].strip() != "":
        parts.append(row["其他，请说明"].strip())

    return "；".join(parts)

CM["用药原因"] = CM.apply(build_reason, axis=1)

CM = CM.drop(columns=["病史名称", "不良事件名称", "帕金森病", "预防用药，请说明", "其他，请说明"])
CM["用药开始日期"] = pd.to_datetime(CM["该用药开始日期"], errors="coerce")
CM["用药结束日期"] = pd.to_datetime(CM["该用药结束日期"], errors="coerce")

CM["合并用药时长（天）"] = (CM["用药结束日期"] - CM["用药开始日期"]).dt.days + 1
CM["合并用药时长（天）"] = (CM["合并用药时长（天）"].astype("Int64").astype("string").fillna(""))

EC = pd.read_excel(raw_path, sheet_name = "EC", header = 0, skiprows = [1], usecols = ["受试者", "服药日期"], dtype = str).fillna("")
EC["服药日期"] = pd.to_datetime(EC["服药日期"], errors="coerce")
EC = (
    EC.groupby("受试者", dropna=False)["服药日期"]
      .agg(["min", "max"])
      .rename(columns={"min": "试验药物首次用药日期", "max": "试验药物末次用药日期"})
)
EC["试验药物首次用药日期"] = EC["试验药物首次用药日期"].dt.strftime("%Y-%m-%d")
EC["试验药物末次用药日期"] = EC["试验药物末次用药日期"].dt.strftime("%Y-%m-%d")


SV = pd.read_excel(raw_path, sheet_name = "SV", header = 0, skiprows = [1], usecols = ["受试者", "访视OID", "访视日期"], dtype = str).fillna("")
SV = SV[SV["访视OID"] == "V50"]

df = (CM.merge(RAND, on = "受试者", how = "left")
        .merge(EC, on = "受试者", how = "left")
        .merge(SV, on = "受试者", how = "left")
        .merge(DS_END, on = "受试者", how = "left")
     )

df["是否为伴发事件"] = "是/否"

df = df[[
 '受试者',
 '随机号',
 '试验药物首次用药日期',
 '试验药物末次用药日期',
 '药物名称（通用名）',
 '该用药开始日期',
 '该用药结束日期',
 '访视日期',
 '合并用药时长（天）',
 '单次给药剂量',
 '单位',
 '给药频率',
 '用药途径',
 '用药原因',
 '是否为伴发事件',
 '是否完成试验_TXT'
]]

df = df.rename(columns = {
    "受试者":"筛选号",
    "该用药开始日期":"开始日期",
    "该用药结束日期":"结束日期",
    "访视日期":"访视5访视日",
    "是否完成试验_TXT":"是否完成试验"})

df.insert(0, "No.", range(1, len(df) + 1))
df

notes = [
    "治疗天数（天）=试验药物末次用药日期-试验药物首次用药日期+1；",
    "合并用药时长（天）=结束日期-开始日期+1。"
]

# save_table_to_docx_threeline(
#         df, 
#         f'{output_path}/table/表7 方案禁止使用的合并药物或治疗.docx', 
#         '表7 方案禁止使用的合并药物或治疗', 
#         notes,
#         row_height_cm=0.6,
#         auto_width=True
#     )

# %% [markdown]
# ## 表 8 方案禁止使用的合并药物或治疗
# #### 1. CRF中没有明确的“因疗效不佳而发生的提前终止治疗”选项，只有“受试者自觉疗效不佳”和“研究者认为受试者继续用药不能获益”，是否均考虑在内
# #### 2. 需确认是否仅考虑治疗结束页面？
# #### 3. 脚注使用表8脚注的第一个?

# %%
DS_END1 = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = ["受试者", "页面名称", "受试者自行退出原因_TXT"], dtype = str).fillna("")
DS_END1 = DS_END1[(DS_END1["页面名称"] == "治疗结束页") & (DS_END1["受试者自行退出原因_TXT"] == "受试者自觉疗效不佳")]  #如果需要考虑研究者判断，这里需要补充条件
DS_END1 = DS_END1.drop(columns = ["页面名称"], axis = 1)

RAND = pd.read_excel(raw_path, sheet_name = "DS_RAND", header = 0, skiprows = [1], usecols = ["受试者", "随机号"], dtype = str).fillna("")

EC = pd.read_excel(raw_path, sheet_name = "EC", header = 0, skiprows = [1], usecols = ["受试者", "服药日期"], dtype = str).fillna("")
EC["服药日期"] = pd.to_datetime(EC["服药日期"], errors="coerce")
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

DS_END2 = pd.read_excel(raw_path, sheet_name = "DS_END", header = 0, skiprows = [1], usecols = ["受试者", "页面名称", "是否完成试验_TXT"], dtype = str).fillna("")
DS_END2 = DS_END2[DS_END2["页面名称"] == "试验完成情况总结"]
DS_END2 = DS_END2.drop(columns = ["页面名称"], axis = 1)

SV = pd.read_excel(raw_path, sheet_name = "SV", header = 0, skiprows = [1], usecols = ["受试者", "访视OID", "访视日期"], dtype = str).fillna("")
SV = SV[SV["访视OID"] == "V50"]

df = (DS_END1.merge(RAND, on = "受试者", how = "left")
        .merge(EC, on = "受试者", how = "left")
        .merge(DS_END2, on = "受试者", how = "left")
     )

df = df.rename(columns = {
    "受试者":"筛选号",
    "受试者自行退出原因_TXT":"提前终止治疗的原因",
    "是否完成试验_TXT":"是否完成试验"})

df = df[[
 '筛选号',
 '随机号',
 '首次用药日期',
 '末次用药日期',
 '治疗天数（天）',
 '提前终止治疗的原因',
 '是否完成试验'
]]

df.insert(0, "No.", range(1, len(df) + 1))

notes = [
    "治疗天数（天）=试验药物末次用药日期-试验药物首次用药日期+1；"
]

# save_table_to_docx_threeline(
#         df, 
#         f'{output_path}/table/表7 方案禁止使用的合并药物或治疗.docx', 
#         '表7 方案禁止使用的合并药物或治疗', 
#         notes,
#         row_height_cm=0.6,
#         auto_width=True
#     )

# %% [markdown]
# # 方案禁止的合并用药清单（XXX例次 XXX例）

# %%
# 先按通用名查找，再按照概括性术语查找

# 通用名列表
drug_names = [
    "米安色林", "米氮平", "奈法唑酮", "赛庚啶", "氟伏沙明", 
    "伊曲康唑", "酮康唑", "克拉霉素", "茚地那韦", "利福平", 
    "卡马西平", "苯妥英", "圣约翰草", "莫达非尼", "甲硫哒嗪", 
    "依非韦伦", "萘夫西林", "苯扎托品", "比哌立登", "苯海索", 
    "奎尼丁", "丙吡胺", "普鲁卡因胺", "普罗帕酮", "胺碘酮", 
    "索他洛尔", "伊布利特", "多非利特", "西沙必利", "红霉素", 
    "阿奇霉素", "莫西沙星", "酮康唑", "氟康唑", "伊曲康唑", 
    "卤泛群", "奎宁", "氯喹", "吩噻嗪", "甲硫哒嗪", "氯丙嗪", 
    "米索哒嗪", "奥氮平", "氟哌利多", "氟哌啶醇", "洛哌丁胺", 
    "阿米替林", "西酞普兰", "氟西汀", "丙咪嗪", "洛夫帕明", 
    "非尔氨酯", "磷苯妥英钠", "氟烷", "多潘立酮", "昂丹司琼", 
    "格拉司琼", "苯海拉明", "阿司咪唑", "特非那定", "美沙酮", 
    "可卡因", "三氧化二砷", "膦甲酸", "索拉非尼", "舒尼替尼",
    "丁酰苯", "氟喹诺酮"
]

# 概括性术语列表
general_terms = [
    "抗精神病药", "抗心律失常药", "促胃动力药", "抗菌素", 
    "抗疟疾药", "抗抑郁药", "抗惊厥药", "麻醉药", 
    "止吐药", "抗组胺药", "强CYP3A4抑制剂", 
    "强CYP3A4诱导剂", "中CYP3A4诱导剂", "中枢性抗胆碱能药",
    "抗逆转录病毒药", "蛋白激酶抑制剂", "5-羟色胺受体拮抗"
]

cols = ["药物名称（通用名）", "首选名称", "药物名称", "ATC1术语", "ATC2术语", "ATC3术语", "ATC4术语"]
df = pd.read_excel(raw_path, sheet_name = "CM", header = 0, skiprows = [1], dtype = str).fillna("")
pattern = "|".join(map(re.escape, drug_names + general_terms))
mask = df[cols].astype(str).apply(lambda s: s.str.contains(pattern, regex=True)).any(axis=1)

CM = df[mask]
cols = [
    "受试者",
    "药物名称（通用名）",
    "该用药开始日期",
    "试验结束时，是否持续_TXT",
    "该用药结束日期",
    "病史名称",
    "不良事件名称",
    "帕金森病",
    "预防用药，请说明",
    "其他，请说明"
]
CM = CM[cols]

pattern = r"\d+(.+?)\s(?:\d{4}|UK|uk)-(?:\d{2}|UK|uk)-(?:\d{2}|UK|uk)"

# 病史名称：提取所有匹配的第 1 捕获组，用 "，" 连接，并在前面加 MH:
def extract_mh(text: str) -> str:
    matches = re.findall(pattern, text)
    if not matches:
        return ""
    # 去掉两端空白再拼接
    names = ["".join(m).strip() if isinstance(m, tuple) else str(m).strip() for m in matches]
    return "MH:" + "，".join(names)

# 不良事件名称：同理，前面加 AE:
def extract_ae(text: str) -> str:
    matches = re.findall(pattern, text)
    if not matches:
        return ""
    names = ["".join(m).strip() if isinstance(m, tuple) else str(m).strip() for m in matches]
    return "AE:" + "，".join(names)

CM["病史名称"] = CM["病史名称"].apply(extract_mh)
CM["不良事件名称"] = CM["不良事件名称"].apply(extract_ae)

# 按规则拼接为“用药原因”，分隔符为 "；"
def build_reason(row):
    parts = []
    if row["病史名称"]:
        parts.append(row["病史名称"])

    if row["不良事件名称"]:
        parts.append(row["不良事件名称"])

    if row["帕金森病"].strip() != "":
        parts.append("帕金森病")

    if row["预防用药，请说明"].strip() != "":
        parts.append(row["预防用药，请说明"].strip())

    if row["其他，请说明"].strip() != "":
        parts.append(row["其他，请说明"].strip())

    return "；".join(parts)

CM["用药原因"] = CM.apply(build_reason, axis=1)

CM = CM.drop(columns=["病史名称", "不良事件名称", "帕金森病", "预防用药，请说明", "其他，请说明"])

# %%
EC = pd.read_excel(raw_path, sheet_name = "EC", header = 0, skiprows = [1], usecols = ["受试者", "服药日期"], dtype = str).fillna("")
EC["服药日期"] = pd.to_datetime(EC["服药日期"], errors="coerce")
EC = (
    EC.groupby("受试者", dropna=False)["服药日期"]
      .agg(["min", "max"])
      .rename(columns={"min": "首次用药日期", "max": "末次用药日期"})
)
EC["首次用药日期"] = EC["首次用药日期"].dt.strftime("%Y-%m-%d")
EC["末次用药日期"] = EC["末次用药日期"].dt.strftime("%Y-%m-%d")

SV = pd.read_excel(raw_path, sheet_name = "SV", header = 0, skiprows = [1], usecols = ["受试者", "访视OID", "访视日期", "访视名称"], dtype = str).fillna("")
SV = SV[SV["访视OID"] == "V20"]
SV["访视日期"] = pd.to_datetime(SV["访视日期"], errors="coerce")


df = (CM.merge(RAND, on = "受试者", how = "left")
        .merge(EC, on = "受试者", how = "left")
        .merge(DS_END, on = "受试者", how = "left")
        .merge(SV, on = "受试者", how = "left")
     )

# 计算与试验用药品同期使用时长（天）
def handle_uk_dates(date_str):
    if isinstance(date_str, str) and 'uk' in date_str:
        return pd.NaT

    return pd.to_datetime(date_str, errors='coerce')

# 应用函数来处理 "uk" 日期
df['首次用药日期_temp'] = df['首次用药日期'].apply(handle_uk_dates)
df['末次用药日期_temp'] = df['末次用药日期'].apply(handle_uk_dates)

df['用药开始日期_temp'] = df['该用药开始日期'].apply(handle_uk_dates)
df['用药结束日期_temp'] = df['该用药结束日期'].apply(handle_uk_dates)

df["合并用药时长（天）"] = (df["用药结束日期_temp"] - df["用药开始日期_temp"]).dt.days + 1
df["合并用药时长（天）"] = (df["合并用药时长（天）"].astype("Int64").astype("string").fillna(""))

df["基线期在合并用药结束日期后多少天"] = (df["访视日期"] - df["用药结束日期_temp"]).dt.days + 1
df["基线期在合并用药结束日期后多少天"] = (df["基线期在合并用药结束日期后多少天"].astype("Int64").astype("string").fillna(""))

# 定义一个函数计算重叠的天数
def calculate_overlap(row):
    overlap_start = max(row['首次用药日期_temp'], row['用药开始日期_temp'])
    overlap_end = min(row['末次用药日期_temp'], row['用药结束日期_temp'])
    
    if overlap_start <= overlap_end:
        return (overlap_end - overlap_start).days + 1
    else:
        return 0

# 应用函数到每一行
df['与试验用药品同期使用时长（天）'] = df.apply(calculate_overlap, axis=1)

df = df.rename(columns = {
    "受试者":"筛选号",
    "试验结束时，是否持续_TXT":"试验结束时，是否持续",
    "是否完成试验_TXT":"是否完成试验"})

df = df[[
 '筛选号',
 '随机号',
 '首次用药日期',
 '末次用药日期',
 '药物名称（通用名）',
 '用药原因',
 '该用药开始日期',
 '试验结束时，是否持续',
 '该用药结束日期',
 '合并用药时长（天）',
 '与试验用药品同期使用时长（天）',
 '是否完成试验',
 '基线期在合并用药结束日期后多少天'
]]

df.insert(0, "No.", range(1, len(df) + 1))

notes = [
    """患者自签署知情同意书后，禁止服用下列药物：
      （1）试验期间，禁止服用抗精神病药物，在基线（D1）之前必须停止使用不少于5个药物半衰期；
      （2）在基线（D1）之前至少21天停止使用以下5-羟色胺受体拮抗药物：米安色林、米氮平、奈法唑酮、赛庚啶、氟伏沙明和其他研究药物；
      （3）在基线（D1）之前至少14天停止使用以下药物：任何强CYP3A4抑制剂（伊曲康唑、酮康唑、克拉霉素、茚地那韦等）或强CYP3A4诱导剂（利福平、卡马西平、苯妥英、圣约翰草等）或中CYP3A4诱导剂（莫达非尼、甲硫哒嗪、依非韦伦、萘夫西林等）；
      （4）试验期间，禁止使用中枢性抗胆碱能药物，且在基线（D1）前必须停止使用不少于2周。这些药物包括，但不限于：苯扎托品、比哌立登和苯海索；
      （5）禁止使用延长QT间期的药物，包括但不限于以下内容：
          · 抗心律失常药物：包括奎尼丁、丙吡胺、普鲁卡因胺、普罗帕酮、胺碘酮、索他洛尔、伊布利特、多非利特；
          · 促胃动力药物：西沙必利；
          · 抗菌素：包括红霉素、克拉霉素、阿奇霉素、氟喹诺酮类（莫西沙星等）、酮康唑、氟康唑、伊曲康唑等；
          · 抗疟疾药物：卤泛群、奎宁、氯喹等；
          · 抗精神病药物：吩噻嗪类（甲硫哒嗪、氯丙嗪、米索哒嗪），奥氮平，丁酰苯类（氟哌利多、氟哌啶醇），洛哌丁胺等；
          · 抗抑郁药物：阿米替林、西酞普兰、氟西汀、丙咪嗪、洛夫帕明等；
          · 抗惊厥药物：包括非尔氨酯，磷苯妥英钠；
          · 麻醉药：氟烷；
          · 止吐药：多潘立酮、昂丹司琼、格拉司琼；
          · 抗组胺药：包括苯海拉明、阿司咪唑、特非那定；
          · 其他：包括美沙酮、可卡因、三氧化二砷，抗逆转录病毒药物（如膦甲酸）、甘草制剂、蛋白激酶抑制剂（索拉非尼、舒尼替尼）。""",
    "治疗天数（天）= 末次用药日期-首次用药日期+1；",
    "合并用药时长（天）= 该用药结束日期-该用药开始日期+1；",
    "与试验用药品同期使用时长（天）= 试验用药品与禁用药物同时使用的天数。",
]

lc = len(df)
ls = len(df.drop_duplicates(subset = "筛选号"))

save_table_to_docx_threeline(
        df, 
        f'{output_path}/table/表24 方案禁止使用的合并药物或治疗.docx', 
        f'表24 方案禁止使用的合并药物或治疗（{lc}例次{ls}例）', 
        notes,
        row_height_cm=0.6,
        auto_width=True
    )

# %%
