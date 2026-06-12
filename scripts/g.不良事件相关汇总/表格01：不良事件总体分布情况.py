# %%
# %run ../../env.py

# %%
index = ["受试者", "记录号"]


# %% [markdown]
# ## 表格： 不良事件总体分布情况

# %%
# 定性资料汇总
def Checkbox_Summary(df: pd.DataFrame, cols: list, key: str, mark: str, cat: str):
    results = []
    for col in cols:
        target_data = df[df[col] == mark]
        ls = target_data[key].nunique()
        lc = len(target_data)

        results.append({
            "项目": cat,
            "类别": col,
            "例数": ls,
            "例次": lc
        })

    res = pd.DataFrame(results)
    return res


# %%
# 治疗措施
cols = ["未采取措施", "药物治疗", "非药物治疗", "其他"]
AE = pd.read_excel(raw_path, sheet_name = "AE", header = 0, skiprows = [1], usecols = index + cols, dtype=str)
res1 = Checkbox_Summary(AE, cols, "受试者", "Y", "是否采取治疗措施")

# %%
#严重不良事件定义
cols = ["导致死亡", "危及生命", "导致住院或延长住院时间", "永久或严重的残疾或者功能丧失", "先天性异常或者出生缺陷", "其他重要的医学事件"]
AE = pd.read_excel(raw_path, sheet_name = "AE", header = 0, skiprows = [1], usecols = index + cols, dtype=str)
res2 = Checkbox_Summary(AE, cols, "受试者", "Y", "严重不良事件定义")


# %%
# 定量资料汇总
def Radio_Summary(df: pd.DataFrame, col: str, key: str, cat: str):
    res = df.groupby(col).agg(
        例数=(key, "nunique"),
        例次=(key, "size")
    ).reset_index()

    res.rename(columns={col: "类别"}, inplace=True)
    res.insert(0, "项目", cat)

    return res

cols = [
    "试验期间是否有不良事件",
    "初始严重程度",
    "严重程度是否有变化",
    "严重程度-1",
    "严重程度-2",
    "严重程度-3",
    "对试验药物采取的措施",
    "与试验药物的关系",
    "是否符合严重不良事件定义",
    "试验结束时，转归",
    "是否因此不良事件退出试验",
    "是否为特别关注不良事件",
]
cols = [col + "_TXT" for col in cols]
AE = pd.read_excel(raw_path, sheet_name = "AE", header = 0, skiprows = [1], usecols = index + cols, dtype=str)
AE = AE[AE["试验期间是否有不良事件_TXT"] == "是"]

res3 = []
for col in cols:
    col_txt = col.replace("_TXT", "")
    res = Radio_Summary(AE, col, "受试者", col_txt)
    res3.append(res)

res3 = pd.concat(res3, ignore_index=True)

# %%
cols = [
    "试验期间是否有不良事件",
    "初始严重程度",
    "严重程度是否有变化",
    "严重程度-1",
    "严重程度-2",
    "严重程度-3",
    "对试验药物采取的措施",
    "与试验药物的关系",
    "是否符合严重不良事件定义",
    "试验结束时，转归",
    "是否因此不良事件退出试验",
    "是否为特别关注不良事件",
]

AE = pd.read_excel(raw_path, sheet_name = "AE", header = 0, skiprows = [1], usecols = index + cols, dtype=str)
AE = AE[AE["试验期间是否有不良事件"] == "1"]
ls = AE["受试者"].nunique()
lc = len(AE)

row = pd.DataFrame({
    "项目":"全部不良事件",
    "类别":"",
    "例数": [ls],
    "例次": [lc]
})

res = pd.concat([res1, res2, res3])
res = pd.concat([res, row], ignore_index=True)

schema = {
    "全部不良事件": [""],
    "初始严重程度": ["1级", "2级", "3级", "4级", "5级"],
    "严重程度是否有变化": ["是", "否"],
    "严重程度-1": ["1级", "2级", "3级", "4级", "5级"],
    "严重程度-2": ["1级", "2级", "3级", "4级", "5级"],
    "严重程度-3": ["1级", "2级", "3级", "4级", "5级"],
    "是否采取治疗措施": ["未采取措施", "药物治疗", "非药物治疗", "其他"],
    "对试验药物采取的措施": ["剂量不变", "减小剂量", "增加剂量", "暂停给药", "永久停药", "不适用"],
    "与试验药物的关系": ["肯定有关", "很可能有关", "可能有关", "可能无关", "无关"],
    "是否符合严重不良事件定义": ["是", "否"],
    "严重不良事件定义": ["导致死亡", "危及生命", "导致住院或延长住院时间", "永久或严重的残疾或者功能丧失", "先天性异常或者出生缺陷", "其他重要的医学事件"],
    "试验结束时，转归": ["已恢复/痊愈", "已恢复/痊愈后有后遗症", "好转", "持续", "死亡", "未知"],
    "是否因此不良事件退出试验": ["是", "否"],
    "是否为特别关注不良事件": ["是", "否"],
}

frame = []
for project, categories in schema.items():
    for cat in categories:
        frame.append({"项目": project, "类别": cat})

frame = pd.DataFrame(frame)
res = pd.merge(frame, res, on=["项目", "类别"], how="left")

res["例数"] = res["例数"].fillna(0).astype(int)
res["例次"] = res["例次"].fillna(0).astype(int)

notes = [
    "受试者不良事件情况以实际发生计例次和例数，不做任何规则处理；",
    "不良事件详细清单见附件“不良事件清单”。"
]
save_table_to_docx_threeline(
        res,
        f'{output_path}/table/表28 不良事件总体分布情况.docx',
        f'表28 不良事件总体分布情况',
        notes,
        row_height_cm=0.6,
        auto_width=True,
        include_notes=True,
        merge_columns=["项目"]
    )
