# %%
# %run ../../env.py
from utils.loaders import load_completion
from utils.loaders import load_rand

# %%
RAND = load_rand(cols=['受试者', '随机号'])

DS_END = load_completion()

# %%
index = ["受试者", "受试者状态", "访视名称","页面名称"]

# %% [markdown]
# # 访视缺失

# %%
cols = ["访视OID","页面名称", "访视日期"]
SV = pd.read_excel(raw_path, sheet_name = "SV", header = 0, skiprows = [1], usecols = index + cols).rename(columns={"访视日期":"评估日期"})
df = SV.sort_values(by = ["受试者", "访视名称", "页面名称", "评估日期"])
df = df.drop_duplicates()

df = df[(df["受试者状态"] != "筛选失败") & (df["访视OID"] != "V90")]
df["类别"] = "访视缺失"

df = (df.merge(RAND, on = "受试者", how = "left")
        .merge(DS_END, on = "受试者", how = "left")
     )

# %% [markdown]
# ## 表格：访视缺失

# %%
visit_missing1 = df.copy()
visit_missing1["提前退出"] = np.where(
    visit_missing1["是否完成试验_TXT"] == "否",
    "是",
    "否"
)

visit_missing1["是否进行"] = np.where(
    visit_missing1["评估日期"].isna(),
    "否",
    "是"
)
visit_missing1 = visit_missing1.sort_values(by = ["受试者", "访视OID"])
visit_missing1 = (visit_missing1.groupby("受试者", group_keys=False).filter(lambda x: x["是否进行"].nunique() > 1)
)
visit_missing1 = visit_missing1.pivot(index = ["受试者", "随机号", "是否完成试验_TXT"], columns = "访视名称", values = "是否进行").reset_index()

visit_missing1["随机号"] = visit_missing1["随机号"].astype(int)
visit_missing1 = visit_missing1.fillna("未激活")

visit_missing1 = visit_missing1.rename(columns = {
    "受试者":"筛选号",
    "是否完成试验_TXT":"是否完成试验"
})

visit = [
    "筛选号",
    "随机号",
    "筛选期（V1，D-15~-13）",
    "基线期（V2，D1）",
    "访视3（V3，D15±3）",
    "访视4（V4，D29±3）",
    "访视5（V5，D43±3）",
    "访视6（V6，D71±3）",
    "提前退出",
    "是否完成试验"
]

visit_missing1 = visit_missing1[visit]
visit_missing1.insert(0, "No.", range(1, len(visit_missing1) + 1))

n = len(visit_missing1.drop_duplicates(subset = ["筛选号"]))

notes = [
    "访视列：“是”代表进行本次访视，“否”代表已激活但未进行本次访视，“未激活”代表该访视未激活"
]

save_table_to_docx_threeline(
        visit_missing1,
        f'{output_path}/table/表22 访视缺失清单.docx',
        f'表22 访视缺失清单（{n}例）',
        notes,
        row_height_cm=0.6,
        auto_width=True,
        include_notes=True,
    )
