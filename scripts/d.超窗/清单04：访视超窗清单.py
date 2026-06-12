# %%
# %run ../../env.py
from utils.loaders import load_first_dose
from utils.loaders import load_completion
from utils.loaders import load_rand

# %%
TW = pd.read_excel(timewin_path, sheet_name = "时间窗", usecols = ["类别", "访视名称", "时间窗下限", "时间窗上限"])
TW["时间窗下限"] = TW["时间窗下限"].astype("Int32")
TW["时间窗上限"] = TW["时间窗上限"].astype("Int32")

# %%
index = ["受试者", "受试者状态", "访视名称","页面名称"]

# %% [markdown]
# ## 表格：访视超窗

# %%
SV = pd.read_excel(raw_path, sheet_name = "SV", header = 0, skiprows = [1], usecols = index + ["访视日期"]).rename(columns={"访视日期":"评估日期"})
df = SV.sort_values(by = ["受试者", "访视名称", "页面名称", "评估日期"])
df = df.drop_duplicates()

df = df[(df["受试者状态"] != "筛选失败")]
df["类别"] = "其他指标超窗"

RAND = load_rand(cols=['受试者', '随机日期'])

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
EC = load_first_dose()

cols = ["受试者", "页面名称", "是否完成试验_TXT"]
DS_END = load_completion()

RAND = load_rand(cols=['受试者', '受试者状态', '随机号'])
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
