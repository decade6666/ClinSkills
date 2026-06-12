# %%
# %run ../../env.py
from utils.loaders import load_rand

# %%
RAND = load_rand(cols=['受试者', '随机号'])

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
