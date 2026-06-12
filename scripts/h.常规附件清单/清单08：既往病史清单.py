# %%
# %run ../../env.py
from utils.loaders import load_rand

# %%
RAND = load_rand(cols=['受试者', '随机号'])

# %% [markdown]
# # 清单：既往史及现病史

# %%
cols = ["受试者", "疾病名称", "开始日期", "筛选期问询时，是否持续_TXT", "结束日期"]
MH = pd.read_excel(raw_path, sheet_name = "MH", header = 0, skiprows = [1], usecols = cols)
MH = MH[MH["疾病名称"].notna()]

MH = MH.merge(RAND, on = "受试者", how = "left")
MH.columns = [col.replace("_TXT", "") for col in MH.columns]

MH = MH.rename(columns = {
    "受试者":"筛选号",
})

stand_cols = ["筛选号", "随机号", "疾病名称", "开始日期", "筛选期问询时，是否持续", "结束日期"]
MH = MH[stand_cols]
MH.insert(0, "No.", range(1, len(MH) + 1))

lc = len(MH)

export_to_excel_with_format(
    MH,
    f"{output_path}/listing/表46 既往病史清单.xlsx",
    "表46 既往病史清单",
    f"表46 既往病史清单（{lc}例）"
)
