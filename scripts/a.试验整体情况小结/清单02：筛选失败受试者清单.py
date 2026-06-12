# %%
# %run ../../env.py
from utils.loaders import load_rand

# %% [markdown]
# ## 清单： 筛选失败受试者

# %%
cols1 = [ "不符合入选标准", "符合排除标准", "撤回知情同意", "失访，尝试联系≥3次均未成功", "其他筛选失败原因"]
index = ["受试者", "受试者状态", "受试者是否随机入组_TXT",]
subj = load_rand(cols=['受试者', '受试者状态', '受试者是否随机入组_TXT', '不符合入选标准', '符合排除标准', '撤回知情同意', '失访，尝试联系≥3次均未成功', '其他筛选失败原因'])
subj = subj.melt(id_vars=index, value_vars=cols1 , var_name='筛选失败原因', value_name='结果')
subj= subj[(subj["受试者状态"] == "筛选失败") & (subj["结果"] == '1')]

cols2 = ["受试者", "知情同意书签署日期"]
DS = pd.read_excel(raw_path, sheet_name = "DS_ICF", header = 0, skiprows = [1], usecols = cols2, dtype = str)

df = subj.merge(DS, on = "受试者", how = "left")
df = df.rename(columns = {
    "受试者":"筛选号"
})
df = df[["筛选号", "知情同意书签署日期", "筛选失败原因"]]
df.insert(0, "No.", range(1, len(df) + 1))
n = len(df)

export_to_excel_with_format(
    df,
    f"{output_path}/listing/表32 筛选失败受试者清单.xlsx",
    "表32 筛选失败受试者清单",
    f"表32 筛选失败受试者清单（{n}例）"
)
df
