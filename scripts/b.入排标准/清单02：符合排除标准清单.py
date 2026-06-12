# %%
# %run ../../env.py
from utils.loaders import load_completion
from utils.loaders import load_rand

# %% [markdown]
# ## 表16 符合排除标准清单（XXX例）

# %%
cols = ["受试者", "受试者状态","不符合的入选标准/符合的排除标准类型_TXT", "标准编号"]
IE = pd.read_excel(raw_path, sheet_name = "IE", header = 0, skiprows = [1], usecols = cols, dtype = str).fillna("")
IE = IE[(IE["标准编号"] != "") & (IE["受试者状态"] != "筛选失败")]

RAND = load_rand(cols=['受试者', '随机号', '随机日期'])

DS_END = load_completion()

df = (IE.merge(RAND, on = "受试者", how = "left")
        .merge(DS_END, on = "受试者", how = "left"))

df = df.rename(columns = {
    "受试者":"筛选号",
    "是否完成试验_TXT":"是否完成试验",
    "不符合的入选标准/符合的排除标准类型_TXT":"不符合入选标准详细描述"
})

df.insert(0, "No.", range(1, len(df) + 1))

lc = len(df)

df = df[["No.", "筛选号", "随机号", "随机日期", "是否完成试验", "不符合入选标准详细描述"]]

notes =[]

save_table_to_docx_threeline(
        df,
        f'{output_path}/table/表16 符合排除标准清单.docx',
        f'表16 符合排除标准清单（{lc}例）',
        notes,
        row_height_cm=0.6,
        auto_width=True,
        include_notes=False
    )
