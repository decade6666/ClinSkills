# %%
# %run ../../env.py
from utils.loaders import load_rand

# %% [markdown]
# ## 表格： 筛选失败原因分类

# %%
subj = load_rand(cols=['受试者是否随机入组_TXT', '不符合入选标准', '符合排除标准', '撤回知情同意', '失访，尝试联系≥3次均未成功', '其他'])
subj = subj[subj["受试者是否随机入组_TXT"] == "否"]
subj = subj.drop(columns = ["受试者是否随机入组_TXT"])

# 需要统计的筛选失败原因
reasons = ["不符合入选标准", "符合排除标准", "撤回知情同意", "失访，尝试联系≥3次均未成功", "其他"]

# 按原因统计"例次"：统计该列非空的行数即可
df = pd.DataFrame({
    "筛选失败原因": reasons,
    "例次": [(subj[col] == "1").sum() for col in reasons]
})

# 加一行"合计"
total = df["例次"].sum()
df.loc[len(df)] = ["合计", total]

# 人数/例次确保为整数
df["例次"] = df["例次"].astype(int)

notes = [
    "根据筛选失败原因，拆分信息按例次计算。"
]
save_table_to_docx_threeline(
        df,
        f'{output_path}/table/表3 筛选失败原因分类.docx',
        '表3 筛选失败原因分类',
        notes,
        row_height_cm=0.6,
        auto_width=True
    )
df
