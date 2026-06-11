# %%
import shutil
import os

# 设置文件夹路径
folder_path = "output/20260129"
zip_file_path = 'output/20260129/output.zip'

# 检查文件夹是否存在
if os.path.exists(folder_path):
    # 创建压缩文件
    shutil.make_archive(zip_file_path.replace('.zip', ''), 'zip', folder_path)
    print(f"文件夹 {folder_path} 已成功压缩为 {zip_file_path}.")
else:
    print(f"文件夹 {folder_path} 不存在，请检查路径。")
