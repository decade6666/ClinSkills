## 强制约束

1. **脚本变动前必须加载 write-script skill**：任何对 `scripts/` 下 Python 脚本的编写或修改，必须先触发 `/write-script`，按其工作流程执行，不得跳过直接动手。
2. **任何时候查询数据形状必须使用 query_metadata.py**：需要确认 Excel 表单的字段名、编码表、列名结构时，优先使用 `.claude/skills/write-script/scripts/query_metadata.py` 查询元数据 JSON；只有当元数据中找不到时，才回退到直接读 Excel 表头。
3. **脚本完成后必须运行验证**：任何 `scripts/` 下的 Python 脚本编写或修改完成后，必须在终端实际运行该脚本，确认无报错且输出文件已生成，才能提醒用户查看复核。如原始数据不可用，至少执行语法检查（`python -c "import ast; ast.parse(open('脚本路径').read())"`）。遇到错误且无法自行判断修复方向时，直接将完整报错输出给用户，由用户指定下一步方向。
