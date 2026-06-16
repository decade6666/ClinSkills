## 工作流程约束

1. **脚本变动前必须加载 write-script skill**：任何对 `scripts/` 下 Python 脚本的编写或修改，必须先触发 `/write-script`，按其工作流程执行，不得跳过直接动手。
2. **查询数据形状必须使用 query_metadata.py**：需要确认 Excel 表单的字段名、编码表、列名结构时，优先使用 `.claude/skills/write-script/scripts/query_metadata.py` 查询元数据 JSON；只有当元数据中找不到时，才回退到直接读 Excel 表头。
