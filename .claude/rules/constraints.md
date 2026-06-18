## 强制约束

常驻不变量。详细操作以 write-script skill 为准——本文件只声明"必须做什么"，不复制"怎么做"的命令与路径，避免两处维护漂移。

1. **改脚本前先加载 write-script skill**：任何对 `04 scripts/` 下 Python 脚本的编写或修改，必须先触发 `/write-script` 并按其流程执行，不得跳过直接动手。例外：仅限单行明确拼写 / 注释修正可跳过 skill。如 skill 不可用（调用失败或被拒绝），暂停并告知用户，不得自行推进。
2. **查数据形状用 query_metadata.py**：确认字段名、编码表、列名结构时，先查元数据（`query_metadata.py`），query_metadata.py 无结果或报错时回退读 Excel 表头；字段名对不上时以 Excel 实际列名为准。用法见 skill Step 2。
3. **改脚本后必须验证**：`04 scripts/` 下脚本编写或修改后，必须实跑（数据不可用时做语法检查）确认通过，才提醒用户复核；同一错误修复尝试达 2 次仍失败时，停止并将完整报错交给用户定夺。命令与降级路径见 skill Step 5。语法地板已由 PostToolUse 钩子（`.claude/hooks/syntax_check.py`）自动执行——你的职责是实跑后核对：无 Python 报错、`03 output/` 下目标文件已生成且非空。
