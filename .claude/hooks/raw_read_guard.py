#!/usr/bin/env python
"""PreToolUse 钩子：拦截"直接读取 raw 原始数据"的探查行为，转为弹窗征求用户确认。

由 .claude/settings.json 的 PreToolUse 钩子调用（matcher: Bash|Read），读取 stdin
的工具事件 JSON。落地 constraints.md #2"查数据形状先用 query_metadata.py"的机制下限：

- 文字约束挡不住"读 Excel 零摩擦"这条惰性路径（allow 里 Read(**)、Bash(python -c *)
  都自动放行）。本钩子把这条路改成"需用户逐次批准"，去掉自动放行的惰性，又保留
  合理回退的口子（区别于硬拒绝）。

判定（命中即 ask）：
- Read 工具：file_path 落在 raw/ 下且为表格文件（.xlsx/.xls/.csv）。
- Bash 工具：命令里出现 read_excel/load_workbook/ExcelFile/openpyxl/raw_path，或
  直指 raw/*.xlsx；但**排除**两类正当用法：运行 scripts/ 下的真实脚本、调用
  query_metadata.py 本身。

输出：命中 → stdout 打印 permissionDecision=ask 的 JSON（exit 0）；未命中 → 静默放行。
"""
import json
import re
import sys
from pathlib import Path

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

PROJECT_ROOT = Path(__file__).resolve().parents[2]

# 表格类原始数据后缀
_DATA_SUFFIXES = {".xlsx", ".xls", ".csv"}

# Bash 命令中"读 raw 原始数据"的特征
_RAW_READ_RE = re.compile(r"read_excel|load_workbook|ExcelFile|openpyxl|raw_path")
_RAW_PATH_RE = re.compile(r"""raw[/\\][^"'\s]*\.(?:xlsx|xls|csv)""", re.IGNORECASE)
# 正当用法：运行 scripts/ 下真实脚本（其内部经 load_sheet 读取，属许可路径）
_RUN_SCRIPT_RE = re.compile(r"""python[\w.]*\s+["']?scripts[/\\]""")

_ASK_REASON = (
    "检测到直接读取 raw 原始数据。按项目约定（constraints.md #2）应先用 "
    "query_metadata.py 无接触查询字段名/编码表/列名结构（用法见 write-script skill "
    "Step 2），查不到才回退读 Excel。确认要直接读原始数据吗？"
)


def _ask(reason):
    print(json.dumps({
        "hookSpecificOutput": {
            "hookEventName": "PreToolUse",
            "permissionDecision": "ask",
            "permissionDecisionReason": reason,
        }
    }, ensure_ascii=False))
    return 0


def _is_raw_data_path(raw_path):
    if not raw_path:
        return False
    p = Path(raw_path)
    if p.suffix.lower() not in _DATA_SUFFIXES:
        return False
    try:
        rel = p.resolve().relative_to(PROJECT_ROOT)
    except ValueError:
        return False
    return bool(rel.parts) and rel.parts[0] == "raw"


def main():
    try:
        event = json.load(sys.stdin)
    except (json.JSONDecodeError, ValueError):
        return 0  # 非预期输入，放行

    tool_name = event.get("tool_name") or ""
    tool_input = event.get("tool_input") or {}

    if tool_name == "Read":
        if _is_raw_data_path(tool_input.get("file_path")):
            return _ask(_ASK_REASON)
        return 0

    if tool_name == "Bash":
        cmd = tool_input.get("command") or ""
        if "query_metadata.py" in cmd or _RUN_SCRIPT_RE.search(cmd):
            return 0  # 元数据工具 / 真实脚本，放行
        if _RAW_READ_RE.search(cmd) or _RAW_PATH_RE.search(cmd):
            return _ask(_ASK_REASON)
        return 0

    return 0


if __name__ == "__main__":
    sys.exit(main())
