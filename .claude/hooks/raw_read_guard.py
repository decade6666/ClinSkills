#!/usr/bin/env python
"""PreToolUse 钩子：拦截直接读取 raw 原始数据；允许读取 metadata。**全局安全**。

由 settings.json 的 PreToolUse 钩子调用（matcher: Bash|Read|PowerShell|Grep），读 stdin 的
工具事件 JSON。落地 constraints.md #2「查数据形状先用 query_metadata.py」的机制下限。

**全局安全**：本 hook 设计为可注册进用户级 `~/.claude/settings.json`（跨项目生效）而不误伤非临床
项目——只针对**确切指向 raw 原始数据**的操作判定，绝不因命令里出现裸 `read_excel(` / openpyxl 就拦。

判定：
- Read 工具：`02 metadata/` 下表格文件 → 放行；`01 rawdata/` 下 → 硬拒绝。
  项目根优先取 `CLAUDE_PROJECT_DIR`（全局安装到 `~/.claude/hooks/` 仍能定位当前项目）。
- Grep 工具：`path` 落在 `01 rawdata/` 下 → 硬拒绝（CSV 等纯文本可被 Grep 直接吐出）。
- Bash / PowerShell：按 `&&` / `||` / `;` / `|` / 换行 **分段逐条判定**，任一段命中 deny 即拒绝整条
  （消解「query_metadata.py / 04 scripts/ 子串短路整条命令」的绕过）。段内：
  (0) 每个读调用都带 `nrows≤2` 的兜底读取（constraints #2）→ 放行；
  (a) 出现字面 `01 rawdata/` 路径（任意后缀，含 cat/head 直读）→ 硬拒绝；
  (b)【仅临床项目】有实际读调用且用了 `raw_path` 配置变量 → 硬拒绝；
  否则该段若是 `query_metadata.py` / `04 scripts/` 脚本运行 → 放行。
  裸 read_excel(（无 raw 指向）一律不拦。

**已接受残留风险**（不做脆弱的脚本内容扫描）：
- `python /tmp/x.py` 内部 `read_excel('01 rawdata/...')`——命令串无字面 raw 路径时本 hook 不拦，
  依赖 Bash 未进 allow 白名单 → 提示用户兜底。
- 字符串拼接（`'01 '+'rawdata'`）规避字面路径正则——需刻意构造，分段修复已堵住命令拼接洗白。
- 朴素分段不解析引号：`echo "a; 01 rawdata/x"` 会被误分段并偏保守 deny（对 guard 可接受）。
"""

import json
import os
import re
import sys
from pathlib import Path

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

# 项目根：优先 Claude Code 注入的 CLAUDE_PROJECT_DIR（本 hook 全局安装到 ~/.claude/hooks/
# 时仍能定位当前项目）；未设置则回退本文件向上三级（本仓库自身作项目时成立）。
PROJECT_ROOT = Path(os.environ.get("CLAUDE_PROJECT_DIR") or Path(__file__).resolve().parents[2])
# 本项目是否为临床数据项目（存在 01 rawdata/）——用于门控只在临床项目才成立的启发式，
# 使本 hook 全局注册后不误伤普通数据项目（如其中恰有名为 raw_path 的变量）。
_CLINICAL_PROJECT = (PROJECT_ROOT / "01 rawdata").is_dir()

_DATA_SUFFIXES = {".xlsx", ".xls", ".csv"}

# 实际读调用：紧跟左括号，只匹配"调用"而非 import / grep / 源码里的裸标识符。
# 含 read_csv，使 CSV 的 nrows≤2 兜底与 Excel 对称。
_READ_CALL_RE = re.compile(r"(?:read_excel|read_csv|load_workbook|ExcelFile)\s*\(")
# 字面 raw 目录引用：任意后缀（含 .sas7bdat/.xpt/.sav/.txt/.json 及 cat/head 直读）。
# 限项目约定的 "01 rawdata/"（带序号+空格，足够特异），全局注册也不误伤普通 raw/ 目录。
_RAW_DIR_REF_RE = re.compile(r"""01\s+rawdata[/\\]\S""", re.IGNORECASE)
# raw_path 配置变量：仅当与实际读调用同现时才视为"读 raw"。
_RAW_PATH_VAR_RE = re.compile(r"\braw_path\b")
_RUN_SCRIPT_RE = re.compile(r"""python[\w.]*\s+["']?(?:04\s+)?scripts[/\\]""")
# constraints #2 兜底：每个读取都带 nrows≤2（含表头≤3 行）的受控读取放行。
_BOUNDED_NROWS_RE = re.compile(r"nrows\s*=\s*[012]\b")
# 复合命令分隔符：&& / || / ; / | / 换行。不解析引号（偏保守 deny 可接受）。
_CMD_SPLIT_RE = re.compile(r"(?:&&|\|\||[;|\n])")

_DENY_RAW_REASON = (
    "严禁直接读取 raw 原始数据。按项目约定（constraints.md #2）应先用 "
    "query_metadata.py 无接触查询字段名/编码表/列名结构（用法见 write-script skill "
    "Step 2），查不到再考虑其他方式。"
)


def _decide(decision, reason=None):
    """输出 hook 决策 JSON，exit 0 表示 hook 本身执行成功。"""
    output = {
        "hookSpecificOutput": {
            "hookEventName": "PreToolUse",
            "permissionDecision": decision,
        }
    }
    if reason:
        output["hookSpecificOutput"]["permissionDecisionReason"] = reason
    print(json.dumps(output, ensure_ascii=False))
    return 0


def _is_table_file(path_str):
    if not path_str:
        return False
    return Path(path_str).suffix.lower() in _DATA_SUFFIXES


def _under_dir(path_str, dir_name):
    """检查路径是否在项目根下指定目录内。"""
    if not path_str:
        return False
    p = Path(path_str)
    try:
        rel = p.resolve().relative_to(PROJECT_ROOT)
    except ValueError:
        return False
    return bool(rel.parts) and rel.parts[0] == dir_name


def _looks_like_raw_path(path_str):
    """path 是否指向 01 rawdata/（resolve 归属或字面相对路径）。"""
    if not path_str:
        return False
    if _under_dir(path_str, "01 rawdata"):
        return True
    normalized = path_str.replace("\\", "/").lstrip("./")
    return bool(re.match(r"01\s+rawdata(/|$)", normalized, re.IGNORECASE))


def _all_reads_bounded(cmd: str) -> bool:
    """命令中每个读取调用都必须配有 nrows≤2；缺一则不放行。"""
    n_reads = len(_READ_CALL_RE.findall(cmd))
    if n_reads == 0:
        return False
    n_bounded = len(_BOUNDED_NROWS_RE.findall(cmd))
    return n_bounded >= n_reads


def _segment_denies(seg: str) -> bool:
    """单段命令是否应被拒绝。

    判定顺序刻意把 raw 路径拒绝放在脚本/元数据白名单之前，避免
    `echo query_metadata.py; cat "01 rawdata/x.csv"` 同类子串短路；
    也堵住同段内 `echo query_metadata.py cat "01 rawdata/x.csv"` 的洗白。
    """
    if _all_reads_bounded(seg):
        return False  # constraints #2 兜底：nrows≤2 受控读
    if _RAW_DIR_REF_RE.search(seg):
        return True  # 字面 01 rawdata/…（任意后缀 / cat/head）
    if "query_metadata.py" in seg or _RUN_SCRIPT_RE.search(seg):
        return False  # 元数据工具 / scripts 下真实脚本（且无字面 raw 路径）
    if _CLINICAL_PROJECT and _READ_CALL_RE.search(seg) and _RAW_PATH_VAR_RE.search(seg):
        return True  # 临床项目：读调用 + raw_path 配置变量
    return False


def _command_denies(cmd: str) -> bool:
    """整条命令是否应被拒绝：分段后任一段 deny 即拒绝。"""
    segments = [s.strip() for s in _CMD_SPLIT_RE.split(cmd) if s.strip()]
    if not segments:
        return False
    return any(_segment_denies(seg) for seg in segments)


def main():
    try:
        event = json.load(sys.stdin)
    except (json.JSONDecodeError, ValueError):
        return 0

    tool_name = event.get("tool_name") or ""
    tool_input = event.get("tool_input") or {}

    if tool_name == "Read":
        file_path = tool_input.get("file_path") or ""
        if not _is_table_file(file_path):
            return 0  # 非表格文件，不干预（全后缀靠 settings deny 兜底）

        if _under_dir(file_path, "02 metadata"):
            return _decide("allow")  # metadata → 显式放行
        if _under_dir(file_path, "01 rawdata"):
            return _decide("deny", _DENY_RAW_REASON)  # raw → 硬拒绝

        return 0

    if tool_name == "Grep":
        # Grep 可直吐 CSV/文本内容；path 落在 01 rawdata/ 下一律拒绝（任意后缀）。
        path = tool_input.get("path") or ""
        if _looks_like_raw_path(path):
            return _decide("deny", _DENY_RAW_REASON)
        return 0

    if tool_name == "Bash" or tool_name == "PowerShell":
        cmd = tool_input.get("command") or ""
        if _command_denies(cmd):
            return _decide("deny", _DENY_RAW_REASON)
        return 0

    return 0


if __name__ == "__main__":
    sys.exit(main())
