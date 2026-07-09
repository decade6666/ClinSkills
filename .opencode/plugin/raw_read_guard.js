// raw_read_guard.js — 禁止直接读取 01 rawdata/ 下的表格文件
// 对应 .claude/hooks/raw_read_guard.py 的 opencode 插件版

const DATA_SUFFIXES = [".xlsx", ".xls", ".csv"];
const RAW_DIR = "01 rawdata";
const DENY_REASON =
  "严禁直接读取 raw 原始数据。按项目约定（constraints.md #2）应先用 " +
  "query_metadata.py 无接触查询字段名/编码表/列名结构（用法见 write-script skill " +
  "Step 2），查不到再考虑其他方式。";

function hasSuffix(filePath) {
  const lower = filePath.toLowerCase();
  return DATA_SUFFIXES.some((s) => lower.endsWith(s));
}

function underDir(filePath, dirName) {
  // 标准化路径分隔符
  const normalized = filePath.replace(/\\/g, "/");
  return normalized.includes("/" + dirName + "/") || normalized.startsWith(dirName + "/");
}

export const RawReadGuard = async ({ project, client, $, directory, worktree }) => {
  return {
    "tool.execute.before": async (input, output) => {
      // 拦截 Read 工具读 rawdata 表格文件
      if (input.tool === "read") {
        const filePath = output.args?.filePath || "";
        if (hasSuffix(filePath) && underDir(filePath, RAW_DIR)) {
          throw new Error(DENY_REASON);
        }
      }

      // 拦截 Bash/PowerShell 命令中直接读 rawdata
      if (input.tool === "bash") {
        const cmd = output.args?.command || "";
        // 放行：query_metadata.py、04 scripts/ 下的脚本
        if (cmd.includes("query_metadata.py")) return;
        if (/python[\w.]*\s+["']?(?:04\s+)?scripts[/\\]/.test(cmd)) return;

        // 拦截：含 read_excel/load_workbook/openpyxl/raw_path 且指向 rawdata
        const rawReadRe = /read_excel|load_workbook|ExcelFile|openpyxl|raw_path/i;
        const rawPathRe = /(?:raw[/\\]|rawdata[/\\]|01\s+rawdata[/\\])[^"'\s]*\.(?:xlsx|xls|csv)/i;
        if (rawReadRe.test(cmd) || rawPathRe.test(cmd)) {
          throw new Error(DENY_REASON);
        }
      }
    },
  };
};
