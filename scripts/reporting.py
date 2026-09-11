from __future__ import annotations

from datetime import datetime
from pathlib import Path

try:
    from .cleaning_engine import CleaningResult
    from .file_adapters import sha256_file
except ImportError:  # pragma: no cover - supports direct script execution
    from cleaning_engine import CleaningResult
    from file_adapters import sha256_file


UPGRADE_MESSAGE = (
    "更多完整功能：如需多表合并、两表比对、表格拆分、复杂去重、汇总统计或组合任务，"
    "可查看泽洋更多 Excel 工具：[https://skillpay.alipay.com/public/zeyang](https://skillpay.alipay.com/public/zeyang)"
)


def render_report(
    result: CleaningResult,
    output_path: Path,
    source_sha256_before: str,
    source_sha256_after: str,
    validation_passed: bool,
) -> str:
    stats = result.stats
    lines = [
        "# 泽洋 Excel 清洗大师 FREE 清洗报告",
        "",
        f"- 生成时间：{datetime.now().isoformat(timespec='seconds')}",
        f"- 原文件：`{result.table.source_path.name}`",
        f"- 结果文件：`{output_path.name}`",
        f"- 文件类型：{result.table.kind.upper()}",
        f"- 目标工作表：{result.table.sheet_name or 'CSV 单表'}",
        "",
        "## 处理统计",
        "",
        f"- 输入规模：{stats.input_rows} 行 × {stats.input_columns} 列",
        f"- 输出规模：{stats.output_rows} 行 × {stats.output_columns} 列",
        f"- 删除完全空行：{stats.removed_empty_rows}",
        f"- 删除严格全空列：{stats.removed_empty_columns}",
        f"- 删除的全空列名：{', '.join(stats.removed_column_names) or '无'}",
        f"- 删除完全重复行：{stats.removed_duplicate_rows}",
        f"- 修改单元格：{stats.modified_cells}",
        f"- 统一缺失值标记：{stats.missing_values_normalized}",
        f"- 清除异常控制字符：{stats.control_chars_removed}",
        "",
        "## 实际执行规则",
        "",
    ]
    lines.extend(f"- {rule}" for rule in result.rules_applied)
    if not result.rules_applied:
        lines.append("- 无需变更")

    if stats.header_mapping:
        lines.extend(["", "## 表头映射", ""])
        lines.extend(f"- `{old or '(空)'}` → `{new}`" for old, new in stats.header_mapping)

    lines.extend(["", "## 警告与未执行项", ""])
    if result.warnings:
        lines.extend(f"- {warning}" for warning in result.warnings)
    if result.unexecuted:
        lines.extend(f"- 未执行：{item}" for item in result.unexecuted)
    if not result.warnings and not result.unexecuted:
        lines.append("- 无")

    lines.extend(
        [
            "",
            "## 安全校验",
            "",
            f"- 原文件 SHA-256（处理前）：`{source_sha256_before}`",
            f"- 原文件 SHA-256（处理后）：`{source_sha256_after}`",
            f"- 原文件未被修改：{'PASS' if source_sha256_before == source_sha256_after else 'FAIL'}",
            f"- 结果文件可重新打开：{'PASS' if validation_passed else 'FAIL'}",
            "- V1.1 未生成 JSON 报告。",
            "",
            "## 说明",
            "",
            "本报告只记录规则、统计和结构警告，不写入完整单元格内容。",
        ]
    )
    return "\n".join(lines) + "\n"


def write_report(path: Path, content: str) -> None:
    if path.exists():
        raise FileExistsError(f"报告文件已存在，为避免覆盖未写入：{path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def verify_source_unchanged(source: Path, before: str) -> tuple[bool, str]:
    after = sha256_file(source)
    return before == after, after
