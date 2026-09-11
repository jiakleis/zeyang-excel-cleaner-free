from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

try:
    from .file_adapters import LoadedTable, UnsafeStructureError
except ImportError:  # pragma: no cover - supports direct script execution
    from file_adapters import LoadedTable, UnsafeStructureError


CONTROL_CHAR_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]")


@dataclass
class CleaningOptions:
    remove_empty_rows: bool = True
    remove_empty_columns: bool = True
    dedupe: bool = False
    missing_values: tuple[str, ...] = ()
    rename_duplicate_headers: bool = False
    clean_control_chars: bool = True


@dataclass
class CleaningStats:
    input_rows: int = 0
    output_rows: int = 0
    input_columns: int = 0
    output_columns: int = 0
    removed_empty_rows: int = 0
    removed_empty_columns: int = 0
    removed_duplicate_rows: int = 0
    modified_cells: int = 0
    missing_values_normalized: int = 0
    control_chars_removed: int = 0
    removed_column_names: list[str] = field(default_factory=list)
    header_mapping: list[tuple[str, str]] = field(default_factory=list)


@dataclass
class CleaningPlan:
    table: LoadedTable
    options: CleaningOptions
    stats: CleaningStats
    warnings: list[str] = field(default_factory=list)
    unexecuted: list[str] = field(default_factory=list)
    structural_actions_blocked: bool = False
    normalized_headers: list[Any] = field(default_factory=list)
    normalized_rows: list[list[Any]] = field(default_factory=list)
    source_row_indices: list[int] = field(default_factory=list)
    source_column_indices: list[int] = field(default_factory=list)


@dataclass
class CleaningResult:
    table: LoadedTable
    headers: list[Any]
    rows: list[list[Any]]
    stats: CleaningStats
    rules_applied: list[str]
    warnings: list[str]
    unexecuted: list[str]
    structural_changed: bool
    template_protected: bool
    source_row_indices: list[int] = field(default_factory=list)
    source_column_indices: list[int] = field(default_factory=list)
    validation_passed: bool = False


def _is_blank(value: Any) -> bool:
    return value is None or value == ""


def _value_key(value: Any) -> tuple[str, str]:
    if value is None:
        return ("none", "")
    return (type(value).__name__, str(value))


def _row_key(row: list[Any]) -> tuple[tuple[str, str], ...]:
    return tuple(_value_key(value) for value in row)


def _clean_text(value: Any, options: CleaningOptions, stats: CleaningStats) -> Any:
    if not isinstance(value, str):
        return value
    original = value
    cleaned = value.replace("\u3000", " ").strip()
    if options.clean_control_chars:
        before = cleaned
        cleaned = CONTROL_CHAR_RE.sub("", cleaned)
        stats.control_chars_removed += len(before) - len(cleaned)
    if options.missing_values and cleaned in options.missing_values:
        stats.missing_values_normalized += 1
        cleaned = None
    if cleaned != original:
        stats.modified_cells += 1
    return cleaned


def _normalized_grid(table: LoadedTable, options: CleaningOptions, stats: CleaningStats) -> tuple[list[Any], list[list[Any]]]:
    width = table.width
    headers = list(table.headers) + [None] * max(0, width - len(table.headers))
    rows = [list(row) + [None] * max(0, width - len(row)) for row in table.rows]
    normalized_headers = [_clean_text(value, options, stats) for value in headers]
    normalized_rows = [[_clean_text(value, options, stats) for value in row] for row in rows]
    return normalized_headers, normalized_rows


def _duplicate_header_names(headers: list[Any]) -> tuple[list[Any], list[tuple[str, str]]]:
    counts: dict[str, int] = {}
    mapping: list[tuple[str, str]] = []
    result: list[Any] = []
    for index, value in enumerate(headers, start=1):
        original = "" if value is None else str(value)
        base = original or f"column_{index}"
        counts[base] = counts.get(base, 0) + 1
        renamed = base if counts[base] == 1 else f"{base}_{counts[base]}"
        result.append(renamed)
        if renamed != original:
            mapping.append((original, renamed))
    return result, mapping


def _has_duplicate_headers(headers: list[Any]) -> bool:
    values = ["" if value is None else str(value) for value in headers]
    return len(values) != len(set(values))


def _empty_row(row: list[Any]) -> bool:
    return all(_is_blank(value) for value in row)


def _empty_column(rows: list[list[Any]], index: int) -> bool:
    return bool(rows) and all(_is_blank(row[index]) for row in rows)


def _column_label(value: Any, index: int) -> str:
    return str(value) if value not in (None, "") else f"第 {index + 1} 列（未命名）"


def _apply_structure(
    table: LoadedTable,
    headers: list[Any],
    rows: list[list[Any]],
    options: CleaningOptions,
    stats: CleaningStats,
    warnings: list[str],
    unexecuted: list[str],
    source_row_indices: list[int],
    source_column_indices: list[int],
) -> tuple[list[Any], list[list[Any]], bool, list[int], list[int]]:
    template_protected = table.has_template_risk
    if table.has_structural_risk:
        raise UnsafeStructureError(
            "目标工作表存在合并单元格，V1 不会对其执行结构清洗；请先转换为普通二维数据表。"
        )
    if template_protected:
        warnings.append("疑似模板/版式结构：已保留疑似布局结构，不执行删除行、删除列和重复行删除。")
        if options.remove_empty_rows:
            unexecuted.append("完全空行删除：因疑似模板/版式结构未执行")
        if options.remove_empty_columns:
            unexecuted.append("全空列删除：因疑似模板/版式结构未执行")
        if options.dedupe:
            unexecuted.append("完全重复行处理：因疑似模板/版式结构未执行")
        return headers, rows, False, source_row_indices, source_column_indices

    structural_changed = False
    if options.remove_empty_rows:
        kept = [
            (row, source_row_index)
            for row, source_row_index in zip(rows, source_row_indices)
            if not _empty_row(row)
        ]
        kept_rows = [row for row, _ in kept]
        source_row_indices = [source_row_index for _, source_row_index in kept]
        stats.removed_empty_rows = len(rows) - len(kept_rows)
        structural_changed = structural_changed or stats.removed_empty_rows > 0
        rows = kept_rows

    if options.remove_empty_columns and rows:
        empty_indices = [index for index in range(len(headers)) if _empty_column(rows, index)]
        if empty_indices:
            stats.removed_empty_columns = len(empty_indices)
            stats.removed_column_names = [_column_label(headers[index], index) for index in empty_indices]
            keep_indices = [index for index in range(len(headers)) if index not in empty_indices]
            headers = [headers[index] for index in keep_indices]
            rows = [[row[index] for index in keep_indices] for row in rows]
            source_column_indices = [source_column_indices[index] for index in keep_indices]
            structural_changed = True

    if options.dedupe:
        seen: set[tuple[tuple[str, str], ...]] = set()
        unique_rows: list[list[Any]] = []
        unique_row_indices: list[int] = []
        for row, source_row_index in zip(rows, source_row_indices):
            key = _row_key(row)
            if key in seen:
                stats.removed_duplicate_rows += 1
                continue
            seen.add(key)
            unique_rows.append(row)
            unique_row_indices.append(source_row_index)
        structural_changed = structural_changed or stats.removed_duplicate_rows > 0
        rows = unique_rows
        source_row_indices = unique_row_indices
    return headers, rows, structural_changed, source_row_indices, source_column_indices


def build_plan(table: LoadedTable, options: CleaningOptions) -> CleaningPlan:
    stats = CleaningStats(
        input_rows=len(table.rows),
        input_columns=table.width,
    )
    warnings = list(table.template_signals)
    unexecuted: list[str] = []
    headers, rows = _normalized_grid(table, options, stats)

    if len(rows) == 0 and options.remove_empty_columns:
        warnings.append("仅有表头：为保护表头，未删除空列。")
    if _has_duplicate_headers(headers):
        if options.rename_duplicate_headers:
            headers, mapping = _duplicate_header_names(headers)
            stats.header_mapping.extend(mapping)
        else:
            warnings.append("检测到重复表头；未重命名。若需要稳定列名，请确认重复表头重命名。")

    preview_headers = list(headers)
    preview_rows = [list(row) for row in rows]
    preview_row_indices = list(range(len(preview_rows)))
    preview_column_indices = list(range(len(preview_headers)))
    (
        preview_headers,
        preview_rows,
        _,
        preview_row_indices,
        preview_column_indices,
    ) = _apply_structure(
        table,
        preview_headers,
        preview_rows,
        options,
        stats,
        warnings,
        unexecuted,
        preview_row_indices,
        preview_column_indices,
    )
    if table.formula_cells and not table.has_template_risk and (
        stats.removed_empty_rows > 0
        or stats.removed_empty_columns > 0
        or stats.removed_duplicate_rows > 0
    ):
        raise UnsafeStructureError(
            "目标工作表包含公式，且本次计划实际需要改变行列结构；"
            "为避免公式引用失真，操作已停止。"
        )
    stats.output_rows = len(preview_rows)
    stats.output_columns = len(preview_headers)
    return CleaningPlan(
        table=table,
        options=options,
        stats=stats,
        warnings=list(dict.fromkeys(warnings)),
        unexecuted=unexecuted,
        structural_actions_blocked=table.has_template_risk,
        normalized_headers=preview_headers,
        normalized_rows=preview_rows,
        source_row_indices=preview_row_indices,
        source_column_indices=preview_column_indices,
    )


def execute_plan(plan: CleaningPlan) -> CleaningResult:
    stats = plan.stats
    headers = list(plan.normalized_headers)
    rows = [list(row) for row in plan.normalized_rows]
    rules = ["清理文本首尾空格和常见全角空格", "基础不可见控制字符清理"]
    if stats.removed_empty_rows:
        rules.append("删除完全空行")
    if stats.removed_empty_columns:
        rules.append("删除普通二维数据表中的严格全空列")
    if stats.removed_duplicate_rows:
        rules.append("精确重复行去重并保留第一条")
    if stats.missing_values_normalized:
        rules.append("按用户指定标记统一缺失值")
    if plan.options.rename_duplicate_headers and stats.header_mapping:
        rules.append("重复表头稳定重命名")
    if plan.structural_actions_blocked:
        rules.append("模板/版式保护")
    return CleaningResult(
        table=plan.table,
        headers=headers,
        rows=rows,
        stats=stats,
        rules_applied=rules,
        warnings=list(dict.fromkeys(plan.warnings)),
        unexecuted=list(plan.unexecuted),
        structural_changed=(
            stats.removed_empty_rows > 0
            or stats.removed_empty_columns > 0
            or stats.removed_duplicate_rows > 0
        ),
        template_protected=plan.structural_actions_blocked,
        source_row_indices=list(plan.source_row_indices),
        source_column_indices=list(plan.source_column_indices),
    )


def format_plan(plan: CleaningPlan) -> str:
    stats = plan.stats
    lines = [
        "清洗计划（尚未写入文件）",
        f"- 输入：{plan.table.source_path.name}",
        f"- 类型：{plan.table.kind.upper()}" + (f" / 工作表：{plan.table.sheet_name}" if plan.table.sheet_name else ""),
        f"- 数据规模：{stats.input_rows} 行 × {stats.input_columns} 列",
        f"- 预计输出：{stats.output_rows} 行 × {stats.output_columns} 列",
        f"- 预计删除完全空行：{stats.removed_empty_rows}",
        f"- 预计删除严格全空列：{stats.removed_empty_columns}（{', '.join(stats.removed_column_names) or '无'}）",
        f"- 预计删除重复行：{stats.removed_duplicate_rows}",
        f"- 预计修改单元格：{stats.modified_cells}",
    ]
    if stats.header_mapping:
        lines.append("- 表头映射：" + ", ".join(f"{old or '(空)'}→{new}" for old, new in stats.header_mapping))
    if plan.warnings:
        lines.append("- 警告：" + "；".join(plan.warnings))
    if plan.unexecuted:
        lines.append("- 未执行：" + "；".join(plan.unexecuted))
    lines.append("- 有损动作需在用户确认后执行；原文件不会被覆盖。")
    return "\n".join(lines)
