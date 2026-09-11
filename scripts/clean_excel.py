from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

try:
    from .cleaning_engine import CleaningOptions, build_plan, execute_plan, format_plan
    from .file_adapters import (
        CleanerError,
        InvalidInputError,
        load_input,
        normalize_delimiter,
        sha256_file,
        write_output,
    )
    from .reporting import UPGRADE_MESSAGE, render_report, verify_source_unchanged, write_report
except ImportError:  # pragma: no cover - supports `python scripts/clean_excel.py`
    from cleaning_engine import CleaningOptions, build_plan, execute_plan, format_plan
    from file_adapters import (
        CleanerError,
        InvalidInputError,
        load_input,
        normalize_delimiter,
        sha256_file,
        write_output,
    )
    from reporting import UPGRADE_MESSAGE, render_report, verify_source_unchanged, write_report


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="泽洋 Excel 清洗大师 FREE V1 deterministic cleaner")
    parser.add_argument("--input", required=True, help="Input .xlsx or .csv path")
    parser.add_argument("--output", help="New output file path; never overwrites an existing file")
    parser.add_argument("--report", help="Markdown report path")
    parser.add_argument("--sheet", help="Target XLSX worksheet")
    parser.add_argument(
        "--delimiter",
        type=_delimiter_arg,
        help="CSV 分隔符：comma、tab、semicolon、pipe，或 ,、\\t、;、|",
    )
    parser.add_argument("--execute", action="store_true", help="Write the result file and report")
    parser.add_argument("--yes", action="store_true", help="Confirm the displayed cleaning plan")
    parser.add_argument("--no-empty-row-removal", action="store_true", help="Keep completely empty rows")
    parser.add_argument("--keep-empty-columns", action="store_true", help="Keep strictly empty columns")
    parser.add_argument("--dedupe", action="store_true", help="Remove exact duplicate data rows, keeping the first")
    parser.add_argument(
        "--missing-values",
        "--missing-value",
        dest="missing_values",
        help="Comma-separated missing-value markers to normalize, for example N/A,NULL",
    )
    parser.add_argument(
        "--rename-duplicate-headers",
        action="store_true",
        help="Rename duplicate headers with stable suffixes",
    )
    parser.add_argument(
        "--no-control-cleaning",
        action="store_true",
        help="Do not remove non-printing control characters",
    )
    return parser


def _delimiter_arg(value: str) -> str:
    try:
        delimiter = normalize_delimiter(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(str(exc)) from exc
    assert delimiter is not None
    return delimiter


def _options(args: argparse.Namespace) -> CleaningOptions:
    missing = tuple(
        marker.strip()
        for marker in (args.missing_values or "").split(",")
        if marker.strip()
    )
    return CleaningOptions(
        remove_empty_rows=not args.no_empty_row_removal,
        remove_empty_columns=not args.keep_empty_columns,
        dedupe=args.dedupe,
        missing_values=missing,
        rename_duplicate_headers=args.rename_duplicate_headers,
        clean_control_chars=not args.no_control_cleaning,
    )


def _default_output(source: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = source.with_name(f"{source.stem}_cleaned_{timestamp}{source.suffix.lower()}")
    index = 1
    while candidate.exists():
        candidate = source.with_name(f"{source.stem}_cleaned_{timestamp}_{index}{source.suffix.lower()}")
        index += 1
    return candidate


def _resolve_paths(args: argparse.Namespace, source: Path) -> tuple[Path, Path]:
    output = Path(args.output).expanduser().resolve() if args.output else _default_output(source)
    if output == source:
        raise InvalidInputError("输出路径不能与原文件相同。")
    if output.suffix.lower() != source.suffix.lower():
        raise InvalidInputError("输出文件扩展名必须与输入文件一致。")
    report = (
        Path(args.report).expanduser().resolve()
        if args.report
        else output.with_name(f"{output.stem}_cleaning_report.md")
    )
    if report == source or report == output:
        raise InvalidInputError("报告路径不能覆盖原文件或结果文件。")
    return output, report


def _verify_output(result, output: Path) -> bool:
    reopened = load_input(output, result.table.sheet_name)
    if len(reopened.headers) != len(result.headers):
        return False
    if len(reopened.rows) != len(result.rows):
        return False
    return True


def run(args: argparse.Namespace) -> int:
    source = Path(args.input).expanduser().resolve()
    table = load_input(source, args.sheet, args.delimiter)
    options = _options(args)
    plan = build_plan(table, options)
    print(format_plan(plan))
    if not args.execute:
        print("\n状态：PLAN_ONLY；确认后请使用 --execute --yes 执行。")
        return 0
    if not args.yes:
        print("\n状态：WAITING_FOR_CONFIRMATION；有损动作必须先确认。", file=sys.stderr)
        return 3

    output, report = _resolve_paths(args, source)
    source_before = sha256_file(source)
    result = execute_plan(plan)
    try:
        write_output(
            table,
            output,
            result.headers,
            result.rows,
            result.structural_changed,
            result.source_row_indices,
            result.source_column_indices,
        )
        validation_passed = _verify_output(result, output)
        unchanged, source_after = verify_source_unchanged(source, source_before)
        if not validation_passed:
            raise InvalidInputError("结果文件重新打开校验失败，未生成成功报告。")
        if not unchanged:
            raise InvalidInputError("检测到原文件 SHA-256 发生变化，已停止报告生成。")
        result.validation_passed = True
        report_content = render_report(result, output, source_before, source_after, validation_passed)
        write_report(report, report_content)
    except Exception:
        if output.exists():
            output.unlink()
        if report.exists():
            report.unlink()
        raise

    print(f"\nRESULT_FILE={output}")
    print(f"REPORT_FILE={report}")
    print("数据清洗已完成。")
    print("UPGRADE_MESSAGE_SHOWN=YES")
    print(UPGRADE_MESSAGE)
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = _parser()
    args = parser.parse_args(argv)
    try:
        return run(args)
    except CleanerError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # pragma: no cover - defensive CLI boundary
        print(f"ERROR: 未预期错误：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
