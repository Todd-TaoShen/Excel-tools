"""Remove non-data rows from receivables reconciliation workbooks."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


REMOVE_MARKERS = (
    "合计",
    "记账人",
    "摘要",
    "主体账簿：北京东苑公寓有限公司-企业会计账簿",
)


def clean_workbook(input_path: Path, output_path: Path) -> int:
    if input_path.resolve() == output_path.resolve():
        raise ValueError("输出文件不能与源文件相同。")
    if not input_path.is_file():
        raise FileNotFoundError(f"找不到输入文件：{input_path}")
    if input_path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError("仅支持 .xlsx 和 .xlsm 文件。")
    if input_path.suffix.lower() != output_path.suffix.lower():
        raise ValueError("输入文件与输出文件的扩展名必须一致。")

    workbook = load_workbook(
        input_path,
        keep_vba=input_path.suffix.lower() == ".xlsm",
    )
    removed_rows = 0
    try:
        for worksheet in workbook.worksheets:
            for row_number in range(worksheet.max_row, 1, -1):
                value = worksheet.cell(row=row_number, column=1).value
                text = "" if value is None else str(value).strip()
                if not text or any(marker in text for marker in REMOVE_MARKERS):
                    worksheet.delete_rows(row_number)
                    removed_rows += 1
        output_path.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(output_path)
    finally:
        workbook.close()
    return removed_rows


def main() -> int:
    parser = argparse.ArgumentParser(description="清理应收账款核对 Excel 中的非数据行。")
    parser.add_argument("input", type=Path, nargs="?", help="待清理的 .xlsx 或 .xlsm 文件")
    parser.add_argument("output", type=Path, nargs="?", help="清理后文件的保存路径")
    parser.add_argument(
        "--search-dir",
        type=Path,
        help="未指定输入文件时，优先在此目录查找默认工作簿",
    )
    arguments = parser.parse_args()

    input_path = arguments.input
    if input_path is None:
        search_directories = []
        if arguments.search_dir is not None:
            search_directories.append(arguments.search_dir)
        script_directory = Path(__file__).resolve().parent
        search_directories.extend((script_directory, script_directory / "dist"))
        default_names = (
            "应收账款核对.xlsx",
            "应收账款核对.xlsm",
            "对应收账款核对.xlsx",
            "对应收账款核对.xlsm",
        )
        input_path = next(
            (
                directory / name
                for directory in search_directories
                for name in default_names
                if (directory / name).is_file()
            ),
            None,
        )
        if input_path is None:
            searched = "、".join(str(directory) for directory in search_directories)
            print(f"清理失败：未找到默认输入文件，已检查目录：{searched}", file=sys.stderr)
            return 1

    output_path = arguments.output
    if output_path is None:
        output_path = input_path.with_name(f"{input_path.stem}_清理后{input_path.suffix}")

    try:
        removed_rows = clean_workbook(input_path, output_path)
    except (OSError, ValueError, InvalidFileException, RuntimeError) as error:
        print(f"清理失败：{error}", file=sys.stderr)
        return 1

    print(f"清理完成：共删除 {removed_rows} 行。")
    print(f"输出文件：{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
