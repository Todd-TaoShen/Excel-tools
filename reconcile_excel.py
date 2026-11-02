"""Compare journal tax amounts with standard taxable amounts in an Excel file."""

from __future__ import annotations

import argparse
import calendar
from datetime import date, datetime
from concurrent.futures import Future, ThreadPoolExecutor
import os
import re
import sys
from time import monotonic
from collections import Counter, defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from clean_receivable import REMOVE_MARKERS
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException


SYNONYM_SHEET = "同义词映射表"
TAX_SHEET = "税额"
RESULT_SHEET = "比对结果"
CENT = Decimal("0.01")
MULTIPLE_RELATIVE_TOLERANCE = Decimal("0.001")
RECEIVABLE_SHEET = "应收实收比对"
DATE_RANGE_PATTERN = re.compile(
    r"(?<!\d)(?P<start_year>\d{4})[./-](?P<start_month>\d{1,2})[./-](?P<start_day>\d{1,2})"
    r"\s*[-~至—–]+\s*"
    r"(?:(?P<end_year>\d{4})[./-])?(?P<end_month>\d{1,2})[./-](?P<end_day>\d{1,2})(?!\d)"
)
DATE_RANGE_CANDIDATE_PATTERN = re.compile(
    r"(?<!\d)\d{3,4}[./-]\d{1,2}[./-]\d{1,2}"
    r"\s*[-~至—–]+\s*(?:\d{3,4}[./-])?\d{1,2}[./-]\d{1,2}(?!\d)"
)
SINGLE_DATE_PATTERN = re.compile(
    r"(?P<year>\d{4})[./-](?P<month>\d{1,2})[./-](?P<day>\d{1,2})"
)
WORKBOOK_ERRORS = (
    OSError,
    ValueError,
    RuntimeError,
    InvalidFileException,
    BadZipFile,
    ParseError,
    KeyError,
)


def keep_clean_receivable_row(name: Any, amount: Any) -> bool:
    name_text = str(name).strip() if name is not None else ""
    return (
        bool(name_text)
        and not any(marker in name_text for marker in REMOVE_MARKERS)
        and amount not in (None, "")
    )


@dataclass
class Entry:
    row: int
    text: str
    amount: Decimal | None
    buyer: str | None
    category: str | None


def normalize(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[\W_]+", "", str(value).casefold(), flags=re.UNICODE)


def parse_amount(value: Any) -> Decimal | None:
    if value is None or value == "":
        return None
    try:
        amount = Decimal(str(value).replace(",", ""))
        if not amount.is_finite():
            return None
        return amount.quantize(CENT, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return None


def add_alias(aliases: dict[str, set[str]], name: Any, standard: Any) -> None:
    key = normalize(name)
    if key and standard is not None and str(standard).strip():
        aliases.setdefault(key, set()).add(str(standard).strip())


def find_names(text: Any, aliases: dict[str, set[str]]) -> set[str]:
    """Find the longest non-overlapping synonym terms in a journal memo."""
    normalized = normalize(text)
    matches = []
    for term in aliases:
        start = normalized.find(term)
        while start >= 0:
            matches.append((start, start + len(term), term))
            start = normalized.find(term, start + 1)

    matches.sort(key=lambda match: (-(match[1] - match[0]), match[0]))
    selected: list[tuple[int, int, str]] = []
    found: set[str] = set()
    for start, end, term in matches:
        if any(start < other_end and other_start < end for other_start, other_end, _ in selected):
            continue
        selected.append((start, end, term))
        found.update(aliases[term])
    return found


def unique_name(text: Any, aliases: dict[str, set[str]]) -> str | None:
    found = find_names(text, aliases)
    return next(iter(found)) if len(found) == 1 else None


def load_synonyms(
    worksheet: Any,
) -> tuple[dict[str, set[str]], dict[str, set[str]]]:
    expected = ("费用类别简称", "费用类别别名", "购买方名称", "缩写")
    actual = tuple(worksheet.cell(1, column).value for column in range(1, 5))
    if actual != expected:
        raise ValueError(
            f"工作表“{SYNONYM_SHEET}”第 1 行表头应为：{'、'.join(expected)}"
        )

    categories: dict[str, set[str]] = {}
    buyers: dict[str, set[str]] = {}
    for row in range(2, worksheet.max_row + 1):
        category, alias, buyer, abbreviation = (
            worksheet.cell(row, column).value for column in range(1, 5)
        )
        if category is not None:
            add_alias(categories, category, category)
            add_alias(categories, alias, category)
        if buyer is not None:
            add_alias(buyers, buyer, buyer)
            add_alias(buyers, abbreviation, buyer)

    if not categories or not buyers:
        raise ValueError(f"工作表“{SYNONYM_SHEET}”缺少费用类别或购买方映射数据。")
    return categories, buyers


def read_tax_entries(
    worksheet: Any,
    categories: dict[str, set[str]],
    buyers: dict[str, set[str]],
) -> tuple[list[Entry], list[Entry]]:
    expected = ("记账摘要", "记账税额", "购买方名称", "标准费用类别", "应税税额")
    actual = tuple(worksheet.cell(1, column).value for column in range(1, 6))
    if actual != expected:
        raise ValueError(
            f"工作表“{TAX_SHEET}”第 1 行表头应为：{'、'.join(expected)}"
        )

    journal: list[Entry] = []
    taxable: list[Entry] = []
    for row in range(2, worksheet.max_row + 1):
        memo, journal_amount, buyer, category, taxable_amount = (
            worksheet.cell(row, column).value for column in range(1, 6)
        )

        memo_text = str(memo).strip() if memo is not None else ""
        if memo_text or journal_amount is not None:
            found_buyers = find_names(memo_text, buyers)
            found_categories = find_names(memo_text, categories)
            journal.append(
                Entry(
                    row,
                    memo_text,
                    parse_amount(journal_amount),
                    next(iter(found_buyers)) if len(found_buyers) == 1 else None,
                    (
                        next(iter(found_categories))
                        if len(found_categories) == 1
                        else None
                    ),
                )
            )

        if buyer is not None or category is not None or taxable_amount is not None:
            buyer_text = str(buyer).strip() if buyer is not None else ""
            category_text = str(category).strip() if category is not None else ""
            taxable.append(
                Entry(
                    row,
                    f"{buyer_text}｜{category_text}",
                    parse_amount(taxable_amount),
                    unique_name(buyer_text, buyers),
                    unique_name(category_text, categories),
                )
            )
    return journal, taxable


def output_row(
    status: str,
    journal: list[Entry],
    taxable: Entry | None,
    reason: str,
) -> list[Any]:
    journal_total = None
    if journal and all(entry.amount is not None for entry in journal):
        journal_total = sum(
            (entry.amount for entry in journal if entry.amount is not None),
            Decimal("0.00"),
        ).quantize(CENT, rounding=ROUND_HALF_UP)
    taxable_amount = taxable.amount if taxable else None
    difference = (
        journal_total - taxable_amount
        if journal_total is not None and taxable_amount is not None
        else ""
    )
    buyer = (
        journal[0].buyer
        if journal and journal[0].buyer
        else taxable.buyer if taxable else ""
    )
    category = (
        journal[0].category
        if journal and journal[0].category
        else taxable.category if taxable else ""
    )
    return [
        status,
        "、".join(str(entry.row) for entry in journal),
        "；".join(entry.text for entry in journal),
        journal_total if journal_total is not None else "",
        buyer or (taxable.text.split("｜", 1)[0] if taxable else ""),
        category or (taxable.text.split("｜", 1)[-1] if taxable else ""),
        taxable.row if taxable else "",
        taxable_amount if taxable_amount is not None else "",
        difference,
        reason,
    ]


def reconcile(journal: list[Entry], taxable: list[Entry]) -> list[list[Any]]:
    journal_groups: dict[tuple[str, str], list[Entry]] = defaultdict(list)
    taxable_groups: dict[tuple[str, str], list[Entry]] = defaultdict(list)
    results: list[list[Any]] = []

    for entry in journal:
        if entry.buyer and entry.category:
            journal_groups[(entry.buyer, entry.category)].append(entry)
        else:
            results.append(
                output_row(
                    "不匹配", [entry], None, "无法从记账摘要中唯一识别购买方或费用类别"
                )
            )

    for entry in taxable:
        if entry.buyer and entry.category:
            taxable_groups[(entry.buyer, entry.category)].append(entry)
        else:
            results.append(
                output_row(
                    "不匹配", [], entry, "无法从标准费用类别记录中唯一识别购买方或费用类别"
                )
            )

    for key in sorted(set(journal_groups) | set(taxable_groups)):
        source_rows = journal_groups.get(key, [])
        target_rows = taxable_groups.get(key, [])
        used_sources: set[int] = set()
        used_targets: set[int] = set()
        source_by_amount: dict[Decimal, list[int]] = defaultdict(list)
        target_by_amount: dict[Decimal, list[int]] = defaultdict(list)
        for index, entry in enumerate(source_rows):
            if entry.amount is not None:
                source_by_amount[entry.amount].append(index)
        for index, entry in enumerate(target_rows):
            if entry.amount is not None:
                target_by_amount[entry.amount].append(index)

        for amount in source_by_amount.keys() & target_by_amount.keys():
            source_indices = source_by_amount[amount]
            target_indices = target_by_amount[amount]
            for source_index, target_index in zip(source_indices, target_indices):
                source = source_rows[source_index]
                target = target_rows[target_index]
                used_sources.add(source_index)
                used_targets.add(target_index)
                results.append(
                    output_row(
                        "完全匹配",
                        [source],
                        target,
                        "购买方、费用类别和税额相同，一对一对应",
                    )
                )

        remaining_sources = [
            entry for index, entry in enumerate(source_rows) if index not in used_sources
        ]
        remaining_targets = [
            entry for index, entry in enumerate(target_rows) if index not in used_targets
        ]
        if (
            len(remaining_sources) > 1
            and len(remaining_targets) == 1
            and all(entry.amount is not None for entry in remaining_sources)
            and remaining_targets[0].amount is not None
        ):
            source_total = sum(
                (entry.amount for entry in remaining_sources if entry.amount is not None),
                Decimal("0.00"),
            ).quantize(CENT, rounding=ROUND_HALF_UP)
            if source_total == remaining_targets[0].amount:
                results.append(
                    output_row(
                        "部分匹配",
                        remaining_sources,
                        remaining_targets[0],
                        "多条记账摘要税额合计与唯一应税税额一致",
                    )
                )
                used_sources.update(range(len(source_rows)))
                used_targets.update(range(len(target_rows)))

        for index, source in enumerate(source_rows):
            if index not in used_sources:
                reason = (
                    "记账税额为空或不是有效数字"
                    if source.amount is None
                    else "没有唯一且税额一致的应税记录"
                )
                results.append(output_row("不匹配", [source], None, reason))
        for index, target in enumerate(target_rows):
            if index not in used_targets:
                reason = (
                    "应税税额为空或不是有效数字"
                    if target.amount is None
                    else "没有唯一且税额一致的记账摘要"
                )
                results.append(output_row("不匹配", [], target, reason))
    return results


def process_workbook(input_path: Path, output_path: Path | None = None) -> Path:
    input_path = input_path.expanduser().resolve()
    if not input_path.is_file():
        raise FileNotFoundError(f"找不到输入文件：{input_path}")
    if input_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("仅支持 .xlsx 和 .xlsm 工作簿。")

    if output_path is None:
        output_path = input_path.with_name(f"{input_path.stem}_比对结果{input_path.suffix}")
    else:
        output_path = output_path.expanduser().resolve()
    if output_path == input_path:
        raise ValueError("输出文件不能与输入文件相同。")
    if output_path.suffix.lower() != input_path.suffix.lower():
        raise ValueError("输出文件扩展名必须与输入文件一致。")
    if output_path.exists() and output_path == input_path.with_name(
        f"{input_path.stem}_比对结果{input_path.suffix}"
    ):
        base = output_path
        suffix = 2
        while output_path.exists():
            output_path = base.with_name(
                f"{base.stem}_{suffix}{base.suffix}"
            )
            suffix += 1

    keep_vba = input_path.suffix.lower() == ".xlsm"
    data_book = load_workbook(input_path, data_only=True)
    output_book = load_workbook(input_path, keep_vba=keep_vba)
    try:
        if SYNONYM_SHEET not in data_book.sheetnames or TAX_SHEET not in data_book.sheetnames:
            raise ValueError(
                f"工作簿必须包含“{SYNONYM_SHEET}”和“{TAX_SHEET}”两个工作表。"
            )
        categories, buyers = load_synonyms(data_book[SYNONYM_SHEET])
        journal, taxable = read_tax_entries(data_book[TAX_SHEET], categories, buyers)
        if not journal and not taxable:
            raise ValueError(f"工作表“{TAX_SHEET}”中没有可比对的数据。")

        result_name = RESULT_SHEET
        suffix = 2
        while result_name in output_book.sheetnames:
            result_name = f"{RESULT_SHEET}_{suffix}"
            suffix += 1
        sheet = output_book.create_sheet(result_name)
        sheet.append(
            [
                "匹配结果",
                "记账摘要行号",
                "记账摘要",
                "记账税额合计",
                "购买方名称",
                "标准费用类别",
                "应税记录行号",
                "应税税额",
                "税额差异",
                "比对说明",
            ]
        )
        results = reconcile(journal, taxable)
        results.sort(key=lambda row: (not bool(str(row[5]).strip()), str(row[5]).strip()))
        for row in results:
            sheet.append(row)

        header_fill = PatternFill("solid", fgColor="1F4E78")
        result_fills = {
            "完全匹配": PatternFill("solid", fgColor="E2F0D9"),
            "部分匹配": PatternFill("solid", fgColor="FFF2CC"),
            "不匹配": PatternFill("solid", fgColor="FCE4D6"),
        }
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
        for row in range(2, sheet.max_row + 1):
            for cell in sheet[row]:
                cell.fill = result_fills[sheet.cell(row, 1).value]
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            for column in (4, 8, 9):
                sheet.cell(row, column).number_format = "0.00"
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column, width in enumerate(
            [14, 16, 48, 16, 36, 22, 14, 14, 14, 42], start=1
        ):
            sheet.column_dimensions[get_column_letter(column)].width = width

        output_book.save(output_path)
        counts = Counter(row[0] for row in results)
        print(
            f"比对完成：{output_path}\n"
            f"完全匹配 {counts['完全匹配']} 条；"
            f"部分匹配 {counts['部分匹配']} 条；"
            f"不匹配 {counts['不匹配']} 条。"
        )
        return output_path
    finally:
        data_book.close()
        output_book.close()


def read_excel_headers(path: Path, sheet_name: str) -> list[str]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"文件中不存在工作表“{sheet_name}”。")
        worksheet = workbook[sheet_name]
        headers = [
            str(cell.value).strip() if cell.value is not None else ""
            for cell in worksheet[1]
        ]
        while headers and not headers[-1]:
            headers.pop()
        if not headers:
            raise ValueError(f"工作表“{sheet_name}”的第一行没有表头。")
        return headers
    finally:
        workbook.close()


def workbook_sheets(path: Path) -> list[str]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()


def room_key(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (int, float, Decimal)):
        try:
            number = Decimal(str(value))
            if number == number.to_integral_value():
                return str(number.quantize(Decimal("1")))
        except InvalidOperation:
            pass
    return normalize(value)


def extract_date_period(value: Any) -> tuple[date, date] | None:
    text = str(value) if value is not None else ""
    for match in DATE_RANGE_PATTERN.finditer(text):
        start_year = int(match.group("start_year"))
        start_month = int(match.group("start_month"))
        start_day = int(match.group("start_day"))
        end_month = int(match.group("end_month"))
        end_day = int(match.group("end_day"))
        end_year_text = match.group("end_year")
        end_year = int(end_year_text) if end_year_text else start_year
        if not end_year_text and (end_year, end_month, end_day) < (
            start_year,
            start_month,
            start_day,
        ):
            end_year += 1
        try:
            start_date = date(start_year, start_month, start_day)
            end_date = date(end_year, end_month, end_day)
        except ValueError:
            continue
        if end_date < start_date:
            continue
        return start_date, end_date
    return None


def parse_date_period(value: Any) -> tuple[date, date] | None:
    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        return value, value
    text = str(value).strip() if value is not None else ""
    if not text:
        return None
    date_range = extract_date_period(text)
    if date_range:
        return date_range
    match = DATE_RANGE_PATTERN.search(text)
    if match:
        start_year = int(match.group("start_year"))
        start_month = int(match.group("start_month"))
        start_day = int(match.group("start_day"))
        end_month = int(match.group("end_month"))
        end_day = int(match.group("end_day"))
        end_year_text = match.group("end_year")
        end_year = int(end_year_text) if end_year_text else start_year
        if not end_year_text and (end_year, end_month, end_day) < (
            start_year,
            start_month,
            start_day,
        ):
            end_year += 1
        if end_month == 2 and end_day == 29 and not calendar.isleap(end_year):
            end_day = calendar.monthrange(end_year, end_month)[1]
            try:
                start_date = date(start_year, start_month, start_day)
                end_date = date(end_year, end_month, end_day)
            except ValueError:
                return None
            if end_date >= start_date:
                return start_date, end_date
    try:
        single_date = date.fromisoformat(text)
    except ValueError:
        return None
    return single_date, single_date


def parse_start_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip() if value is not None else ""
    month_match = re.fullmatch(r"(\d{4})[./-](\d{1,2})", text)
    if month_match:
        try:
            return date(int(month_match.group(1)), int(month_match.group(2)), 1)
        except ValueError:
            return None
    month_match = re.fullmatch(r"(\d{4})年(\d{1,2})月", text)
    if month_match:
        try:
            return date(int(month_match.group(1)), int(month_match.group(2)), 1)
        except ValueError:
            return None
    match = SINGLE_DATE_PATTERN.fullmatch(text)
    if not match:
        return None
    try:
        return date(
            int(match.group("year")),
            int(match.group("month")),
            int(match.group("day")),
        )
    except ValueError:
        return None


def periods_overlap(
    first: tuple[date, date], second: tuple[date, date]
) -> bool:
    return first[0] <= second[1] and first[1] >= second[0]


def parse_comparison_period(
    start_text: str, end_text: str
) -> tuple[date, date] | None:
    start_value = start_text.strip()
    end_value = end_text.strip()
    if not start_value and not end_value:
        return None
    if not start_value or not end_value:
        raise ValueError("请同时填写核对时间段的开始日期和结束日期。")
    try:
        start_date = date.fromisoformat(start_value)
        end_date = date.fromisoformat(end_value)
    except ValueError as exc:
        raise ValueError("日期格式应为 YYYY-MM-DD，例如 2026-10-01。") from exc
    if end_date < start_date:
        raise ValueError("核对时间段的结束日期不能早于开始日期。")
    return start_date, end_date


@dataclass(frozen=True)
class RoomMapping:
    company: str
    room_display: str
    rent_free: bool
    lease_period: tuple[date, date] | None
    start_date: date | None = None
    active: bool = True
    start_date_source: Any = None


def load_room_mapping(
    path: Path,
    sheet_name: str,
    room_header: str,
    company_header: str,
    rent_free_header: str,
    lease_header: str,
    comparison_period: tuple[date, date] | None = None,
    start_date_header: str | None = None,
) -> dict[str, list[RoomMapping]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"映射文件中不存在工作表“{sheet_name}”。")
        worksheet = workbook[sheet_name]
        headers = [
            str(cell.value).strip() if cell.value is not None else ""
            for cell in worksheet[1]
        ]
        try:
            room_column = headers.index(room_header) + 1
            company_column = headers.index(company_header) + 1
            rent_free_column = headers.index(rent_free_header) + 1
            lease_column = headers.index(lease_header) + 1
            start_date_column = (
                headers.index(start_date_header) + 1
                if start_date_header
                else None
            )
        except ValueError as exc:
            raise ValueError(
                "映射表所选的房间号、承租方、免租期、实际租期或起租日列不在表头中，请重新选择。"
            ) from exc

        mapping: dict[str, list[RoomMapping]] = defaultdict(list)
        company_names: dict[str, str] = {}
        for values in worksheet.iter_rows(min_row=2, values_only=True):
            room_display_value = values[room_column - 1] if len(values) >= room_column else None
            company_value = values[company_column - 1] if len(values) >= company_column else None
            rent_free_value = values[rent_free_column - 1] if len(values) >= rent_free_column else None
            lease_value = values[lease_column - 1] if len(values) >= lease_column else None
            start_date_value = (
                values[start_date_column - 1]
                if start_date_column and len(values) >= start_date_column
                else None
            )
            room = room_key(room_display_value)
            company = str(company_value).strip() if company_value is not None else ""
            if not room or not company:
                continue
            lease_period = parse_date_period(lease_value)
            active = comparison_period is None or (
                lease_period is not None
                and periods_overlap(lease_period, comparison_period)
            )
            company = company_names.setdefault(normalize(company), company)
            item = RoomMapping(
                company,
                str(room_display_value).strip(),
                str(rent_free_value).strip() == "有",
                lease_period,
                parse_start_date(start_date_value),
                active,
                start_date_value,
            )
            if item not in mapping[room]:
                mapping[room].append(item)
        if start_date_header:
            for room, room_mappings in mapping.items():
                active_mappings = [item for item in room_mappings if item.active]
                dated_mappings = [
                    item for item in active_mappings if item.start_date is not None
                ]
                if dated_mappings:
                    latest_start_date = max(
                        item.start_date
                        for item in dated_mappings
                        if item.start_date is not None
                    )
                    latest_mappings = [
                        item
                        for item in dated_mappings
                        if item.start_date == latest_start_date
                    ]
                    if len(latest_mappings) != 1:
                        raise ValueError(
                            f"房间号“{latest_mappings[0].room_display}”有多条记录使用相同的最新起租日"
                            f"（{latest_start_date:%Y-%m-%d}），无法唯一确定映射；请检查映射表。"
                        )
                    selected_mapping = latest_mappings[0]
                elif len(active_mappings) > 1:
                    raise ValueError(
                        f"房间号“{active_mappings[0].room_display}”有多条有效记录，但起租日均为空或无法解析，"
                        "无法确定唯一承租方；请检查映射表。"
                    )
                elif active_mappings:
                    selected_mapping = active_mappings[0]
                else:
                    continue
                mapping[room] = [
                    item
                    if not item.active or item == selected_mapping
                    else RoomMapping(
                        item.company,
                        item.room_display,
                        item.rent_free,
                        item.lease_period,
                        item.start_date,
                        False,
                        item.start_date_source,
                    )
                    for item in room_mappings
                ]
        if not mapping:
            if comparison_period:
                raise ValueError(
                    "映射表中没有有效的房间号和公司全称对应数据。"
                )
            raise ValueError("映射表中没有有效的房间号和公司全称对应数据。")
        free_statuses: dict[tuple[str, str], set[bool]] = defaultdict(set)
        for room, mappings in mapping.items():
            for item in mappings:
                if item.active:
                    free_statuses[(room, normalize(item.company))].add(item.rent_free)
        conflicting_statuses = [
            key for key, statuses in free_statuses.items() if len(statuses) > 1
        ]
        if conflicting_statuses:
            raise ValueError(
                "映射表中有同一房间和公司在核对时段内同时标记“有”和“无”免租期的记录，请先核实。"
            )
        return dict(mapping)
    finally:
        workbook.close()


def resolve_company(
    value: Any, mapping: dict[str, list[RoomMapping]]
) -> tuple[str, str, bool]:
    text = str(value).strip() if value is not None else ""
    normalized_text = normalize(text)
    if not normalized_text:
        return "", "", False

    direct_matches = [
        item.company
        for room_mappings in mapping.values()
        for item in room_mappings
        if item.active and normalize(item.company) and normalize(item.company) in normalized_text
    ]
    direct_company = ""
    if direct_matches:
        longest_company = max(len(normalize(company)) for company in direct_matches)
        companies = {
            company
            for company in direct_matches
            if len(normalize(company)) == longest_company
        }
        if len(companies) == 1:
            direct_company = next(iter(companies))
    room_keys = [
        key
        for key, room_mappings in mapping.items()
        if key
        and (
            bool(re.search(rf"(?<!\d){re.escape(key)}(?!\d)", normalized_text))
            if key.isdigit()
            else key in normalized_text
        )
    ]
    longest_room_key = max(room_keys, key=len) if room_keys else ""
    room_display = next(
        (
            item.room_display
            for item in mapping.get(longest_room_key, [])
            if item.room_display
        ),
        longest_room_key,
    )
    room_matches = [
        (key, item)
        for key, room_mappings in mapping.items()
        if key
        and (
            bool(re.search(rf"(?<!\d){re.escape(key)}(?!\d)", normalized_text))
            if key.isdigit()
            else key in normalized_text
        )
        for item in room_mappings
        if item.active
    ]
    room = room_display if longest_room_key else ""
    room_company = ""
    if room_matches:
        longest = max(len(key) for key, _ in room_matches)
        longest_matches = [
            match for match in room_matches if len(match[0]) == longest
        ]
        room_displays = {item.room_display for _, item in longest_matches}
        if len(room_displays) == 1:
            room = next(iter(room_displays))
        companies = {item.company for _, item in longest_matches}
        if len(companies) == 1:
            room_company = next(iter(companies))
        room_rent_free = all(item.rent_free for _, item in longest_matches)
    else:
        room_rent_free = False
    if room_matches:
        if room_company:
            rent_free = all(
                item.rent_free
                for key, item in room_matches
                if len(key) == max(len(match[0]) for match in room_matches)
                and item.company == room_company
            )
            return room_company, room, rent_free
        return "", room, room_rent_free
    if direct_company:
        return direct_company, "", False
    return "", room, room_rent_free


@dataclass
class PaymentRecord:
    excel_row: int
    raw_name: Any
    raw_amount: Any
    company: str
    room: str
    amount: Decimal | None
    tax_status: str
    factor: Decimal
    date_period: tuple[date, date] | None = None
    rounding_tolerance: Decimal = Decimal("0")


def filter_records_by_period(
    records: list[PaymentRecord],
    comparison_period: tuple[date, date],
) -> list[PaymentRecord]:
    selected: list[PaymentRecord] = []
    for record in records:
        if record.date_period is None and DATE_RANGE_CANDIDATE_PATTERN.search(
            str(record.raw_name or "")
        ):
            continue
        summary_period = record.date_period or comparison_period
        if summary_period[0] <= comparison_period[1] and summary_period[1] >= comparison_period[0]:
            selected.append(record)
    return selected


def load_payment_records(
    path: Path,
    sheet_name: str,
    name_header: str,
    amount_header: str,
    mapping: dict[str, list[RoomMapping]],
    factor: Decimal,
    apply_rent_free_adjustment: bool = False,
    include_non_rent: bool = False,
    include_unresolved_room: bool = False,
    clean_receivable: bool = False,
) -> list[PaymentRecord]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"文件中不存在工作表“{sheet_name}”。")
        worksheet = workbook[sheet_name]
        headers = [
            str(cell.value).strip() if cell.value is not None else ""
            for cell in worksheet[1]
        ]
        try:
            name_column = headers.index(name_header) + 1
            amount_column = headers.index(amount_header) + 1
        except ValueError as exc:
            raise ValueError("购买方名称列或金额列不在表头中，请重新选择。") from exc

        records: list[PaymentRecord] = []
        for row, values in enumerate(
            worksheet.iter_rows(min_row=2, values_only=True), start=2
        ):
            name = values[name_column - 1] if len(values) >= name_column else None
            raw_amount = values[amount_column - 1] if len(values) >= amount_column else None
            if clean_receivable and not keep_clean_receivable_row(name, raw_amount):
                continue
            if (name is None or str(name).strip() == "") and (
                raw_amount is None or raw_amount == ""
            ):
                continue
            if raw_amount is None or raw_amount == "":
                continue
            if not include_non_rent and "租金" not in str(name or ""):
                continue
            company, room, rent_free = resolve_company(name, mapping)
            if not room and not include_unresolved_room:
                continue
            amount = parse_amount(raw_amount)
            effective_factor = factor
            rounding_tolerance = Decimal("0")
            tax_status = (
                "税后" if factor > Decimal("1")
                else "税前" if factor < Decimal("1")
                else "原金额"
            )
            if apply_rent_free_adjustment and rent_free:
                effective_factor *= Decimal(11) / Decimal(12)
                tax_status += "；免租期调整 11/12"
            if amount is not None and abs(effective_factor) > Decimal("1"):
                source_amount = Decimal(str(raw_amount).replace(",", ""))
                source_quantum = min(
                    Decimal(1).scaleb(source_amount.as_tuple().exponent),
                    CENT,
                )
                rounding_tolerance = (
                    source_quantum * abs(effective_factor) / Decimal("2")
                )
            converted = (
                (amount * effective_factor).quantize(CENT, rounding=ROUND_HALF_UP)
                if amount is not None
                else None
            )
            records.append(
                PaymentRecord(
                    row,
                    name,
                    raw_amount,
                    company,
                    room,
                    converted,
                    tax_status,
                    effective_factor,
                    extract_date_period(name),
                    rounding_tolerance,
                )
            )
        return records
    finally:
        workbook.close()


def select_actual_rent_candidates(
    receivables: list[PaymentRecord],
    actuals: list[PaymentRecord],
) -> list[PaymentRecord]:
    selected: list[PaymentRecord] = []
    for actual in actuals:
        if "应收" in str(actual.raw_name or ""):
            continue
        if "租金" in str(actual.raw_name or ""):
            selected.append(actual)
            continue
        for receivable in receivables:
            if (
                not receivable.company
                or receivable.company != actual.company
                or not receivable.room
                or receivable.room != actual.room
                or (
                    receivable.date_period is not None
                    and actual.date_period is not None
                    and not periods_overlap(
                        receivable.date_period, actual.date_period
                    )
                )
            ):
                continue
            if amounts_match_with_rounding(
                receivable, actual
            ) or multiple_relation(
                receivable.amount,
                actual.amount,
                receivable.rounding_tolerance,
                actual.rounding_tolerance,
                allow_half_multiple="免租期调整 11/12" in actual.tax_status,
            ):
                selected.append(actual)
                break
    return selected


def multiple_relation(
    receivable: Decimal | None,
    actual: Decimal | None,
    receivable_tolerance: Decimal = Decimal("0"),
    actual_tolerance: Decimal = Decimal("0"),
    allow_half_multiple: bool = False,
) -> str | None:
    if receivable is None or actual is None or receivable == 0 or actual == 0:
        return None
    if receivable * actual < 0:
        return None
    receivable_abs = abs(receivable)
    actual_abs = abs(actual)
    larger, smaller, larger_tolerance, smaller_tolerance = (
        (actual_abs, receivable_abs, actual_tolerance, receivable_tolerance)
        if actual_abs > receivable_abs
        else (receivable_abs, actual_abs, receivable_tolerance, actual_tolerance)
    )
    ratio = (larger / smaller)
    if allow_half_multiple:
        ratio = (
            (ratio * 2).to_integral_value(rounding=ROUND_HALF_UP) / 2
        )
    else:
        ratio = ratio.to_integral_value(rounding=ROUND_HALF_UP)
    difference = abs(larger - smaller * ratio)
    rounding_tolerance = (
        CENT * (ratio + 1) / 2
        + larger_tolerance
        + smaller_tolerance * ratio
        + smaller * ratio * MULTIPLE_RELATIVE_TOLERANCE
    )
    if ratio > 1 and difference <= rounding_tolerance:
        ratio_text = f"{ratio:g}"
        return (
            f"实收是应收的 {ratio_text} 倍"
            if actual_abs > receivable_abs
            else f"应收是实收的 {ratio_text} 倍"
        )
    return None


def amounts_match_with_rounding(
    receivable: PaymentRecord, actual: PaymentRecord
) -> bool:
    if receivable.amount is None or actual.amount is None:
        return False
    tolerance = CENT + receivable.rounding_tolerance + actual.rounding_tolerance
    return abs(receivable.amount - actual.amount) <= tolerance


def reconcile_payments(
    receivables: list[PaymentRecord], actuals: list[PaymentRecord]
) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    used_receivables: set[int] = set()
    used_actuals: set[int] = set()
    companies = sorted(
        {record.company for record in receivables + actuals if record.company}
    )

    for company in companies:
        r_indices = [
            index for index, record in enumerate(receivables)
            if record.company == company and record.amount is not None
        ]
        a_indices = [
            index for index, record in enumerate(actuals)
            if record.company == company and record.amount is not None
        ]
        exact_edges = [
            (ri, ai)
            for ri in r_indices
            for ai in a_indices
            if amounts_match_with_rounding(receivables[ri], actuals[ai])
        ]
        r_edge_counts = Counter(ri for ri, _ in exact_edges)
        a_edge_counts = Counter(ai for _, ai in exact_edges)
        exact_pairs = [
            (ri, ai)
            for ri, ai in exact_edges
            if r_edge_counts[ri] == 1 and a_edge_counts[ai] == 1
        ]
        for ri, ai in exact_pairs:
            used_receivables.add(ri)
            used_actuals.add(ai)
            results.append(
                payment_result(
                    "完全匹配", receivables[ri], actuals[ai], "唯一的一对一匹配"
                )
            )

        ambiguous_exact_edges = [
            (ri, ai)
            for ri, ai in exact_edges
            if (ri, ai) not in exact_pairs
            and ri not in used_receivables
            and ai not in used_actuals
        ]
        for ri, ai in ambiguous_exact_edges:
            details = []
            if r_edge_counts[ri] > 1:
                details.append(
                    f"该应收记录有 {r_edge_counts[ri]} 条金额在舍入容差内的实收候选"
                )
            if a_edge_counts[ai] > 1:
                details.append(
                    f"该实收记录有 {a_edge_counts[ai]} 条金额在舍入容差内的应收候选"
                )
            results.append(
                payment_result(
                    "待确认",
                    receivables[ri],
                    actuals[ai],
                    "；".join(details) + "，请结合租赁期间或付款日期确认",
                )
            )
            used_receivables.add(ri)
            used_actuals.add(ai)

        remaining_r = [i for i in r_indices if i not in used_receivables]
        remaining_a = [i for i in a_indices if i not in used_actuals]
        edges = [
            (
                ri,
                ai,
                multiple_relation(
                    receivables[ri].amount,
                    actuals[ai].amount,
                    receivables[ri].rounding_tolerance,
                    actuals[ai].rounding_tolerance,
                    allow_half_multiple=(
                        "免租期调整 11/12" in actuals[ai].tax_status
                    ),
                ),
            )
            for ri in remaining_r
            for ai in remaining_a
        ]
        edges = [edge for edge in edges if edge[2] is not None]
        r_edge_counts = Counter(ri for ri, _, _ in edges)
        a_edge_counts = Counter(ai for _, ai, _ in edges)
        ambiguous_edges: list[tuple[int, int, str]] = []
        for ri, ai, relation in edges:
            if r_edge_counts[ri] == 1 and a_edge_counts[ai] == 1:
                used_receivables.add(ri)
                used_actuals.add(ai)
                results.append(
                    payment_result(
                        "条件匹配", receivables[ri], actuals[ai], relation or ""
                    )
                )
            else:
                ambiguous_edges.append((ri, ai, relation or ""))

        for ri, ai, relation in ambiguous_edges:
            source_candidates = r_edge_counts[ri]
            target_candidates = a_edge_counts[ai]
            details = []
            if source_candidates > 1:
                details.append(f"该应收记录有 {source_candidates} 条实收候选")
            if target_candidates > 1:
                details.append(f"该实收记录有 {target_candidates} 条应收候选")
            results.append(
                payment_result(
                    "待确认",
                    receivables[ri],
                    actuals[ai],
                    "；".join(details) + "，请结合租赁期间或付款日期确认",
                    relation=relation,
                )
            )
            used_receivables.add(ri)
            used_actuals.add(ai)

    unresolved_company_edges = [
        (
            ri,
            ai,
            multiple_relation(
                receivable.amount,
                actual.amount,
                receivable.rounding_tolerance,
                actual.rounding_tolerance,
                allow_half_multiple=(
                    "免租期调整 11/12" in actual.tax_status
                ),
            ),
        )
        for ri, receivable in enumerate(receivables)
        for ai, actual in enumerate(actuals)
        if ri not in used_receivables
        and ai not in used_actuals
        and receivable.room
        and receivable.room == actual.room
        and bool(receivable.company) != bool(actual.company)
    ]
    unresolved_company_edges = [
        edge for edge in unresolved_company_edges if edge[2] is not None
    ]
    r_edge_counts = Counter(ri for ri, _, _ in unresolved_company_edges)
    a_edge_counts = Counter(ai for _, ai, _ in unresolved_company_edges)
    for ri, ai, relation in unresolved_company_edges:
        details = [
            "公司未能唯一识别，按相同房间号和整数倍金额列为候选"
        ]
        if r_edge_counts[ri] > 1:
            details.append(f"该应收记录有 {r_edge_counts[ri]} 条实收候选")
        if a_edge_counts[ai] > 1:
            details.append(f"该实收记录有 {a_edge_counts[ai]} 条应收候选")
        results.append(
            payment_result(
                "待确认",
                receivables[ri],
                actuals[ai],
                "；".join(details) + "，请人工确认",
                relation=relation or "",
            )
        )
        used_receivables.add(ri)
        used_actuals.add(ai)

    for index, record in enumerate(receivables):
        if index not in used_receivables:
            reason = (
                "无法唯一识别购买方公司，可能是同一房间号对应多个历史租户"
                if not record.company
                else "金额为空或不是有效数字"
                if record.amount is None
                else "没有唯一且金额相同或成倍数的实收记录"
            )
            results.append(payment_result("未匹配", record, None, reason))
    for index, record in enumerate(actuals):
        if index not in used_actuals:
            reason = (
                "无法唯一识别购买方公司，可能是同一房间号对应多个历史租户"
                if not record.company
                else "金额为空或不是有效数字"
                if record.amount is None
                else "没有唯一且金额相同或成倍数的应收记录"
            )
            results.append(payment_result("未匹配", None, record, reason))
    return results


def payment_result(
    status: str,
    receivable: PaymentRecord | None,
    actual: PaymentRecord | None,
    note: str,
    relation: str = "",
) -> dict[str, Any]:
    return {
        "结果": status,
        "应收原始行号": receivable.excel_row if receivable else "",
        "应收原始购买方": receivable.raw_name if receivable else "",
        "应收原始金额": receivable.raw_amount if receivable else "",
        "应收换算金额": receivable.amount if receivable and receivable.amount is not None else "",
        "应收税前/税后": receivable.tax_status if receivable else "",
        "应收换算系数": receivable.factor if receivable else "",
        "房间号": (receivable.room if receivable else "") or (actual.room if actual else ""),
        "识别公司": (receivable.company if receivable else "") or (actual.company if actual else ""),
        "实收原始行号": actual.excel_row if actual else "",
        "实收原始购买方": actual.raw_name if actual else "",
        "实收原始金额": actual.raw_amount if actual else "",
        "实收换算金额": actual.amount if actual and actual.amount is not None else "",
        "实收税前/税后": actual.tax_status if actual else "",
        "实收换算系数": actual.factor if actual else "",
        "倍数关系": note if status == "条件匹配" else relation,
        "说明": note if status != "条件匹配" else "",
    }


def save_payment_results(
    results: list[dict[str, Any]],
    output_path: Path,
    source_sheets: list[tuple[str, Path, str]],
    comparison_period: tuple[date, date] | None = None,
    room_mappings: dict[str, list[RoomMapping]] | None = None,
    start_date_header: str = "起租日",
    receivable_clean_headers: tuple[str, str] | None = None,
) -> Path:
    from openpyxl import Workbook

    output_path = output_path.expanduser().resolve()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = RECEIVABLE_SHEET
    headers = [
        "结果", "房间号", "识别公司",
        "应收原始行号", "应收原始购买方", "应收原始金额",
        "应收换算金额", "应收税前/税后", "应收换算系数",
        "实收原始行号", "实收原始购买方", "实收原始金额",
        "实收换算金额", "实收税前/税后", "实收换算系数",
        "倍数关系", "说明",
    ]
    sections = [
        (
            "比对及公共信息"
            + (
                f"（核对时段：{comparison_period[0]:%Y-%m-%d} 至 "
                f"{comparison_period[1]:%Y-%m-%d}）"
                if comparison_period
                else "（全部时段）"
            ),
            1,
            3,
            "44546A",
        ),
        ("应收款信息", 4, 9, "2F75B5"),
        ("实收款信息", 10, 15, "548235"),
        ("匹配说明", 16, 17, "8064A2"),
    ]
    for title, start_column, end_column, color in sections:
        sheet.merge_cells(
            start_row=1,
            start_column=start_column,
            end_row=1,
            end_column=end_column,
        )
        cell = sheet.cell(1, start_column, title)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill("solid", fgColor=color)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for column in range(start_column, end_column + 1):
            sheet.cell(1, column).fill = PatternFill("solid", fgColor=color)
    for column, header in enumerate(headers, start=1):
        sheet.cell(2, column, header)
    fills = {
        "完全匹配": PatternFill("solid", fgColor="E2F0D9"),
        "条件匹配": PatternFill("solid", fgColor="FFF2CC"),
        "待确认": PatternFill("solid", fgColor="DDEBF7"),
        "未匹配": PatternFill("solid", fgColor="FCE4D6"),
    }
    for result in results:
        sheet.append([result.get(header, "") for header in headers])
    section_header_colors = {
        1: "D9E2F3",
        2: "D9E2F3",
        3: "D9E2F3",
        4: "DDEBF7",
        5: "DDEBF7",
        6: "DDEBF7",
        7: "DDEBF7",
        8: "DDEBF7",
        9: "DDEBF7",
        10: "E2F0D9",
        11: "E2F0D9",
        12: "E2F0D9",
        13: "E2F0D9",
        14: "E2F0D9",
        15: "E2F0D9",
        16: "E4DFEC",
        17: "E4DFEC",
    }
    for column, color in section_header_colors.items():
        cell = sheet.cell(2, column)
        cell.font = Font(bold=True, color="1F1F1F")
        cell.fill = PatternFill("solid", fgColor=color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    section_body_colors = {
        1: "F2F5F9",
        2: "F2F5F9",
        3: "F2F5F9",
        4: "F4F9FD",
        5: "F4F9FD",
        6: "F4F9FD",
        7: "F4F9FD",
        8: "F4F9FD",
        9: "F4F9FD",
        10: "F4F8F2",
        11: "F4F8F2",
        12: "F4F8F2",
        13: "F4F8F2",
        14: "F4F8F2",
        15: "F4F8F2",
        16: "F8F6FA",
        17: "F8F6FA",
    }
    separator = Side(style="medium", color="A6A6A6")
    for row in range(3, sheet.max_row + 1):
        status_fill = fills.get(sheet.cell(row, 1).value, fills["未匹配"])
        for column in range(1, len(headers) + 1):
            cell = sheet.cell(row, column)
            cell.fill = status_fill if column == 1 else PatternFill(
                "solid", fgColor=section_body_colors[column]
            )
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        for column in (6, 7, 12, 13):
            sheet.cell(row, column).number_format = "0.00"
        for column in (9, 15):
            sheet.cell(row, column).number_format = "0.00000"
        for column in (3, 9, 15, 17):
            sheet.cell(row, column).border = Border(right=separator)
    sheet.row_dimensions[1].height = 24
    sheet.row_dimensions[2].height = 34
    sheet.freeze_panes = "D3"
    sheet.auto_filter.ref = f"A2:Q{sheet.max_row}"
    widths = [14, 16, 28, 14, 30, 16, 16, 14, 14, 14, 30, 16, 16, 14, 14, 24, 40]
    for column, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width
    if room_mappings is not None:
        mapping_sheet = workbook.create_sheet("房间承租人映射")
        mapping_headers = [
            "房间号",
            "承租人",
            "免租期",
            "实际租期",
            "起租日",
        ]
        mapping_sheet.append(mapping_headers)
        for column, cell in enumerate(mapping_sheet[1], start=1):
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
        mapping_rows = sorted(
            (
                (item.room_display, item)
                for items in room_mappings.values()
                for item in items
                if item.active
            ),
            key=lambda row: (
                room_key(row[0]),
                normalize(row[1].company),
                row[1].start_date or date.min,
            ),
        )
        for room_display, item in mapping_rows:
            lease_period = item.lease_period
            mapping_sheet.append(
                [
                    room_display,
                    item.company,
                    "有" if item.rent_free else "无",
                    (
                        f"{lease_period[0]:%Y-%m-%d} 至 "
                        f"{lease_period[1]:%Y-%m-%d}"
                        if lease_period
                        else ""
                    ),
                    item.start_date_source
                    if item.start_date_source is not None
                    else "",
                ]
            )
        mapping_sheet.freeze_panes = "A2"
        mapping_sheet.auto_filter.ref = mapping_sheet.dimensions
        for column, width in enumerate((16, 42, 12, 28, 18), start=1):
            mapping_sheet.column_dimensions[get_column_letter(column)].width = width
        for row in mapping_sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            row[4].number_format = "yyyy-mm-dd"
    for output_title, source_path, source_title in source_sheets:
        source_workbook = load_workbook(source_path, read_only=True, data_only=False)
        try:
            if source_title not in source_workbook.sheetnames:
                raise ValueError(f"文件中不存在工作表“{source_title}”。")
            source_sheet = source_workbook[source_title]
            raw_sheet = workbook.create_sheet(output_title)
            source_rows = source_sheet.iter_rows(values_only=True)
            if output_title == "应收原始数据" and receivable_clean_headers:
                headers = next(source_rows, ())
                raw_sheet.append(list(headers))
                normalized_headers = [
                    str(value).strip() if value is not None else ""
                    for value in headers
                ]
                try:
                    name_index = normalized_headers.index(
                        receivable_clean_headers[0]
                    )
                    amount_index = normalized_headers.index(
                        receivable_clean_headers[1]
                    )
                except ValueError as exc:
                    raise ValueError(
                        "应收款原始数据表缺少清理所需的名称列或金额列。"
                    ) from exc
                for row in source_rows:
                    name = row[name_index] if len(row) > name_index else None
                    amount = row[amount_index] if len(row) > amount_index else None
                    if keep_clean_receivable_row(name, amount):
                        raw_sheet.append(list(row))
            else:
                for row in source_rows:
                    raw_sheet.append(list(row))
            raw_sheet.freeze_panes = "A2"
            raw_sheet.auto_filter.ref = raw_sheet.dimensions
            for column in range(1, raw_sheet.max_column + 1):
                raw_sheet.column_dimensions[get_column_letter(column)].width = 20
            for cell in raw_sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(horizontal="center", wrap_text=True)
        finally:
            source_workbook.close()
    workbook.save(output_path)
    workbook.close()
    return output_path


def process_payment_files(
    mapping_path: Path,
    mapping_sheet: str,
    room_header: str,
    company_header: str,
    rent_free_header: str,
    lease_header: str,
    start_date_header: str,
    receivable_path: Path,
    receivable_sheet: str,
    receivable_name_header: str,
    receivable_amount_header: str,
    receivable_factor: Decimal,
    actual_path: Path,
    actual_sheet: str,
    actual_name_header: str,
    actual_amount_header: str,
    actual_factor: Decimal,
    output_path: Path,
    comparison_period: tuple[date, date] | None = None,
    clean_receivable: bool = False,
) -> tuple[Path, Counter[str]]:
    mapping_data = load_room_mapping(
        mapping_path,
        mapping_sheet,
        room_header,
        company_header,
        rent_free_header,
        lease_header,
        comparison_period,
        start_date_header,
    )
    receivables = load_payment_records(
        receivable_path,
        receivable_sheet,
        receivable_name_header,
        receivable_amount_header,
        mapping_data,
        receivable_factor,
        include_non_rent=True,
        include_unresolved_room=True,
        clean_receivable=clean_receivable,
    )
    actuals = load_payment_records(
        actual_path,
        actual_sheet,
        actual_name_header,
        actual_amount_header,
        mapping_data,
        actual_factor,
        apply_rent_free_adjustment=True,
        include_non_rent=True,
    )
    excluded_receivables: list[PaymentRecord] = []
    if comparison_period is not None:
        period_receivables = filter_records_by_period(receivables, comparison_period)
        period_receivable_rows = {
            record.excel_row for record in period_receivables
        }
        excluded_receivables = [
            record
            for record in receivables
            if record.excel_row not in period_receivable_rows
        ]
        receivables = period_receivables
        actuals = filter_records_by_period(actuals, comparison_period)
    actuals = select_actual_rent_candidates(receivables, actuals)
    if not receivables and not actuals and not excluded_receivables:
        if comparison_period:
            raise ValueError(
                f"应收款表和实收款表中均没有包含 "
                f"{comparison_period[0]:%Y-%m-%d} 至 "
                f"{comparison_period[1]:%Y-%m-%d} 的可比对数据。"
            )
        raise ValueError("应收款表和实收款表中均没有包含“租金”的可比对数据。")
    results = reconcile_payments(receivables, actuals)
    results.extend(
        payment_result(
            "未匹配",
            record,
            None,
            "摘要日期范围未与核对时段重叠或日期无法解析，未参与配对",
        )
        for record in excluded_receivables
    )
    saved_path = save_payment_results(
        results,
        output_path,
        [
            ("应收原始数据", receivable_path, receivable_sheet),
            ("实收原始数据", actual_path, actual_sheet),
        ],
        comparison_period,
        mapping_data,
        start_date_header,
        (
            receivable_name_header,
            receivable_amount_header,
        )
        if clean_receivable
        else None,
    )
    return saved_path, Counter(result["结果"] for result in results)


def run_gui() -> int:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("应收款与实收款比对工具")
    root.geometry("1160x760")
    root.minsize(1000, 700)

    main_frame = ttk.Frame(root, padding=16)
    main_frame.pack(fill="both", expand=True)
    ttk.Label(
        main_frame,
        text="应收款与实收款比对",
        font=("", 15, "bold"),
    ).pack(anchor="w")
    ttk.Label(
        main_frame,
        text="选择房间号映射表和两份收款数据。选择起租日列后，同一房间仅保留租期有效记录中起租日最新的一条；如果只有一条有效记录但起租日为空，则保留该记录。Excel 第一行需为表头；仅支持 .xlsx / .xlsm。",
        wraplength=1080,
    ).pack(anchor="w", pady=(5, 12))

    def file_panel(parent: Any, title: str, field_names: tuple[str, ...]) -> dict[str, Any]:
        panel = ttk.LabelFrame(parent, text=title, padding=10)
        state: dict[str, Any] = {
            "panel": panel,
            "path": tk.StringVar(),
            "file_label": tk.StringVar(value="尚未选择"),
        }
        ttk.Label(panel, textvariable=state["file_label"], width=30).grid(
            row=0, column=0, columnspan=3, sticky="ew", pady=(0, 6)
        )

        def choose_file() -> None:
            selected = filedialog.askopenfilename(
                parent=root,
                title=f"选择{title} Excel 文件",
                filetypes=[("Excel 工作簿", "*.xlsx *.xlsm")],
            )
            if not selected:
                return
            path = Path(selected)
            try:
                sheets = workbook_sheets(path)
            except WORKBOOK_ERRORS as exc:
                messagebox.showerror("读取文件失败", str(exc), parent=root)
                return
            state["path"].set(str(path))
            state["file_label"].set(str(path))
            state["sheet"]["values"] = sheets
            state["sheet"].current(0)
            refresh_columns()

        ttk.Button(panel, text="选择文件…", command=choose_file).grid(
            row=0, column=3, sticky="e", padx=(8, 0), pady=(0, 6)
        )

        ttk.Label(panel, text="工作表").grid(row=1, column=0, sticky="w")
        state["sheet_var"] = tk.StringVar()
        state["sheet"] = ttk.Combobox(
            panel, textvariable=state["sheet_var"], state="readonly", width=17
        )
        state["sheet"].grid(row=1, column=1, columnspan=3, sticky="ew", padx=(6, 0))
        state["sheet"].bind("<<ComboboxSelected>>", lambda _event: refresh_columns())
        state["columns"] = {}
        for index, field_name in enumerate(field_names):
            field_row = 2 + index
            ttk.Label(panel, text=field_name).grid(
                row=field_row, column=0, sticky="w", pady=(5, 0)
            )
            variable = tk.StringVar()
            combo = ttk.Combobox(
                panel, textvariable=variable, state="readonly", width=17
            )
            combo.grid(
                row=field_row,
                column=1,
                columnspan=3,
                sticky="ew",
                padx=(6, 0),
                pady=(5, 0),
            )
            state["columns"][field_name] = (variable, combo)

        def refresh_columns() -> None:
            path_text = state["path"].get()
            sheet_name = state["sheet_var"].get()
            if not path_text or not sheet_name:
                return
            try:
                headers = read_excel_headers(Path(path_text), sheet_name)
            except WORKBOOK_ERRORS as exc:
                messagebox.showerror("读取表头失败", str(exc), parent=root)
                return
            for variable, combo in state["columns"].values():
                combo["values"] = headers
                if headers:
                    variable.set(headers[0])
        state["refresh_columns"] = refresh_columns
        for column in range(4):
            panel.columnconfigure(column, weight=1 if column % 2 else 0)
        return state

    input_panels = ttk.Frame(main_frame)
    input_panels.pack(fill="x")
    for column in range(3):
        input_panels.columnconfigure(column, weight=1, uniform="file_panel")
    mapping = file_panel(
        input_panels,
        "1. 房间号与公司全称映射表",
        ("房间号列", "承租方列", "免租期列", "实际租期列", "起租日列"),
    )
    mapping["panel"].grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=5)
    receivable = file_panel(input_panels, "2. 应收款表", ("购买方名称列", "金额列"))
    receivable["panel"].grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
    actual = file_panel(input_panels, "3. 实收款表", ("购买方名称列", "金额列"))
    actual["panel"].grid(row=0, column=2, sticky="nsew", padx=(5, 0), pady=5)

    factors = ttk.LabelFrame(main_frame, text="金额换算与时间段筛选", padding=10)
    factors.pack(fill="x", pady=5)
    ttk.Label(factors, text="应收款换算系数").grid(row=0, column=0, sticky="w")
    receivable_factor = tk.StringVar(value="1")
    ttk.Entry(factors, textvariable=receivable_factor, width=12).grid(
        row=0, column=1, sticky="w", padx=(6, 18)
    )
    ttk.Label(factors, text="实收款换算系数").grid(row=0, column=2, sticky="w")
    actual_factor = tk.StringVar(value="1")
    ttk.Entry(factors, textvariable=actual_factor, width=12).grid(
        row=0, column=3, sticky="w", padx=(6, 18)
    )
    ttk.Label(
        factors,
        text="默认 1（不换算）；例如含税金额按 1.09 换算。映射表免租期为“有”时，实收金额再乘 11/12。",
        wraplength=1040,
    ).grid(row=1, column=0, columnspan=6, sticky="w", pady=(7, 0))
    ttk.Label(factors, text="核对时间段").grid(row=2, column=0, sticky="w", pady=(8, 0))
    comparison_start_text = tk.StringVar()
    comparison_end_text = tk.StringVar()
    ttk.Entry(factors, textvariable=comparison_start_text, width=12).grid(
        row=2, column=1, sticky="w", padx=(6, 4), pady=(8, 0)
    )
    ttk.Label(factors, text="至").grid(row=2, column=1, sticky="e", padx=(0, 4), pady=(8, 0))
    ttk.Entry(factors, textvariable=comparison_end_text, width=12).grid(
        row=2, column=2, sticky="w", padx=(6, 18), pady=(8, 0)
    )
    clean_receivable = tk.BooleanVar(value=False)
    ttk.Checkbutton(
        factors,
        text="清理应收表（删除 A 列为空或含指定标记的行；导出的应收原始数据仅保留有金额的有效记录）",
        variable=clean_receivable,
    ).grid(row=3, column=0, columnspan=6, sticky="w", pady=(8, 0))
    ttk.Label(
        factors,
        text="填写开始和结束日期（YYYY-MM-DD）；摘要日期范围与此时段重叠时参与配对，无日期范围时按输入时段处理。应收表中所有金额非空的行都会列入结果；超出时段或日期无法解析的行列为未匹配。",
        wraplength=1040,
    ).grid(row=4, column=0, columnspan=6, sticky="w", pady=(8, 0))

    status = tk.StringVar(value="就绪。请按顺序导入三份 Excel 并选择对应列。")
    ttk.Label(main_frame, textvariable=status, wraplength=1080).pack(
        anchor="w", pady=(8, 8)
    )
    progress = ttk.Progressbar(main_frame, mode="indeterminate")
    progress.pack(fill="x", pady=(0, 8))
    buttons = ttk.Frame(main_frame)
    buttons.pack(fill="x", pady=(2, 0))
    run_button = ttk.Button(buttons, text="开始比对并导出结果")
    run_button.pack(side="left")
    open_button = ttk.Button(buttons, text="打开结果文件", state="disabled")
    open_button.pack(side="left", padx=(10, 0))
    last_output: Path | None = None
    executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="excel-reconcile")
    pending_future: Future[tuple[Path, Counter[str]]] | None = None
    comparison_started_at: float | None = None

    def open_result() -> None:
        if last_output is None:
            return
        try:
            os.startfile(last_output)
        except OSError as exc:
            messagebox.showerror("无法打开结果文件", str(exc), parent=root)

    def run_comparison() -> None:
        nonlocal last_output, pending_future, comparison_started_at
        configs = (mapping, receivable, actual)
        if any(not config["path"].get() for config in configs):
            messagebox.showwarning("缺少文件", "请先导入映射表、应收款表和实收款表。", parent=root)
            return
        for config in configs:
            if not config["sheet_var"].get():
                messagebox.showwarning("缺少工作表", "请为每份文件选择工作表。", parent=root)
                return
            if any(not variable.get() for variable, _combo in config["columns"].values()):
                messagebox.showwarning("缺少列", "请为每份文件选择所需的数据列。", parent=root)
                return
        try:
            r_factor = Decimal(receivable_factor.get().strip())
            a_factor = Decimal(actual_factor.get().strip())
            if not r_factor.is_finite() or not a_factor.is_finite() or r_factor <= 0 or a_factor <= 0:
                raise InvalidOperation
        except InvalidOperation:
            messagebox.showerror("换算系数无效", "请输入大于 0 的有效数字，例如 1 或 1.09。", parent=root)
            return
        try:
            comparison_period = parse_comparison_period(
                comparison_start_text.get(),
                comparison_end_text.get(),
            )
        except ValueError as exc:
            messagebox.showerror("核对时间段无效", str(exc), parent=root)
            return

        default_path = Path(receivable["path"].get())
        output_default = default_path.with_name(f"{default_path.stem}_应收实收比对结果.xlsx")
        output_selected = filedialog.asksaveasfilename(
            parent=root,
            title="保存比对结果",
            initialdir=str(output_default.parent),
            initialfile=output_default.name,
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not output_selected:
            return
        output_path = Path(output_selected)
        input_paths = {Path(config["path"].get()).resolve() for config in configs}
        if output_path.resolve() in input_paths:
            messagebox.showerror("输出路径无效", "结果文件不能覆盖任何输入文件。", parent=root)
            return

        run_button.configure(state="disabled")
        open_button.configure(state="disabled")
        status.set("正在后台读取 Excel 并比对，请稍候；窗口仍可响应。")
        progress.start(12)
        comparison_started_at = monotonic()
        pending_future = executor.submit(
            process_payment_files,
            Path(mapping["path"].get()),
            mapping["sheet_var"].get(),
            mapping["columns"]["房间号列"][0].get(),
            mapping["columns"]["承租方列"][0].get(),
            mapping["columns"]["免租期列"][0].get(),
            mapping["columns"]["实际租期列"][0].get(),
            mapping["columns"]["起租日列"][0].get(),
            Path(receivable["path"].get()),
            receivable["sheet_var"].get(),
            receivable["columns"]["购买方名称列"][0].get(),
            receivable["columns"]["金额列"][0].get(),
            r_factor,
            Path(actual["path"].get()),
            actual["sheet_var"].get(),
            actual["columns"]["购买方名称列"][0].get(),
            actual["columns"]["金额列"][0].get(),
            a_factor,
            output_path,
            comparison_period,
            clean_receivable.get(),
        )
        root.after(100, check_comparison)

    def check_comparison() -> None:
        nonlocal last_output, pending_future, comparison_started_at
        if pending_future is None:
            return
        if not pending_future.done():
            elapsed_seconds = (
                int(monotonic() - comparison_started_at)
                if comparison_started_at is not None
                else 0
            )
            status.set(
                "正在后台读取 Excel 并比对，请稍候；窗口仍可响应。"
                f"\n已用时：{elapsed_seconds} 秒。"
            )
            root.after(500, check_comparison)
            return
        try:
            last_output, counts = pending_future.result()
        except WORKBOOK_ERRORS + (InvalidOperation,) as exc:
            last_output = None
            status.set("比对失败。")
            messagebox.showerror("比对失败", str(exc), parent=root)
        except Exception as exc:
            last_output = None
            status.set("后台处理发生未预期错误，详情请查看错误提示。")
            messagebox.showerror(
                "后台处理失败",
                f"{type(exc).__name__}: {exc}",
                parent=root,
            )
        else:
            open_button.configure(state="normal")
            status.set(
                f"比对完成：\n{last_output}\n"
                f"完全匹配 {counts['完全匹配']} 条；条件匹配 {counts['条件匹配']} 条；"
                f"待确认 {counts['待确认']} 条；"
                f"未匹配 {counts['未匹配']} 条。"
            )
        finally:
            pending_future = None
            comparison_started_at = None
            progress.stop()
            run_button.configure(state="normal")

    def close_window() -> None:
        executor.shutdown(wait=False, cancel_futures=True)
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", close_window)

    run_button.configure(command=run_comparison)
    open_button.configure(command=open_result)
    root.mainloop()
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="比对记账摘要税额与购买方标准费用类别的应税税额。"
    )
    parser.add_argument("input", nargs="?", help="输入 Excel 文件路径")
    parser.add_argument("-o", "--output", help="输出 Excel 文件路径")
    args = parser.parse_args()
    if not args.input and not args.output:
        return run_gui()
    if not args.input:
        parser.error("指定输出文件时也必须提供输入 Excel 文件。")
    try:
        process_workbook(Path(args.input), Path(args.output) if args.output else None)
        return 0
    except WORKBOOK_ERRORS as exc:
        print(f"处理失败：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
