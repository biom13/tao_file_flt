#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import sys
import unicodedata
from copy import copy, deepcopy
from datetime import date
from pathlib import Path
from typing import Any
from uuid import uuid4
from xml.etree.ElementTree import tostring

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from openpyxl import load_workbook
from openpyxl.formula.tokenizer import Tokenizer, TokenizerError
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


ACTION_BY_SHEET = {
    "new": "CREATE-NEW",
    "new employee": "CREATE-NEW",
    "new dw": "CREATE-NEW",
    "transfer": "TRANSFER",
    "transfer & rehire": "TRANSFER",
    "transfer&rehire": "TRANSFER",
    "transfer & rehire & đổi tên": "TRANSFER",
    "transfer&rehire&đổi tên": "TRANSFER",
    "old dw": "TRANSFER",
    "resigned": "RESIGNED",
    "resign": "RESIGNED",
    "quit": "RESIGNED",
    "info-change": "INFO-CHANGED",
    "info change": "INFO-CHANGED",
    "info-changed": "INFO-CHANGED",
    "info changed": "INFO-CHANGED",
}
ACTION_VALUES = ("CREATE-NEW", "TRANSFER", "RESIGNED", "INFO-CHANGED")
ACTION_FORMULA = '"' + ",".join(ACTION_VALUES) + '"'
RESIGNED_COLUMNS = {
    "Action",
    "Effected Date",
    "Employee Code",
    "FullName",
    "ITCode",
    "SetCode",
    "Level",
    "Location",
}
TARGET_SHEET = "EMPLOYEES"
MAX_RECORDS_PER_FILE = 30
HEADER_ROW = 5
DATA_START_ROW = HEADER_ROW + 1
BUNDLED_TEMPLATE = Path(__file__).resolve().parent.parent / "resources" / "7009_Employee_Template.xlsx"
BUNDLED_TEMPLATE_SHA256 = "3242d0ac486b87b48f71d553f9317c0fc9834dc350a6c44f6167ffa1cebe1a5a"
EXPECTED_TEMPLATE_SHEETS = [
    "EMPLOYEES",
    "SetCodes",
    "Locations",
    "Departments",
    "Sections",
    "Groups",
    "Levels",
]

EXPECTED_TARGET_HEADERS = [
    "Action",
    "Effected Date",
    "Employee Code",
    "FullName",
    "Bio-Code",
    "ITCode",
    "SetCode",
    "Level",
    "Location",
    "Department",
    "Section",
    "Group",
    "Note",
    "ID Card No.",
    "Gender",
    "Email",
    "Phone Number",
]

REQUIRED_SOURCE_COLUMNS = [
    "IT. No",
    "Code",
    "Name",
    "Location",
    "Departments",
    "Section",
    "Group",
    "Setcode",
    "Position",
]
DATE_COLUMNS_BY_ACTION = {
    "CREATE-NEW": ("Ngày lấy vân tay", "Starting Date", "Effected Date", "Effective date", "Transfer/rehire date", "Transfer / rehire date"),
    "TRANSFER": ("Transfer From", "Transfer/rehire date", "Transfer / rehire date", "Starting Date", "Effected Date", "Effective date", "Ngày lấy vân tay"),
    "RESIGNED": ("Stopping date", "Quit Date", "Effected Date", "Effective date", "Transfer/rehire date", "Transfer / rehire date"),
    "INFO-CHANGED": ("Effected Date", "Effective date", "Transfer/rehire date", "Transfer / rehire date"),
}


class ImportBuildError(Exception):
    pass


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def validate_template_file(template_path: Path) -> None:
    if not template_path.exists():
        raise ImportBuildError(
            f"Không tìm thấy workbook mẫu: {template_path}. "
            "Không được tự tạo lại workbook mẫu thủ công."
        )
    if template_path.resolve() == BUNDLED_TEMPLATE.resolve():
        actual_hash = sha256_file(template_path)
        if actual_hash != BUNDLED_TEMPLATE_SHA256:
            raise ImportBuildError(
                "Kiểm tra toàn vẹn SHA-256 của file mẫu 7009_Employee_Template.xlsx thất bại. "
                "Vui lòng khôi phục đúng file mẫu chuẩn."
            )


def remove_broken_template_links(workbook) -> None:
    for name in ("Staff", "Worker"):
        item = workbook.defined_names.get(name)
        if item is not None and "[1]Dept Code" in (item.attr_text or ""):
            del workbook.defined_names[name]
    workbook._external_links = []


def has_external_formula(cell) -> bool:
    if cell.data_type != "f":
        return False
    if not isinstance(cell.value, str):
        raise ImportBuildError(f"Unsupported formula type at {cell.parent.title}!{cell.coordinate}")
    try:
        tokens = Tokenizer(cell.value).items
    except TokenizerError as exc:
        raise ImportBuildError(f"Cannot inspect formula at {cell.parent.title}!{cell.coordinate}") from exc
    return any(
        token.type in ("OPERAND", "FUNC")
        and token.subtype != "TEXT"
        and "!" in token.value
        and re.search(r"\[(?:\d+|[^\[\]]+\.(?:xlsx?|xlsm|xlsb|xlam|xla))\]", token.value, re.IGNORECASE)
        for token in tokens
    )


def freeze_external_source_formulas(source_wb, source_path: Path) -> list[dict[str, Any]]:
    cells = [
        cell
        for sheet in source_wb
        if sheet.sheet_state == "visible"
        and action_from_sheet(sheet.title) is None
        and sheet.title != TARGET_SHEET
        for cell in sheet._cells.values()
        if has_external_formula(cell)
    ]
    if not cells:
        return []
    cached_wb = load_workbook(source_path, data_only=True, keep_links=False)
    replacements: list[dict[str, Any]] = []
    try:
        for cell in cells:
            cached = cached_wb[cell.parent.title][cell.coordinate]
            if cached.value is None:
                raise ImportBuildError(
                    f"Công thức tham chiếu ngoài tại {cell.parent.title}!{cell.coordinate} chưa có giá trị đã lưu. "
                    "Hãy mở file nguồn trong Excel để tính lại, lưu rồi chạy lại."
                )
            replacements.append(
                {
                    "sheet": cell.parent.title,
                    "cell": cell.coordinate,
                    "cached_error": cached.data_type == "e",
                }
            )
            cell.value = cached.value
            cell.data_type = cached.data_type
    finally:
        cached_wb.close()
    return replacements


def validate_no_external_formulas(workbook) -> None:
    for sheet in workbook:
        for cell in sheet._cells.values():
            if has_external_formula(cell):
                raise ImportBuildError(f"Output contains an external formula at {sheet.title}!{cell.coordinate}")


def validate_template_structure(template_wb) -> None:
    if template_wb.sheetnames != EXPECTED_TEMPLATE_SHEETS:
        raise ImportBuildError(
            "Template sheet structure differs from the approved 7009 template. Expected: "
            + ", ".join(EXPECTED_TEMPLATE_SHEETS)
        )
    if template_wb.defined_names.get("Staff") or template_wb.defined_names.get("Worker"):
        raise ImportBuildError("Template still contains unsupported Staff/Worker named ranges")
    if template_wb._external_links:
        raise ImportBuildError("Template still contains an external workbook link")


def normalize_header(value: Any) -> str:
    return "" if value is None else str(value).strip()


def action_from_sheet(sheet_name: str) -> str | None:
    key = " ".join(sheet_name.strip().casefold().replace("_", " ").split())
    exact = ACTION_BY_SHEET.get(key)
    if exact is not None:
        return exact
    compact = location_key(sheet_name)
    if "resign" in compact or "nghiviec" in compact:
        return "RESIGNED"
    return None


def source_sheet_names(workbook) -> list[str]:
    return [
        sheet.title
        for sheet in workbook
        if sheet.sheet_state == "visible" and action_from_sheet(sheet.title) is not None
    ]


def effected_date(row: dict[str, Any], action: str) -> Any:
    row_lower = {k.casefold(): v for k, v in row.items() if isinstance(k, str)}
    for name in DATE_COLUMNS_BY_ACTION[action]:
        val = row.get(name)
        if val in (None, ""):
            val = row_lower.get(name.casefold())
        if val not in (None, ""):
            return val
    if action == "INFO-CHANGED":
        return date.today()
    return None


def location_key(value: str) -> str:
    text = unicodedata.normalize("NFD", value.casefold().replace("Đ", "d").replace("đ", "d"))
    return "".join(char for char in text if not char.isspace() and not unicodedata.combining(char))


def location_reference(workbook) -> dict[str, set[str]]:
    sheet = workbook["Locations"]
    headers = build_header_map(sheet, HEADER_ROW)
    if "Code" not in headers:
        raise ImportBuildError("Locations reference sheet is missing its Code header")
    reference: dict[str, set[str]] = {}
    for cells in sheet.iter_rows(min_row=DATA_START_ROW):
        value = cells[headers["Code"] - 1].value
        if value in (None, ""):
            continue
        if not isinstance(value, str) or not location_key(value):
            raise ImportBuildError(f"Locations reference has an invalid Code: {value!r}")
        reference.setdefault(location_key(value), set()).add(value)
    if not reference:
        raise ImportBuildError("Locations reference has no location codes")
    return reference


def normalize_location(row: dict[str, Any], reference: dict[str, set[str]]) -> str:
    value = row.get("Location")
    matches = reference.get(location_key(value), set()) if isinstance(value, str) else set()
    if value in matches:
        return value
    if len(matches) == 1:
        return next(iter(matches))
    reason = "mơ hồ (khớp nhiều mã)" if matches else "trống hoặc không khớp danh mục"
    raise ImportBuildError(
        f"Sheet {row['_source_sheet']!r} dòng {row['_source_row']} có Location {reason}: {value!r}. "
        "Phải khớp với cột Code trong sheet Locations của mẫu."
    )


def level_from_position(row: dict[str, Any]) -> str:
    position = row.get("Position")
    normalized = str(position).strip().casefold() if position is not None else ""
    levels = {"daily worker": "DW", "dw": "DW", "staff": "Staff", "worker": "Worker"}
    if normalized not in levels:
        raise ImportBuildError(
            f"Sheet {row['_source_sheet']!r} dòng {row['_source_row']} có Position không hợp lệ: {position!r}. "
            "Chỉ chấp nhận Daily Worker, DW, staff hoặc worker."
        )
    return levels[normalized]


def detect_header_row(ws: Worksheet, max_scan_row: int = 5) -> int:
    key_headers = {
        "it. no",
        "it code",
        "code",
        "name",
        "names",
        "full name",
        "location",
        "departments",
        "department",
        "position",
        "level",
        "setcode",
        "set code",
    }
    best_row = 1
    best_hits = 0
    for r in range(1, min(ws.max_row, max_scan_row) + 1):
        hits = 0
        for col in range(1, min(ws.max_column, 30) + 1):
            val = normalize_header(ws.cell(r, col).value)
            if val and " ".join(val.casefold().split()) in key_headers:
                hits += 1
        if hits > best_hits:
            best_hits = hits
            best_row = r
        if best_hits >= 3:
            break
    return best_row


def build_header_map(ws: Worksheet, row: int = 1, *, stop_on_secondary_section: bool = False) -> dict[str, int]:
    result: dict[str, int] = {}
    lower_map: dict[str, int] = {}
    seen_code_col = False
    for cell in ws[row]:
        if stop_on_secondary_section and row > 1:
            banner_val = normalize_header(ws.cell(row - 1, cell.column).value)
            if banner_val and "phutrach" in location_key(banner_val) and seen_code_col:
                break
        header = normalize_header(cell.value)
        if not header:
            continue
        norm_key = " ".join(header.casefold().split())
        if stop_on_secondary_section and norm_key == "code" and seen_code_col:
            break
        if norm_key == "code":
            seen_code_col = True
        result.setdefault(header, cell.column)
        lower_map.setdefault(norm_key, cell.column)

    canonical_headers = (
        "IT. No",
        "Code",
        "Name",
        "Location",
        "Departments",
        "Section",
        "Group",
        "Setcode",
        "Position",
        "Gender",
    )
    for target in canonical_headers:
        if target not in result:
            norm_target = " ".join(target.casefold().split())
            if norm_target in lower_map:
                result[target] = lower_map[norm_target]

    aliases = {
        "it code": "IT. No",
        "no.": "IT. No",
        "full name": "Name",
        "names": "Name",
        "department": "Departments",
        "group/actual": "Group",
        "group / actual": "Group",
        "set code": "Setcode",
        "level": "Position",
    }
    for source_key, target_header in aliases.items():
        if target_header not in result and source_key in lower_map:
            result[target_header] = lower_map[source_key]
    return result


def ensure_source_columns(ws: Worksheet, sheet_name: str) -> tuple[dict[str, int], int]:
    header_row = detect_header_row(ws)
    header_map = build_header_map(ws, header_row, stop_on_secondary_section=True)
    action = action_from_sheet(sheet_name)
    if action is None:
        raise ImportBuildError(f"Unrecognized source sheet: {sheet_name!r}")
    required = (
        ("IT. No", "Code", "Name", "Position")
        if action == "RESIGNED"
        else REQUIRED_SOURCE_COLUMNS
    )
    missing = [name for name in required if name not in header_map]
    date_columns = DATE_COLUMNS_BY_ACTION[action]
    header_keys_lower = {k.casefold() for k in header_map}
    if action != "INFO-CHANGED" and not any(name.casefold() in header_keys_lower for name in date_columns):
        missing.append("cột ngày (" + " hoặc ".join(date_columns) + ")")
    if missing:
        raise ImportBuildError(f"Sheet {sheet_name!r} thiếu cột bắt buộc: {', '.join(missing)}")
    return header_map, header_row


def row_has_identity(row: dict[str, Any]) -> bool:
    return any(row.get(key) not in (None, "") for key in ("IT. No", "Code", "Name"))


def is_ignored_remark(row: dict[str, Any]) -> bool:
    for key, val in row.items():
        if isinstance(key, str) and key.strip().casefold() == "remark" and val is not None:
            if location_key(str(val)) == "doiten":
                return True
    return False


def read_source_rows(ws: Worksheet, sheet_name: str) -> list[dict[str, Any]]:
    header_map, header_row = ensure_source_columns(ws, sheet_name)
    tracked_cols = set(header_map.values())
    candidate_rows = sorted(
        {
            r
            for (r, c), cell in ws._cells.items()
            if r > header_row and c in tracked_cols and cell.value not in (None, "")
        }
    )
    rows: list[dict[str, Any]] = []
    for row_index in candidate_rows:
        item = {header: ws.cell(row_index, col).value for header, col in header_map.items()}
        if row_has_identity(item) and not is_ignored_remark(item):
            item["_source_sheet"] = sheet_name
            item["_source_row"] = row_index
            rows.append(item)
    return rows


def mapped_employee_rows(source_wb, template_wb) -> list[dict[str, Any]]:
    locations = location_reference(template_wb)
    output_rows: list[dict[str, Any]] = []
    sources = source_sheet_names(source_wb)
    if not sources:
        raise ImportBuildError("Không tìm thấy sheet nghiệp vụ nào đang hiển thị trong workbook nguồn.")
    for sheet_name in sources:
        for row in read_source_rows(source_wb[sheet_name], sheet_name):
            action = action_from_sheet(sheet_name)
            is_resigned = action == "RESIGNED"
            output_rows.append(
                {
                    "Action": action,
                    "Effected Date": effected_date(row, action),
                    "Employee Code": row.get("Code"),
                    "FullName": row.get("Name"),
                    "Bio-Code": None,
                    "ITCode": row.get("IT. No"),
                    "SetCode": "Resignation" if is_resigned else row.get("Setcode"),
                    "Level": level_from_position(row),
                    "Location": "Resignation" if is_resigned else normalize_location(row, locations),
                    "Department": None if is_resigned else row.get("Departments"),
                    "Section": None if is_resigned else row.get("Section"),
                    "Group": None if is_resigned else row.get("Group"),
                    "Note": None,
                    "ID Card No.": None,
                    "Gender": None if is_resigned else row.get("Gender"),
                    "Email": None,
                    "Phone Number": None,
                    "_source_sheet": row["_source_sheet"],
                    "_source_row": row["_source_row"],
                }
            )
    return output_rows


def copy_cell(src, dst) -> None:
    dst.value = src.value
    if src.has_style:
        dst._style = copy(src._style)
    if src.number_format:
        dst.number_format = src.number_format
    if src.font:
        dst.font = copy(src.font)
    if src.fill:
        dst.fill = copy(src.fill)
    if src.border:
        dst.border = copy(src.border)
    if src.alignment:
        dst.alignment = copy(src.alignment)
    if src.protection:
        dst.protection = copy(src.protection)
    if src.hyperlink:
        dst._hyperlink = copy(src.hyperlink)
    if src.comment:
        dst.comment = copy(src.comment)


def copy_dimension(src, dst) -> None:
    for attr in (
        "hidden",
        "outlineLevel",
        "collapsed",
        "thickTop",
        "thickBottom",
    ):
        if hasattr(src, attr) and hasattr(dst, attr):
            setattr(dst, attr, getattr(src, attr))

    for attr in ("width", "height", "bestFit", "min", "max"):
        if hasattr(src, attr) and hasattr(dst, attr):
            setattr(dst, attr, getattr(src, attr))

    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)


def copy_sheet_contents(src: Worksheet, dst: Worksheet) -> None:
    for row in src.iter_rows(min_row=1, max_row=src.max_row, min_col=1, max_col=src.max_column):
        for src_cell in row:
            copy_cell(src_cell, dst.cell(src_cell.row, src_cell.column))

    for key, dim in src.column_dimensions.items():
        copy_dimension(dim, dst.column_dimensions[key])

    for key, dim in src.row_dimensions.items():
        copy_dimension(dim, dst.row_dimensions[key])

    for merged_range in src.merged_cells.ranges:
        dst.merge_cells(str(merged_range))

    dst.freeze_panes = src.freeze_panes
    for attr in (
        "showGridLines",
        "showRowColHeaders",
        "showZeros",
        "rightToLeft",
        "tabSelected",
        "topLeftCell",
        "view",
        "zoomScale",
        "zoomScaleNormal",
        "zoomScalePageLayoutView",
        "zoomScaleSheetLayoutView",
    ):
        setattr(dst.sheet_view, attr, getattr(src.sheet_view, attr))
    dst.sheet_view.pane = deepcopy(src.sheet_view.pane)
    dst.sheet_view.selection = deepcopy(src.sheet_view.selection)
    dst.page_margins = deepcopy(src.page_margins)
    dst.page_setup = deepcopy(src.page_setup)
    dst.print_options = deepcopy(src.print_options)
    dst.sheet_properties = deepcopy(src.sheet_properties)
    dst.sheet_format = deepcopy(src.sheet_format)
    dst.protection = deepcopy(src.protection)
    dst.auto_filter = deepcopy(src.auto_filter)
    dst.data_validations = deepcopy(src.data_validations)
    dst.conditional_formatting = deepcopy(src.conditional_formatting)
    dst.sheet_state = src.sheet_state
    dst.print_area = src.print_area
    dst.print_title_rows = src.print_title_rows
    dst.print_title_cols = src.print_title_cols

    for table in src.tables.values():
        copied_table = deepcopy(table)
        for attr in (
            "headerRowDxfId",
            "dataDxfId",
            "totalsRowDxfId",
            "headerRowBorderDxfId",
            "tableBorderDxfId",
            "totalsRowBorderDxfId",
        ):
            setattr(copied_table, attr, None)
        for column in copied_table.tableColumns:
            column.headerRowDxfId = None
            column.dataDxfId = None
            column.totalsRowDxfId = None
        dst.add_table(copied_table)
    for chart in src._charts:
        dst.add_chart(deepcopy(chart), deepcopy(chart.anchor))
    for image in src._images:
        dst.add_image(deepcopy(image), deepcopy(image.anchor))


def append_source_sheets(output_wb, source_wb, template_sheet_names: list[str]) -> None:
    reserved = set(template_sheet_names)
    collisions = [
        name
        for name in source_wb.sheetnames
        if source_wb[name].sheet_state == "visible"
        and action_from_sheet(name) is None
        and name != TARGET_SHEET
        and name in reserved
    ]
    if collisions:
        raise ImportBuildError(
            "Workbook nguồn có sheet hiển thị trùng tên với sheet tham chiếu của mẫu: " + ", ".join(collisions)
        )

    for source_ws in source_wb.worksheets:
        if (
            source_ws.sheet_state != "visible"
            or source_ws.title == TARGET_SHEET
            or action_from_sheet(source_ws.title) is not None
        ):
            continue
        target_ws = output_wb.create_sheet(source_ws.title)
        copy_sheet_contents(source_ws, target_ws)


def target_header_map(ws: Worksheet) -> dict[str, int]:
    headers = [ws.cell(HEADER_ROW, col).value for col in range(1, len(EXPECTED_TARGET_HEADERS) + 1)]
    if headers != EXPECTED_TARGET_HEADERS:
        raise ImportBuildError(
            f"{TARGET_SHEET!r} header row does not match expected template headers: {headers!r}"
        )
    return {header: idx + 1 for idx, header in enumerate(EXPECTED_TARGET_HEADERS)}


def clear_existing_data(ws: Worksheet, header_map: dict[str, int]) -> None:
    template_data_row = DATA_START_ROW
    max_col = max(ws.max_column, max(header_map.values()))
    for row in range(DATA_START_ROW, ws.max_row + 1):
        for col in range(1, max_col + 1):
            ws.cell(row, col).value = None
    for row in range(DATA_START_ROW + 1, ws.max_row + 1):
        ws.row_dimensions[row].height = ws.row_dimensions[template_data_row].height


def apply_data_row_style(ws: Worksheet, row_index: int, max_col: int) -> None:
    for col in range(1, max_col + 1):
        copy_cell(ws.cell(DATA_START_ROW, col), ws.cell(row_index, col))
        ws.cell(row_index, col).value = None
    ws.row_dimensions[row_index].height = ws.row_dimensions[DATA_START_ROW].height


def write_employee_rows(ws: Worksheet, rows: list[dict[str, Any]]) -> None:
    header_map = target_header_map(ws)
    clear_existing_data(ws, header_map)
    max_col = max(header_map.values())

    for offset, row in enumerate(rows):
        row_index = DATA_START_ROW + offset
        apply_data_row_style(ws, row_index, max_col)
        for header in EXPECTED_TARGET_HEADERS:
            ws.cell(row_index, header_map[header]).value = row.get(header)

    action_col = get_column_letter(header_map["Action"])
    action_range = f"{action_col}{DATA_START_ROW}:{action_col}{DATA_START_ROW + len(rows) - 1}"
    for validation in ws.data_validations.dataValidation:
        if validation.type == "list" and validation.formula1 == ACTION_FORMULA:
            validation.sqref = action_range
            validation.showDropDown = False
            validation.showErrorMessage = True
            validation.errorStyle = "stop"


def validate_output(ws: Worksheet, rows: list[dict[str, Any]]) -> None:
    target_header_map(ws)
    if not rows:
        raise ImportBuildError("Không tìm thấy bản ghi nhân viên hợp lệ nào.")
    if len(rows) > MAX_RECORDS_PER_FILE:
        raise ImportBuildError(f"Mỗi file import chỉ chứa tối đa {MAX_RECORDS_PER_FILE} nhân viên.")

    required_output = [
        "Action",
        "Effected Date",
        "Employee Code",
        "FullName",
        "ITCode",
        "SetCode",
        "Level",
        "Location",
        "Department",
    ]
    duplicates: dict[str, set[Any]] = {"Employee Code": set(), "ITCode": set()}
    seen: dict[str, set[Any]] = {"Employee Code": set(), "ITCode": set()}
    problems: list[str] = []
    allowed_locations = set().union(*location_reference(ws.parent).values())

    for idx, row in enumerate(rows, start=DATA_START_ROW):
        is_resigned = row.get("Action") == "RESIGNED"
        if is_resigned:
            for header in ("SetCode", "Location"):
                if row.get(header) != "Resignation":
                    problems.append(f"Row {idx} resigned {header} must be Resignation")
            for header in EXPECTED_TARGET_HEADERS:
                value = ws.cell(idx, EXPECTED_TARGET_HEADERS.index(header) + 1).value
                if header in RESIGNED_COLUMNS:
                    if value != row.get(header):
                        problems.append(f"Row {idx} resigned {header} differs from mapped source")
                elif row.get(header) not in (None, "") or value not in (None, ""):
                    problems.append(f"Row {idx} resigned {header} must be blank")
        expected_action = action_from_sheet(row.get("_source_sheet", ""))
        if row.get("Action") not in ACTION_VALUES or row.get("Action") != expected_action:
            problems.append(f"Row {idx} Action does not match its source sheet")
        if ws.cell(idx, 1).value != row.get("Action"):
            problems.append(f"Row {idx} written Action differs from its source sheet action")
        if row.get("Location") not in allowed_locations:
            problems.append(f"Row {idx} Location is not an exact Locations Code: {row.get('Location')!r}")
        if ws.cell(idx, EXPECTED_TARGET_HEADERS.index("Location") + 1).value != row.get("Location"):
            problems.append(f"Row {idx} written Location differs from its mapped source")
        if row.get("Level") not in ("DW", "Staff", "Worker"):
            problems.append(f"Row {idx} has invalid Level: {row.get('Level')!r}")
        for header in required_output:
            if is_resigned and header == "Department":
                continue
            if row.get(header) in (None, ""):
                problems.append(
                    f"Sheet {row.get('_source_sheet')!r} dòng {row.get('_source_row')} thiếu {header}"
                )
        for header in seen:
            value = row.get(header)
            if value in (None, ""):
                continue
            if value in seen[header]:
                duplicates[header].add(value)
            seen[header].add(value)

    for header, values in duplicates.items():
        if values:
            problems.append(f"Trùng {header}: {sorted(values)!r}")

    written_count = 0
    header_map = target_header_map(ws)
    for row_index in range(DATA_START_ROW, ws.max_row + 1):
        if any(ws.cell(row_index, header_map[h]).value not in (None, "") for h in EXPECTED_TARGET_HEADERS):
            written_count += 1
    if written_count != len(rows):
        problems.append(f"Expected {len(rows)} generated rows, found {written_count} written rows")

    first_blank_row = DATA_START_ROW + len(rows)
    for cell in ws._cells.values():
        if cell.row >= first_blank_row and cell.value is not None:
            problems.append(f"Unexpected content after employee records: {cell.coordinate}")
        elif cell.row >= DATA_START_ROW and cell.column > len(EXPECTED_TARGET_HEADERS) and cell.value is not None:
            problems.append(f"Unexpected content outside import columns: {cell.coordinate}")

    if problems:
        raise ImportBuildError("Kiểm tra thất bại:\n" + "\n".join(f"  - {p}" for p in problems))


def cell_style_signature(cell) -> tuple[Any, ...]:
    def xml(component) -> str:
        return tostring(component.to_tree(), encoding="unicode")

    return (
        xml(cell.font),
        xml(cell.fill),
        xml(cell.border),
        xml(cell.alignment),
        cell.number_format,
        xml(cell.protection),
    )


def dimension_signature(dimension) -> tuple[Any, ...]:
    width = getattr(dimension, "width", None)
    if dimension.hidden and width in (0, 13.0):
        width = "hidden"
    return (
        width,
        getattr(dimension, "height", None),
        dimension.hidden,
        dimension.outlineLevel,
        dimension.collapsed,
        getattr(dimension, "bestFit", None),
        cell_style_signature(dimension) if dimension.has_style else None,
    )


def validate_sheet_matches(
    expected: Worksheet,
    actual: Worksheet,
    *,
    compare_values: bool = True,
    max_value_row: int | None = None,
    max_style_row: int | None = None,
) -> None:
    problems: list[str] = []
    max_row = max(expected.max_row, actual.max_row)
    max_col = max(expected.max_column, actual.max_column)

    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            expected_cell = expected.cell(row, col)
            actual_cell = actual.cell(row, col)
            if compare_values and (max_value_row is None or row <= max_value_row):
                if expected_cell.value != actual_cell.value or expected_cell.data_type != actual_cell.data_type:
                    problems.append(f"cell {actual_cell.coordinate} value/formula differs")
            if (max_style_row is None or row <= max_style_row) and cell_style_signature(
                expected_cell
            ) != cell_style_signature(actual_cell):
                problems.append(f"cell {actual_cell.coordinate} formatting differs")
            if len(problems) >= 20:
                break
        if len(problems) >= 20:
            break

    if {str(item) for item in expected.merged_cells.ranges} != {
        str(item) for item in actual.merged_cells.ranges
    }:
        problems.append("merged ranges differ")

    for key in set(expected.column_dimensions) | set(actual.column_dimensions):
        if dimension_signature(expected.column_dimensions[key]) != dimension_signature(actual.column_dimensions[key]):
            problems.append(f"column {key} dimensions differ")
    for key in set(expected.row_dimensions) | set(actual.row_dimensions):
        if dimension_signature(expected.row_dimensions[key]) != dimension_signature(actual.row_dimensions[key]):
            problems.append(f"row {key} dimensions differ")

    if len(expected.data_validations.dataValidation) != len(actual.data_validations.dataValidation):
        problems.append("data validation count differs")
    if len(expected.conditional_formatting) != len(actual.conditional_formatting):
        problems.append("conditional formatting count differs")
    if len(expected.tables) != len(actual.tables):
        problems.append("table count differs")
    if len(expected._charts) != len(actual._charts) or len(expected._images) != len(actual._images):
        problems.append("drawing count differs")

    if problems:
        raise ImportBuildError(
            f"Sheet {actual.title!r} does not match its reference:\n"
            + "\n".join(f"  - {problem}" for problem in problems[:20])
        )


def validate_template_and_preserved_sheets(
    output_wb, template_wb, source_wb, generated_row_count: int
) -> None:
    template_names = template_wb.sheetnames
    if output_wb.sheetnames[: len(template_names)] != template_names:
        raise ImportBuildError(
            "Template sheets are not in the required order: " + ", ".join(template_names)
        )

    validate_sheet_matches(
        template_wb[TARGET_SHEET],
        output_wb[TARGET_SHEET],
        compare_values=True,
        max_value_row=HEADER_ROW,
        max_style_row=HEADER_ROW,
    )
    template_data_row = template_wb[TARGET_SHEET]
    output_employee_sheet = output_wb[TARGET_SHEET]
    for output_row in range(DATA_START_ROW, DATA_START_ROW + generated_row_count):
        for col in range(1, len(EXPECTED_TARGET_HEADERS) + 1):
            if cell_style_signature(template_data_row.cell(DATA_START_ROW, col)) != cell_style_signature(
                output_employee_sheet.cell(output_row, col)
            ):
                raise ImportBuildError(
                    f"Sheet {TARGET_SHEET!r} row {output_row} does not match the template data-row formatting"
                )
        if dimension_signature(template_data_row.row_dimensions[DATA_START_ROW]) != dimension_signature(
            output_employee_sheet.row_dimensions[output_row]
        ):
            raise ImportBuildError(
                f"Sheet {TARGET_SHEET!r} row {output_row} height/outline differs from the template"
            )
    for name in template_names[1:]:
        validate_sheet_matches(template_wb[name], output_wb[name])

    for name in source_wb.sheetnames:
        if source_wb[name].sheet_state != "visible" or action_from_sheet(name) is not None or name == TARGET_SHEET:
            continue
        validate_sheet_matches(source_wb[name], output_wb[name])


def validate_action_validation(ws: Worksheet, row_count: int) -> None:
    expected_range = (
        f"A{DATA_START_ROW}"
        if row_count == 1
        else f"A{DATA_START_ROW}:A{DATA_START_ROW + row_count - 1}"
    )
    valid = any(
        item.type == "list"
        and item.formula1 == ACTION_FORMULA
        and str(item.sqref) == expected_range
        and item.showDropDown is False
        and item.showErrorMessage is True
        and item.errorStyle == "stop"
        for item in ws.data_validations.dataValidation
    )
    if not valid:
        raise ImportBuildError(f"Action validation does not cover {expected_range}")


def remove_source_sheets(output_wb) -> None:
    for sheet_name in source_sheet_names(output_wb):
        if sheet_name in output_wb.sheetnames:
            output_wb.remove(output_wb[sheet_name])


def trim_unused_employee_rows(ws: Worksheet, row_count: int) -> None:
    last_row = HEADER_ROW + row_count
    if ws.max_row > last_row:
        ws.delete_rows(last_row + 1, ws.max_row - last_row)
    for row in list(ws.row_dimensions):
        if row > last_row:
            del ws.row_dimensions[row]


def build_import_part(source_wb, template_wb, template_path: Path, temporary_path: Path, rows) -> list[str]:
    final_sheet_names: list[str] = []
    output_wb = None
    saved_wb = None
    try:
        shutil.copy2(template_path, temporary_path)
        output_wb = load_workbook(temporary_path, data_only=False, keep_links=False)
        remove_broken_template_links(output_wb)
        append_source_sheets(output_wb, source_wb, template_wb.sheetnames)

        employees_ws = output_wb[TARGET_SHEET]
        write_employee_rows(employees_ws, rows)
        validate_output(employees_ws, rows)
        validate_action_validation(employees_ws, len(rows))
        validate_template_and_preserved_sheets(output_wb, template_wb, source_wb, len(rows))

        remove_source_sheets(output_wb)
        trim_unused_employee_rows(employees_ws, len(rows))
        output_wb.save(temporary_path)
        output_wb.close()
        output_wb = None

        saved_wb = load_workbook(temporary_path, data_only=False, keep_links=False)
        saved_employees = saved_wb[TARGET_SHEET]
        last_row = HEADER_ROW + len(rows)
        if saved_employees.max_row != last_row or any(row > last_row for row in saved_employees.row_dimensions):
            raise ImportBuildError(f"Saved {TARGET_SHEET} must end at row {last_row}")
        remove_broken_template_links(saved_wb)
        if saved_wb.defined_names.get("Staff") or saved_wb.defined_names.get("Worker"):
            raise ImportBuildError("Output still contains unsupported Staff/Worker named ranges")
        if saved_wb._external_links:
            raise ImportBuildError("Output still contains an external workbook link")
        validate_no_external_formulas(saved_wb)
        validate_output(saved_wb[TARGET_SHEET], rows)
        validate_action_validation(saved_wb[TARGET_SHEET], len(rows))
        validate_template_and_preserved_sheets(saved_wb, template_wb, source_wb, len(rows))
        if source_sheet_names(saved_wb):
            raise ImportBuildError("Source sheets are still present after final validation")
        final_sheet_names = saved_wb.sheetnames
        saved_wb.close()
        saved_wb = None
    finally:
        if output_wb is not None:
            output_wb.close()
        if saved_wb is not None:
            saved_wb.close()
    return final_sheet_names


def split_employee_rows(rows):
    return [rows[start : start + MAX_RECORDS_PER_FILE] for start in range(0, len(rows), MAX_RECORDS_PER_FILE)]


def build_import_file(source_path: Path, template_path: Path, output_path: Path) -> dict[str, Any]:
    if not source_path.exists():
        raise ImportBuildError(f"Không tìm thấy workbook nguồn: {source_path}")
    validate_template_file(template_path)
    source_wb = load_workbook(source_path, data_only=False, keep_links=False)
    template_wb = load_workbook(template_path, data_only=False, keep_links=False)
    staged_paths: list[Path] = []
    try:
        remove_broken_template_links(template_wb)
        validate_template_structure(template_wb)
        rows = mapped_employee_rows(source_wb, template_wb)
        if not rows:
            raise ImportBuildError("Không tìm thấy bản ghi nhân viên hợp lệ nào.")
        for header in ("Employee Code", "ITCode"):
            seen = set()
            for row in rows:
                value = row.get(header)
                if value in (None, ""):
                    continue
                if value in seen:
                    raise ImportBuildError(f"Trùng {header} trên toàn bộ nguồn: {value!r}")
                seen.add(value)
        batches = split_employee_rows(rows)
        final_paths = [
            output_path
            if len(batches) == 1
            else output_path.with_name(f"{output_path.stem}_{index:02d}{output_path.suffix}")
            for index in range(1, len(batches) + 1)
        ]
        if any(path.resolve() == source_path.resolve() for path in [output_path, *final_paths]):
            raise ImportBuildError("Đường dẫn xuất không được trùng với file nguồn.")
        external_formula_changes = freeze_external_source_formulas(source_wb, source_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        outputs = []
        for path, batch in zip(final_paths, batches):
            temporary_path = path.with_name(f".{path.stem}.{uuid4().hex}.building{path.suffix}")
            staged_paths.append(temporary_path)
            sheets = build_import_part(source_wb, template_wb, template_path, temporary_path, batch)
            outputs.append({"output": str(path), "generated_rows": len(batch), "sheets": sheets})
        if sum(part["generated_rows"] for part in outputs) != len(rows):
            raise ImportBuildError("Tổng số dòng sau khi tách file không khớp số bản ghi nguồn.")
        for staged, final in zip(staged_paths, final_paths):
            staged.replace(final)
        return {
            "output": str(final_paths[0]) if len(final_paths) == 1 else None,
            "outputs": outputs,
            "generated_rows": len(rows),
            "source_rows": {
                name: sum(1 for row in rows if row["_source_sheet"] == name)
                for name in source_sheet_names(source_wb)
            },
            "sheets": outputs[0]["sheets"],
            "ignored_hidden_sheets": [sheet.title for sheet in source_wb if sheet.sheet_state != "visible"],
            "external_formulas_frozen": len(external_formula_changes),
            "external_formula_changes": external_formula_changes,
        }
    finally:
        source_wb.close()
        template_wb.close()
        for staged in staged_paths:
            if staged.exists():
                staged.unlink()


def process_folder(folder: Path, template_path: Path) -> int:
    print(f"Thư mục làm việc: {folder}")
    print(f"File mẫu 7009   : {template_path}")
    print("=" * 60)

    try:
        validate_template_file(template_path)
    except ImportBuildError as exc:
        print(f"[LỖI FILE MẪU] {exc}", file=sys.stderr)
        return 1

    template_resolved = template_path.resolve()
    excel_files = [
        p
        for p in sorted(folder.iterdir(), key=lambda item: item.name.casefold())
        if p.is_file()
        and p.suffix.lower() == ".xlsx"
        and not p.name.startswith("~$")
        and not p.name.startswith(".")
        and p.resolve() != template_resolved
        and p.name.casefold() != "7009_employee_template.xlsx"
    ]

    if not excel_files:
        print("Không tìm thấy file Excel (.xlsx) nào trong thư mục này.")
        return 0

    out_dir = folder / "result"
    ok_count = 0
    fail_count = 0

    for src in excel_files:
        print(f"\n▶ Đang xử lý: {src.name}")
        target_out = out_dir / src.name
        try:
            summary = build_import_file(src, template_path, target_out)
            ok_count += 1
            for part in summary["outputs"]:
                rel = Path(part["output"]).relative_to(folder)
                print(f"  ✅ Đã tạo: {rel} ({part['generated_rows']} nhân viên)")
            if summary["external_formulas_frozen"]:
                print(f"  ℹ️ Đã đóng băng {summary['external_formulas_frozen']} công thức ngoài sang giá trị đã lưu.")
        except ImportBuildError as exc:
            fail_count += 1
            print(f"  ❌ LỖI: {exc}", flush=True)
        except Exception as exc:
            fail_count += 1
            print(f"  ❌ LỖI KHÔNG MONG ĐỢI: {exc}", flush=True)

    print("\n" + "=" * 60)
    print(f"HOÀN TẤT! Thành công: {ok_count}/{len(excel_files)} file | Lỗi: {fail_count} file.")
    if ok_count > 0:
        print(f"Thư mục kết quả: {out_dir}")
    return 0 if fail_count == 0 else 1


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Tạo file import 7009 từ 1 file Excel hoặc toàn bộ file Excel trong 1 thư mục."
    )
    parser.add_argument(
        "target",
        nargs="?",
        type=Path,
        default=None,
        help="Đường dẫn thư mục chứa các file .xlsx hoặc 1 file .xlsx cụ thể. Mặc định là thư mục hiện tại.",
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=BUNDLED_TEMPLATE,
        help="Đường dẫn file mẫu 7009_Employee_Template.xlsx.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Đường dẫn file .xlsx đầu ra (chỉ dùng khi target là 1 file .xlsx).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    target = args.target if args.target is not None else Path.cwd()

    if target.is_dir():
        return process_folder(target, args.template)

    if target.is_file():
        out_path = args.output if args.output is not None else (target.parent / "result" / target.name)
        try:
            summary = build_import_file(target, args.template, out_path)
        except ImportBuildError as exc:
            print(str(exc), file=sys.stderr)
            return 1
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return 0

    print(f"[LỖI] Không tìm thấy đường dẫn: {target}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
