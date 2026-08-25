"""Local inventory organizer: Excel validation, export engine, and web UI API."""

from __future__ import annotations

import io
import json
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import threading
import uuid
import webbrowser
from copy import copy
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from itertools import islice
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from inventory_logging import configure_logging as configure_diagnostic_logging, log_event


SOURCE_ROOT = Path(__file__).resolve().parent
BUNDLE_ROOT = Path(getattr(sys, "_MEIPASS", SOURCE_ROOT))
WEB_ROOT = BUNDLE_ROOT / "web"
TEMPLATE_PATH = BUNDLE_ROOT / "模板总模板.xlsx"
INSTALL_ROOT = Path(os.environ["INVENTORY_INSTALL_ROOT"]).resolve() if os.environ.get("INVENTORY_INSTALL_ROOT") else None

# Packaged builds keep user files out of the installation directory. Electron
# supplies INVENTORY_DATA_ROOT so upgrades preserve mappings and export history.
if getattr(sys, "frozen", False):
    default_data_root = Path(os.environ.get("LOCALAPPDATA", str(SOURCE_ROOT))) / "库存工作台"
    DATA_ROOT = Path(os.environ.get("INVENTORY_DATA_ROOT", default_data_root))
else:
    DATA_ROOT = SOURCE_ROOT / ".inventory_data"

UPLOAD_ROOT = DATA_ROOT / "uploads"
CONFIG_PATH = DATA_ROOT / "config.json"
DEFAULT_EXPORT_DIR = DATA_ROOT / "exports"
LOG_ROOT = DATA_ROOT / "logs"
LOG_PATH = LOG_ROOT / "inventory-workbench.log"
CONFIG_LOCK = threading.Lock()

REQUIRED_COLUMNS = [
    "物料编码",
    "物料名称",
    "规格型号",
    "仓库名称",
    "批号",
    "库存主单位",
    "库存量(主单位)",
    "库存状态",
    "库存组织",
    "数量（箱）",
]
NUMBER_COLUMNS = ["库存量(主单位)", "数量（箱）"]


def default_config():
    return {
        "settings": {
            "export_path": str(DEFAULT_EXPORT_DIR),
            "compare_filename": "{品牌}_单品牌库存对比_{本周日期}.xlsx",
            "summary_filename": "库存汇总_{生成日期}.xlsx",
            "summary_compare_filename": "汇总对比_{生成日期}.xlsx",
            "special_brand_filename": "{品牌}_品牌拆分_{生成日期}.xlsx",
            "log_enabled": False,
        },
        "mappings": [],
        "export_records": [],
    }


def ensure_directories():
    for directory in (DATA_ROOT, UPLOAD_ROOT, DEFAULT_EXPORT_DIR, LOG_ROOT):
        directory.mkdir(parents=True, exist_ok=True)


def configure_logging(enabled):
    """Configure the opt-in diagnostic logger at the configured log path."""
    configure_diagnostic_logging(enabled, LOG_PATH, ensure_directories)


def load_config():
    ensure_directories()
    if not CONFIG_PATH.exists():
        config = default_config()
        save_config(config)
        return config
    try:
        config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        if CONFIG_PATH.exists():
            backup = CONFIG_PATH.with_name(f"config.corrupt-{datetime.now():%Y%m%d%H%M%S}.json")
            try:
                shutil.copy2(CONFIG_PATH, backup)
            except OSError:
                pass
        config = default_config()
    config.setdefault("settings", default_config()["settings"])
    config.setdefault("mappings", [])
    config.setdefault("export_records", [])
    config_changed = False
    for key, value in default_config()["settings"].items():
        if key not in config["settings"]:
            config["settings"][key] = value
            config_changed = True
    if isinstance(config["mappings"], dict):
        config["mappings"] = [
            {
                "id": uuid.uuid4().hex,
                "brand": normalize_text(brand),
                "source_warehouse": normalize_text(warehouse),
                "target_warehouse": "德清仓",
            }
            for brand, warehouse in config["mappings"].items()
            if normalize_text(brand) and normalize_text(warehouse)
        ]
        config_changed = True
    elif not isinstance(config["mappings"], list):
        config["mappings"] = []
        config_changed = True
    else:
        normalized_mappings = []
        mapping_keys = set()
        for mapping in config["mappings"]:
            if not isinstance(mapping, dict):
                config_changed = True
                continue
            brand = normalize_text(mapping.get("brand"))
            source_warehouse = normalize_text(mapping.get("source_warehouse", mapping.get("warehouse")))
            target_warehouse = normalize_text(mapping.get("target_warehouse")) or "德清仓"
            if not brand or not source_warehouse:
                config_changed = True
                continue
            key = (brand, source_warehouse)
            if key in mapping_keys:
                config_changed = True
                continue
            mapping_keys.add(key)
            normalized_mappings.append({
                "id": normalize_text(mapping.get("id")) or uuid.uuid4().hex,
                "brand": brand,
                "source_warehouse": source_warehouse,
                "target_warehouse": target_warehouse,
            })
        if normalized_mappings != config["mappings"]:
            config["mappings"] = normalized_mappings
            config_changed = True
    if config["settings"]["compare_filename"] == "{品牌}_库存对比_{本周日期}.xlsx":
        config["settings"]["compare_filename"] = default_config()["settings"]["compare_filename"]
        config_changed = True
    for record in config["export_records"]:
        if record.get("type") == "库存对比":
            record["type"] = "单品牌库存对比"
            config_changed = True
    if config_changed:
        save_config(config)
    configure_logging(config["settings"]["log_enabled"])
    return config


def save_config(config):
    ensure_directories()
    temporary = CONFIG_PATH.with_suffix(".tmp")
    with CONFIG_LOCK:
        temporary.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(CONFIG_PATH)


def safe_filename(value):
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip().rstrip(".")
    return cleaned or "库存导出"


def normalize_text(value):
    return "" if value is None else str(value).strip()


class UploadedFormFile:
    """The small uploaded-file interface used by the inventory import methods."""

    def __init__(self, filename, content):
        self.filename = filename
        self.file = io.BytesIO(content)


class MultipartForm:
    """Compatibility wrapper for the FieldStorage operations used by this app."""

    def __init__(self):
        self.values = {}

    def add(self, name, value):
        self.values.setdefault(name, []).append(value)

    def getfirst(self, name, default=None):
        values = self.values.get(name)
        return values[0] if values else default

    def __contains__(self, name):
        return name in self.values

    def __getitem__(self, name):
        values = self.values[name]
        return values[0] if len(values) == 1 else values


def parse_multipart_form(body, content_type):
    """Parse browser multipart uploads without the removed Python 3.14 cgi module."""
    if "multipart/form-data" not in content_type.lower():
        raise ValueError("上传请求格式不正确")
    message = BytesParser(policy=policy.default).parsebytes(
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
    )
    if not message.is_multipart():
        raise ValueError("上传请求格式不正确")
    form = MultipartForm()
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        content = part.get_payload(decode=True) or b""
        filename = part.get_filename()
        if filename is not None:
            form.add(name, UploadedFormFile(filename, content))
        else:
            form.add(name, content.decode(part.get_content_charset() or "utf-8", errors="replace"))
    return form


def decimal_number(value):
    if value is None or (isinstance(value, str) and not value.strip()):
        raise ValueError("数量不能为空")
    try:
        result = Decimal(str(value).strip())
        if not result.is_finite():
            raise ValueError("数量必须是有限数字")
        return result
    except (InvalidOperation, ValueError):
        raise ValueError("数量必须是数字")


def round_decimal(value, places=2):
    """Round decimal values consistently for values shown as boxes or percentages."""
    return Decimal(value).quantize(Decimal("1").scaleb(-places), rounding=ROUND_HALF_UP)


def json_number(value):
    value = Decimal(value)
    if value == value.to_integral():
        return int(value)
    return float(value)


def display_number(value):
    return json_number(value) if value is not None else None


def format_issue(filename, issue):
    """Attach the source workbook location to a validation message."""
    location = f"文件“{filename}”"
    if issue.get("sheet"):
        location += f"，工作表“{issue['sheet']}”"
    if issue.get("row"):
        location += f"，第 {issue['row']} 行"
    return {**issue, "message": f"{location}：{issue['message']}"}


TOTAL_ROW_MARKERS = ("总计", "合计", "小计")


def total_marker(value):
    text = re.sub(r"\s+", "", normalize_text(value))
    for marker in TOTAL_ROW_MARKERS:
        if text == marker or re.fullmatch(rf"{marker}(?:[:：]|[（(].*[）)])", text):
            return marker
    return None


def find_last_total_row(path):
    """Return metadata when the final non-empty row contains a total marker.

    Some Excel clients preserve a trailing styled range or place the total label
    outside the first two columns, so inspect every cell in the final value row.
    """
    try:
        workbook = openpyxl.load_workbook(path, data_only=False, read_only=True)
    except Exception:
        return None
    try:
        return find_last_total_row_in_sheet(workbook.active)
    finally:
        workbook.close()
    return None


def find_last_total_row_in_sheet(sheet):
    for row_number in range(sheet.max_row, 1, -1):
        values = [sheet.cell(row_number, column).value for column in range(1, sheet.max_column + 1)]
        if not any(value is not None and str(value).strip() for value in values):
            continue
        marker = next((total_marker(value) for value in values if total_marker(value)), None)
        return {"row": row_number, "label": marker} if marker else None
    return None


def remove_last_total_row(path):
    """Delete the detected total row from an application-owned upload copy."""
    total_row = find_last_total_row(path)
    if not total_row:
        raise ValueError("文件最后一行未识别为总计行，无需清除")
    workbook = openpyxl.load_workbook(path)
    try:
        sheet = workbook.active
        sheet.delete_rows(total_row["row"], 1)
        workbook.save(path)
    finally:
        workbook.close()
    return total_row


def parse_inventory_workbook(path, brand):
    """Read one workbook and return normalized rows plus validation findings."""
    errors = []
    warnings = []
    try:
        workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:
        return {"rows": [], "errors": [{"message": f"无法读取 Excel 文件：{exc}"}], "warnings": []}

    sheet = workbook.active
    header_values = [normalize_text(cell.value) for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
    missing = [name for name in REQUIRED_COLUMNS if name not in header_values]
    unexpected = [name for name in header_values if name and name not in REQUIRED_COLUMNS]
    if missing or unexpected or len(header_values) != len(REQUIRED_COLUMNS) or header_values != REQUIRED_COLUMNS:
        details = []
        if missing:
            details.append("缺少列：" + "、".join(missing))
        if unexpected:
            details.append("不允许的列：" + "、".join(unexpected))
        if len(header_values) != len(REQUIRED_COLUMNS):
            details.append("列数量与样例不一致")
        elif header_values != REQUIRED_COLUMNS:
            details.append("列顺序与样例不一致，请在“库存量(主单位)”后增加“库存状态”列")
        workbook.close()
        return {"rows": [], "errors": [{"message": "；".join(details), "sheet": sheet.title, "row": 1}], "warnings": []}

    index = {name: header_values.index(name) for name in REQUIRED_COLUMNS}
    total_row = find_last_total_row_in_sheet(sheet)
    rows = []
    specs_by_product = {}
    codes_by_name_spec = {}
    for excel_row, values in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        if total_row and excel_row == total_row["row"]:
            continue
        if not any(value is not None and str(value).strip() for value in values):
            continue
        raw = {name: values[index[name]] if index[name] < len(values) else None for name in REQUIRED_COLUMNS}
        key_fields = ["物料编码", "物料名称", "规格型号", "仓库名称"]
        missing_values = [name for name in key_fields if not normalize_text(raw[name])]
        if missing_values:
            errors.append({"message": "必填字段为空：" + "、".join(missing_values), "sheet": sheet.title, "row": excel_row})
            continue
        numbers = {}
        for column in NUMBER_COLUMNS:
            try:
                numbers[column] = decimal_number(raw[column])
            except ValueError as exc:
                errors.append({"message": f"{column}{exc}", "sheet": sheet.title, "row": excel_row, "field": column, "value": raw[column]})
        if len(numbers) != len(NUMBER_COLUMNS):
            continue
        code = normalize_text(raw["物料编码"])
        name = normalize_text(raw["物料名称"])
        spec = normalize_text(raw["规格型号"])
        warehouse = normalize_text(raw["仓库名称"])
        product_key = (code, name)
        specs_by_product.setdefault(product_key, {}).setdefault(spec, []).append(excel_row)
        codes_by_name_spec.setdefault((name, spec), set()).add(code)
        rows.append(
            {
                "brand": brand,
                "code": code,
                "name": name,
                "spec": spec,
                "warehouse": warehouse,
                "main_qty": numbers["库存量(主单位)"],
                "box_qty": numbers["数量（箱）"],
                "sheet": sheet.title,
                "row": excel_row,
            }
        )
    workbook.close()

    for (code, name), specs in specs_by_product.items():
        if len(specs) > 1:
            row_list = sorted(row for rows in specs.values() for row in rows)
            errors.append({
                "message": f"同一商品键“{code} + {name}”存在不同规格：{'、'.join(sorted(specs))}，请修正源数据",
                "field": "规格型号", "sheet": sheet.title, "row": row_list[0],
                "rows": row_list,
            })
    for (name, spec), codes in codes_by_name_spec.items():
        if len(codes) > 1:
            warnings.append({"message": f"名称“{name}”、规格“{spec}”对应多个物料编码：{'、'.join(sorted(codes))}；确认后将按多行保留", "kind": "same_name_spec"})
    return {"rows": rows, "errors": errors, "warnings": warnings}


def aggregate_rows(rows, include_brand=False):
    grouped = {}
    for row in rows:
        key = (row["code"], row["name"], row["warehouse"])
        if include_brand:
            key = (row["brand"],) + key
        item = grouped.setdefault(
            key,
            {
                "brand": row["brand"], "code": row["code"], "name": row["name"],
                "spec": row["spec"], "warehouse": row["warehouse"],
                "main_qty": Decimal("0"), "box_qty": Decimal("0"),
            },
        )
        item["main_qty"] += row["main_qty"]
        item["box_qty"] += row["box_qty"]
    return grouped


def make_compare_result(brand, previous_rows, current_rows):
    previous = aggregate_rows(previous_rows)
    current = aggregate_rows(current_rows)
    keys = sorted(set(previous) | set(current), key=lambda key: (key[1], key[0], key[2]))
    records = []
    for key in keys:
        before = previous.get(key)
        after = current.get(key)
        sample = after or before
        before_main = before["main_qty"] if before else Decimal("0")
        before_box = before["box_qty"] if before else Decimal("0")
        after_main = after["main_qty"] if after else Decimal("0")
        after_box = after["box_qty"] if after else Decimal("0")
        main_change = after_main - before_main
        before_box = round_decimal(before_box)
        after_box = round_decimal(after_box)
        box_change = round_decimal(after_box - before_box)
        if after and not before:
            change_type = "新增商品"
        elif before and not after:
            change_type = "删除商品"
        elif main_change == 0 and box_change == 0:
            change_type = "无变化"
        elif main_change > 0 or (main_change == 0 and box_change > 0):
            change_type = "库存增加"
        else:
            change_type = "库存减少"
        records.append(
            {
                "brand": brand, "code": sample["code"], "name": sample["name"], "spec": sample["spec"], "warehouse": sample["warehouse"],
                "previous_main": before_main, "previous_box": before_box,
                "current_main": after_main, "current_box": after_box,
                "main_change": main_change, "box_change": box_change,
                "main_mom_change": percentage_change(before_main, after_main),
                "change_type": change_type,
            }
        )
    return records


def percentage_change(previous, current):
    """Return a signed relative change, or None when the base value is zero."""
    if previous == 0:
        return None
    return round_decimal(((current - previous) / abs(previous)), places=4)


def find_compare_spec_errors(previous_rows, current_rows, previous_filename, current_filename):
    """Reject a specification change that would otherwise be hidden by the compare key."""
    previous_specs = {}
    current_specs = {}
    for row in previous_rows:
        previous_specs.setdefault((row["code"], row["name"]), {}).setdefault(row["spec"], []).append(row)
    for row in current_rows:
        current_specs.setdefault((row["code"], row["name"]), {}).setdefault(row["spec"], []).append(row)
    errors = []
    for key in sorted(set(previous_specs) & set(current_specs)):
        if set(previous_specs[key]) != set(current_specs[key]):
            previous_locations = "、".join(f"{row['sheet']} 第 {row['row']} 行" for rows in previous_specs[key].values() for row in rows)
            current_locations = "、".join(f"{row['sheet']} 第 {row['row']} 行" for rows in current_specs[key].values() for row in rows)
            errors.append(
                {
                    "message": (
                        f"商品键“{key[0]} + {key[1]}”在上周与本周的规格不一致："
                        f"上周文件“{previous_filename}”（{previous_locations}）为“{'、'.join(sorted(previous_specs[key]))}”，"
                        f"本周文件“{current_filename}”（{current_locations}）为“{'、'.join(sorted(current_specs[key]))}”，请修正源数据后再比较"
                    ),
                    "field": "规格型号",
                }
            )
    return errors


def find_cross_code_warnings(imports):
    values = {}
    for item in imports:
        for row in item["rows"]:
            values.setdefault((row["name"], row["spec"]), set()).add(row["code"])
    messages = []
    for (name, spec), codes in values.items():
        if len(codes) > 1:
            messages.append({"message": f"名称“{name}”、规格“{spec}”对应多个物料编码：{'、'.join(sorted(codes))}；确认后会分行展示", "kind": "same_name_spec"})
    return messages


def build_summary_result(imports, mappings):
    mapping_targets = {
        (mapping["brand"], mapping["source_warehouse"]): mapping["target_warehouse"]
        for mapping in mappings
    }
    warehouse_order = []
    result_rows = []
    for imported in imports:
        brand = imported["brand"]
        product_rows = {}
        for row in imported["rows"]:
            warehouse = mapping_targets.get((brand, row["warehouse"]), row["warehouse"])
            if warehouse not in warehouse_order:
                warehouse_order.append(warehouse)
            product_key = (row["code"], row["name"])
            product = product_rows.setdefault(
                product_key,
                {"brand": brand, "code": row["code"], "name": row["name"], "spec": row["spec"], "warehouses": {}},
            )
            cell = product["warehouses"].setdefault(warehouse, {"main_qty": Decimal("0"), "box_qty": Decimal("0")})
            cell["main_qty"] += row["main_qty"]
            cell["box_qty"] += row["box_qty"]
        for serial, product in enumerate(sorted(product_rows.values(), key=lambda item: (item["name"], item["code"])), start=1):
            product["serial"] = serial
            for cell in product["warehouses"].values():
                cell["main_qty"] = round_decimal(cell["main_qty"])
                cell["box_qty"] = round_decimal(cell["box_qty"])
            product["total_main"] = sum((item["main_qty"] for item in product["warehouses"].values()), Decimal("0"))
            product["total_box"] = sum((item["box_qty"] for item in product["warehouses"].values()), Decimal("0"))
            result_rows.append(product)
    if "德清仓" in warehouse_order:
        warehouse_order.remove("德清仓")
        warehouse_order.insert(0, "德清仓")
    unconfigured = []
    for item in imports:
        source_warehouses = {row["warehouse"] for row in item["rows"]}
        configured_sources = {
            source
            for (brand, source), _target in mapping_targets.items()
            if brand == item["brand"]
        }
        if not configured_sources:
            unconfigured.append(
                {
                    "brand": item["brand"],
                    "warehouses": sorted(source_warehouses),
                    "reason": "尚未配置仓库映射",
                }
            )
            continue
        stale_sources = sorted(configured_sources - source_warehouses)
        if stale_sources:
            unconfigured.append(
                {
                    "brand": item["brand"],
                    "warehouses": sorted(source_warehouses),
                    "reason": f"已配置的源仓库“{'、'.join(stale_sources)}”未在本次导入文件中找到",
                }
            )
    return {"warehouses": warehouse_order, "rows": result_rows, "unconfigured": unconfigured}


def new_workbook_with_title(title, column_count, sheet_name="单品牌库存对比"):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    sheet.sheet_view.showGridLines = False
    sheet.append([title])
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=column_count)
    title_cell = sheet["A1"]
    title_cell.font = Font(name="Microsoft YaHei", size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill("solid", fgColor="124E3B")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 29
    return workbook, sheet


def apply_table_styles(sheet, header_row, last_row, column_count):
    header_fill = PatternFill("solid", fgColor="DCECE3")
    alt_fill = PatternFill("solid", fgColor="F6F8F6")
    border = Border(
        left=Side(style="thin", color="DCE2DD"), right=Side(style="thin", color="DCE2DD"),
        top=Side(style="thin", color="DCE2DD"), bottom=Side(style="thin", color="DCE2DD"),
    )
    for cell in sheet[header_row]:
        if cell.column <= column_count:
            cell.font = Font(name="Microsoft YaHei", bold=True, color="124E3B")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
    for row in sheet.iter_rows(min_row=header_row + 1, max_row=last_row, max_col=column_count):
        for cell in row:
            cell.font = Font(name="Microsoft YaHei", size=10)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if cell.row % 2 == 0:
                cell.fill = alt_fill
    sheet.freeze_panes = f"A{header_row + 1}"
    sheet.auto_filter.ref = f"A{header_row}:{get_column_letter(column_count)}{last_row}"


def export_compare(result, target):
    headers = [
        "品牌", "物料编码", "物料名称", "规格型号", "仓库名称",
        "上周主单位数量", "本周主单位数量", "主单位变化值",
        "上周箱数", "本周箱数", "箱数变化值", "环比变化", "变动类型",
    ]
    workbook, sheet = new_workbook_with_title(f"{result['brand']} 单品牌库存对比表", len(headers))
    sheet.append(headers)
    for row in result["records"]:
        sheet.append([
            row["brand"], row["code"], row["name"], row["spec"], row["warehouse"],
            display_number(round_decimal(row["previous_main"])), display_number(round_decimal(row["current_main"])), display_number(row["main_change"]),
            display_number(row["previous_box"]), display_number(row["current_box"]), display_number(row["box_change"]),
            display_number(row["main_mom_change"]), row["change_type"],
        ])
    apply_table_styles(sheet, 2, sheet.max_row, len(headers))
    widths = [14, 14, 22, 18, 16, 15, 15, 15, 12, 12, 12, 15, 14]
    for column, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width
    for row in range(3, sheet.max_row + 1):
        for column in (6, 7, 9, 10, 11):
            sheet.cell(row, column).number_format = "0.00"
        sheet.cell(row, 12).number_format = "0.00%"
    for row in range(3, sheet.max_row + 1):
        type_cell = sheet.cell(row, len(headers))
        if type_cell.value in ("新增商品", "库存增加"):
            type_cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color="127145")
        elif type_cell.value in ("删除商品", "库存减少"):
            type_cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color="B8392E")
    workbook.save(target)


def export_summary(result, target):
    if not TEMPLATE_PATH.is_file():
        raise ValueError("汇总模板不可用，请重新安装库存工作台后再导出")
    workbook = openpyxl.load_workbook(TEMPLATE_PATH)
    sheet = workbook.active
    template_title = normalize_text(sheet.cell(1, 1).value) or "运营公司库存项目明细表"
    sheet.sheet_view.showGridLines = False
    for merged in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged))
    if sheet.max_row >= 4:
        sheet.delete_rows(4, sheet.max_row - 3)
    last_column = 4 + len(result["warehouses"]) * 2 + 2
    for row in range(1, 4):
        for column in range(1, max(last_column, 14) + 1):
            cell = sheet.cell(row, column)
            cell.value = None
            cell._style = copy(sheet.cell(3 if row == 3 else 2, min(column, 14))._style)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_column)
    title = sheet.cell(1, 1)
    title.value = template_title
    title.font = Font(name="Microsoft YaHei", size=15, bold=True, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor="124E3B")
    title.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 28
    base_headers = ["品牌", "序号", "名称", "规格"]
    for column, value in enumerate(base_headers, start=1):
        sheet.cell(3, column).value = value
        sheet.cell(3, column).font = Font(name="Microsoft YaHei", bold=True, color="124E3B")
        sheet.cell(3, column).fill = PatternFill("solid", fgColor="DCECE3")
    for column in range(1, 5):
        sheet.cell(2, column).fill = PatternFill("solid", fgColor="DCECE3")
    data_column = 5
    for warehouse in result["warehouses"] + ["合计"]:
        sheet.merge_cells(start_row=2, start_column=data_column, end_row=2, end_column=data_column + 1)
        group_cell = sheet.cell(2, data_column)
        group_cell.value = warehouse
        group_cell.font = Font(name="Microsoft YaHei", bold=True, color="124E3B")
        group_cell.fill = PatternFill("solid", fgColor="DCECE3" if warehouse != "德清仓" else "CDE8D8")
        group_cell.alignment = Alignment(horizontal="center", vertical="center")
        for offset, label in enumerate(["数量（主单位）", "数量（箱）"]):
            cell = sheet.cell(3, data_column + offset)
            cell.value = label
            cell.font = Font(name="Microsoft YaHei", bold=True, color="124E3B")
            cell.fill = PatternFill("solid", fgColor="DCECE3")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        data_column += 2
    border = Border(
        left=Side(style="thin", color="DCE2DD"), right=Side(style="thin", color="DCE2DD"),
        top=Side(style="thin", color="DCE2DD"), bottom=Side(style="thin", color="DCE2DD"),
    )
    start_row = 4
    grouped_rows = {}
    for row in result["rows"]:
        grouped_rows.setdefault(row["brand"], []).append(row)
    current_row = start_row
    for brand, items in grouped_rows.items():
        brand_start = current_row
        for item in items:
            values = [brand, item["serial"], item["name"], item["spec"]]
            for warehouse in result["warehouses"]:
                values.extend([
                    display_number(round_decimal(item["warehouses"].get(warehouse, {}).get("main_qty"))) if warehouse in item["warehouses"] else None,
                    display_number(item["warehouses"].get(warehouse, {}).get("box_qty")) if warehouse in item["warehouses"] else None,
                ])
            values.extend([display_number(round_decimal(item["total_main"])), display_number(item["total_box"])])
            for column, value in enumerate(values, start=1):
                cell = sheet.cell(current_row, column, value)
                cell.font = Font(name="Microsoft YaHei", size=10)
                cell.border = border
                cell.alignment = Alignment(horizontal="center" if column != 3 else "left", vertical="center", wrap_text=True)
                if current_row % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="F6F8F6")
            current_row += 1
        if current_row - brand_start > 1:
            sheet.merge_cells(start_row=brand_start, start_column=1, end_row=current_row - 1, end_column=1)
            sheet.cell(brand_start, 1).alignment = Alignment(horizontal="center", vertical="center")
    for row in range(2, 4):
        for column in range(1, last_column + 1):
            sheet.cell(row, column).border = border
    widths = [14, 8, 24, 18] + [14, 11] * (len(result["warehouses"]) + 1)
    for column, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width
    for row in range(start_row, current_row):
        for column in range(5, last_column + 1):
            sheet.cell(row, column).number_format = "0.00"
    sheet.freeze_panes = "E4"
    workbook.save(target)


def find_summary_total_column(sheet):
    """Locate required summary columns in either grouped or flat header layouts."""
    required_labels = {"品牌", "序号", "名称", "规格"}
    for header_row in range(1, min(sheet.max_row, 20) + 1):
        labels = [normalize_text(sheet.cell(header_row, column).value) for column in range(1, sheet.max_column + 1)]
        if not required_labels.issubset(set(labels)):
            continue
        brand_column = labels.index("品牌") + 1
        name_column = labels.index("名称") + 1
        spec_column = labels.index("规格") + 1
        group_row = header_row - 1
        if group_row >= 1:
            for column in range(1, sheet.max_column + 1):
                if normalize_text(sheet.cell(group_row, column).value) != "合计":
                    continue
                if "主单位" in normalize_text(sheet.cell(header_row, column).value):
                    return {
                        "header_row": header_row,
                        "brand_column": brand_column,
                        "name_column": name_column,
                        "spec_column": spec_column,
                        "total_main_column": column,
                        "total_last_column": column + 1,
                        "header_layout": "grouped",
                    }
        if "合计" in labels:
            total_column = labels.index("合计") + 1
            return {
                "header_row": header_row,
                "brand_column": brand_column,
                "name_column": name_column,
                "spec_column": spec_column,
                "total_main_column": total_column,
                "total_last_column": total_column,
                "header_layout": "flat",
            }
    raise ValueError("未识别到汇总表关键列，请确保表头包含“品牌、序号、名称、规格、合计”")


def find_summary_total_column_in_rows(rows):
    """Locate summary columns from a small sequential header slice."""
    required_labels = {"品牌", "序号", "名称", "规格"}
    normalized_rows = [[normalize_text(value) for value in row] for row in rows]
    for index, labels in enumerate(normalized_rows):
        if not required_labels.issubset(set(labels)):
            continue
        brand_column = labels.index("品牌")
        name_column = labels.index("名称")
        spec_column = labels.index("规格")
        if index:
            group_labels = normalized_rows[index - 1]
            for column, group_label in enumerate(group_labels):
                if group_label == "合计" and column < len(labels) and "主单位" in labels[column]:
                    return {
                        "header_row": index + 1,
                        "brand_column": brand_column,
                        "name_column": name_column,
                        "spec_column": spec_column,
                        "total_main_column": column,
                        "header_layout": "grouped",
                    }
        if "合计" in labels:
            return {
                "header_row": index + 1,
                "brand_column": brand_column,
                "name_column": name_column,
                "spec_column": spec_column,
                "total_main_column": labels.index("合计"),
                "header_layout": "flat",
            }
    raise ValueError("未识别到汇总表关键列，请确保表头包含“品牌、序号、名称、规格、合计”")


def parse_summary_compare_workbook(path):
    """Read totals from a workbook with the required summary columns.

    The workbook is intentionally consumed as a forward-only value stream. Calling
    ``sheet.cell`` for every value on a read-only worksheet repeatedly seeks within
    its XML stream and becomes prohibitively slow for large summary workbooks.
    """
    try:
        workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:
        raise ValueError(f"无法读取汇总表：{exc}") from exc
    try:
        sheet = workbook.active
        # Some third-party workbooks omit worksheet dimensions, leaving max_row
        # unknown in read-only mode. A bounded stream works for both cases.
        header_rows = list(islice(sheet.iter_rows(values_only=True), 20))
        columns = find_summary_total_column_in_rows(header_rows)
        records = []
        current_brand = ""
        for row_number, values in enumerate(
            sheet.iter_rows(min_row=columns["header_row"] + 1, values_only=True),
            start=columns["header_row"] + 1,
        ):
            if any(total_marker(value) for value in values):
                continue
            brand = normalize_text(values[columns["brand_column"]] if columns["brand_column"] < len(values) else None) or current_brand
            name = normalize_text(values[columns["name_column"]] if columns["name_column"] < len(values) else None)
            spec = normalize_text(values[columns["spec_column"]] if columns["spec_column"] < len(values) else None)
            if not any((brand, name, spec)):
                continue
            if not all((brand, name, spec)):
                raise ValueError(f"汇总表第 {row_number} 行缺少品牌、名称或规格，无法进行汇总对比")
            current_brand = brand
            try:
                total_main = decimal_number(
                    values[columns["total_main_column"]]
                    if columns["total_main_column"] < len(values)
                    else None
                )
            except ValueError as exc:
                raise ValueError(f"汇总表第 {row_number} 行合计主单位{exc}") from exc
            records.append({"brand": brand, "name": name, "spec": spec, "total_main": total_main})
        return records
    finally:
        workbook.close()


def aggregate_summary_totals(records):
    totals = {}
    for row in records:
        key = (row["brand"], row["name"], row["spec"])
        totals[key] = totals.get(key, Decimal("0")) + row["total_main"]
    return totals


def make_summary_compare_result(previous_records, current_records):
    previous = aggregate_summary_totals(previous_records)
    current = aggregate_summary_totals(current_records)
    rows = []
    for key in sorted(set(previous) | set(current), key=lambda value: (value[0], value[1], value[2])):
        before = previous.get(key, Decimal("0"))
        after = current.get(key, Decimal("0"))
        is_new = key not in previous and key in current
        rows.append(
            {
                "brand": key[0], "name": key[1], "spec": key[2],
                "previous_total_main": before, "current_total_main": after,
                "total_main_change": after - before,
                "main_mom_change": percentage_change(before, after),
                "remark": "新增物料" if is_new else "",
            }
        )
    return rows


def export_summary_compare(result, target):
    """Append comparison columns to the current summary workbook without losing its layout."""
    source_path = Path(result["current_path"])
    workbook = openpyxl.load_workbook(source_path)
    try:
        sheet = workbook.active
        columns = find_summary_total_column(sheet)
        comparison_rows = {
            (row["brand"], row["name"], row["spec"]): row
            for row in result["records"]
        }
        original_last_column = sheet.max_column
        total_main_column = columns["total_main_column"]
        total_last_column = columns["total_last_column"]
        ratio_column = total_last_column + 1
        remark_column = total_last_column + 2

        # The summary title spans the full table width, so extend it for the two new columns.
        title_range = next(
            (
                str(merged)
                for merged in sheet.merged_cells.ranges
                if merged.min_row == merged.max_row == 1 and merged.min_col == 1
            ),
            None,
        )
        if title_range:
            sheet.unmerge_cells(title_range)
        sheet.insert_cols(ratio_column, amount=2)
        if title_range:
            sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=original_last_column + 2)
            sheet.cell(1, 1).value = "库存汇总对比表"

        header_row = columns["header_row"]
        for column, label in ((ratio_column, "环比"), (remark_column, "备注")):
            sheet.cell(header_row, column)._style = copy(sheet.cell(header_row, total_main_column)._style)
            if columns["header_layout"] == "grouped":
                group_row = header_row - 1
                sheet.cell(group_row, column)._style = copy(sheet.cell(group_row, total_main_column)._style)
                sheet.cell(group_row, column).value = label
                sheet.cell(group_row, column).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                sheet.merge_cells(start_row=group_row, start_column=column, end_row=header_row, end_column=column)
            else:
                sheet.cell(header_row, column).value = label
                sheet.cell(header_row, column).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            sheet.column_dimensions[get_column_letter(column)].width = 15 if label == "环比" else 16

        current_brand = ""
        for row_number in range(header_row + 1, sheet.max_row + 1):
            brand = normalize_text(sheet.cell(row_number, columns["brand_column"]).value) or current_brand
            name = normalize_text(sheet.cell(row_number, columns["name_column"]).value)
            spec = normalize_text(sheet.cell(row_number, columns["spec_column"]).value)
            if not any((brand, name, spec)):
                continue
            if not all((brand, name, spec)):
                continue
            current_brand = brand
            comparison = comparison_rows.get((brand, name, spec))
            if not comparison:
                continue
            ratio_cell = sheet.cell(row_number, ratio_column)
            remark_cell = sheet.cell(row_number, remark_column)
            ratio_cell._style = copy(sheet.cell(row_number, total_main_column)._style)
            remark_cell._style = copy(sheet.cell(row_number, total_main_column)._style)
            ratio_cell.value = display_number(comparison["main_mom_change"])
            ratio_cell.number_format = "0.00%"
            remark_cell.value = comparison["remark"]
            if comparison["remark"] == "新增物料":
                remark_cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color="127145")
        sheet.freeze_panes = sheet.freeze_panes or "E4"
        workbook.save(target)
    finally:
        workbook.close()


def find_warehouse_column(sheet):
    for row_number in range(1, min(sheet.max_row, 10) + 1):
        for column in range(1, sheet.max_column + 1):
            if normalize_text(sheet.cell(row_number, column).value) == "仓库名称":
                return {"header_row": row_number, "warehouse_column": column}
    raise ValueError("未识别到“仓库名称”列，无法拆分特殊品牌表")


def split_special_brand_workbook(path, export_directory, filename_pattern="{品牌}_品牌拆分_{生成日期}.xlsx"):
    """Split a combined source workbook while preserving its columns and styles."""
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if len(workbook.worksheets) != 1:
            raise ValueError("特殊品牌拆分仅支持单工作表 Excel，避免其他工作表保留未拆分的数据")
        sheet = workbook.active
        columns = find_warehouse_column(sheet)
        grouped_rows = {}
        rows_without_separator = []
        for row_number in range(columns["header_row"] + 1, sheet.max_row + 1):
            values = [sheet.cell(row_number, column).value for column in range(1, sheet.max_column + 1)]
            if not any(value is not None and str(value).strip() for value in values):
                continue
            warehouse = normalize_text(sheet.cell(row_number, columns["warehouse_column"]).value)
            if "-" not in warehouse:
                rows_without_separator.append(row_number)
                continue
            brand, short_warehouse = (normalize_text(value) for value in warehouse.split("-", 1))
            if not brand or not short_warehouse:
                raise ValueError(f"第 {row_number} 行仓库名称“{warehouse}”格式不完整，应为“品牌名-仓库名”")
            grouped_rows.setdefault(brand, []).append((row_number, short_warehouse))
    finally:
        workbook.close()
    if not grouped_rows:
        raise ValueError("未发现“品牌名-仓库名”格式的仓库名称，无需进行特殊品牌拆分")
    if rows_without_separator:
        row_labels = "、".join(str(row) for row in rows_without_separator[:10])
        suffix = "等" if len(rows_without_separator) > 10 else ""
        raise ValueError(f"第 {row_labels}{suffix} 行仓库名称不含“-”，为避免遗漏数据，本次未拆分；请先整理为统一格式后再试")

    targets = []
    for brand, row_values in grouped_rows.items():
        output = unique_export_path(
            export_directory,
            filename_from_pattern(
                filename_pattern,
                {"品牌": brand, "生成日期": datetime.now().strftime("%Y-%m-%d")},
            ),
        )
        source_book = openpyxl.load_workbook(path)
        try:
            source_sheet = source_book.active
            selected = {row_number: short_warehouse for row_number, short_warehouse in row_values}
            for row_number in range(source_sheet.max_row, columns["header_row"], -1):
                if row_number not in selected:
                    source_sheet.delete_rows(row_number, 1)
            for offset, short_warehouse in enumerate(selected.values(), start=columns["header_row"] + 1):
                source_sheet.cell(offset, columns["warehouse_column"]).value = short_warehouse
            source_book.save(output)
        finally:
            source_book.close()
        targets.append({"brand": brand, "path": output, "rows": len(row_values)})
    return targets


def validate_export_directory(directory):
    directory = Path(directory).resolve()
    if not directory.exists() or not directory.is_dir():
        raise ValueError("导出路径不存在，请在设置中选择一个已存在的文件夹")
    if INSTALL_ROOT:
        try:
            directory.relative_to(INSTALL_ROOT)
        except ValueError:
            pass
        else:
            raise ValueError("导出路径不能位于软件安装目录，请在设置中选择其他文件夹")
    try:
        probe = directory / f".write_probe_{uuid.uuid4().hex}"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
    except OSError as exc:
        raise ValueError(f"导出路径不可写：{exc}") from exc
    return directory


def unique_export_path(directory, filename):
    directory = validate_export_directory(directory)
    base = safe_filename(filename)
    if not base.lower().endswith(".xlsx"):
        base += ".xlsx"
    candidate = directory / base
    index = 1
    while candidate.exists():
        candidate = directory / f"{Path(base).stem}_{datetime.now():%H%M%S}_{index}{Path(base).suffix}"
        index += 1
    return candidate


def filename_from_pattern(pattern, variables):
    output = pattern
    for key, value in variables.items():
        output = output.replace("{" + key + "}", value)
    return safe_filename(output)
