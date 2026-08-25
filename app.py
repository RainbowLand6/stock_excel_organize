"""Local inventory organizer: Excel validation, export engine, and web UI API."""

from __future__ import annotations

import io
import json
import logging
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
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

from logging.handlers import RotatingFileHandler

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


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
LOGGER = logging.getLogger("inventory_workbench")

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
    """Write diagnostic events only when the user has opted in."""
    for handler in LOGGER.handlers[:]:
        LOGGER.removeHandler(handler)
        handler.close()
    LOGGER.propagate = False
    LOGGER.setLevel(logging.INFO)
    if not enabled:
        return
    ensure_directories()
    handler = RotatingFileHandler(LOG_PATH, maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    LOGGER.addHandler(handler)


def log_event(message, *args):
    if LOGGER.handlers:
        LOGGER.info(message, *args)


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
        stale_sources = sorted(
            source for (brand, source), _target in mapping_targets.items()
            if brand == item["brand"] and source not in source_warehouses
        )
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
                    display_number(item["warehouses"].get(warehouse, {}).get("main_qty")) if warehouse in item["warehouses"] else None,
                    display_number(item["warehouses"].get(warehouse, {}).get("box_qty")) if warehouse in item["warehouses"] else None,
                ])
            values.extend([display_number(item["total_main"]), display_number(item["total_box"])])
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
        for column in range(6, last_column + 1, 2):
            sheet.cell(row, column).number_format = "0.00"
    sheet.freeze_panes = "E4"
    workbook.save(target)


def find_summary_total_column(sheet):
    """Locate the total main-unit column in an exported summary workbook."""
    for header_row in range(2, min(sheet.max_row, 10) + 1):
        labels = [normalize_text(sheet.cell(header_row, column).value) for column in range(1, sheet.max_column + 1)]
        if not {"品牌", "名称", "规格"}.issubset(set(labels)):
            continue
        brand_column = labels.index("品牌") + 1
        name_column = labels.index("名称") + 1
        spec_column = labels.index("规格") + 1
        group_row = header_row - 1
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
                }
    raise ValueError("未识别到汇总表的“品牌、名称、规格、合计（主单位）”列，请上传本软件导出的库存汇总表")


def parse_summary_compare_workbook(path):
    """Read totals from a workbook produced by the inventory summary export."""
    try:
        workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:
        raise ValueError(f"无法读取汇总表：{exc}") from exc
    try:
        sheet = workbook.active
        columns = find_summary_total_column(sheet)
        records = []
        current_brand = ""
        for row_number in range(columns["header_row"] + 1, sheet.max_row + 1):
            brand = normalize_text(sheet.cell(row_number, columns["brand_column"]).value) or current_brand
            name = normalize_text(sheet.cell(row_number, columns["name_column"]).value)
            spec = normalize_text(sheet.cell(row_number, columns["spec_column"]).value)
            if not any((brand, name, spec)):
                continue
            if not all((brand, name, spec)):
                raise ValueError(f"汇总表第 {row_number} 行缺少品牌、名称或规格，无法进行汇总对比")
            current_brand = brand
            try:
                total_main = decimal_number(sheet.cell(row_number, columns["total_main_column"]).value)
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
        total_last_column = columns["total_main_column"] + 1
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
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=original_last_column + 2)
        sheet.cell(1, 1).value = "库存汇总对比表"

        header_row = columns["header_row"]
        group_row = header_row - 1
        for column, label in ((ratio_column, "同比"), (remark_column, "备注")):
            sheet.cell(group_row, column)._style = copy(sheet.cell(group_row, total_last_column)._style)
            sheet.cell(header_row, column)._style = copy(sheet.cell(header_row, total_last_column)._style)
            sheet.cell(group_row, column).value = label
            sheet.cell(group_row, column).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            sheet.merge_cells(start_row=group_row, start_column=column, end_row=header_row, end_column=column)
            sheet.column_dimensions[get_column_letter(column)].width = 15 if label == "同比" else 16

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
            ratio_cell._style = copy(sheet.cell(row_number, total_last_column)._style)
            remark_cell._style = copy(sheet.cell(row_number, total_last_column)._style)
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


class InventoryApp:
    def __init__(self):
        self.config = load_config()
        self.summary_imports = []
        self.compare_results = {}
        self.summary_results = {}
        self.summary_compare_uploads = {}
        self.summary_compare_results = {}
        self.compare_uploads = {}
        self.special_brand_upload = None
        self.special_brand_results = []
        self.operation_lock = threading.RLock()

    def state(self):
        return {
            "settings": self.config["settings"],
            "mappings": sorted(
                self.config["mappings"],
                key=lambda item: (item["brand"], item["source_warehouse"], item["target_warehouse"]),
            ),
            "exports": self.config["export_records"],
            "summary_compare_uploads": {
                slot: self.summary_compare_upload_overview(slot)
                for slot in ("previous", "current")
                if slot in self.summary_compare_uploads
            },
            "summary_imports": [self.import_overview(item) for item in self.summary_imports],
            "compare_uploads": {
                slot: self.compare_upload_overview(slot)
                for slot in ("previous", "current")
                if slot in self.compare_uploads
            },
            "special_brand_upload": self.special_brand_upload_overview(),
            "special_brand_results": self.special_brand_results,
        }

    @staticmethod
    def import_overview(item):
        return {
            "brand": item["brand"], "filename": item["filename"], "warehouses": sorted({row["warehouse"] for row in item["rows"]}),
            "rows": len(item["rows"]),
            "errors": [format_issue(item["filename"], error) for error in item["errors"]],
            "warnings": [format_issue(item["filename"], warning) for warning in item["warnings"]],
            "total_row": item.get("total_row"),
            "total_row_removed": item.get("total_row_removed"),
        }

    def save_upload(self, file_item):
        original = Path(file_item.filename or "库存文件.xlsx").name
        suffix = Path(original).suffix.lower()
        if suffix != ".xlsx":
            raise ValueError("仅支持 .xlsx Excel 文件")
        saved = UPLOAD_ROOT / f"{uuid.uuid4().hex}_{safe_filename(original)}"
        with saved.open("wb") as destination:
            shutil.copyfileobj(file_item.file, destination)
        log_event("uploaded file=%s", original)
        return saved, original

    @staticmethod
    def delete_uploads(paths):
        for value in paths:
            path = Path(value)
            try:
                path.relative_to(UPLOAD_ROOT)
            except ValueError:
                continue
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass

    def import_summary_files(self, file_items):
        brands = [Path(Path(file_item.filename or "").name).stem for file_item in file_items if file_item.filename]
        duplicates = sorted({brand for brand in brands if brands.count(brand) > 1})
        if duplicates:
            raise ValueError("同一次汇总不能重复导入同一品牌文件：" + "、".join(duplicates))
        existing_brands = {item["brand"] for item in self.summary_imports}
        conflicts = sorted(existing_brands & set(brands))
        if conflicts:
            raise ValueError("以下品牌已在本次汇总中，请先删除旧文件后再重新导入：" + "、".join(conflicts))
        imported = []
        for file_item in file_items:
            path, filename = self.save_upload(file_item)
            brand = Path(filename).stem
            parsed = parse_inventory_workbook(path, brand)
            imported.append(
                {
                    "brand": brand, "filename": filename, "path": str(path),
                    "total_row": find_last_total_row(path), "total_row_removed": None,
                    **parsed,
                }
            )
        self.summary_imports.extend(imported)
        self.summary_results.clear()
        return [self.import_overview(item) for item in self.summary_imports]

    def delete_summary_import(self, brand):
        for index, item in enumerate(self.summary_imports):
            if item["brand"] == brand:
                self.summary_imports.pop(index)
                self.delete_uploads([item["path"]])
                self.summary_results.clear()
                return [self.import_overview(value) for value in self.summary_imports]
        raise ValueError(f"本次汇总中未找到品牌“{brand}”")

    def rename_summary_import(self, brand, new_brand):
        new_brand = normalize_text(new_brand)
        if not new_brand:
            raise ValueError("品牌名称不能为空")
        item = next((value for value in self.summary_imports if value["brand"] == brand), None)
        if not item:
            raise ValueError(f"本次汇总中未找到品牌“{brand}”")
        if new_brand == brand:
            return [self.import_overview(value) for value in self.summary_imports]
        if any(value["brand"] == new_brand for value in self.summary_imports):
            raise ValueError(f"本次汇总中已存在品牌“{new_brand}”")
        item["brand"] = new_brand
        for row in item["rows"]:
            row["brand"] = new_brand
        self.summary_results.clear()
        return [self.import_overview(value) for value in self.summary_imports]

    def reorder_summary_imports(self, brands):
        existing = [item["brand"] for item in self.summary_imports]
        if len(brands) != len(existing) or set(brands) != set(existing) or len(set(brands)) != len(brands):
            raise ValueError("品牌排序数据已失效，请刷新后重试")
        by_brand = {item["brand"]: item for item in self.summary_imports}
        self.summary_imports = [by_brand[brand] for brand in brands]
        self.summary_results.clear()
        return [self.import_overview(item) for item in self.summary_imports]

    def remove_summary_total_row(self, brand):
        item = next((value for value in self.summary_imports if value["brand"] == brand), None)
        if not item:
            raise ValueError("本次汇总中未找到该品牌文件")
        removed = remove_last_total_row(item["path"])
        item.update(parse_inventory_workbook(item["path"], item["brand"]))
        item["total_row"] = None
        item["total_row_removed"] = removed
        self.summary_results.clear()
        return [self.import_overview(value) for value in self.summary_imports]

    def stage_compare_file(self, slot, file_item):
        if slot not in {"previous", "current"}:
            raise ValueError("对比文件类型不正确")
        path, filename = self.save_upload(file_item)
        previous = self.compare_uploads.get(slot)
        if previous:
            self.delete_uploads([previous["path"]])
        self.compare_uploads[slot] = {
            "filename": filename, "path": str(path), "total_row": find_last_total_row(path), "total_row_removed": None,
        }
        self.compare_results.clear()
        return self.compare_upload_overview(slot)

    def stage_summary_compare_file(self, slot, file_item):
        if slot not in {"previous", "current"}:
            raise ValueError("汇总对比文件类型不正确")
        path, filename = self.save_upload(file_item)
        previous = self.summary_compare_uploads.get(slot)
        if previous:
            self.delete_uploads([previous["path"]])
        self.summary_compare_uploads[slot] = {"filename": filename, "path": str(path)}
        self.summary_compare_results.clear()
        return self.summary_compare_upload_overview(slot)

    def summary_compare_upload_overview(self, slot):
        item = self.summary_compare_uploads.get(slot)
        return {"slot": slot, "filename": item["filename"]} if item else None

    def compare_summaries(self):
        previous = self.summary_compare_uploads.get("previous")
        current = self.summary_compare_uploads.get("current")
        if not previous or not current:
            raise ValueError("请先上传上周和本周两张库存汇总表")
        records = make_summary_compare_result(
            parse_summary_compare_workbook(previous["path"]),
            parse_summary_compare_workbook(current["path"]),
        )
        result_id = uuid.uuid4().hex
        self.summary_compare_results[result_id] = {"records": records, "current_path": current["path"]}
        log_event("summary compared previous=%s current=%s records=%d", previous["filename"], current["filename"], len(records))
        return {"result_id": result_id, "records": serialize_summary_compare(records)}

    def export_summary_compare(self, result_id):
        result = self.summary_compare_results.get(result_id)
        if not result:
            raise ValueError("汇总对比结果已失效，请重新上传并生成")
        now = datetime.now()
        settings = self.config["settings"]
        filename = filename_from_pattern(
            settings["summary_compare_filename"],
            {"生成日期": now.strftime("%Y-%m-%d")},
        )
        target = unique_export_path(settings["export_path"], filename)
        export_summary_compare(result, target)
        log_event("exported summary compare file=%s", target)
        return self.add_export_record(target, "库存汇总对比", "上周与本周汇总表", result_id)

    def clear_summary_compare(self):
        self.summary_compare_results.clear()
        self.delete_uploads(item["path"] for item in self.summary_compare_uploads.values())
        self.summary_compare_uploads = {}

    def stage_special_brand_file(self, file_item):
        path, filename = self.save_upload(file_item)
        if self.special_brand_upload:
            self.delete_uploads([self.special_brand_upload["path"]])
        self.special_brand_upload = {
            "filename": filename,
            "path": str(path),
            "total_row": find_last_total_row(path),
            "total_row_removed": None,
        }
        self.special_brand_results = []
        return self.special_brand_upload_overview()

    def special_brand_upload_overview(self):
        item = self.special_brand_upload
        if not item:
            return None
        return {
            "filename": item["filename"],
            "total_row": item["total_row"],
            "total_row_removed": item["total_row_removed"],
        }

    def remove_special_brand_total_row(self):
        if not self.special_brand_upload:
            raise ValueError("请先上传需要拆分的特殊品牌库存表")
        removed = remove_last_total_row(self.special_brand_upload["path"])
        self.special_brand_upload["total_row"] = None
        self.special_brand_upload["total_row_removed"] = removed
        self.special_brand_results = []
        return self.special_brand_upload_overview()

    def split_special_brand(self):
        if not self.special_brand_upload:
            raise ValueError("请先上传需要拆分的特殊品牌库存表")
        targets = split_special_brand_workbook(
            self.special_brand_upload["path"],
            self.config["settings"]["export_path"],
            self.config["settings"]["special_brand_filename"],
        )
        records = [self.add_export_record(item["path"], "特殊品牌拆分", item["brand"], "") for item in targets]
        self.special_brand_results = [
            {"brand": item["brand"], "rows": item["rows"], "record": record}
            for item, record in zip(targets, records)
        ]
        log_event("split special brand source=%s files=%d", self.special_brand_upload["filename"], len(records))
        return self.special_brand_results

    def clear_special_brand(self):
        if self.special_brand_upload:
            self.delete_uploads([self.special_brand_upload["path"]])
        self.special_brand_upload = None
        self.special_brand_results = []

    def special_brand_export(self, record_id):
        record = next(
            (
                item for item in self.config["export_records"]
                if item["id"] == record_id and item["type"] == "特殊品牌拆分"
            ),
            None,
        )
        if not record:
            raise ValueError("未找到特殊品牌拆分导出文件")
        path = Path(record["path"])
        if not path.is_file():
            raise ValueError("导出文件已不存在")
        return path

    def compare_upload_overview(self, slot):
        item = self.compare_uploads.get(slot)
        if not item:
            return None
        return {"slot": slot, "filename": item["filename"], "total_row": item["total_row"], "total_row_removed": item["total_row_removed"]}

    def remove_compare_total_row(self, slot):
        item = self.compare_uploads.get(slot)
        if not item:
            raise ValueError("请先上传该对比文件")
        removed = remove_last_total_row(item["path"])
        item["total_row"] = None
        item["total_row_removed"] = removed
        self.compare_results.clear()
        return self.compare_upload_overview(slot)

    def compare_files(self, brand):
        previous_upload = self.compare_uploads.get("previous")
        current_upload = self.compare_uploads.get("current")
        if not previous_upload or not current_upload:
            raise ValueError("请先上传上周和本周文件")
        previous_path = Path(previous_upload["path"])
        current_path = Path(current_upload["path"])
        previous_filename = previous_upload["filename"]
        current_filename = current_upload["filename"]
        previous = parse_inventory_workbook(previous_path, brand)
        current = parse_inventory_workbook(current_path, brand)
        errors = (
            [format_issue(previous_filename, error) for error in previous["errors"]]
            + [format_issue(current_filename, error) for error in current["errors"]]
            + find_compare_spec_errors(previous["rows"], current["rows"], previous_filename, current_filename)
        )
        warnings = [format_issue(previous_filename, warning) for warning in previous["warnings"]] + [format_issue(current_filename, warning) for warning in current["warnings"]]
        if errors:
            return {"errors": errors, "warnings": warnings}
        result_id = uuid.uuid4().hex
        records = make_compare_result(brand, previous["rows"], current["rows"])
        needs_confirmation = bool(warnings)
        self.compare_results[result_id] = {"brand": brand, "records": records, "confirmed": not needs_confirmation}
        return {
            "result_id": result_id,
            "errors": [],
            "warnings": warnings,
            "needs_confirmation": needs_confirmation,
            "records": serialize_compare(records),
        }

    def clear_compare(self):
        self.compare_results.clear()
        self.delete_uploads(item["path"] for item in self.compare_uploads.values())
        self.compare_uploads = {}

    def clear_summary(self):
        self.delete_uploads(item["path"] for item in self.summary_imports)
        self.summary_imports = []
        self.summary_results.clear()

    def build_summary(self):
        if not self.summary_imports:
            return {"errors": [{"message": "请先导入至少一个品牌文件"}], "warnings": []}
        errors = [format_issue(item["filename"], error) for item in self.summary_imports for error in item["errors"]]
        warnings = [format_issue(item["filename"], warning) for item in self.summary_imports for warning in item["warnings"]]
        warnings.extend(find_cross_code_warnings(self.summary_imports))
        if errors:
            return {"errors": errors, "warnings": warnings}
        result = build_summary_result(self.summary_imports, self.config["mappings"])
        result_id = uuid.uuid4().hex
        result["confirmed"] = not bool(result["unconfigured"] or warnings)
        self.summary_results[result_id] = result
        return {"result_id": result_id, "errors": [], "warnings": warnings, "unconfigured": result["unconfigured"], "preview": serialize_summary(result)}

    def confirm_compare(self, result_id):
        result = self.compare_results.get(result_id)
        if not result:
            raise ValueError("对比结果已失效，请重新导入并生成")
        result["confirmed"] = True

    def confirm_summary(self, result_id):
        result = self.summary_results.get(result_id)
        if not result:
            raise ValueError("汇总结果已失效，请重新导入并生成")
        result["confirmed"] = True

    def export_compare(self, result_id):
        result = self.compare_results.get(result_id)
        if not result:
            raise ValueError("对比结果已失效，请重新导入并生成")
        if not result.get("confirmed"):
            raise ValueError("请先确认需要核对的数据，再导出对比表")
        now = datetime.now()
        settings = self.config["settings"]
        filename = filename_from_pattern(settings["compare_filename"], {"品牌": result["brand"], "本周日期": now.strftime("%Y-%m-%d"), "生成日期": now.strftime("%Y-%m-%d")})
        target = unique_export_path(settings["export_path"], filename)
        export_compare(result, target)
        return self.add_export_record(target, "单品牌库存对比", result["brand"], result_id)

    def export_summary(self, result_id):
        result = self.summary_results.get(result_id)
        if not result:
            raise ValueError("汇总结果已失效，请重新导入并生成")
        if not result.get("confirmed"):
            raise ValueError("请先确认需要核对的数据，再导出汇总表")
        now = datetime.now()
        settings = self.config["settings"]
        filename = filename_from_pattern(settings["summary_filename"], {"生成日期": now.strftime("%Y-%m-%d")})
        target = unique_export_path(settings["export_path"], filename)
        export_summary(result, target)
        brands = "、".join(dict.fromkeys(row["brand"] for row in result["rows"]))
        return self.add_export_record(target, "库存汇总", brands, result_id)

    def add_export_record(self, target, export_type, brands, result_id):
        record = {
            "id": uuid.uuid4().hex, "filename": target.name, "type": export_type, "brands": brands,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "path": str(target), "result_id": result_id,
        }
        self.config["export_records"].insert(0, record)
        save_config(self.config)
        return record

    def update_settings(self, settings):
        export_path = Path(str(settings.get("export_path", "")).strip())
        if not export_path:
            raise ValueError("默认导出路径不能为空")
        validate_export_directory(export_path)
        compare_filename = str(settings.get("compare_filename", "")).strip()
        summary_filename = str(settings.get("summary_filename", "")).strip()
        summary_compare_filename = str(settings.get("summary_compare_filename", "")).strip()
        special_brand_filename = str(settings.get("special_brand_filename", "")).strip()
        if not all((compare_filename, summary_filename, summary_compare_filename, special_brand_filename)):
            raise ValueError("所有导出文件命名规则都不能为空")
        log_enabled = bool(settings.get("log_enabled"))
        self.config["settings"] = {
            "export_path": str(export_path), "compare_filename": compare_filename,
            "summary_filename": summary_filename,
            "summary_compare_filename": summary_compare_filename,
            "special_brand_filename": special_brand_filename,
            "log_enabled": log_enabled,
        }
        save_config(self.config)
        configure_logging(log_enabled)
        log_event("settings updated export_path=%s", export_path)
        return self.config["settings"]

    def open_diagnostic_log(self):
        ensure_directories()
        LOG_PATH.touch(exist_ok=True)
        os.startfile(str(LOG_PATH))


def serialize_compare(records):
    return [{key: display_number(value) if isinstance(value, Decimal) else value for key, value in row.items()} for row in records]


def serialize_summary(result):
    return {
        "warehouses": result["warehouses"],
        "rows": [
            {
                "brand": row["brand"], "serial": row["serial"], "name": row["name"], "spec": row["spec"],
                "warehouses": {warehouse: {key: display_number(value) for key, value in values.items()} for warehouse, values in row["warehouses"].items()},
                "total_main": display_number(row["total_main"]), "total_box": display_number(row["total_box"]),
            }
            for row in result["rows"]
        ],
    }


def serialize_summary_compare(records):
    return [{key: display_number(value) if isinstance(value, Decimal) else value for key, value in row.items()} for row in records]


APP = InventoryApp()


class RequestHandler(BaseHTTPRequestHandler):
    server_version = "InventoryStudio/1.0"

    def log_message(self, format, *args):
        print(f"[{self.log_date_time_string()}] {format % args}")

    def send_json(self, payload, status=HTTPStatus.OK):
        data = json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_error_json(self, message, status=HTTPStatus.BAD_REQUEST):
        self.send_json({"ok": False, "error": message}, status)

    def read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        return json.loads(self.rfile.read(length).decode("utf-8")) if length else {}

    def read_form(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError as exc:
            raise ValueError("上传请求长度不正确") from exc
        return parse_multipart_form(self.rfile.read(length), self.headers.get("Content-Type", ""))

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/state":
            self.send_json({"ok": True, "data": APP.state()})
            return
        if parsed.path == "/api/special-brand/download":
            record_id = normalize_text(parse_qs(parsed.query).get("record_id", [""])[0])
            try:
                path = APP.special_brand_export(record_id)
            except ValueError as exc:
                self.send_error_json(str(exc), HTTPStatus.NOT_FOUND)
                return
            self.send_download(path)
            return
        if parsed.path == "/":
            self.serve_file(WEB_ROOT / "index.html")
            return
        if parsed.path == "/assets/lucide.js":
            self.serve_file(BUNDLE_ROOT / "lucide.js")
            return
        if parsed.path.startswith("/assets/"):
            relative = Path(parsed.path.removeprefix("/assets/"))
            if relative.is_absolute() or ".." in relative.parts:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            asset_path = (WEB_ROOT / relative).resolve()
            try:
                asset_path.relative_to(WEB_ROOT.resolve())
            except ValueError:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            self.serve_file(asset_path)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def serve_file(self, path):
        if not path.exists() or not path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", mimetypes.guess_type(path.name)[0] or "application/octet-stream")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_download(self, path):
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(path.name)}")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        try:
            # Import, generation, and clear requests share a local session.
            # Serializing mutations prevents a slow request from restoring data after clear.
            with APP.operation_lock:
                self.route_post(urlparse(self.path).path)
        except ValueError as exc:
            log_event("request validation error path=%s error=%s", self.path, exc)
            self.send_error_json(str(exc))
        except Exception as exc:
            log_event("request failure path=%s error=%r", self.path, exc)
            print("Unhandled request error:", repr(exc), file=sys.stderr)
            self.send_error_json(f"操作失败：{exc}", HTTPStatus.INTERNAL_SERVER_ERROR)

    def route_post(self, path):
        if path == "/api/compare/upload":
            form = self.read_form()
            slot = normalize_text(form.getfirst("slot"))
            file_item = form["file"] if "file" in form else None
            if file_item is None or not file_item.filename:
                raise ValueError("请选择一个 Excel 文件")
            self.send_json({"ok": True, "data": APP.stage_compare_file(slot, file_item)})
            return
        if path == "/api/compare":
            form = self.read_form()
            brand = normalize_text(form.getfirst("brand"))
            if not brand:
                raise ValueError("请填写品牌名称")
            self.send_json({"ok": True, "data": APP.compare_files(brand)})
            return
        if path == "/api/summary/import":
            form = self.read_form()
            files = form["files"] if "files" in form else []
            if not isinstance(files, list):
                files = [files]
            if not files or not files[0].filename:
                raise ValueError("请选择至少一个品牌 Excel 文件")
            self.send_json({"ok": True, "data": APP.import_summary_files(files)})
            return
        if path == "/api/summary-compare/upload":
            form = self.read_form()
            slot = normalize_text(form.getfirst("slot"))
            file_item = form["file"] if "file" in form else None
            if file_item is None or not file_item.filename:
                raise ValueError("请选择一张库存汇总表")
            self.send_json({"ok": True, "data": APP.stage_summary_compare_file(slot, file_item)})
            return
        if path == "/api/summary-compare/build":
            self.send_json({"ok": True, "data": APP.compare_summaries()})
            return
        if path == "/api/summary-compare/export":
            self.send_json({"ok": True, "data": APP.export_summary_compare(self.read_json().get("result_id", ""))})
            return
        if path == "/api/summary-compare/clear":
            APP.clear_summary_compare()
            self.send_json({"ok": True})
            return
        if path == "/api/special-brand/upload":
            form = self.read_form()
            file_item = form["file"] if "file" in form else None
            if file_item is None or not file_item.filename:
                raise ValueError("请选择需要拆分的特殊品牌库存表")
            self.send_json({"ok": True, "data": APP.stage_special_brand_file(file_item)})
            return
        if path == "/api/special-brand/remove-total-row":
            self.send_json({"ok": True, "data": APP.remove_special_brand_total_row()})
            return
        if path == "/api/special-brand/split":
            self.send_json({"ok": True, "data": APP.split_special_brand()})
            return
        if path == "/api/special-brand/clear":
            APP.clear_special_brand()
            self.send_json({"ok": True})
            return
        if path == "/api/summary/delete-import":
            payload = self.read_json()
            self.send_json({"ok": True, "data": APP.delete_summary_import(normalize_text(payload.get("brand")))})
            return
        if path == "/api/summary/remove-total-row":
            payload = self.read_json()
            self.send_json({"ok": True, "data": APP.remove_summary_total_row(normalize_text(payload.get("brand")))})
            return
        if path == "/api/compare/remove-total-row":
            payload = self.read_json()
            self.send_json({"ok": True, "data": APP.remove_compare_total_row(normalize_text(payload.get("slot")))})
            return
        if path == "/api/summary/rename-import":
            payload = self.read_json()
            self.send_json({
                "ok": True,
                "data": APP.rename_summary_import(
                    normalize_text(payload.get("brand")), normalize_text(payload.get("new_brand"))
                ),
            })
            return
        if path == "/api/summary/reorder-imports":
            payload = self.read_json()
            brands = payload.get("brands")
            if not isinstance(brands, list) or not all(isinstance(brand, str) for brand in brands):
                raise ValueError("品牌排序格式不正确")
            self.send_json({"ok": True, "data": APP.reorder_summary_imports([normalize_text(brand) for brand in brands])})
            return
        if path == "/api/summary/build":
            self.send_json({"ok": True, "data": APP.build_summary()})
            return
        if path == "/api/compare/clear":
            APP.clear_compare()
            self.send_json({"ok": True})
            return
        if path == "/api/summary/clear":
            APP.clear_summary()
            self.send_json({"ok": True})
            return
        if path == "/api/compare/export":
            record = APP.export_compare(self.read_json().get("result_id", ""))
            self.send_json({"ok": True, "data": record})
            return
        if path == "/api/compare/confirm":
            APP.confirm_compare(self.read_json().get("result_id", ""))
            self.send_json({"ok": True})
            return
        if path == "/api/summary/export":
            record = APP.export_summary(self.read_json().get("result_id", ""))
            self.send_json({"ok": True, "data": record})
            return
        if path == "/api/summary/confirm":
            APP.confirm_summary(self.read_json().get("result_id", ""))
            self.send_json({"ok": True})
            return
        if path == "/api/settings":
            self.send_json({"ok": True, "data": APP.update_settings(self.read_json())})
            return
        if path == "/api/mapping":
            payload = self.read_json()
            brand = normalize_text(payload.get("brand"))
            source_warehouse = normalize_text(payload.get("source_warehouse"))
            target_warehouse = normalize_text(payload.get("target_warehouse")) or "德清仓"
            if not brand or not source_warehouse:
                raise ValueError("品牌和源仓库名称不能为空")
            imported = next((item for item in APP.summary_imports if item["brand"] == brand), None)
            if imported:
                warehouses = {row["warehouse"] for row in imported["rows"]}
                if source_warehouse not in warehouses:
                    raise ValueError(f"品牌“{brand}”本次导入的文件中不存在源仓库“{source_warehouse}”")
            if any(
                item["brand"] == brand and item["source_warehouse"] == source_warehouse
                for item in APP.config["mappings"]
            ):
                raise ValueError("该品牌和源仓库名称的映射已存在，请使用编辑功能修改")
            APP.config["mappings"].append({
                "id": uuid.uuid4().hex,
                "brand": brand,
                "source_warehouse": source_warehouse,
                "target_warehouse": target_warehouse,
            })
            APP.summary_results.clear()
            save_config(APP.config)
            self.send_json({"ok": True, "data": APP.state()["mappings"]})
            return
        if path == "/api/mapping/update":
            payload = self.read_json()
            mapping_id = normalize_text(payload.get("id"))
            brand = normalize_text(payload.get("brand"))
            source_warehouse = normalize_text(payload.get("source_warehouse"))
            target_warehouse = normalize_text(payload.get("target_warehouse")) or "德清仓"
            mapping = next((item for item in APP.config["mappings"] if item["id"] == mapping_id), None)
            if not mapping:
                raise ValueError("未找到需要编辑的仓库映射")
            if not brand or not source_warehouse:
                raise ValueError("品牌和源仓库名称不能为空")
            imported = next((item for item in APP.summary_imports if item["brand"] == brand), None)
            if imported and source_warehouse not in {row["warehouse"] for row in imported["rows"]}:
                raise ValueError(f"品牌“{brand}”本次导入的文件中不存在源仓库“{source_warehouse}”")
            if any(
                item["id"] != mapping_id and item["brand"] == brand and item["source_warehouse"] == source_warehouse
                for item in APP.config["mappings"]
            ):
                raise ValueError("该品牌和源仓库名称的映射已存在")
            mapping.update({"brand": brand, "source_warehouse": source_warehouse, "target_warehouse": target_warehouse})
            APP.summary_results.clear()
            save_config(APP.config)
            self.send_json({"ok": True, "data": APP.state()["mappings"]})
            return
        if path == "/api/mapping/delete":
            mapping_id = normalize_text(self.read_json().get("id"))
            original_count = len(APP.config["mappings"])
            APP.config["mappings"] = [item for item in APP.config["mappings"] if item["id"] != mapping_id]
            if len(APP.config["mappings"]) == original_count:
                raise ValueError("未找到需要删除的仓库映射")
            APP.summary_results.clear()
            save_config(APP.config)
            self.send_json({"ok": True, "data": APP.state()["mappings"]})
            return
        if path == "/api/open-folder":
            path_value = Path(str(self.read_json().get("path", "")))
            if not path_value.exists():
                raise ValueError("文件或文件夹已不存在")
            target = path_value if path_value.is_dir() else path_value.parent
            os.startfile(str(target))
            self.send_json({"ok": True})
            return
        if path == "/api/log/open":
            APP.open_diagnostic_log()
            self.send_json({"ok": True})
            return
        if path == "/api/choose-directory":
            self.send_json({"ok": True, "data": {"path": choose_directory(APP.config["settings"]["export_path"])}})
            return
        self.send_error_json("未找到接口", HTTPStatus.NOT_FOUND)


def choose_directory(initial_directory):
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askdirectory(initialdir=initial_directory, title="选择默认导出文件夹")
        root.destroy()
        return selected
    except Exception as exc:
        raise ValueError(f"无法打开文件夹选择窗口：{exc}") from exc


def main():
    ensure_directories()
    port = int(os.environ.get("INVENTORY_PORT", "4174"))
    server = ThreadingHTTPServer(("127.0.0.1", port), RequestHandler)
    address = f"http://127.0.0.1:{port}"
    print(f"库存工作台已启动：{address}")
    if os.environ.get("INVENTORY_NO_BROWSER") != "1":
        threading.Timer(0.5, lambda: webbrowser.open(address)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
