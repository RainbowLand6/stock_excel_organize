# Session-level inventory workflows.

from __future__ import annotations

import inventory_core as core
from inventory_core import *
from inventory_logging import audited, log_operation

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
        if not log_enabled:
            log_operation("settings.update", "disabling_logging")
        save_config(self.config)
        configure_logging(log_enabled)
        if log_enabled:
            log_operation("diagnostic_log", "enabled", path=str(core.LOG_PATH))
        log_operation("settings.update", "completed", export_path=str(export_path), log_enabled=log_enabled)
        return self.config["settings"]

    def open_diagnostic_log(self):
        if not self.config["settings"].get("log_enabled"):
            raise ValueError("诊断日志尚未开启，请先勾选“开启日志记录”并保存设置")
        ensure_directories()
        core.LOG_PATH.touch(exist_ok=True)
        log_operation("diagnostic_log.open", "completed", path=str(core.LOG_PATH))
        os.startfile(str(core.LOG_PATH))


def _audit_inventory_operations():
    """Attach audit logging to the user-visible workflow methods."""
    names = (
        "import_summary_files", "delete_summary_import", "rename_summary_import", "reorder_summary_imports",
        "remove_summary_total_row", "stage_compare_file", "stage_summary_compare_file", "compare_summaries",
        "export_summary_compare", "clear_summary_compare", "stage_special_brand_file",
        "remove_special_brand_total_row", "split_special_brand", "clear_special_brand", "special_brand_export",
        "remove_compare_total_row", "compare_files", "clear_compare", "clear_summary", "build_summary",
        "confirm_compare", "confirm_summary", "export_compare", "export_summary", "add_export_record",
        "update_settings", "open_diagnostic_log",
    )
    for name in names:
        setattr(InventoryApp, name, audited(f"inventory.{name}")(getattr(InventoryApp, name)))


_audit_inventory_operations()


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
