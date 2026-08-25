# Local HTTP API and static-file server.

from __future__ import annotations

import json
import mimetypes
import os
import sys
import threading
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

from inventory_core import *
from inventory_logging import install_exception_hooks, log_exception, log_operation
from inventory_service import InventoryApp

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
        request_id = os.urandom(4).hex()
        parsed = urlparse(self.path)
        log_operation("http.request", "started", request_id=request_id, method="GET", path=parsed.path)
        try:
            self.route_get(parsed)
            log_operation("http.request", "completed", request_id=request_id, method="GET", path=parsed.path)
        except ValueError as exc:
            log_exception("http.request", exc, request_id=request_id, method="GET", path=parsed.path, status=400)
            self.send_error_json(str(exc))
        except Exception as exc:
            log_exception("http.request", exc, request_id=request_id, method="GET", path=parsed.path, status=500)
            print("Unhandled request error:", repr(exc), file=sys.stderr)
            self.send_error_json(f"操作失败：{exc}", HTTPStatus.INTERNAL_SERVER_ERROR)

    def route_get(self, parsed):
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
        request_id = os.urandom(4).hex()
        path = urlparse(self.path).path
        log_operation("http.request", "started", request_id=request_id, method="POST", path=path)
        try:
            # Import, generation, and clear requests share a local session.
            # Serializing mutations prevents a slow request from restoring data after clear.
            with APP.operation_lock:
                self.route_post(path)
            log_operation("http.request", "completed", request_id=request_id, method="POST", path=path)
        except ValueError as exc:
            log_exception("http.request", exc, request_id=request_id, method="POST", path=path, status=400)
            self.send_error_json(str(exc))
        except Exception as exc:
            log_exception("http.request", exc, request_id=request_id, method="POST", path=path, status=500)
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
    install_exception_hooks()
    port = int(os.environ.get("INVENTORY_PORT", "4174"))
    server = ThreadingHTTPServer(("127.0.0.1", port), RequestHandler)
    address = f"http://127.0.0.1:{port}"
    log_operation("application", "started", address=address)
    print(f"库存工作台已启动：{address}")
    if os.environ.get("INVENTORY_NO_BROWSER") != "1":
        threading.Timer(0.5, lambda: webbrowser.open(address)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        log_operation("application", "stopped", address=address)


if __name__ == "__main__":
    main()
