from decimal import Decimal
from pathlib import Path
import sys

import openpyxl

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import app


OUT = PROJECT_ROOT / "output" / "verification"
FIXTURE_A = OUT / "fixture-a.xlsx"
FIXTURE_B = OUT / "fixture-b.xlsx"
TOTAL_FIXTURE = OUT / "fixture-with-total.xlsx"
TOTAL_MARKER_FIXTURE = OUT / "fixture-with-total-marker.xlsx"
TOTAL_LIKE_PRODUCT_FIXTURE = OUT / "fixture-with-total-like-product.xlsx"
SPECIAL_BRAND_FIXTURE = OUT / "special-brand.xlsx"


def write_fixture(path, rows):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Inventory"
    sheet.append(app.REQUIRED_COLUMNS)
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    workbook.close()


def main():
    app.ensure_directories()
    OUT.mkdir(parents=True, exist_ok=True)

    multipart_body = (
        b"--boundary\r\n"
        b"Content-Disposition: form-data; name=\"slot\"\r\n\r\n"
        b"previous\r\n"
        b"--boundary\r\n"
        b"Content-Disposition: form-data; name=\"files\"; filename=\"one.xlsx\"\r\n"
        b"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
        b"first\r\n"
        b"--boundary\r\n"
        b"Content-Disposition: form-data; name=\"files\"; filename=\"two.xlsx\"\r\n"
        b"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
        b"second\r\n"
        b"--boundary--\r\n"
    )
    multipart_form = app.parse_multipart_form(multipart_body, "multipart/form-data; boundary=boundary")
    assert multipart_form.getfirst("slot") == "previous"
    assert [item.filename for item in multipart_form["files"]] == ["one.xlsx", "two.xlsx"]
    assert multipart_form["files"][0].file.read() == b"first"

    original_config_path = app.CONFIG_PATH
    legacy_config_path = OUT / "legacy-config.json"
    legacy_config = app.default_config()
    legacy_config["settings"]["compare_filename"] = "{品牌}_库存对比_{本周日期}.xlsx"
    legacy_config["export_records"] = [{"type": "库存对比"}]
    app.CONFIG_PATH = legacy_config_path
    try:
        app.save_config(legacy_config)
        migrated_config = app.load_config()
    finally:
        app.CONFIG_PATH = original_config_path
    assert migrated_config["settings"]["compare_filename"] == "{品牌}_单品牌库存对比_{本周日期}.xlsx"
    assert migrated_config["export_records"][0]["type"] == "单品牌库存对比"

    write_fixture(
        FIXTURE_A,
        [
            ["A-001", "Product A", "Spec A", "Warehouse A", "Lot-1", "Piece", 18, "Available", "Org A", 3],
            ["A-002", "Product Zero", "Spec Z", "Warehouse A", "Lot-2", "Piece", 0, "Available", "Org A", 0],
            ["A-003", "Removed Product", "Spec R", "Warehouse B", "Lot-3", "Piece", 6, "Available", "Org A", 1],
        ],
    )
    write_fixture(
        TOTAL_MARKER_FIXTURE,
        [
            ["T-001", "Total Test Product", "Spec T", "Warehouse T", "Lot-1", "Piece", 12, "Available", "Org T", 2],
            ["", "", "", "", "", "", 12, "合计", "", 2],
            ["", "", "", "", "", "", "", "", "", ""],
        ],
    )
    write_fixture(
        FIXTURE_B,
        [
            ["A-001", "Product A", "Spec A", "Warehouse A", "Lot-9", "Piece", 15, "Locked", "Org A", 2.5],
            ["A-002", "Product Zero", "Spec Z", "Warehouse A", "Lot-8", "Piece", 5, "Available", "Org A", 1],
            ["B-001", "Added Product", "Spec B", "Warehouse C", "Lot-4", "Piece", 10, "Available", "Org B", 2],
        ],
    )
    write_fixture(
        TOTAL_FIXTURE,
        [
            ["T-001", "Total Test Product", "Spec T", "Warehouse T", "Lot-1", "Piece", 12, "Available", "Org T", 2],
            ["", "总计", "", "", "", "", 12, "", "", 2],
        ],
    )
    write_fixture(
        TOTAL_LIKE_PRODUCT_FIXTURE,
        [["T-002", "合计装商品", "Spec T", "Warehouse T", "Lot-2", "Piece", 9, "Available", "Org T", 1]],
    )

    total_row = app.find_last_total_row(TOTAL_FIXTURE)
    assert total_row == {"row": 3, "label": "总计"}
    assert app.find_last_total_row(TOTAL_MARKER_FIXTURE) == {"row": 3, "label": "合计"}
    assert app.find_last_total_row(TOTAL_LIKE_PRODUCT_FIXTURE) is None
    original_total_book = openpyxl.load_workbook(TOTAL_FIXTURE, data_only=True)
    assert original_total_book.active.max_row == 3
    original_total_book.close()

    previous = app.parse_inventory_workbook(FIXTURE_A, "fixture-a")
    current = app.parse_inventory_workbook(FIXTURE_B, "fixture-a")
    assert not previous["errors"], previous["errors"]
    assert not current["errors"], current["errors"]
    assert len(previous["rows"]) == 3
    assert len(current["rows"]) == 3

    compare = app.make_compare_result("fixture-a", previous["rows"], current["rows"])
    by_code = {row["code"]: row for row in compare}
    assert len(compare) == 4
    assert by_code["A-001"]["main_mom_change"] == Decimal("-0.1667")
    assert by_code["A-002"]["main_mom_change"] is None
    assert by_code["B-001"]["change_type"] == "新增商品"
    assert by_code["A-003"]["change_type"] == "删除商品"

    old_format = OUT / "old-format.xlsx"
    write_fixture(old_format, [])
    book = openpyxl.load_workbook(old_format)
    book.active.delete_cols(8)
    book.save(old_format)
    book.close()
    rejected = app.parse_inventory_workbook(old_format, "old-format")
    assert rejected["errors"]
    assert "库存状态" in rejected["errors"][0]["message"]

    summary = app.build_summary_result(
        [
            {"brand": "fixture-a", "rows": previous["rows"]},
            {"brand": "fixture-b", "rows": current["rows"]},
        ],
        {"fixture-a": "Warehouse A", "fixture-b": "Warehouse C"},
    )
    assert summary["warehouses"][0] == "德清仓"
    assert not summary["unconfigured"]
    assert summary["rows"]

    compare_path = OUT / "compare.xlsx"
    summary_path = OUT / "summary.xlsx"
    export_records = [dict(compare[0], previous_main=Decimal("1.234"), current_main=Decimal("5.678"))]
    app.export_compare({"brand": "fixture-test", "records": export_records}, compare_path)
    app.export_summary(summary, summary_path)
    assert compare_path.exists()
    assert summary_path.exists()

    compare_book = openpyxl.load_workbook(compare_path, data_only=True)
    compare_sheet = compare_book.active
    assert compare_sheet.title == "单品牌库存对比"
    assert compare_sheet["A1"].value == "fixture-test 单品牌库存对比表"
    assert compare_sheet["L2"].value == "环比变化"
    assert compare_sheet["M2"].value == "变动类型"
    assert compare_sheet["F3"].number_format == "0.00"
    assert compare_sheet["G3"].number_format == "0.00"
    assert compare_sheet["F3"].value == 1.23
    assert compare_sheet["G3"].value == 5.68
    assert compare_sheet["L3"].number_format == "0.00%"

    summary_book = openpyxl.load_workbook(summary_path, data_only=True)
    summary_sheet = summary_book.active
    template_book = openpyxl.load_workbook(PROJECT_ROOT / "模板总模板.xlsx", data_only=True)
    assert summary_sheet["A1"].value == template_book.active["A1"].value
    assert summary_sheet["E2"].value == "德清仓"
    assert summary_sheet.max_row >= 4

    previous_summary_path = OUT / "previous-summary.xlsx"
    current_summary_path = OUT / "current-summary.xlsx"
    previous_summary = {
        "warehouses": ["德清仓"],
        "rows": [{
            "brand": "Brand A", "serial": 1, "name": "Product A", "spec": "Spec A",
            "warehouses": {"德清仓": {"main_qty": Decimal("10"), "box_qty": Decimal("2")}},
            "total_main": Decimal("10"), "total_box": Decimal("2"),
        }],
    }
    current_summary = {
        "warehouses": ["德清仓"],
        "rows": [
            {
                "brand": "Brand A", "serial": 1, "name": "Product A", "spec": "Spec A",
                "warehouses": {"德清仓": {"main_qty": Decimal("15"), "box_qty": Decimal("3")}},
                "total_main": Decimal("15"), "total_box": Decimal("3"),
            },
            {
                "brand": "Brand A", "serial": 2, "name": "Product New", "spec": "Spec N",
                "warehouses": {"德清仓": {"main_qty": Decimal("4"), "box_qty": Decimal("1")}},
                "total_main": Decimal("4"), "total_box": Decimal("1"),
            },
        ],
    }
    app.export_summary(previous_summary, previous_summary_path)
    app.export_summary(current_summary, current_summary_path)
    summary_compare = app.make_summary_compare_result(
        app.parse_summary_compare_workbook(previous_summary_path), app.parse_summary_compare_workbook(current_summary_path)
    )
    compare_by_name = {row["name"]: row for row in summary_compare}
    assert compare_by_name["Product A"]["main_mom_change"] == Decimal("0.5000")
    assert compare_by_name["Product New"]["remark"] == "新增物料"
    summary_compare_path = OUT / "summary-compare.xlsx"
    app.export_summary_compare({"records": summary_compare}, summary_compare_path)
    summary_compare_book = openpyxl.load_workbook(summary_compare_path, data_only=True)
    assert summary_compare_book.active.title == "汇总对比"
    assert summary_compare_book.active["G2"].value == "环比变化"
    assert summary_compare_book.active["H4"].value == "新增物料"
    summary_compare_book.close()

    special_book = openpyxl.Workbook()
    special_sheet = special_book.active
    special_sheet.append(app.REQUIRED_COLUMNS)
    special_sheet.append(["SP-001", "Special A", "Spec A", "Brand One-East Warehouse", "Lot", "Piece", 2, "Available", "Org", 1])
    special_sheet.append(["SP-002", "Special B", "Spec B", "Brand One-West Warehouse", "Lot", "Piece", 3, "Available", "Org", 1])
    special_sheet.append(["SP-003", "Special C", "Spec C", "Brand Two-South Warehouse", "Lot", "Piece", 4, "Available", "Org", 2])
    special_sheet["A1"].fill = openpyxl.styles.PatternFill("solid", fgColor="00FF00")
    special_book.save(SPECIAL_BRAND_FIXTURE)
    special_book.close()
    split_results = app.split_special_brand_workbook(SPECIAL_BRAND_FIXTURE, OUT)
    assert {(item["brand"], item["rows"]) for item in split_results} == {("Brand One", 2), ("Brand Two", 1)}
    split_brand_one = next(item["path"] for item in split_results if item["brand"] == "Brand One")
    split_book = openpyxl.load_workbook(split_brand_one, data_only=True)
    assert split_book.active["D2"].value == "East Warehouse"
    assert split_book.active["D3"].value == "West Warehouse"
    assert split_book.active["A1"].fill.fgColor.rgb == "0000FF00"
    split_book.close()

    multi_sheet_book = openpyxl.load_workbook(SPECIAL_BRAND_FIXTURE)
    multi_sheet_book.create_sheet("Other")
    multi_sheet_path = OUT / "special-brand-multi-sheet.xlsx"
    multi_sheet_book.save(multi_sheet_path)
    multi_sheet_book.close()
    try:
        app.split_special_brand_workbook(multi_sheet_path, OUT)
        raise AssertionError("multi-sheet special brand workbook should be rejected")
    except ValueError as exc:
        assert "单工作表" in str(exc)

    reexport_app = app.InventoryApp()
    reexport_app.config = app.default_config()
    reexport_app.config["settings"]["export_path"] = str(OUT)
    reexport_app.config["export_records"] = [{
        "id": "fixture-record", "filename": compare_path.name, "type": "单品牌库存对比",
        "brands": "fixture-test", "created_at": "2026-08-21 00:00:00",
        "path": str(compare_path), "result_id": "expired-result",
    }]
    original_save_config = app.save_config
    app.save_config = lambda config: None
    try:
        reexported = reexport_app.reexport("fixture-record")
    finally:
        app.save_config = original_save_config
    assert Path(reexported["path"]).is_file()
    assert reexported["path"] != str(compare_path)

    try:
        app.decimal_number("NaN")
        raise AssertionError("NaN should be rejected")
    except ValueError:
        pass
    changed_spec = [dict(previous["rows"][0], spec="Changed")]
    assert app.find_compare_spec_errors(previous["rows"], changed_spec, "fixture-a.xlsx", "fixture-b.xlsx")
    invalid_mapping = app.build_summary_result(
        [{"brand": "fixture-a", "rows": previous["rows"]}], {"fixture-a": "Missing Warehouse"}
    )
    assert invalid_mapping["unconfigured"][0]["reason"].startswith("已配置的源仓库")

    invalid_rows = [
        ["BAD-001", "Bad Product", "Spec", "Warehouse A", "Lot", "Piece", "", "Available", "Org", 1],
    ]
    invalid_path = OUT / "invalid-row.xlsx"
    write_fixture(invalid_path, invalid_rows)
    invalid = app.parse_inventory_workbook(invalid_path, "invalid")
    assert invalid["errors"][0]["row"] == 2
    detailed = app.format_issue(invalid_path.name, invalid["errors"][0])
    assert invalid_path.name in detailed["message"] and "第 2 行" in detailed["message"]

    session = app.InventoryApp()
    session.summary_imports = [
        {"brand": "fixture-a", "filename": FIXTURE_A.name, "path": str(FIXTURE_A), **previous},
    ]
    class Upload:
        def __init__(self, path):
            self.filename = path.name
            self.file = path.open("rb")

    second_upload = Upload(FIXTURE_B)
    try:
        imported = session.import_summary_files([second_upload])
    finally:
        second_upload.file.close()
    assert [item["brand"] for item in imported] == ["fixture-a", "fixture-b"]
    reordered = session.reorder_summary_imports(["fixture-b", "fixture-a"])
    assert [item["brand"] for item in reordered] == ["fixture-b", "fixture-a"]
    ordered_summary = app.build_summary_result(session.summary_imports, {})
    assert ordered_summary["rows"][0]["brand"] == "fixture-b"
    session.config["mappings"] = {"fixture-b": "Warehouse C"}
    renamed = session.rename_summary_import("fixture-b", "Renamed Brand")
    assert [item["brand"] for item in renamed] == ["Renamed Brand", "fixture-a"]
    assert session.config["mappings"] == {"fixture-b": "Warehouse C"}
    remaining = session.delete_summary_import("Renamed Brand")
    assert [item["brand"] for item in remaining] == ["fixture-a"]

    total_upload = Upload(TOTAL_FIXTURE)
    try:
        total_session = app.InventoryApp()
        imported_total = total_session.import_summary_files([total_upload])
    finally:
        total_upload.file.close()
    assert imported_total[0]["total_row"] == {"row": 3, "label": "总计"}
    cleared_total = total_session.remove_summary_total_row("fixture-with-total")
    assert cleared_total[0]["total_row"] is None
    assert cleared_total[0]["total_row_removed"] == {"row": 3, "label": "总计"}
    assert len(total_session.summary_imports[0]["rows"]) == 1
    original_total_book = openpyxl.load_workbook(TOTAL_FIXTURE, data_only=True)
    assert original_total_book.active.max_row == 3
    assert original_total_book.active.cell(3, 2).value == "总计"
    original_total_book.close()
    total_session.clear_summary()
    print("engine verification passed")


if __name__ == "__main__":
    main()
