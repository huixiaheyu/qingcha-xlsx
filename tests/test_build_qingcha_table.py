from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from scripts.build_qingcha_table import (
    BuildDiagnostics,
    MissingSourceFilesError,
    annual_rent,
    build_workbook,
    contract_years,
    normalize_resource_period,
    parse_context,
    required_source_filenames,
    resolve_balance_sheet_name,
)
from scripts.xlsx_xml import NS, read_sheet_rows, write_table_rows

BASE_DIR = Path(__file__).resolve().parents[1]
COMPLETE_DIR = BASE_DIR / "完整"


class BuildQingchaTableTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = Path(tempfile.mkdtemp(prefix="qingcha-test-"))
        self.complete_dir = self.temp_dir / "完整"
        shutil.copytree(COMPLETE_DIR, self.complete_dir)
        self.output_path = self.complete_dir / "清查表（已填充）.xlsx"

    def tearDown(self) -> None:
        shutil.rmtree(self.temp_dir)

    def test_builds_output_workbook(self) -> None:
        build_workbook(self.complete_dir, self.output_path)
        self.assertTrue(self.output_path.exists())

    def test_balance_sheet_name_is_resolved_from_uploaded_workbook(self) -> None:
        diagnostics = BuildDiagnostics.create()
        sheet_name = resolve_balance_sheet_name(self.complete_dir, diagnostics)

        self.assertTrue(sheet_name.endswith("-科目余额"))
        self.assertIn(f"INFO: 已识别科目余额源工作表：{sheet_name}", diagnostics.logs)

    def test_builds_with_server_template_and_upload_source_dir(self) -> None:
        source_dir = self.temp_dir / "uploads"
        source_dir.mkdir()
        for filename in required_source_filenames():
            shutil.copy2(self.complete_dir / filename, source_dir / filename)

        output_path = self.temp_dir / "outputs" / "清查表（已填充）.xlsx"
        output_path.parent.mkdir()
        diagnostics = BuildDiagnostics.create()
        build_workbook(
            source_dir,
            output_path,
            template_path=self.complete_dir / "清查表.xlsx",
            diagnostics=diagnostics,
        )

        self.assertTrue(output_path.exists())
        self.assertIn("INFO: 已找到全部必需源工作簿。", diagnostics.logs)
        self.assertTrue(any(log.startswith("INFO: 已生成工作簿：") for log in diagnostics.logs))
        rows = read_sheet_rows(output_path, "村集体经济组织资源信息表")
        data_rows = [row for row in rows if row["row_number"] >= 5 and row.get("E")]
        self.assertEqual(len(data_rows), 36)

    def test_missing_source_files_are_reported_before_building(self) -> None:
        source_dir = self.temp_dir / "uploads-missing"
        source_dir.mkdir()
        missing_filename = "合同明细数据.xlsx"
        for filename in required_source_filenames():
            if filename != missing_filename:
                shutil.copy2(self.complete_dir / filename, source_dir / filename)

        diagnostics = BuildDiagnostics.create()
        with self.assertRaises(MissingSourceFilesError) as raised:
            build_workbook(
                source_dir,
                self.temp_dir / "missing-output.xlsx",
                template_path=self.complete_dir / "清查表.xlsx",
                diagnostics=diagnostics,
            )

        self.assertEqual(raised.exception.missing_files, [missing_filename])
        self.assertEqual(diagnostics.missing_files, [missing_filename])
        self.assertIn(f"WARN: 缺少上传文件：{missing_filename}", diagnostics.logs)

    def test_diagnostics_do_not_warn_for_unmatched_leased_resources(self) -> None:
        diagnostics = BuildDiagnostics.create()
        build_workbook(self.complete_dir, self.output_path, diagnostics=diagnostics)

        self.assertEqual(diagnostics.unresolved_resources, [])
        self.assertFalse(any("出租经营资源未能确定性匹配合同" in log for log in diagnostics.logs))

    def test_missing_template_cells_are_reported_by_write_helper(self) -> None:
        copied_template = self.temp_dir / "missing-cell-template.xlsx"
        shutil.copy2(self.complete_dir / "清查表.xlsx", copied_template)
        missing_cells: list[str] = []

        write_table_rows(
            copied_template,
            "用户信息表",
            4,
            [{"A": "襄城县", "ZZ": "无法写入"}],
            template_row_number=4,
            clear_existing_from_row=4,
            missing_cell_callback=lambda sheet, row, col, value: missing_cells.append(f"{sheet}:{row}:{col}:{value}"),
        )

        self.assertEqual(missing_cells, ["用户信息表:4:ZZ:无法写入"])

    def test_resource_sheet_has_36_rows_from_complete_dir(self) -> None:
        build_workbook(self.complete_dir, self.output_path)
        rows = read_sheet_rows(self.output_path, "村集体经济组织资源信息表")
        data_rows = [row for row in rows if row["row_number"] >= 5 and row.get("E")]
        self.assertEqual(len(data_rows), 36)

    def test_resource_sheet_uses_source_or_deterministic_contract_lease_fields(self) -> None:
        build_workbook(self.complete_dir, self.output_path)
        rows = read_sheet_rows(self.output_path, "村集体经济组织资源信息表")
        by_code = {row["E"]: row for row in rows if row["row_number"] >= 5 and row.get("E")}

        self.assertEqual(by_code["Y4110251012010012"]["J"], "襄城县居美达机械厂")
        self.assertEqual(by_code["Y4110251012010012"]["K"], "2010-06-30~2030-06-30")
        self.assertEqual(by_code["Y4110251012010012"]["L"], "3400.00")

        self.assertEqual(by_code["Y4110251012010013"]["J"], "赵建垒")
        self.assertEqual(by_code["Y4110251012010013"]["K"], "2012-08-31~2032-08-31")
        self.assertEqual(by_code["Y4110251012010013"]["L"], "1000.00")

        self.assertEqual(by_code["Y4110251012010006"]["J"], "巴海棠")
        self.assertEqual(by_code["Y4110251012010006"]["K"], "2025-07-18~2041-10-31")
        self.assertEqual(by_code["Y4110251012010006"]["L"], "0.00")

        self.assertEqual(by_code["Y4110251012010008"]["J"], "")
        self.assertEqual(by_code["Y4110251012010008"]["K"], "")
        self.assertEqual(by_code["Y4110251012010008"]["L"], "")
        self.assertTrue(all(row.get("O", "") == "" for row in by_code.values()))

    def test_contract_sheet_has_all_contract_rows(self) -> None:
        build_workbook(self.complete_dir, self.output_path)
        rows = read_sheet_rows(self.output_path, "村集体经济组织资产合同表")
        data_rows = [row for row in rows if row["row_number"] >= 4 and row.get("E")]
        self.assertEqual(len(data_rows), 5)
        self.assertEqual(data_rows[0]["I"], "襄城县居美达机械厂")
        self.assertEqual(data_rows[2]["I"], "巴海棠")

    def test_contract_sheet_maps_transaction_payment_and_subject_fields(self) -> None:
        build_workbook(self.complete_dir, self.output_path)
        rows = read_sheet_rows(self.output_path, "村集体经济组织资产合同表")
        by_code = {row["E"]: row for row in rows if row["row_number"] >= 4 and row.get("E")}

        first = by_code["2006411025101201C001"]
        self.assertEqual(first["Q"], "银行转账")
        self.assertEqual(first["R"], "否")
        self.assertEqual(first["S"], "年付")
        self.assertEqual(first["T"], "银行存款")
        self.assertEqual(first["U"], "集体建设用地承包收入")
        self.assertEqual(first["N"], "亩")
        self.assertEqual(first["V"], "工业用地12")

        third = by_code["2025411025101201C001"]
        self.assertEqual(third["Q"], "银行转账")
        self.assertEqual(third["R"], "是")
        self.assertEqual(third["S"], "一次付清")
        self.assertEqual(third["T"], "银行存款")
        self.assertEqual(third["U"], "集体建设用地承包收入")
        self.assertEqual(third["V"], "科教文卫用地6")

        asset = by_code["HT2025411025101201C001"]
        self.assertEqual(asset["V"], "2024年东街村厂房建设项目")

        grouped = by_code["200641102510120104C001"]
        self.assertEqual(grouped["V"], "")

    def test_account_book_start_period_keeps_numeric_excel_date_cell(self) -> None:
        build_workbook(self.complete_dir, self.output_path)
        with ZipFile(self.output_path) as zf:
            workbook = ET.fromstring(zf.read("xl/workbook.xml"))
            rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall("pkgrel:Relationship", NS)}
            rid = next(
                sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
                for sheet in workbook.findall("main:sheets/main:sheet", NS)
                if sheet.attrib["name"] == " 组织账套表 "
            )
            sheet = ET.fromstring(zf.read("xl/" + relmap[rid]))
            cell = sheet.find('.//main:sheetData/main:row[@r="4"]/main:c[@r="F4"]', NS)
            self.assertIsNotNone(cell)
            self.assertNotEqual(cell.attrib.get("t"), "inlineStr")
            value_node = cell.find("main:v", NS)
            self.assertIsNotNone(value_node)
            self.assertEqual(value_node.text, "45292")

    def test_balance_sheet_appends_total_row(self) -> None:
        build_workbook(self.complete_dir, self.output_path)
        rows = read_sheet_rows(self.output_path, " 科目余额表")
        total_row = next(row for row in rows if row.get("F") == "合计")
        self.assertEqual(total_row.get("C", ""), "合计")
        self.assertEqual(total_row.get("G", ""), "6383493.24")
        self.assertEqual(total_row.get("H", ""), "6383493.24")
        self.assertEqual(total_row.get("M", ""), "149948.48")
        self.assertEqual(total_row.get("N", ""), "149948.48")

    def test_asset_summary_sheet_preserves_two_header_rows(self) -> None:
        build_workbook(self.complete_dir, self.output_path)
        rows = read_sheet_rows(self.output_path, "村集体经济组织固定资产表（汇总表）")
        row4 = next(row for row in rows if row["row_number"] == 4)
        row5 = next(row for row in rows if row["row_number"] == 5)
        self.assertEqual(row4.get("U", ""), "出租对象")
        self.assertEqual(row4.get("X", ""), "资产数量/面积")
        self.assertEqual(row5.get("A", ""), "N2411025MF3559576G")
        self.assertEqual(row5.get("G", ""), "2020年襄城县颍桥回族镇东街道路建设项目")

    def test_well_sheet_preserves_multi_row_headers(self) -> None:
        build_workbook(self.complete_dir, self.output_path)
        rows = read_sheet_rows(self.output_path, "村集体经济组织固定资产表（水井）")
        row4 = next(row for row in rows if row["row_number"] == 4)
        row5 = next(row for row in rows if row["row_number"] == 5)
        row6 = next(row for row in rows if row["row_number"] == 6)
        self.assertEqual(row4.get("I", ""), "是否有水")
        self.assertEqual(row5.get("P", ""), "姓名")
        self.assertEqual(row6.get("A", ""), "N2411025MF3559576G")
        self.assertEqual(row6.get("F", ""), "机井28")
        self.assertEqual(row6.get("G", ""), "机井28")

    def test_subject_sheet_uses_source_balance_direction_and_aux_flag(self) -> None:
        build_workbook(self.complete_dir, self.output_path)
        rows = read_sheet_rows(self.output_path, "科目表")
        by_code = {row["E"]: row for row in rows if row["row_number"] >= 4 and row.get("E")}

        self.assertEqual(by_code["101"]["K"], "借")
        self.assertEqual(by_code["101"]["L"], "否")
        self.assertEqual(by_code["102"]["K"], "借")
        self.assertEqual(by_code["102"]["L"], "是")
        self.assertEqual(by_code["152"]["K"], "贷")
        self.assertEqual(by_code["152001"]["K"], "贷")

    def test_resource_sheet_maps_usage_status_and_area_columns(self) -> None:
        build_workbook(self.complete_dir, self.output_path)
        rows = read_sheet_rows(self.output_path, "村集体经济组织资源信息表")
        by_code = {row["E"]: row for row in rows if row["row_number"] >= 5 and row.get("E")}

        self.assertEqual(by_code["Y4110251012010006"]["I"], "出租经营")
        self.assertEqual(by_code["Y4110251012010006"]["M"], "5.20")
        self.assertEqual(by_code["Y4110251012010006"]["N"], "亩")
        self.assertEqual(by_code["Y4110251012010006"]["P"], "村内")
        self.assertEqual(by_code["Y4110251012010006"]["Q"], "2025-12-31")
        self.assertEqual(by_code["Y4110251012010007"]["I"], "自主经营")
        self.assertEqual(by_code["Y4110251012010007"]["M"], "7.04")

    def test_well_sheet_maps_caretaker_and_repairer_fields(self) -> None:
        build_workbook(self.complete_dir, self.output_path)
        rows = read_sheet_rows(self.output_path, "村集体经济组织固定资产表（水井）")
        by_code = {row["E"]: row for row in rows if row["row_number"] >= 6 and row.get("E")}

        first = by_code["C4110251012010089"]
        self.assertEqual(first["P"], "陈广讯")
        self.assertEqual(first["Q"], "15037449998")
        self.assertEqual(first["R"], "王雪平")
        self.assertEqual(first["S"], "18937485150")
        self.assertEqual(first["T"], "崔二恒")
        self.assertEqual(first["U"], "13837403975")

        second = by_code["C4110251012010081"]
        self.assertEqual(second["R"], "关海军")
        self.assertEqual(second["S"], "1599364814")
        self.assertEqual(second["T"], "关海军")
        self.assertEqual(second["U"], "1599364814")

    def test_asset_summary_sheet_maps_lease_and_financial_tail_columns(self) -> None:
        build_workbook(self.complete_dir, self.output_path)
        rows = read_sheet_rows(self.output_path, "村集体经济组织固定资产表（汇总表）")
        by_code = {row["E"]: row for row in rows if row["row_number"] >= 5 and row.get("E")}

        self_use = by_code["C4110251012010009"]
        self.assertEqual(self_use["U"], "")
        self.assertEqual(self_use["V"], "")
        self.assertEqual(self_use["W"], "")
        self.assertEqual(self_use["X"], "630.00")
        self.assertEqual(self_use["Z"], "200800.00")
        self.assertEqual(self_use["AA"], "200800.00")
        self.assertEqual(self_use["AB"], "0.00")

        leased = by_code["C4110251012010055"]
        self.assertEqual(leased["U"], "襄城县恒美发制品有限公司")
        self.assertEqual(leased["V"], "240")
        self.assertEqual(leased["W"], "5.15")
        self.assertEqual(leased["X"], "800.00")
        self.assertEqual(leased["Z"], "1035944.44")
        self.assertEqual(leased["AA"], "1035944.44")
        self.assertEqual(leased["AB"], "0.00")

    def test_master_and_balance_sheets_keep_semantic_values(self) -> None:
        build_workbook(self.complete_dir, self.output_path)

        org_rows = read_sheet_rows(self.output_path, "组织单位信息表")
        org = next(row for row in org_rows if row["row_number"] == 4)
        self.assertEqual(org["A"], "襄城县")
        self.assertEqual(org["B"], "颍桥回族镇")
        self.assertEqual(org["C"], "东街村")
        self.assertEqual(org["G"], "陈广勋")
        self.assertEqual(org["H"], "15037449998")

        user_rows = read_sheet_rows(self.output_path, "用户信息表")
        by_login = {row["D"]: row for row in user_rows if row["row_number"] >= 4 and row.get("D")}
        self.assertEqual(by_login["15037449998"]["H"], "村审核员")
        self.assertEqual(by_login["13663740812"]["H"], "村记账员")

        balance_rows = read_sheet_rows(self.output_path, " 科目余额表")
        bank = next(row for row in balance_rows if row.get("E") == "102")
        self.assertEqual(bank["F"], "银行存款")
        self.assertEqual(bank["G"], "12949.70")
        self.assertEqual(bank["H"], "0.00")
        self.assertEqual(bank["I"], "0.00")
        self.assertEqual(bank["J"], "0.00")
        self.assertEqual(bank["K"], "12949.70")
        self.assertEqual(bank["L"], "0.00")
        self.assertEqual(bank["M"], "10237.12")
        self.assertEqual(bank["N"], "10237.12")

    def test_documented_context_and_period_rules(self) -> None:
        context = parse_context(self.complete_dir)
        self.assertEqual(context.village, "东街村")
        self.assertEqual(context.contact_name, "陈广勋")
        self.assertEqual(context.contact_phone, "15037449998")

        self.assertEqual(normalize_resource_period("2025-07-18-2041-10-31"), "2025-07-18~2041-10-31")
        self.assertEqual(normalize_resource_period("2025/07/18-2041/10/31"), "2025/07/18-2041/10/31")

    def test_documented_annual_rent_rules(self) -> None:
        self.assertEqual(contract_years("2010-06-30", "2030-06-30"), 20)
        self.assertIsNone(contract_years("2025-01-01", "2025-12-31"))
        self.assertEqual(annual_rent("68000.00", "2010-06-30", "2030-06-30"), "3400.00")
        self.assertEqual(annual_rent("0.00", "2025-01-01", "2025-12-31"), "0.00")
        self.assertEqual(annual_rent("100.00", "2025-01-01", "2025-12-31"), "")

if __name__ == "__main__":
    unittest.main()
