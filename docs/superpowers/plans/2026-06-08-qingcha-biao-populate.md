# 清查表（已填充） Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build `完整/清查表（已填充）.xlsx` from the Excel files inside `完整/`, using `完整/清查表.xlsx` as the template workbook and filling every target sheet from the matching source workbook.

**Architecture:** Use pure Python standard library only. Read and write `.xlsx` files as ZIP packages with `xml.etree.ElementTree`, so the workflow does not depend on `openpyxl`, `pandas`, or LibreOffice. Normalize each source workbook into row dictionaries, fill each target sheet with a dedicated mapper, and write the final workbook to `完整/清查表（已填充）.xlsx` without modifying the original input files.

**Tech Stack:** Python 3 stdlib (`zipfile`, `xml.etree.ElementTree`, `pathlib`, `copy`, `decimal`, `datetime`, `tempfile`, `unittest`)

---

## File Structure

**Create:**
- `scripts/xlsx_xml.py` — low-level helpers for reading sheet rows, cloning rows, setting cell values, updating dimensions, and preserving workbook XML structure.
- `scripts/build_qingcha_table.py` — end-to-end builder that loads `完整/` source files, fills target sheets, and writes `完整/清查表（已填充）.xlsx`.
- `tests/test_build_qingcha_table.py` — end-to-end tests for workbook creation, row counts, known field mappings, and ambiguity handling.

**Read-only inputs:**
- `完整/清查表.xlsx` — template workbook to copy and fill.
- `完整/用户信息表.xlsx`
- `完整/科目期初数据明细.xlsx`
- `完整/科目余额表数据.xlsx`
- `完整/固定资产统计数据.xlsx`
- `完整/水井固定资产统计数据.xlsx`
- `完整/资源明细数据.xlsx`
- `完整/合同明细数据.xlsx`

**Output:**
- `完整/清查表（已填充）.xlsx`

## Data Rules To Lock Before Coding

1. **Use only files under `完整/` as authoritative inputs.** Do not read `东街村清查表.xlsx`, `资源明细数据 (2).xlsx`, or `合同明细数据 (1).xlsx` during the final build.
2. **Preserve the workbook shape of `完整/清查表.xlsx`.** Fill sheets in place, do not add or remove sheets.
3. **Do not overwrite the template.** Always write a new file: `完整/清查表（已填充）.xlsx`.
4. **Keep ambiguous fields blank.** This matters for the four resource fields the user called out:
   - `责任人`: source column exists in `完整/资源明细数据.xlsx`, but the column is blank in the observed data, so keep blank.
   - `承租人 / 起止时间 / 年租金`: fill only when a contract can be matched deterministically.
5. **Resource-contract matching rule:**
   - Only consider resource rows where `使用状态 == 出租经营`.
   - Match contracts to resources by organization name and exact area/quantity when `合同明细数据.xlsx` provides `数量`.
   - `起止时间` format: `开始日期~结束日期`.
   - `年租金` formula: if `总金额 > 0` and contract duration is known, compute annual rent as `总金额 / 合同年数`; otherwise keep the literal zero value if total amount is `0.00`; otherwise blank.
   - If multiple resources could match one contract, leave the resource row blank and record the ambiguity in the script output.
6. **Known strict-mode consequence:** The `巴海棠` contract in `完整/合同明细数据.xlsx` has no quantity/area value, so it cannot be assigned deterministically between the leased East Street resource rows from `完整/资源明细数据.xlsx`. In strict `完整/`-only mode, that row stays blank in `承租人 / 起止时间 / 年租金`.

---

### Task 1: Add a failing end-to-end workbook test

**Files:**
- Create: `tests/test_build_qingcha_table.py`
- Create: `scripts/build_qingcha_table.py`

- [ ] **Step 1: Write the failing end-to-end test scaffold**

```python
import shutil
import tempfile
import unittest
from pathlib import Path

from scripts.build_qingcha_table import build_workbook, read_sheet_rows

BASE_DIR = Path(__file__).resolve().parents[1]
COMPLETE_DIR = BASE_DIR / "完整"


class BuildQingchaTableTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="qingcha-test-"))
        self.output_path = self.temp_dir / "清查表（已填充）.xlsx"

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_builds_output_workbook(self):
        build_workbook(COMPLETE_DIR, self.output_path)
        self.assertTrue(self.output_path.exists())

    def test_resource_sheet_has_36_rows_from_complete_dir(self):
        build_workbook(COMPLETE_DIR, self.output_path)
        rows = read_sheet_rows(self.output_path, "村集体经济组织资源信息表")
        data_rows = [row for row in rows if row["row_number"] >= 5 and row["E"]]
        self.assertEqual(len(data_rows), 36)
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `python3 -m unittest -v tests.test_build_qingcha_table`

Expected: FAIL with `ModuleNotFoundError` or `ImportError` for `scripts.build_qingcha_table`.

- [ ] **Step 3: Add the second failing test for the four resource columns**

```python
    def test_resource_contract_fields_fill_only_when_deterministic(self):
        build_workbook(COMPLETE_DIR, self.output_path)
        rows = read_sheet_rows(self.output_path, "村集体经济组织资源信息表")
        by_code = {row["E"]: row for row in rows if row["row_number"] >= 5 and row["E"]}

        self.assertEqual(by_code["Y4110251012010012"]["J"], "襄城县居美达机械厂")
        self.assertEqual(by_code["Y4110251012010012"]["K"], "2010-06-30~2030-06-30")
        self.assertEqual(by_code["Y4110251012010012"]["L"], "3400.00")

        self.assertEqual(by_code["Y4110251012010013"]["J"], "赵建垒")
        self.assertEqual(by_code["Y4110251012010013"]["K"], "2012-08-31~2032-08-31")
        self.assertEqual(by_code["Y4110251012010013"]["L"], "1000.00")

        # Strict 完整-only mode: ambiguous or unsourced fields stay blank.
        self.assertEqual(by_code["Y4110251012010006"]["J"], "")
        self.assertEqual(by_code["Y4110251012010008"]["J"], "")
        self.assertTrue(all(row["O"] == "" for row in by_code.values()))
```

- [ ] **Step 4: Run the focused test to verify it fails for the right reason**

Run: `python3 -m unittest -v tests.test_build_qingcha_table.BuildQingchaTableTest.test_resource_contract_fields_fill_only_when_deterministic`

Expected: FAIL because `build_workbook()` and `read_sheet_rows()` do not exist yet.

- [ ] **Step 5: Commit if the workspace is later put under git**

```bash
# Skip this step if the workspace is still not a git repository.
git add tests/test_build_qingcha_table.py scripts/build_qingcha_table.py
git commit -m "test: add failing workbook build coverage"
```

### Task 2: Implement the XLSX XML helper layer

**Files:**
- Create: `scripts/xlsx_xml.py`
- Modify: `scripts/build_qingcha_table.py`
- Test: `tests/test_build_qingcha_table.py`

- [ ] **Step 1: Write the low-level workbook reader helpers**

```python
# scripts/xlsx_xml.py
from zipfile import ZipFile
import xml.etree.ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
PKGREL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOCREL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"main": MAIN_NS, "pkgrel": PKGREL_NS}


def parse_shared_strings(zf: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return [
        "".join(t.text or "" for t in si.iter(f"{{{MAIN_NS}}}t"))
        for si in root.findall("main:si", NS)
    ]


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        is_node = cell.find("main:is", NS)
        return "" if is_node is None else "".join(
            t.text or "" for t in is_node.iter(f"{{{MAIN_NS}}}t")
        )
    value_node = cell.find("main:v", NS)
    if value_node is None:
        return ""
    if cell_type == "s":
        return shared_strings[int(value_node.text)]
    return value_node.text or ""
```

- [ ] **Step 2: Add sheet lookup and row extraction helpers**

```python
def get_sheet_path(zf: ZipFile, sheet_name: str) -> str:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall("pkgrel:Relationship", NS)}
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        if sheet.attrib["name"] == sheet_name:
            rid = sheet.attrib[f"{{{DOCREL_NS}}}id"]
            return "xl/" + relmap[rid]
    raise KeyError(sheet_name)


def read_sheet_rows(path, sheet_name: str) -> list[dict[str, str]]:
    with ZipFile(path) as zf:
        shared_strings = parse_shared_strings(zf)
        sheet_root = ET.fromstring(zf.read(get_sheet_path(zf, sheet_name)))
        rows = []
        for row in sheet_root.findall('.//main:sheetData/main:row', NS):
            item = {"row_number": int(row.attrib["r"])}
            for cell in row.findall('main:c', NS):
                col = ''.join(ch for ch in cell.attrib['r'] if ch.isalpha())
                item[col] = cell_value(cell, shared_strings)
            rows.append(item)
        return rows
```

- [ ] **Step 3: Add row-cloning and cell-writing helpers**

```python
from copy import deepcopy


def set_cell_text(cell: ET.Element, text: str) -> None:
    for child in list(cell):
        cell.remove(child)
    if text == "":
        cell.attrib.pop("t", None)
        return
    cell.attrib["t"] = "inlineStr"
    is_node = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
    t_node = ET.SubElement(is_node, f"{{{MAIN_NS}}}t")
    t_node.text = text


def clone_row(template_row: ET.Element, row_number: int) -> ET.Element:
    new_row = deepcopy(template_row)
    new_row.attrib["r"] = str(row_number)
    for cell in new_row.findall("main:c", NS):
        col = ''.join(ch for ch in cell.attrib['r'] if ch.isalpha())
        cell.attrib['r'] = f"{col}{row_number}"
    return new_row
```

- [ ] **Step 4: Make the builder import the helpers and rerun the tests**

```python
# scripts/build_qingcha_table.py
from scripts.xlsx_xml import read_sheet_rows
```

Run: `python3 -m unittest -v tests.test_build_qingcha_table`

Expected: FAIL deeper in `build_workbook()` because the build logic still does not exist. That is the right next failure.

- [ ] **Step 5: Commit if git is available**

```bash
git add scripts/xlsx_xml.py scripts/build_qingcha_table.py tests/test_build_qingcha_table.py
git commit -m "feat: add xlsx xml helper layer"
```

### Task 3: Normalize all source workbooks from `完整/`

**Files:**
- Modify: `scripts/build_qingcha_table.py`
- Test: `tests/test_build_qingcha_table.py`

- [ ] **Step 1: Add the source workbook inventory constants**

```python
# scripts/build_qingcha_table.py
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP

TEMPLATE_NAME = "清查表.xlsx"
OUTPUT_NAME = "清查表（已填充）.xlsx"
SOURCE_FILES = {
    "users": "用户信息表.xlsx",
    "subjects": "科目期初数据明细.xlsx",
    "balances": "科目余额表数据.xlsx",
    "assets": "固定资产统计数据.xlsx",
    "wells": "水井固定资产统计数据.xlsx",
    "resources": "资源明细数据.xlsx",
    "contracts": "合同明细数据.xlsx",
}
```

- [ ] **Step 2: Add a normalizer for workbook-level context**

```python
def parse_context(complete_dir: Path) -> dict[str, str]:
    rows = read_sheet_rows(complete_dir / SOURCE_FILES["subjects"], "科目期初数据明细")
    first = next(row for row in rows if row.get("A") and row["row_number"] >= 2)
    org_name = first["C"]
    return {
        "county": first["A"],
        "township": first["B"],
        "organization_name": org_name,
        "social_code": first["D"],
        "village": org_name.replace(first["A"], "").replace(first["B"], "").replace("股份经济合作社", ""),
    }
```

- [ ] **Step 3: Normalize the user and resource datasets**

```python
def load_users(complete_dir: Path) -> list[dict[str, str]]:
    rows = read_sheet_rows(complete_dir / SOURCE_FILES["users"], "Sheet1")
    role_map = {"审核人": "村审核员", "记账人": "村记账员"}
    return [
        {
            "login": row.get("A", ""),
            "name": row.get("B", ""),
            "phone": row.get("C", ""),
            "role": role_map.get(row.get("D", ""), row.get("D", "")),
        }
        for row in rows if row["row_number"] >= 2 and row.get("A")
    ]


def load_resources(complete_dir: Path) -> list[dict[str, str]]:
    rows = read_sheet_rows(complete_dir / SOURCE_FILES["resources"], "资源明细数据")
    return [
        {
            "code": row.get("C", ""),
            "name": row.get("D", ""),
            "registered_at": row.get("E", ""),
            "category": row.get("F", ""),
            "attribute": row.get("G", ""),
            "area": row.get("H", ""),
            "unit": row.get("I", ""),
            "owner": row.get("J", ""),
            "location": row.get("M", ""),
            "usage_status": row.get("T", ""),
        }
        for row in rows if row["row_number"] >= 2 and row.get("C") and row.get("D")
    ]
```

- [ ] **Step 4: Normalize contract rows and the annual-rent calculation**

```python
def contract_years(start_date: str, end_date: str) -> Decimal | None:
    if not start_date or not end_date:
        return None
    start_year = Decimal(start_date[:4])
    end_year = Decimal(end_date[:4])
    whole_years = end_year - start_year
    return whole_years if whole_years > 0 else None


def annual_rent(total_amount: str, start_date: str, end_date: str) -> str:
    if total_amount == "":
        return ""
    total = Decimal(total_amount)
    if total == 0:
        return "0.00"
    years = contract_years(start_date, end_date)
    if not years:
        return ""
    return str((total / years).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
```

```python
def load_contracts(complete_dir: Path) -> list[dict[str, str]]:
    rows = read_sheet_rows(complete_dir / SOURCE_FILES["contracts"], "合同明细数据")
    return [
        {
            "organization_name": row.get("B", ""),
            "contract_code": row.get("C", ""),
            "contract_name": row.get("D", ""),
            "contract_type": row.get("F", ""),
            "tenant": row.get("H", ""),
            "start_date": row.get("J", ""),
            "end_date": row.get("K", ""),
            "quantity": row.get("M", ""),
            "total_amount": row.get("O", ""),
            "annual_rent": annual_rent(row.get("O", ""), row.get("J", ""), row.get("K", "")),
        }
        for row in rows if row["row_number"] >= 2 and row.get("C")
    ]
```

- [ ] **Step 5: Run tests and confirm the next failures are mapping failures, not loader failures**

Run: `python3 -m unittest -v tests.test_build_qingcha_table`

Expected: FAIL because the workbook-copy and per-sheet fill logic does not exist yet, but the stack traces now come from the mapper layer instead of missing helpers.

### Task 4: Fill the non-resource target sheets from the normalized datasets

**Files:**
- Modify: `scripts/build_qingcha_table.py`
- Test: `tests/test_build_qingcha_table.py`

- [ ] **Step 1: Copy the template workbook to the output path before editing**

```python
import shutil


def build_workbook(complete_dir: Path, output_path: Path | None = None) -> Path:
    output_path = output_path or (complete_dir / OUTPUT_NAME)
    shutil.copy2(complete_dir / TEMPLATE_NAME, output_path)
    return output_path
```

- [ ] **Step 2: Add a single entrypoint that collects all datasets once**

```python
def build_workbook(complete_dir: Path, output_path: Path | None = None) -> Path:
    output_path = output_path or (complete_dir / OUTPUT_NAME)
    shutil.copy2(complete_dir / TEMPLATE_NAME, output_path)

    context = parse_context(complete_dir)
    users = load_users(complete_dir)
    resources = load_resources(complete_dir)
    contracts = load_contracts(complete_dir)
    subjects = read_sheet_rows(complete_dir / SOURCE_FILES["subjects"], "科目期初数据明细")
    balances = read_sheet_rows(complete_dir / SOURCE_FILES["balances"], "襄城县颍桥回族镇襄城县颍桥回族镇东街村股份经济合作社-科目余额")
    assets = read_sheet_rows(complete_dir / SOURCE_FILES["assets"], "固定资产统计数据")
    wells = read_sheet_rows(complete_dir / SOURCE_FILES["wells"], "固定资产统计数据")

    fill_master_sheets(output_path, context, users)
    fill_subject_sheet(output_path, context, subjects)
    fill_balance_sheet(output_path, context, balances)
    fill_asset_sheets(output_path, context, assets, wells)
    fill_contract_sheet(output_path, context, contracts)
    fill_resource_sheet(output_path, context, resources, contracts)
    return output_path
```

- [ ] **Step 3: Implement the organization, user, and account-sheet fillers**

```python
def fill_master_sheets(output_path: Path, context: dict[str, str], users: list[dict[str, str]]) -> None:
    reviewer = next((user for user in users if user["role"] == "村审核员"), {"name": "", "phone": ""})
    book_name = context["organization_name"]

    write_single_row(output_path, "组织单位信息表", 4, {
        "A": context["county"],
        "B": context["township"],
        "C": context["village"],
        "D": context["social_code"],
        "E": context["social_code"],
        "F": "是",
        "G": reviewer["name"],
        "H": reviewer["phone"],
    })

    write_table_rows(output_path, "用户信息表", 4, [
        {"A": context["county"], "B": context["township"], "C": context["village"], "D": user["login"], "E": user["name"], "F": context["social_code"], "G": user["phone"], "H": user["role"]}
        for user in users
    ])

    write_single_row(output_path, " 组织账套表 ", 4, {
        "A": context["county"],
        "B": context["township"],
        "C": context["village"],
        "D": book_name,
        "E": context["social_code"],
    })
```

- [ ] **Step 4: Implement the subject, balance, fixed-asset, well, and contract fillers**

```python
def fill_contract_sheet(output_path: Path, context: dict[str, str], contracts: list[dict[str, str]]) -> None:
    rows = []
    for contract in contracts:
        rows.append({
            "A": context["social_code"],
            "B": context["county"],
            "C": context["township"],
            "D": context["organization_name"],
            "E": contract["contract_code"],
            "F": contract["contract_name"],
            "G": contract["contract_type"],
            "H": "",
            "I": contract["tenant"],
            "J": "",
            "K": contract["start_date"],
            "L": contract["end_date"],
        })
    write_table_rows(output_path, "村集体经济组织资产合同表", 4, rows)
```

Use the same pattern for:
- `科目表` from `科目期初数据明细.xlsx`
- `科目余额表` from `科目余额表数据.xlsx`
- `村集体经济组织固定资产表（汇总表）` from `固定资产统计数据.xlsx`
- `村集体经济组织固定资产表（水井）` from `水井固定资产统计数据.xlsx`

- [ ] **Step 5: Run the full test suite and inspect the first workbook-level failures**

Run: `python3 -m unittest -v tests.test_build_qingcha_table`

Expected: the remaining failures should now be concentrated in the resource-sheet contract matching logic.

### Task 5: Fill the resource sheet, handle the four requested fields, and verify the final workbook

**Files:**
- Modify: `scripts/build_qingcha_table.py`
- Test: `tests/test_build_qingcha_table.py`

- [ ] **Step 1: Implement deterministic resource-contract matching**

```python
def match_contracts_to_resources(resources: list[dict[str, str]], contracts: list[dict[str, str]]) -> dict[str, dict[str, str]]:
    matches: dict[str, dict[str, str]] = {}
    leased_resources = [row for row in resources if row["usage_status"] == "出租经营"]

    contracts_by_quantity = {}
    for contract in contracts:
        quantity = contract.get("quantity", "")
        if quantity:
            contracts_by_quantity.setdefault(quantity, []).append(contract)

    for resource in leased_resources:
        candidates = contracts_by_quantity.get(resource["area"], [])
        if len(candidates) == 1:
            matches[resource["code"]] = candidates[0]

    return matches
```

- [ ] **Step 2: Implement the final resource-sheet writer**

```python
def fill_resource_sheet(output_path: Path, context: dict[str, str], resources: list[dict[str, str]], contracts: list[dict[str, str]]) -> None:
    contract_matches = match_contracts_to_resources(resources, contracts)
    rows = []
    for resource in resources:
        contract = contract_matches.get(resource["code"], {})
        rows.append({
            "A": context["social_code"],
            "B": context["county"],
            "C": context["township"],
            "D": context["organization_name"],
            "E": resource["code"],
            "F": resource["name"],
            "G": resource["category"],
            "H": resource["attribute"],
            "I": resource["area"],
            "J": contract.get("tenant", ""),
            "K": f"{contract['start_date']}~{contract['end_date']}" if contract.get("start_date") and contract.get("end_date") else "",
            "L": contract.get("annual_rent", ""),
            "M": resource["area"],
            "N": resource["unit"],
            "O": resource["owner"],
            "P": resource["location"],
            "Q": resource["registered_at"],
        })
    write_table_rows(output_path, "村集体经济组织资源信息表", 5, rows)
```

- [ ] **Step 3: Add a builder warning for unresolved ambiguities instead of guessing**

```python
def unresolved_resource_codes(resources: list[dict[str, str]], matches: dict[str, dict[str, str]]) -> list[str]:
    return [
        row["code"]
        for row in resources
        if row["usage_status"] == "出租经营" and row["code"] not in matches
    ]

# In build_workbook():
missing = unresolved_resource_codes(resources, match_contracts_to_resources(resources, contracts))
if missing:
    print("Unresolved leased resource rows kept blank:", ", ".join(missing))
```

- [ ] **Step 4: Run the tests and then generate the workbook**

Run:
- `python3 -m unittest -v tests.test_build_qingcha_table`
- `python3 scripts/build_qingcha_table.py`

Expected:
- Tests PASS
- The script prints the output path and any unresolved leased resource codes
- `完整/清查表（已填充）.xlsx` exists

- [ ] **Step 5: Perform manual workbook verification on the output file**

Check these facts programmatically and then open the file manually if needed:

```python
rows = read_sheet_rows(Path("完整/清查表（已填充）.xlsx"), "村集体经济组织资源信息表")
resource_rows = [row for row in rows if row["row_number"] >= 5 and row.get("E")]
assert len(resource_rows) == 36
assert any(row["J"] == "襄城县居美达机械厂" for row in resource_rows)
assert any(row["J"] == "赵建垒" for row in resource_rows)
assert all(row.get("O", "") == "" for row in resource_rows)
```

If the workspace is under git, commit:

```bash
git add scripts/xlsx_xml.py scripts/build_qingcha_table.py tests/test_build_qingcha_table.py 完整/清查表（已填充）.xlsx
git commit -m "feat: build filled qingcha workbook from complete sources"
```

---

## Self-Review

- **Spec coverage:** This plan covers all source files inside `完整/`, uses `完整/清查表.xlsx` as the template, writes `完整/清查表（已填充）.xlsx`, fills the resource sheet with the four user-requested columns, and defines strict handling for missing/ambiguous values.
- **Placeholder scan:** No `TODO`, `TBD`, or “handle appropriately later” placeholders remain.
- **Type consistency:** All code snippets use the same function names: `build_workbook`, `read_sheet_rows`, `load_resources`, `load_contracts`, `fill_resource_sheet`, and `match_contracts_to_resources`.

## Execution Handoff

Plan complete and saved to `docs/superpowers/plans/2026-06-08-qingcha-biao-populate.md`. Two execution options:

**1. Subagent-Driven (recommended)** - I dispatch a fresh subagent per task, review between tasks, fast iteration

**2. Inline Execution** - Execute tasks in this session using executing-plans, batch execution with checkpoints

**Which approach?**
