from __future__ import annotations

"""根据模板工作簿和源数据文件构建已填充的村级清查表工作簿。

本脚本有意绑定当前模板。它不是通用 Excel 导入器：下面的源工作簿名称、
工作表名称、表头行、数据起始行以及 Excel 列字母，都是当前清查表模板的约定。

硬编码清单，按存在原因分组：
- 文件/工作簿约定：``SOURCE_FILES``、``TEMPLATE_NAME``、``OUTPUT_NAME``。
- 工作表名称约定：``SOURCE_SHEETS`` 和 ``TARGET_SHEETS``。部分目标工作表
  名称包含前后空格，因为真实工作簿中也有这些空格。
- 版式约定：``TARGET_LAYOUTS`` 记录起始行、模板行、清理起始行，以便在替换
  旧数据行时保留多行表头。
- 业务词汇：``审核人``、``出租经营`` 等角色/状态标签。
- 当前业务规则：村名推导、年租金计算，以及资源直接出租信息的使用。
- 占位行为：不支持的模板列会有意留空，无形资产工作表会输出为空白模板行。

``tests/test_build_qingcha_table.py`` 中的测试是行为锁。修改注释、常量或映射
后，在重新生成输出工作簿之前应运行工作簿测试。
"""

from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import shutil
import sys

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from scripts.xlsx_xml import apply_text_color_to_cells, list_sheet_names, read_sheet_rows, write_single_row, write_table_rows

BASE_DIR = Path(__file__).resolve().parents[1]
COMPLETE_DIR = BASE_DIR / "完整"
TEMPLATE_NAME = "清查表.xlsx"
OUTPUT_NAME = "清查表（已填充）.xlsx"

# 源工作簿文件名是当前 ``完整/`` 文件夹的输入约定。
SOURCE_FILES = {
    "users": "用户信息表.xlsx",
    "subjects": "科目期初数据明细.xlsx",
    "balances": "科目余额表数据.xlsx",
    "assets": "固定资产统计数据.xlsx",
    "wells": "水井固定资产统计数据.xlsx",
    "resources": "资源明细数据.xlsx",
    "contracts": "合同明细数据.xlsx",
}

SOURCE_SHEETS = {
    "users": "Sheet1",
    "subjects": "科目期初数据明细",
    "assets": "固定资产统计数据",
    "wells": "固定资产统计数据",
    "resources": "资源明细数据",
    "contracts": "合同明细数据",
}

# 输出工作表名称必须与模板完全一致。不要去掉 ``account_book`` 或 ``balance``
# 中的空格：这些空格是真实工作表名称的一部分。
TARGET_SHEETS = {
    "organization": "组织单位信息表",
    "users": "用户信息表",
    "account_book": " 组织账套表 ",
    "subjects": "科目表",
    "balance": " 科目余额表",
    "asset_summary": "村集体经济组织固定资产表（汇总表）",
    "wells": "村集体经济组织固定资产表（水井）",
    "intangibles": "村集体经济组织无形资产表",
    "contracts": "村集体经济组织资产合同表",
    "resources": "村集体经济组织资源信息表",
}

# 科目余额表的源工作表名称包含组织名，不同村的数据会变化，因此按名称包含标记自动定位。
BALANCE_SHEET_MARKER = "-科目余额"

REVIEWER_SOURCE_ROLE = "审核人"
BOOKKEEPER_SOURCE_ROLE = "记账人"
ROLE_NAME_MAP = {
    REVIEWER_SOURCE_ROLE: "村审核员",
    BOOKKEEPER_SOURCE_ROLE: "村记账员",
}

COOPERATIVE_SUFFIX = "股份经济合作社"
YES_VALUE = "是"
TOTAL_LABEL = "合计"
LEASED_USAGE_STATUS = "出租经营"
ACCOUNTING_SYSTEM_NAME = "村集体经济组织会计制度(2024)"
ZERO_AMOUNT = "0.00"
HIGHLIGHT_FONT_RGB = "FFFF0000"
INTANGIBLE_BLANK_ROWS = 18
INTANGIBLE_COLUMNS = "ABCDEFGHIJKLMNOPQRST"


@dataclass(frozen=True)
class SheetLayout:
    """替换基于模板的输出工作表时使用的行号。

    ``start_row`` 是生成数据开始写入的位置。``template_row_number`` 是用于克隆
    以保留单元格样式和合并表头结构的行。``clear_existing_from_row`` 是写入新行
    之前开始删除旧模板/示例数据的位置。
    """

    start_row: int
    template_row_number: int | None = None
    clear_existing_from_row: int | None = None


TARGET_LAYOUTS = {
    "organization": SheetLayout(start_row=4),
    "users": SheetLayout(start_row=4),
    "account_book": SheetLayout(start_row=4),
    "subjects": SheetLayout(start_row=4, template_row_number=3, clear_existing_from_row=4),
    "balance": SheetLayout(start_row=5, template_row_number=4, clear_existing_from_row=5),
    "asset_summary": SheetLayout(start_row=5, template_row_number=4, clear_existing_from_row=5),
    "wells": SheetLayout(start_row=6, template_row_number=5, clear_existing_from_row=6),
    "intangibles": SheetLayout(start_row=5, template_row_number=5, clear_existing_from_row=5),
    "contracts": SheetLayout(start_row=4, template_row_number=3, clear_existing_from_row=4),
    "resources": SheetLayout(start_row=5, template_row_number=4, clear_existing_from_row=5),
}

# 固定资产汇总模板中存在、但当前 ``完整/`` 数据没有可信来源字段的目标列。
# 这些列按设计保持为空。
ASSET_SUMMARY_UNSUPPORTED_COLUMNS = ("O", "P", "Q", "R", "S", "T", "Y")


@dataclass
class Context:
    county: str
    township: str
    village: str
    organization_name: str
    social_code: str
    contact_name: str
    contact_phone: str


@dataclass
class BuildDiagnostics:
    """一次构建过程收集给前端展示的日志与诊断信息。"""

    logs: list[str]
    missing_files: list[str]
    unresolved_resources: list[str]
    missing_fill_cells: list[str]

    @classmethod
    def create(cls) -> "BuildDiagnostics":
        return cls(logs=[], missing_files=[], unresolved_resources=[], missing_fill_cells=[])

    def info(self, message: str) -> None:
        self.logs.append(f"INFO: {message}")

    def warning(self, message: str) -> None:
        self.logs.append(f"WARN: {message}")


class MissingSourceFilesError(Exception):
    """上传临时目录缺少必需源工作簿。"""

    def __init__(self, missing_files: list[str]) -> None:
        self.missing_files = missing_files
        super().__init__("缺少必需源工作簿：" + ", ".join(missing_files))


def required_source_filenames() -> list[str]:
    """返回 Web 上传模式要求用户提供的源工作簿文件名。"""
    return list(SOURCE_FILES.values())


def find_missing_source_files(source_dir: Path) -> list[str]:
    """检查上传源目录中缺少哪些必需工作簿；模板不属于上传项。"""
    return [
        filename
        for filename in required_source_filenames()
        if not (source_dir / filename).is_file()
    ]


def resolve_balance_sheet_name(complete_dir: Path, diagnostics: BuildDiagnostics | None = None) -> str:
    """自动找到科目余额源工作簿中名称包含 ``-科目余额`` 的工作表。"""
    path = complete_dir / SOURCE_FILES["balances"]
    sheet_names = list_sheet_names(path)
    candidates = [name for name in sheet_names if BALANCE_SHEET_MARKER in name]
    if len(candidates) == 1:
        sheet_name = candidates[0]
        if diagnostics is not None:
            diagnostics.info(f"已识别科目余额源工作表：{sheet_name}")
        return sheet_name
    if not candidates and len(sheet_names) == 1:
        sheet_name = sheet_names[0]
        if diagnostics is not None:
            diagnostics.warning(f"科目余额源工作簿没有包含 {BALANCE_SHEET_MARKER} 的工作表，使用唯一工作表：{sheet_name}")
        return sheet_name
    raise KeyError(f"无法唯一识别科目余额源工作表，候选：{', '.join(sheet_names)}")


def validate_source_files(source_dir: Path, diagnostics: BuildDiagnostics | None = None) -> None:
    """缺少源工作簿时写入日志并抛出可被 API 捕获的异常。"""
    missing = find_missing_source_files(source_dir)
    if not missing:
        if diagnostics is not None:
            diagnostics.info("已找到全部必需源工作簿。")
        return

    if diagnostics is not None:
        diagnostics.missing_files.extend(missing)
        for filename in missing:
            diagnostics.warning(f"缺少上传文件：{filename}")
    raise MissingSourceFilesError(missing)


def report_missing_fill(
    diagnostics: BuildDiagnostics | None,
    sheet_name: str,
    row_number: int,
    col: str,
    value: str,
) -> None:
    """记录模板中缺少目标单元格导致非空值无法写入的情况。"""
    if diagnostics is None:
        return
    message = f"模板缺少目标单元格，无法写入：sheet={sheet_name}, row={row_number}, col={col}, value={value}"
    diagnostics.missing_fill_cells.append(message)
    diagnostics.warning(message)


def write_layout_rows(
    output_path: Path,
    sheet_key: str,
    rows: list[dict[str, str]],
    diagnostics: BuildDiagnostics | None = None,
) -> None:
    """使用集中记录的目标版式写入生成行。"""
    layout = TARGET_LAYOUTS[sheet_key]
    write_table_rows(
        output_path,
        TARGET_SHEETS[sheet_key],
        layout.start_row,
        rows,
        template_row_number=layout.template_row_number,
        clear_existing_from_row=layout.clear_existing_from_row,
        missing_cell_callback=lambda sheet, row, col, value: report_missing_fill(diagnostics, sheet, row, col, value),
    )


def normalize_resource_period(period: str) -> str:
    """将源资源期间从 ``YYYY-MM-DD-YYYY-MM-DD`` 转换为 ``start~end``。

    资源源文件会把部分出租期间存成一个 21 字符的字符串。只有这个精确形态会被
    规范化；其他值会原样保留，让意外的源格式保持可见，而不是被猜测处理。
    """
    if len(period) == 21 and period[4] == "-" and period[7] == "-" and period[10] == "-" and period[15] == "-":
        return f"{period[:10]}~{period[11:]}"
    return period


def parse_context(complete_dir: Path) -> Context:
    """从科目和用户源数据推导整个工作簿通用的组织上下文。

    第一条科目数据行携带县、乡镇、组织、统一社会信用代码等值。联系人优先使用
    源角色为 ``审核人`` 的用户；如果不存在，则使用第一条用户行。``village`` 是
    从组织名称推导出的村名，该规则较脆弱但与当前源数据兼容。
    """
    subject_rows = read_sheet_rows(complete_dir / SOURCE_FILES["subjects"], SOURCE_SHEETS["subjects"])
    first = next(row for row in subject_rows if row.get("A") and row["row_number"] >= 2)

    user_rows = read_sheet_rows(complete_dir / SOURCE_FILES["users"], SOURCE_SHEETS["users"])
    users = [row for row in user_rows if row["row_number"] >= 2 and row.get("A")]
    reviewer = next((row for row in users if row.get("D") == REVIEWER_SOURCE_ROLE), users[0] if users else {})

    county = first.get("A", "")
    township = first.get("B", "")
    organization_name = first.get("C", "")
    # 当前源名称形态为：县名 + 乡镇名 + 村名 + 股份经济合作社。
    # 如果该命名约定发生变化，这条规则应优先复核。
    village = organization_name.replace(county, "").replace(township, "").replace(COOPERATIVE_SUFFIX, "")

    return Context(
        county=county,
        township=township,
        village=village,
        organization_name=organization_name,
        social_code=first.get("D", ""),
        contact_name=reviewer.get("B", ""),
        contact_phone=reviewer.get("C", ""),
    )


def load_users(complete_dir: Path) -> list[dict[str, str]]:
    """从 ``用户信息表.xlsx`` 加载用户，并将角色转换为目标词汇。"""
    rows = read_sheet_rows(complete_dir / SOURCE_FILES["users"], SOURCE_SHEETS["users"])
    return [
        {
            "login": row.get("A", ""),
            "name": row.get("B", ""),
            "phone": row.get("C", ""),
            "role": ROLE_NAME_MAP.get(row.get("D", ""), row.get("D", "")),
        }
        for row in rows if row["row_number"] >= 2 and row.get("A")
    ]


def load_subjects(complete_dir: Path) -> list[dict[str, str]]:
    """加载科目行。

    源列由导入工作簿固定约定：G/H/I/J/K/L/M/P 会成为 ``fill_subject_sheet``
    使用的语义字段。余额方向和辅助核算标志从源数据复制，不做推断。
    """
    rows = read_sheet_rows(complete_dir / SOURCE_FILES["subjects"], SOURCE_SHEETS["subjects"])
    return [
        {
            "subject_name": row.get("G", ""),
            "subject_type": row.get("H", ""),
            "subject_level": row.get("I", ""),
            "subject_direction": row.get("J", ""),
            "subject_code": row.get("K", ""),
            "opening_balance": row.get("L", "") or ZERO_AMOUNT,
            "balance_direction": row.get("M", ""),
            "aux_enabled": row.get("P", ""),
        }
        for row in rows if row["row_number"] >= 2 and row.get("K")
    ]


def load_balances(complete_dir: Path, diagnostics: BuildDiagnostics | None = None) -> list[dict[str, str]]:
    """加载组织专属的科目余额工作表，包括底部的合计行。"""
    rows = read_sheet_rows(complete_dir / SOURCE_FILES["balances"], resolve_balance_sheet_name(complete_dir, diagnostics))
    return [
        row
        for row in rows
        if row["row_number"] >= 5 and (row.get("A") or row.get("B"))
    ]


def load_assets(complete_dir: Path) -> list[dict[str, str]]:
    """加载第 3 行及之后的固定资产汇总源数据行。"""
    rows = read_sheet_rows(complete_dir / SOURCE_FILES["assets"], SOURCE_SHEETS["assets"])
    return [
        {
            "asset_category": row.get("C", ""),
            "asset_name": row.get("D", ""),
            "asset_code": row.get("E", ""),
            "asset_attribute": row.get("F", ""),
            "specification": row.get("G", ""),
            "keeper": row.get("H", ""),
            "location": row.get("I", ""),
            "built_at": row.get("J", ""),
            "poverty_asset": row.get("K", ""),
            "usage_status": row.get("N", ""),
            "lease_target": row.get("O", ""),
            "lease_term_months": row.get("P", ""),
            "annual_rent": row.get("Q", ""),
            "quantity_or_area": row.get("R", ""),
            "original_value": row.get("S", ""),
            "depreciated": row.get("T", ""),
            "net_value": row.get("U", ""),
            "cleanup_value": row.get("V", ""),
        }
        for row in rows if row["row_number"] >= 3 and row.get("E")
    ]


def load_wells(complete_dir: Path) -> list[dict[str, str]]:
    """加载水井固定资产行，包括一长两员联系人字段。"""
    rows = read_sheet_rows(complete_dir / SOURCE_FILES["wells"], SOURCE_SHEETS["wells"])
    return [
        {
            "asset_name": row.get("D", ""),
            "asset_code": row.get("E", ""),
            "well_code": row.get("P", ""),
            "is_high_standard": row.get("Q", ""),
            "has_water": row.get("R", ""),
            "has_well_house": row.get("S", ""),
            "has_electricity": row.get("T", ""),
            "has_pump": row.get("U", ""),
            "has_caretaker": row.get("V", ""),
            "has_funding": row.get("W", ""),
            "has_supervision": row.get("X", ""),
            "well_head_name": row.get("Y", ""),
            "well_head_phone": row.get("Z", ""),
            "caretaker_name": row.get("AA", ""),
            "caretaker_phone": row.get("AB", ""),
            "repairer_name": row.get("AC", ""),
            "repairer_phone": row.get("AD", ""),
            "cleanup_value": row.get("AJ", ""),
        }
        for row in rows if row["row_number"] >= 3 and row.get("E")
    ]


def load_resources(complete_dir: Path) -> list[dict[str, str]]:
    """加载资源行，并在存在时保留直接出租明细。

    ``display_area`` 优先使用源列 U，因为它位于源数据的使用情况区域；否则回退
    到登记资源面积 H。后续会优先使用直接填写的承租人、期间、租金字段，而不是
    合同推导出的匹配值。
    """
    rows = read_sheet_rows(complete_dir / SOURCE_FILES["resources"], SOURCE_SHEETS["resources"])
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
            "display_area": row.get("U", "") or row.get("H", ""),
            "direct_annual_rent": row.get("V", ""),
            "direct_tenant": row.get("W", ""),
            "direct_period": row.get("X", ""),
            "organization_name": row.get("B", ""),
        }
        for row in rows if row["row_number"] >= 2 and row.get("C") and row.get("D")
    ]


def contract_years(start_date: str, end_date: str) -> Decimal | None:
    """返回当前年租金规则使用的仅按年份计算的期间。

    这里有意忽略月份和日期精度，因为现有源数据预期按 ``end_year - start_year``
    相除。同一年或起止颠倒的日期视为不可计算。
    """
    if not start_date or not end_date:
        return None
    start_year = Decimal(start_date[:4])
    end_year = Decimal(end_date[:4])
    whole_years = end_year - start_year
    return whole_years if whole_years > 0 else None


def annual_rent(total_amount: str, start_date: str, end_date: str) -> str:
    """使用仅按年份计算的期间规则，从总金额推导年租金。"""
    if total_amount == "":
        return ""
    total = Decimal(total_amount)
    if total == 0:
        return ZERO_AMOUNT
    years = contract_years(start_date, end_date)
    if not years:
        return ""
    return str((total / years).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def load_contracts(complete_dir: Path) -> list[dict[str, str]]:
    """加载合同行，并计算年租金供资源匹配兜底使用。"""
    rows = read_sheet_rows(complete_dir / SOURCE_FILES["contracts"], SOURCE_SHEETS["contracts"])
    return [
        {
            "organization_name": row.get("B", ""),
            "contract_code": row.get("C", ""),
            "contract_name": row.get("D", ""),
            "contract_type": row.get("F", ""),
            "party_a": row.get("G", ""),
            "tenant": row.get("H", ""),
            "signed_at": row.get("I", ""),
            "start_date": row.get("J", ""),
            "end_date": row.get("K", ""),
            "period": row.get("L", ""),
            "quantity": row.get("M", ""),
            "unit_price": row.get("N", ""),
            "total_amount": row.get("O", ""),
            "unit": row.get("P", ""),
            "transaction_method": row.get("Q", ""),
            "payment_method": row.get("R", ""),
            "debit_subject": row.get("S", ""),
            "credit_subject": row.get("T", ""),
            "bidding": row.get("U", ""),
            "asset_or_resource_name": row.get("V", ""),
            "annual_rent": annual_rent(row.get("O", ""), row.get("J", ""), row.get("K", "")),
        }
        for row in rows if row["row_number"] >= 2 and row.get("C")
    ]


def fill_master_sheets(
    output_path: Path,
    context: Context,
    users: list[dict[str, str]],
    diagnostics: BuildDiagnostics | None = None,
) -> None:
    """根据共享上下文填充组织、用户、账套工作表。"""
    reviewer = next((user for user in users if user["role"] == ROLE_NAME_MAP[REVIEWER_SOURCE_ROLE]), {"name": context.contact_name, "phone": context.contact_phone})
    write_single_row(
        output_path,
        TARGET_SHEETS["organization"],
        TARGET_LAYOUTS["organization"].start_row,
        {
            "A": context.county,
            "B": context.township,
            "C": context.village,
            "D": context.social_code,
            "E": context.social_code,
            "F": YES_VALUE,
            "G": reviewer["name"],
            "H": reviewer["phone"],
        },
        missing_cell_callback=lambda sheet, row, col, value: report_missing_fill(diagnostics, sheet, row, col, value),
    )
    write_layout_rows(output_path, "users", [
        {
            "A": context.county,
            "B": context.township,
            "C": context.village,
            "D": user["login"],
            "E": user["name"],
            "F": context.social_code,
            "G": user["phone"],
            "H": user["role"],
        }
        for user in users
    ], diagnostics)
    # F4 是模板原生的 Excel 数值日期单元格。不要把它改写成字符串，
    # 否则 Excel 会以不同方式显示原始序列值。
    write_single_row(
        output_path,
        TARGET_SHEETS["account_book"],
        TARGET_LAYOUTS["account_book"].start_row,
        {
            "A": context.county,
            "B": context.township,
            "C": context.village,
            "D": context.organization_name,
            "E": context.social_code,
            "G": ACCOUNTING_SYSTEM_NAME,
        },
        missing_cell_callback=lambda sheet, row, col, value: report_missing_fill(diagnostics, sheet, row, col, value),
    )


def fill_subject_sheet(output_path: Path, context: Context, subjects: list[dict[str, str]], diagnostics: BuildDiagnostics | None = None) -> None:
    """使用源数据提供的方向和辅助核算标志填充科目表。"""
    rows = []
    for row in subjects:
        rows.append({
            "A": context.county,
            "B": context.township,
            "C": context.organization_name,
            "D": context.social_code,
            "E": row["subject_code"],
            "F": row["subject_name"],
            "G": row["subject_type"],
            "H": row["subject_level"],
            "I": row["subject_direction"],
            "J": row["opening_balance"],
            "K": row["balance_direction"],
            "L": row["aux_enabled"],
        })
    write_layout_rows(output_path, "subjects", rows, diagnostics)


def fill_balance_sheet(output_path: Path, context: Context, balances: list[dict[str, str]], diagnostics: BuildDiagnostics | None = None) -> None:
    """填充科目余额表，并带上源数据底部的合计行。"""
    rows = []
    for row in balances:
        # 源合计行的科目代码 (A) 为空，科目名称 (B) 为合计。
        is_total = row.get("B", "") == TOTAL_LABEL and not row.get("A", "")
        rows.append({
            "A": "" if is_total else context.county,
            "B": "" if is_total else context.township,
            "C": TOTAL_LABEL if is_total else context.village,
            "D": "" if is_total else context.social_code,
            "E": row.get("A", ""),
            "F": row.get("B", ""),
            "G": row.get("C", ZERO_AMOUNT),
            "H": row.get("D", ZERO_AMOUNT),
            "I": row.get("E", ZERO_AMOUNT),
            "J": row.get("F", ZERO_AMOUNT),
            "K": row.get("I", ZERO_AMOUNT),
            "L": row.get("J", ZERO_AMOUNT),
            "M": row.get("G", ZERO_AMOUNT),
            "N": row.get("H", ZERO_AMOUNT),
        })
    write_layout_rows(output_path, "balance", rows, diagnostics)


def fill_asset_summary_sheet(output_path: Path, context: Context, assets: list[dict[str, str]], diagnostics: BuildDiagnostics | None = None) -> None:
    """填充固定资产汇总行，同时保留两行表头。

    ``ASSET_SUMMARY_UNSUPPORTED_COLUMNS`` 中的列存在于模板中，但当前数据集中
    没有可信来源字段，因此保持为空。
    """
    rows = []
    for row in assets:
        rows.append({
            "A": context.social_code,
            "B": context.county,
            "C": context.township,
            "D": context.organization_name,
            "E": row["asset_code"],
            "F": row["asset_category"],
            "G": row["asset_name"],
            "H": row["asset_attribute"],
            "I": row["specification"],
            "J": row["keeper"],
            "K": row["location"],
            "L": row["built_at"],
            "M": row["poverty_asset"],
            "N": row["usage_status"],
            "O": "",
            "P": "",
            "Q": "",
            "R": "",
            "S": "",
            "T": "",
            "U": row["lease_target"],
            "V": row["lease_term_months"],
            "W": row["annual_rent"],
            "X": row["quantity_or_area"],
            "Y": "",
            "Z": row["original_value"],
            "AA": row["net_value"],
            "AB": row["cleanup_value"],
        })
    write_layout_rows(output_path, "asset_summary", rows, diagnostics)


def fill_well_sheet(output_path: Path, context: Context, wells: list[dict[str, str]], diagnostics: BuildDiagnostics | None = None) -> None:
    """填充水井行；数据从三行表头下方开始。"""
    rows = []
    highlight_cells: dict[str, str] = {}
    for index, row in enumerate(wells):
        row_number = TARGET_LAYOUTS["wells"].start_row + index
        rows.append({
            "A": context.social_code,
            "B": context.county,
            "C": context.township,
            "D": context.organization_name,
            "E": row["asset_code"],
            "F": row["asset_name"],
            "G": row["well_code"],
            "H": row["is_high_standard"],
            "I": row["has_water"],
            "J": row["has_electricity"],
            "K": row["has_well_house"],
            "L": row["has_pump"],
            "M": row["has_caretaker"],
            "N": row["has_funding"],
            "O": row["has_supervision"],
            "P": row["well_head_name"],
            "Q": row["well_head_phone"],
            "R": row["caretaker_name"],
            "S": row["caretaker_phone"],
            "T": row["repairer_name"],
            "U": row["repairer_phone"],
        })
        if row["cleanup_value"] not in {"", "0", ZERO_AMOUNT}:
            highlight_cells[f"E{row_number}"] = HIGHLIGHT_FONT_RGB
    write_layout_rows(output_path, "wells", rows, diagnostics)
    apply_text_color_to_cells(output_path, TARGET_SHEETS["wells"], highlight_cells)


def fill_intangible_sheet(output_path: Path, diagnostics: BuildDiagnostics | None = None) -> None:
    """用空白模板行清理并填充无形资产工作表。

    当前尚未接入专门的无形资产源工作簿。目标工作表仍然需要主体行，因此用
    A-T 列范围内的 18 条空白行替换模板主体。
    """
    blank_rows = [{col: "" for col in INTANGIBLE_COLUMNS} for _ in range(INTANGIBLE_BLANK_ROWS)]
    write_layout_rows(output_path, "intangibles", blank_rows, diagnostics)


def fill_contract_sheet(output_path: Path, context: Context, contracts: list[dict[str, str]], diagnostics: BuildDiagnostics | None = None) -> None:
    """直接使用合同明细数据源列填充合同行。"""
    rows = []
    for contract in contracts:
        rows.append({
            "A": context.social_code,
            "B": context.county,
            "C": context.township,
            "D": context.organization_name,
            "E": contract["contract_code"],
            "F": contract["contract_name"],
            "G": contract["contract_type"],
            "H": contract["party_a"],
            "I": contract["tenant"],
            "J": contract["signed_at"],
            "K": contract["start_date"],
            "L": contract["end_date"],
            "M": contract["quantity"],
            "N": contract["unit"],
            "O": contract["unit_price"],
            "P": contract["total_amount"],
            "Q": contract["transaction_method"],
            "R": contract["bidding"],
            "S": contract["payment_method"],
            "T": contract["debit_subject"],
            "U": contract["credit_subject"],
            "V": contract["asset_or_resource_name"],
        })
    write_layout_rows(output_path, "contracts", rows, diagnostics)


def fill_resource_sheet(
    output_path: Path,
    context: Context,
    resources: list[dict[str, str]],
    diagnostics: BuildDiagnostics | None = None,
) -> None:
    """填充资源行；出租字段只使用资源明细数据中的直接值。

    即使资源状态为 ``出租经营``，也不再尝试用合同明细数据推导承租人、期间、
    年租金。资源源表没有提供这些字段时，目标列按业务要求保持为空。
    """
    rows = []
    for resource in resources:
        rows.append({
            "A": context.social_code,
            "B": context.county,
            "C": context.township,
            "D": context.organization_name,
            "E": resource["code"],
            "F": resource["name"],
            "G": resource["category"],
            "H": resource["attribute"],
            "I": resource["usage_status"],
            "J": resource["direct_tenant"],
            "K": normalize_resource_period(resource["direct_period"]),
            "L": resource["direct_annual_rent"],
            "M": resource["display_area"],
            "N": resource["unit"],
            "O": resource["owner"],
            "P": resource["location"],
            "Q": resource["registered_at"],
        })
    write_layout_rows(output_path, "resources", rows, diagnostics)


def build_workbook(
    complete_dir: Path = COMPLETE_DIR,
    output_path: Path | None = None,
    *,
    template_path: Path | None = None,
    diagnostics: BuildDiagnostics | None = None,
) -> Path:
    """根据模板和源工作簿构建 ``清查表（已填充）.xlsx``。

    CLI 兼容模式下，模板仍来自 ``complete_dir / 清查表.xlsx``。Web 上传模式下，
    ``complete_dir`` 是只包含源工作簿的临时目录，模板由 ``template_path`` 指向
    服务端固定模板。
    """
    diagnostics = diagnostics or BuildDiagnostics.create()
    output_path = output_path or (complete_dir / OUTPUT_NAME)
    template_path = template_path or (complete_dir / TEMPLATE_NAME)

    validate_source_files(complete_dir, diagnostics)
    if not template_path.is_file():
        diagnostics.warning(f"服务端模板不存在：{template_path}")
        raise FileNotFoundError(template_path)

    shutil.copy2(template_path, output_path)
    diagnostics.info(f"已复制模板：{template_path}")

    context = parse_context(complete_dir)
    users = load_users(complete_dir)
    subjects = load_subjects(complete_dir)
    balances = load_balances(complete_dir, diagnostics)
    assets = load_assets(complete_dir)
    wells = load_wells(complete_dir)
    resources = load_resources(complete_dir)
    contracts = load_contracts(complete_dir)

    fill_master_sheets(output_path, context, users, diagnostics)
    fill_subject_sheet(output_path, context, subjects, diagnostics)
    fill_balance_sheet(output_path, context, balances, diagnostics)
    fill_asset_summary_sheet(output_path, context, assets, diagnostics)
    fill_well_sheet(output_path, context, wells, diagnostics)
    fill_intangible_sheet(output_path, diagnostics)
    fill_contract_sheet(output_path, context, contracts, diagnostics)
    fill_resource_sheet(output_path, context, resources, diagnostics)

    diagnostics.info(f"已生成工作簿：{output_path}")

    print(f"Wrote {output_path}")
    return output_path


def main() -> None:
    build_workbook()


if __name__ == "__main__":
    main()
