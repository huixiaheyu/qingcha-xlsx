from __future__ import annotations

"""清查表构建脚本使用的轻量级 XLSX XML 辅助函数。

项目运行环境不能依赖已安装 Excel 库，因此本模块把 ``.xlsx`` 文件当作包含
OOXML 文件的 ZIP 包进行编辑。它只实现当前模板所需的行为：
- 将工作表行读取为以 Excel 列字母为键的字典；
- 解析共享字符串和内联字符串；
- 克隆现有模板行，以保留样式和合并表头结构；
- 将内联字符串值写入现有单元格。

重要限制：写入函数不会补造缺失的单元格。如果克隆出的行本身不包含目标列，
该列的值会被跳过。因此，工作簿构建逻辑依赖模板行已经包含计划填充的列。
"""

from copy import deepcopy
from pathlib import Path
from typing import Callable
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

# OOXML 协议常量，不是业务配置。
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
PKGREL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOCREL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"main": MAIN_NS, "pkgrel": PKGREL_NS}

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", DOCREL_NS)


def parse_shared_strings(zf: ZipFile) -> list[str]:
    """返回工作簿的共享字符串表；如果不存在则返回空列表。"""
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return [
        "".join(t.text or "" for t in si.iter(f"{{{MAIN_NS}}}t"))
        for si in root.findall("main:si", NS)
    ]


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    """读取这些工作簿中所用单元格类型的显示文本或值。"""
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


def get_workbook_rels(zf: ZipFile) -> dict[str, str]:
    """将工作簿关系 ID 映射到目标 XML 路径。"""
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    return {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall("pkgrel:Relationship", NS)
    }


def get_sheet_path(zf: ZipFile, sheet_name: str) -> str:
    """将精确的工作簿工作表名称解析为对应的 ``xl/worksheets/...`` 路径。"""
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    relmap = get_workbook_rels(zf)
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        if sheet.attrib["name"] == sheet_name:
            rid = sheet.attrib[f"{{{DOCREL_NS}}}id"]
            return "xl/" + relmap[rid]
    raise KeyError(sheet_name)


def list_sheet_names(path: Path) -> list[str]:
    """返回工作簿中所有工作表名称，保留名称中的真实空格。"""
    with ZipFile(path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        return [sheet.attrib["name"] for sheet in workbook.findall("main:sheets/main:sheet", NS)]


def read_sheet_rows(path: Path, sheet_name: str) -> list[dict[str, str]]:
    """将工作表读取为以 Excel 列字母为键的行字典。"""
    with ZipFile(path) as zf:
        shared_strings = parse_shared_strings(zf)
        sheet_root = ET.fromstring(zf.read(get_sheet_path(zf, sheet_name)))
        rows: list[dict[str, str]] = []
        for row in sheet_root.findall('.//main:sheetData/main:row', NS):
            item: dict[str, str] = {"row_number": int(row.attrib["r"])}
            for cell in row.findall('main:c', NS):
                col = ''.join(ch for ch in cell.attrib['r'] if ch.isalpha())
                item[col] = cell_value(cell, shared_strings)
            rows.append(item)
        return rows


def clear_cell(cell: ET.Element) -> None:
    """清除单元格内容和字符串类型，同时保留样式与引用属性。"""
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)


def set_cell_text(cell: ET.Element, text: str) -> None:
    """将单元格设为内联文本；若文本为 ``""``，清空后保持为空。"""
    clear_cell(cell)
    if text == "":
        return
    cell.attrib["t"] = "inlineStr"
    is_node = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
    t_node = ET.SubElement(is_node, f"{{{MAIN_NS}}}t")
    if text.startswith(" ") or text.endswith(" "):
        t_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t_node.text = text


def clone_row(template_row: ET.Element, row_number: int) -> ET.Element:
    """克隆模板行，并把每个单元格引用改写到 ``row_number``。"""
    new_row = deepcopy(template_row)
    new_row.attrib["r"] = str(row_number)
    for cell in new_row.findall("main:c", NS):
        col = ''.join(ch for ch in cell.attrib['r'] if ch.isalpha())
        cell.attrib['r'] = f"{col}{row_number}"
    return new_row


def load_workbook_package(path: Path) -> tuple[list, dict[str, bytes]]:
    """加载所有 ZIP 成员，以便替换一个工作表 XML 后重新保存。"""
    with ZipFile(path, 'r') as zin:
        infos = zin.infolist()
        contents = {info.filename: zin.read(info.filename) for info in infos}
    return infos, contents


def save_workbook_package(path: Path, infos: list, contents: dict[str, bytes]) -> None:
    """按原成员顺序和元数据写回工作簿包。"""
    with ZipFile(path, 'w', compression=ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info, contents[info.filename])


def append_unique_xml_child(parent: ET.Element, child: ET.Element) -> int:
    """追加语义等价的 XML 子节点前先去重，并返回其索引。"""
    serialized = ET.tostring(child, encoding='unicode')
    for index, existing in enumerate(list(parent)):
        if ET.tostring(existing, encoding='unicode') == serialized:
            return index
    parent.append(child)
    parent.attrib['count'] = str(len(parent))
    return len(parent) - 1


def ensure_cell_font_color_style(styles_root: ET.Element, base_style_id: int, rgb: str) -> int:
    """基于现有单元格样式克隆出指定字体颜色的新样式。"""
    fonts = styles_root.find('main:fonts', NS)
    cell_xfs = styles_root.find('main:cellXfs', NS)
    assert fonts is not None
    assert cell_xfs is not None

    base_xf = list(cell_xfs)[base_style_id]
    base_font_id = int(base_xf.attrib.get('fontId', '0'))
    colored_font = deepcopy(list(fonts)[base_font_id])
    for color in list(colored_font.findall('main:color', NS)):
        colored_font.remove(color)
    ET.SubElement(colored_font, f'{{{MAIN_NS}}}color', {'rgb': rgb})
    font_id = append_unique_xml_child(fonts, colored_font)

    colored_xf = deepcopy(base_xf)
    colored_xf.attrib['fontId'] = str(font_id)
    colored_xf.attrib['applyFont'] = '1'
    return append_unique_xml_child(cell_xfs, colored_xf)


def apply_text_color_to_cells(path: Path, sheet_name: str, color_by_cell_ref: dict[str, str]) -> None:
    """按单元格引用批量设置字体颜色，同时保留原样式其余部分。"""
    if not color_by_cell_ref:
        return

    infos, contents = load_workbook_package(path)
    styles_root = ET.fromstring(contents['xl/styles.xml'])
    with ZipFile(path, 'r') as zf:
        sheet_path = get_sheet_path(zf, sheet_name)
    sheet_root = ET.fromstring(contents[sheet_path])

    style_cache: dict[tuple[int, str], int] = {}
    for cell in sheet_root.findall('.//main:sheetData/main:row/main:c', NS):
        cell_ref = cell.attrib.get('r', '')
        rgb = color_by_cell_ref.get(cell_ref)
        if rgb is None:
            continue
        base_style_id = int(cell.attrib.get('s', '0'))
        cache_key = (base_style_id, rgb)
        style_id = style_cache.get(cache_key)
        if style_id is None:
            style_id = ensure_cell_font_color_style(styles_root, base_style_id, rgb)
            style_cache[cache_key] = style_id
        cell.attrib['s'] = str(style_id)

    contents['xl/styles.xml'] = ET.tostring(styles_root, encoding='utf-8', xml_declaration=True)
    contents[sheet_path] = ET.tostring(sheet_root, encoding='utf-8', xml_declaration=True)
    save_workbook_package(path, infos, contents)


def write_table_rows(
    path: Path,
    sheet_name: str,
    start_row: int,
    rows_data: list[dict[str, str]],
    *,
    template_row_number: int | None = None,
    clear_existing_from_row: int | None = None,
    missing_cell_callback: Callable[[str, int, str, str], None] | None = None,
) -> None:
    """用从模板行克隆出的行替换表格主体。

    模板行选择遵循“显式优先”：如果提供了 ``template_row_number`` 就使用它，
    否则使用 ``start_row``，再否则使用 ``start_row`` 之前的最后一行。
    现有行会从 ``clear_existing_from_row`` 或 ``start_row`` 开始删除。

    写入值之前会清空每个克隆出的单元格，避免旧表头或示例文本漏入新数据行。
    模板行中缺失列对应的值会被忽略，因为创建新的带样式单元格超出了此辅助
    模块当前的范围。
    """
    infos, contents = load_workbook_package(path)
    with ZipFile(path, 'r') as zf:
        sheet_path = get_sheet_path(zf, sheet_name)
    sheet_root = ET.fromstring(contents[sheet_path])
    sheet_data = sheet_root.find('main:sheetData', NS)
    assert sheet_data is not None

    all_rows = sheet_data.findall('main:row', NS)
    template_row = None
    if template_row_number is not None:
        template_row = next((row for row in all_rows if int(row.attrib['r']) == template_row_number), None)
    if template_row is None:
        template_row = next((row for row in all_rows if int(row.attrib['r']) == start_row), None)
    if template_row is None:
        template_row = next((row for row in reversed(all_rows) if int(row.attrib['r']) < start_row), None)
    if template_row is None:
        raise ValueError(f"Missing usable template row for {sheet_name} starting at row {start_row}")

    remove_from = clear_existing_from_row if clear_existing_from_row is not None else start_row
    for row in list(all_rows):
        if int(row.attrib['r']) >= remove_from:
            sheet_data.remove(row)

    max_col_idx = 1
    for offset, row_data in enumerate(rows_data):
        row_number = start_row + offset
        row = clone_row(template_row, row_number)
        cell_map: dict[str, ET.Element] = {}
        for cell in row.findall('main:c', NS):
            col = ''.join(ch for ch in cell.attrib['r'] if ch.isalpha())
            cell_map[col] = cell
            clear_cell(cell)
            max_col_idx = max(max_col_idx, col_to_index(col))
        for col, value in row_data.items():
            if col not in cell_map:
                if value != "" and missing_cell_callback is not None:
                    missing_cell_callback(sheet_name, row_number, col, str(value))
                continue
            set_cell_text(cell_map[col], str(value))
            max_col_idx = max(max_col_idx, col_to_index(col))
        sheet_data.append(row)

    dimension = sheet_root.find('main:dimension', NS)
    if dimension is not None:
        end_row = max(start_row - 1 + len(rows_data), remove_from - 1)
        end_col = index_to_col(max_col_idx)
        dimension.attrib['ref'] = f"A1:{end_col}{end_row}"

    contents[sheet_path] = ET.tostring(sheet_root, encoding='utf-8', xml_declaration=True)
    save_workbook_package(path, infos, contents)


def write_single_row(
    path: Path,
    sheet_name: str,
    row_number: int,
    row_data: dict[str, str],
    *,
    missing_cell_callback: Callable[[str, int, str, str], None] | None = None,
) -> None:
    """更新现有行中的指定单元格，同时保留未触及的单元格。

    该函数用于主数据工作表：模板中可能包含原生 Excel 日期或数字单元格，
    这些单元格不能被转换为内联文本。与 ``write_table_rows`` 相同，缺失的
    目标列会因模板单元格限制而被忽略。
    """
    infos, contents = load_workbook_package(path)
    with ZipFile(path, 'r') as zf:
        sheet_path = get_sheet_path(zf, sheet_name)
    sheet_root = ET.fromstring(contents[sheet_path])
    sheet_data = sheet_root.find('main:sheetData', NS)
    assert sheet_data is not None

    row = next((item for item in sheet_data.findall('main:row', NS) if int(item.attrib['r']) == row_number), None)
    if row is None:
        raise ValueError(f"Missing row {row_number} in {sheet_name}")

    cell_map: dict[str, ET.Element] = {}
    for cell in row.findall('main:c', NS):
        col = ''.join(ch for ch in cell.attrib['r'] if ch.isalpha())
        cell_map[col] = cell

    for col, value in row_data.items():
        if col not in cell_map:
            if value != "" and missing_cell_callback is not None:
                missing_cell_callback(sheet_name, row_number, col, str(value))
            continue
        set_cell_text(cell_map[col], str(value))

    contents[sheet_path] = ET.tostring(sheet_root, encoding='utf-8', xml_declaration=True)
    save_workbook_package(path, infos, contents)


def col_to_index(col: str) -> int:
    """将 ``AA`` 等 Excel 列字母转换为从 1 开始的编号。"""
    value = 0
    for ch in col:
        value = value * 26 + (ord(ch.upper()) - 64)
    return value


def index_to_col(index: int) -> str:
    """将从 1 开始的列编号转换回 Excel 列字母。"""
    out = []
    while index > 0:
        index, rem = divmod(index - 1, 26)
        out.append(chr(65 + rem))
    return ''.join(reversed(out))
