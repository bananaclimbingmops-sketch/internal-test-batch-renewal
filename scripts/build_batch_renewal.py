#!/usr/bin/env python3
"""Portable batch-renewal workbook builder using only Python's standard library."""

import argparse
import json
import os
import posixpath
import re
import sys
import tempfile
import zipfile
from collections import OrderedDict
from pathlib import Path
from xml.etree import ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"m": MAIN_NS, "r": DOC_REL_NS, "pr": PKG_REL_NS}
EXPECTED_HEADERS = ["用户卡ID", "成员姓名", "手机号码", "卡项类别", "新增天数", "新增次数", "续卡原因"]
DEFAULT_DURATION_CARDS = ["新人月卡", "月卡", "季卡", "年卡"]
DEFAULT_COUNT_CARDS = ["10次卡", "20次卡"]
PLACEHOLDERS = {"-", "—", "－"}
FORMULA_ERROR = re.compile(r"^#(?:REF|DIV/0|VALUE|NAME\?|N/A|NUM|NULL|SPILL|CALC)!?$")

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", DOC_REL_NS)


def normalize(value):
    return str(value if value is not None else "").strip()


def is_placeholder(value):
    return normalize(value) in PLACEHOLDERS


def column_index(reference):
    match = re.match(r"^([A-Za-z]+)", reference or "")
    if not match:
        return 0
    result = 0
    for character in match.group(1).upper():
        result = result * 26 + ord(character) - 64
    return result - 1


def column_name(index):
    value = index + 1
    result = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(65 + remainder) + result
    return result


def read_shared_strings(archive):
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings = []
    for item in root.findall("m:si", NS):
        strings.append("".join(node.text or "" for node in item.findall(".//m:t", NS)))
    return strings


def read_cell(cell, shared_strings):
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//m:t", NS))
    value_node = cell.find("m:v", NS)
    raw = "" if value_node is None or value_node.text is None else value_node.text
    if cell_type == "s":
        try:
            return shared_strings[int(raw)]
        except (ValueError, IndexError):
            return raw
    if cell_type == "b":
        return "TRUE" if raw == "1" else "FALSE"
    return raw


def workbook_sheets(archive):
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        relation.get("Id"): relation.get("Target")
        for relation in relationships.findall("pr:Relationship", NS)
    }
    sheets = []
    for sheet in workbook.findall("m:sheets/m:sheet", NS):
        relationship_id = sheet.get("{%s}id" % DOC_REL_NS)
        target = targets.get(relationship_id)
        if not target:
            continue
        normalized_target = target.lstrip("/")
        sheet_path = posixpath.normpath(
            normalized_target if normalized_target.startswith("xl/")
            else posixpath.join("xl", normalized_target)
        )
        sheets.append({"name": sheet.get("name", ""), "path": sheet_path})
    return sheets


def read_sheet_rows(archive, sheet_path, shared_strings):
    root = ET.fromstring(archive.read(sheet_path))
    rows = []
    for row in root.findall("m:sheetData/m:row", NS):
        row_number = int(row.get("r") or len(rows) + 1)
        cells = {}
        max_column = -1
        for cell in row.findall("m:c", NS):
            index = column_index(cell.get("r"))
            max_column = max(max_column, index)
            cells[index] = read_cell(cell, shared_strings)
        values = [cells.get(index, "") for index in range(max_column + 1)] if max_column >= 0 else []
        rows.append((row_number, values))
    return rows


def load_workbook_data(path):
    with zipfile.ZipFile(path, "r") as archive:
        shared_strings = read_shared_strings(archive)
        sheets = workbook_sheets(archive)
        return [
            {
                "name": sheet["name"],
                "path": sheet["path"],
                "rows": read_sheet_rows(archive, sheet["path"], shared_strings),
            }
            for sheet in sheets
        ]


def find_header_index(header, aliases):
    for alias in aliases:
        try:
            return header.index(alias)
        except ValueError:
            pass
    return -1


def find_source_data(path):
    for sheet in load_workbook_data(path):
        nonempty_rows = [(number, values) for number, values in sheet["rows"] if any(normalize(value) for value in values)]
        if not nonempty_rows:
            continue
        header_row, header_values = nonempty_rows[0]
        header = [normalize(value) for value in header_values]
        indexes = {
            "id": find_header_index(header, ["用户卡ID"]),
            "name": find_header_index(header, ["成员姓名"]),
            "phone": find_header_index(header, ["手机号码"]),
            "card": find_header_index(header, ["签到卡", "签到卡类"]),
        }
        if all(index >= 0 for index in indexes.values()):
            data_rows = [(number, values) for number, values in sheet["rows"] if number > header_row]
            return {"name": sheet["name"], "path": sheet["path"], "header": header, "indexes": indexes, "rows": data_rows}
    raise ValueError("没有找到同时包含“用户卡ID、成员姓名、手机号码、签到卡/签到卡类”的工作表")


def row_value(values, index):
    return normalize(values[index] if index < len(values) else "")


def make_entry(source_row, source_order, values, indexes):
    return {
        "sourceRow": source_row,
        "sourceOrder": source_order,
        "id": row_value(values, indexes["id"]),
        "name": row_value(values, indexes["name"]),
        "phone": row_value(values, indexes["phone"]),
        "card": row_value(values, indexes["card"]),
    }


def same_entry(left, right):
    return left["card"] == right["card"] and left["name"] == right["name"] and left["phone"] == right["phone"]


def inspect_source(source_data):
    card_stats = OrderedDict()
    seen = {}
    missing_fields = []
    conflicts = []
    blank_rows_skipped = 0
    placeholder_rows_skipped = 0
    duplicates_skipped = 0

    for source_order, (source_row, values) in enumerate(source_data["rows"]):
        entry = make_entry(source_row, source_order, values, source_data["indexes"])
        if not any(entry[key] for key in ("id", "name", "phone", "card")):
            blank_rows_skipped += 1
            continue
        if is_placeholder(entry["card"]):
            placeholder_rows_skipped += 1
            continue
        missing = []
        if not entry["id"] or is_placeholder(entry["id"]):
            missing.append("用户卡ID")
        if not entry["name"] or is_placeholder(entry["name"]):
            missing.append("成员姓名")
        if not entry["phone"] or is_placeholder(entry["phone"]):
            missing.append("手机号码")
        if not entry["card"]:
            missing.append("签到卡/签到卡类")
        if missing:
            missing_fields.append({"row": source_row, "missing": missing})
        if not entry["card"]:
            continue

        stats = card_stats.setdefault(entry["card"], {
            "card": entry["card"], "firstSourceRow": source_row, "occurrences": 0, "uniqueIds": set()
        })
        stats["occurrences"] += 1
        if entry["id"] and not is_placeholder(entry["id"]):
            stats["uniqueIds"].add(entry["id"])

        if not entry["id"] or is_placeholder(entry["id"]):
            continue
        if entry["id"] in seen:
            prior = seen[entry["id"]]
            if same_entry(prior, entry):
                duplicates_skipped += 1
            else:
                conflicts.append({"id": entry["id"], "first": prior, "duplicate": entry})
            continue
        seen[entry["id"]] = entry

    return {
        "mode": "inspect",
        "engine": "python-stdlib",
        "sourceSheet": source_data["name"],
        "sourceRows": len(source_data["rows"]),
        "nonblankUniqueIds": len(seen),
        "blankRowsSkipped": blank_rows_skipped,
        "placeholderRowsSkipped": placeholder_rows_skipped,
        "duplicatesSkipped": duplicates_skipped,
        "cardItems": [
            {
                "card": item["card"],
                "occurrences": item["occurrences"],
                "uniqueUserCardIds": len(item["uniqueIds"]),
                "firstSourceRow": item["firstSourceRow"],
            }
            for item in card_stats.values()
        ],
        "issues": {
            "missingFieldRows": len(missing_fields),
            "missingFields": missing_fields[:20],
            "conflictingIds": len(conflicts),
            "conflicts": [
                {
                    "id": item["id"],
                    "firstRow": item["first"]["sourceRow"],
                    "duplicateRow": item["duplicate"]["sourceRow"],
                    "firstCard": item["first"]["card"],
                    "duplicateCard": item["duplicate"]["card"],
                    "firstName": item["first"]["name"],
                    "duplicateName": item["duplicate"]["name"],
                    "firstPhone": item["first"]["phone"],
                    "duplicatePhone": item["duplicate"]["phone"],
                }
                for item in conflicts[:20]
            ],
        },
    }


def parse_card_list(raw, fallback):
    if raw is None:
        return list(fallback)
    cards = [item.strip() for item in raw.split(",") if item.strip()]
    if not cards:
        raise ValueError("卡项列表不能为空")
    if len(set(cards)) != len(cards):
        raise ValueError("卡项列表包含重复值：%s" % "、".join(cards))
    return cards


def validate_rule(raw_rule, index):
    if not isinstance(raw_rule, dict):
        raise ValueError("规则第 %d 项必须是对象" % (index + 1))
    card = normalize(raw_rule.get("card"))
    days = raw_rule.get("days")
    times = raw_rule.get("times")
    if not card:
        raise ValueError("规则第 %d 项缺少 card" % (index + 1))
    if isinstance(days, bool) or not isinstance(days, int) or days < -1:
        raise ValueError("卡项“%s”的 days 必须是大于等于 -1 的整数" % card)
    if isinstance(times, bool) or not isinstance(times, int) or times < 0:
        raise ValueError("卡项“%s”的 times 必须是大于等于 0 的整数" % card)
    if days == 0 and times == 0:
        raise ValueError("卡项“%s”的 days 和 times 不能同时为 0" % card)
    return {"card": card, "days": days, "times": times}


def load_rules(args):
    if args.rules_json:
        if args.duration_cards is not None or args.count_cards is not None:
            raise ValueError("--rules-json 不能与 --duration-cards 或 --count-cards 同时使用")
        with open(args.rules_json, "r", encoding="utf-8") as handle:
            parsed = json.load(handle)
        raw_rules = parsed if isinstance(parsed, list) else parsed.get("rules") if isinstance(parsed, dict) else None
        if not isinstance(raw_rules, list) or not raw_rules:
            raise ValueError("规则文件必须是非空数组，或包含非空 rules 数组")
        rules = [validate_rule(rule, index) for index, rule in enumerate(raw_rules)]
        if len({rule["card"] for rule in rules}) != len(rules):
            raise ValueError("规则文件包含重复卡项")
        return rules, "custom"

    duration_cards = parse_card_list(args.duration_cards, DEFAULT_DURATION_CARDS)
    count_cards = parse_card_list(args.count_cards, DEFAULT_COUNT_CARDS)
    overlap = [card for card in duration_cards if card in set(count_cards)]
    if overlap:
        raise ValueError("期限卡与次卡列表重叠：%s" % "、".join(overlap))
    rules = ([{"card": card, "days": 1, "times": 0} for card in duration_cards]
             + [{"card": card, "days": -1, "times": 1} for card in count_cards])
    return rules, "legacy"


def collect_rows(source_data, rules):
    rules_by_card = {rule["card"]: rule for rule in rules}
    category_order = {rule["card"]: index for index, rule in enumerate(rules)}
    seen = {}
    selected = []
    conflicts = []
    duplicates_skipped = 0

    for source_order, (source_row, values) in enumerate(source_data["rows"]):
        entry = make_entry(source_row, source_order, values, source_data["indexes"])
        rule = rules_by_card.get(entry["card"])
        if not rule:
            continue
        entry["rule"] = rule
        for field, label in (("id", "用户卡ID"), ("name", "成员姓名"), ("phone", "手机号码")):
            if not entry[field] or is_placeholder(entry[field]):
                raise ValueError("源表“%s”第 %d 行缺少%s" % (source_data["name"], source_row, label))
        if entry["id"] in seen:
            prior = seen[entry["id"]]
            if same_entry(prior, entry):
                duplicates_skipped += 1
            else:
                conflicts.append({"id": entry["id"], "first": prior, "duplicate": entry})
            continue
        seen[entry["id"]] = entry
        selected.append(entry)

    if conflicts:
        preview = [
            {
                "id": item["id"], "firstRow": item["first"]["sourceRow"],
                "duplicateRow": item["duplicate"]["sourceRow"],
                "firstCard": item["first"]["card"], "duplicateCard": item["duplicate"]["card"],
                "firstName": item["first"]["name"], "duplicateName": item["duplicate"]["name"],
                "firstPhone": item["first"]["phone"], "duplicatePhone": item["duplicate"]["phone"],
            }
            for item in conflicts[:10]
        ]
        raise ValueError("发现 %d 个用户卡ID冲突：%s" % (len(conflicts), json.dumps(preview, ensure_ascii=False)))
    if not selected:
        raise ValueError("没有任何记录精确命中用户确认的续卡卡项")
    matched_cards = {entry["card"] for entry in selected}
    unmatched = [rule["card"] for rule in rules if rule["card"] not in matched_cards]
    if unmatched:
        raise ValueError("用户确认的卡项在源表中没有匹配记录：%s" % "、".join(unmatched))
    selected.sort(key=lambda entry: (category_order[entry["card"]], entry["sourceOrder"]))
    return selected, duplicates_skipped


def clone_row_attributes(row, row_number):
    attributes = {}
    if row is not None:
        attributes.update({key: value for key, value in row.attrib.items() if key not in {"r", "spans"}})
    attributes["r"] = str(row_number)
    return attributes


def style_by_column(sheet_root):
    styles = {}
    for row in sheet_root.findall("m:sheetData/m:row", NS)[:2]:
        for cell in row.findall("m:c", NS):
            index = column_index(cell.get("r"))
            if index not in styles and cell.get("s") is not None:
                styles[index] = cell.get("s")
    return styles


def add_text_cell(row, address, value, style=None):
    attributes = {"r": address, "t": "inlineStr"}
    if style is not None:
        attributes["s"] = style
    cell = ET.SubElement(row, "{%s}c" % MAIN_NS, attributes)
    inline = ET.SubElement(cell, "{%s}is" % MAIN_NS)
    text_node = ET.SubElement(inline, "{%s}t" % MAIN_NS)
    text = str(value)
    if text[:1].isspace() or text[-1:].isspace():
        text_node.set("{%s}space" % XML_NS, "preserve")
    text_node.text = text


def add_number_cell(row, address, value, style=None):
    attributes = {"r": address, "t": "n"}
    if style is not None:
        attributes["s"] = style
    cell = ET.SubElement(row, "{%s}c" % MAIN_NS, attributes)
    ET.SubElement(cell, "{%s}v" % MAIN_NS).text = str(int(value))


def build_sheet_xml(original_xml, output_rows):
    root = ET.fromstring(original_xml)
    sheet_data = root.find("m:sheetData", NS)
    if sheet_data is None:
        raise ValueError("模板主工作表缺少 sheetData")
    original_rows = sheet_data.findall("m:row", NS)
    header_template = original_rows[0] if original_rows else None
    data_template = original_rows[1] if len(original_rows) > 1 else header_template
    styles = style_by_column(root)
    for row in list(sheet_data):
        sheet_data.remove(row)

    for row_index, values in enumerate(output_rows, start=1):
        row = ET.SubElement(
            sheet_data,
            "{%s}row" % MAIN_NS,
            clone_row_attributes(header_template if row_index == 1 else data_template, row_index),
        )
        for column, value in enumerate(values):
            address = "%s%d" % (column_name(column), row_index)
            if column in (4, 5) and row_index > 1:
                add_number_cell(row, address, value, styles.get(column))
            else:
                add_text_cell(row, address, value, styles.get(column))

    last_row = max(1, len(output_rows))
    dimension = root.find("m:dimension", NS)
    if dimension is None:
        dimension = ET.Element("{%s}dimension" % MAIN_NS)
        root.insert(0, dimension)
    dimension.set("ref", "A1:G%d" % last_row)

    columns = root.find("m:cols", NS)
    if columns is None:
        columns = ET.Element("{%s}cols" % MAIN_NS)
        root.insert(list(root).index(sheet_data), columns)
    else:
        for child in list(columns):
            columns.remove(child)
    widths = [40.8, 14.0, 16.0, 28.0, 12.0, 12.0, 32.0]
    for index, width in enumerate(widths, start=1):
        ET.SubElement(columns, "{%s}col" % MAIN_NS, {
            "min": str(index), "max": str(index), "width": str(width), "customWidth": "1"
        })

    ignored_errors = root.find("m:ignoredErrors", NS)
    if ignored_errors is None:
        ignored_errors = ET.SubElement(root, "{%s}ignoredErrors" % MAIN_NS)
    for child in list(ignored_errors):
        ignored_errors.remove(child)
    ET.SubElement(ignored_errors, "{%s}ignoredError" % MAIN_NS, {
        "numberStoredAsText": "1", "sqref": "A1:G%d" % last_row
    })
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def target_template_sheet(template_path):
    with zipfile.ZipFile(template_path, "r") as archive:
        sheets = workbook_sheets(archive)
    if not sheets:
        raise ValueError("模板没有可写入的工作表")
    for sheet in sheets:
        if sheet["name"] == "批量续卡模板":
            return sheet
    return sheets[0]


def write_output(template_path, output_path, output_rows):
    target = target_template_sheet(template_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = tempfile.NamedTemporaryFile(prefix="renewal-", suffix=".xlsx", dir=str(output_path.parent), delete=False)
    temporary_path = Path(temporary.name)
    temporary.close()
    try:
        with zipfile.ZipFile(template_path, "r") as source_archive:
            replacement = build_sheet_xml(source_archive.read(target["path"]), output_rows)
            with zipfile.ZipFile(temporary_path, "w") as output_archive:
                for info in source_archive.infolist():
                    data = replacement if info.filename == target["path"] else source_archive.read(info.filename)
                    output_archive.writestr(info, data)
        os.replace(str(temporary_path), str(output_path))
    finally:
        if temporary_path.exists():
            temporary_path.unlink()
    return target


def compare_other_parts(template_path, output_path, changed_path):
    with zipfile.ZipFile(template_path, "r") as template, zipfile.ZipFile(output_path, "r") as output:
        template_names = set(template.namelist())
        output_names = set(output.namelist())
        if template_names != output_names:
            return False
        return all(
            template.read(name) == output.read(name)
            for name in template_names
            if name != changed_path
        )


def verify_output(output_path, template_path, target, selected, rules, reason):
    sheets = load_workbook_data(output_path)
    output_sheet = next((sheet for sheet in sheets if sheet["name"] == target["name"]), None)
    if output_sheet is None or not output_sheet["rows"]:
        raise ValueError("导出后找不到主工作表")
    header = [normalize(value) for value in output_sheet["rows"][0][1][:7]]
    if header != EXPECTED_HEADERS:
        raise ValueError("导出表头不正确：%s" % json.dumps(header, ensure_ascii=False))
    rules_by_card = {rule["card"]: rule for rule in rules}
    data_rows = [values for _, values in output_sheet["rows"][1:] if row_value(values, 0)]
    if len(data_rows) != len(selected):
        raise ValueError("导出记录数不一致：预期 %d，实际 %d" % (len(selected), len(data_rows)))
    seen = set()
    invalid = []
    for index, row in enumerate(data_rows, start=2):
        values = [row_value(row, column) for column in range(7)]
        card = values[3]
        rule = rules_by_card.get(card)
        try:
            days = int(values[4])
            times = int(values[5])
        except ValueError:
            days = times = None
        duplicate = values[0] in seen
        seen.add(values[0])
        valid_rule = rule is not None and days == rule["days"] and times == rule["times"]
        if (not all(values[:4]) or not valid_rule or values[6] != reason or duplicate):
            invalid.append({"row": index, "values": values, "duplicate": duplicate})
        if any(FORMULA_ERROR.match(value) for value in values):
            invalid.append({"row": index, "formulaError": values})
    if invalid:
        raise ValueError("导出校验失败：%s" % json.dumps(invalid[:20], ensure_ascii=False))
    if not compare_other_parts(template_path, output_path, target["path"]):
        raise ValueError("模板主工作表以外的内容在导出后发生了变化")


def build_parser():
    parser = argparse.ArgumentParser(description="跨 Agent 的内测批量续卡 Excel 处理器")
    parser.add_argument("--mode", choices=("inspect", "build"), default="build")
    parser.add_argument("--source", required=True)
    parser.add_argument("--output")
    parser.add_argument("--reason")
    parser.add_argument("--rules-json")
    parser.add_argument("--duration-cards")
    parser.add_argument("--count-cards")
    parser.add_argument("--template", default=str(Path(__file__).resolve().parent.parent / "assets" / "batch-renewal-template.xlsx"))
    return parser


def main():
    args = build_parser().parse_args()
    source_path = Path(args.source).resolve()
    if source_path.suffix.lower() != ".xlsx":
        raise ValueError("原始签到记录必须是 .xlsx 文件")
    if not source_path.is_file():
        raise ValueError("找不到原始签到记录：%s" % source_path)
    source_data = find_source_data(source_path)
    if args.mode == "inspect":
        print(json.dumps(inspect_source(source_data), ensure_ascii=False, indent=2))
        return

    if not args.output:
        raise ValueError("build 模式缺少 --output")
    if not normalize(args.reason):
        raise ValueError("build 模式缺少 --reason")
    output_path = Path(args.output).resolve()
    if output_path.suffix.lower() != ".xlsx":
        raise ValueError("输出路径必须以 .xlsx 结尾")
    template_path = Path(args.template).resolve()
    if not template_path.is_file():
        raise ValueError("找不到续卡模板：%s" % template_path)

    rules, rules_mode = load_rules(args)
    selected, duplicates_skipped = collect_rows(source_data, rules)
    output_rows = [EXPECTED_HEADERS] + [
        [entry["id"], entry["name"], entry["phone"], entry["card"],
         entry["rule"]["days"], entry["rule"]["times"], normalize(args.reason)]
        for entry in selected
    ]
    target = write_output(template_path, output_path, output_rows)
    verify_output(output_path, template_path, target, selected, rules, normalize(args.reason))
    card_counts = OrderedDict((rule["card"], sum(entry["card"] == rule["card"] for entry in selected)) for rule in rules)
    summary = {
        "mode": "build",
        "engine": "python-stdlib",
        "outputPath": str(output_path),
        "sourceSheet": source_data["name"],
        "reason": normalize(args.reason),
        "outputRows": len(selected),
        "duplicatesSkipped": duplicates_skipped,
        "rulesMode": rules_mode,
        "rules": rules,
        "cardCounts": card_counts,
        "formulaErrors": [],
        "otherTemplatePartsPreserved": True,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    try:
        main()
    except (ValueError, OSError, zipfile.BadZipFile, json.JSONDecodeError) as error:
        print("错误：%s" % error, file=sys.stderr)
        sys.exit(1)
