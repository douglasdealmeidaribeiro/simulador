import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def col_to_num(col):
    value = 0
    for char in col:
        value = value * 26 + ord(char.upper()) - 64
    return value


def num_to_col(num):
    out = ""
    while num:
        num, rem = divmod(num - 1, 26)
        out = chr(65 + rem) + out
    return out


def split_cell(ref):
    match = re.match(r"^\$?([A-Z]{1,3})\$?(\d+)$", ref, re.I)
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")
    return col_to_num(match.group(1)), int(match.group(2))


def translate_a1_token(token, d_col, d_row):
    match = re.match(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)$", token, re.I)
    if not match:
        return token
    col_abs, col, row_abs, row = match.groups()
    col_num = col_to_num(col)
    row_num = int(row)
    if not col_abs:
        col_num += d_col
    if not row_abs:
        row_num += d_row
    return f"{col_abs}{num_to_col(max(1, col_num))}{row_abs}{max(1, row_num)}"


CELL_RE = re.compile(r"(?<![A-Za-z0-9_])(\$?[A-Z]{1,3}\$?\d+)(?![A-Za-z0-9_])")


def translate_formula(formula, master_ref, target_ref):
    master_col, master_row = split_cell(master_ref)
    target_col, target_row = split_cell(target_ref)
    d_col = target_col - master_col
    d_row = target_row - master_row

    parts = formula.split('"')
    for idx in range(0, len(parts), 2):
        parts[idx] = CELL_RE.sub(
            lambda match: translate_a1_token(match.group(1), d_col, d_row),
            parts[idx],
        )
    return '"'.join(parts)


def normalize_formula(formula, sheet_name):
    formula = formula.replace("_xlfn._xlws.", "")
    formula = formula.replace("_xlfn.", "")
    formula = re.sub(r"(?:'[^']+'|[A-Za-zÀ-ÿ0-9_]+)!#REF!", "#REF!", formula)
    formula = re.sub(r"IFERROR\(INDEX\(#REF!,MATCH\(.*?\)\),0\)", "0", formula)
    escaped_sheet = sheet_name.replace("'", "''")
    formula = formula.replace(f"'{escaped_sheet}'!", "")
    formula = formula.replace(f"{sheet_name}!", "")
    return formula


def read_shared_strings(zf):
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    values = []
    for si in root.findall(NS + "si"):
        values.append("".join(t.text or "" for t in si.iter(NS + "t")))
    return values


def cell_value(cell, shared_strings):
    typ = cell.attrib.get("t")
    value_node = cell.find(NS + "v")
    inline_node = cell.find(NS + "is")
    if typ == "s" and value_node is not None:
        return shared_strings[int(value_node.text)]
    if typ == "inlineStr" and inline_node is not None:
        return "".join(t.text or "" for t in inline_node.iter(NS + "t"))
    if typ == "b" and value_node is not None:
        return value_node.text == "1"
    if typ == "e" and value_node is not None:
        return value_node.text
    if value_node is None:
        return ""
    raw = value_node.text or ""
    try:
        number = float(raw)
        if number.is_integer():
            return int(number)
        return number
    except ValueError:
        return raw


def extract_workbook(path):
    with zipfile.ZipFile(path) as zf:
        shared_strings = read_shared_strings(zf)
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        relationships = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rels = {rel.attrib["Id"]: rel.attrib["Target"] for rel in relationships}

        sheets = []
        for sheet in workbook.find(NS + "sheets"):
            rid = sheet.attrib[REL_NS + "id"]
            sheets.append(
                {
                    "name": sheet.attrib["name"],
                    "path": "xl/" + rels[rid],
                    "state": sheet.attrib.get("state", "visible"),
                }
            )

        model = {"source": os.path.basename(path), "sheets": []}

        for sheet in sheets:
            root = ET.fromstring(zf.read(sheet["path"]))
            cells = {}
            cache = {}
            formula_masters = {}
            max_row = 0
            max_col = 0

            for row in root.find(NS + "sheetData").findall(NS + "row"):
                for cell in row.findall(NS + "c"):
                    ref = cell.attrib.get("r")
                    if not ref:
                        continue
                    col, row_num = split_cell(ref.replace("$", ""))
                    max_row = max(max_row, row_num)
                    max_col = max(max_col, col)

                    formula_node = cell.find(NS + "f")
                    if formula_node is not None:
                        formula_text = formula_node.text
                        shared_id = formula_node.attrib.get("si")
                        if formula_node.attrib.get("t") == "shared":
                            if formula_text:
                                formula_masters[shared_id] = {
                                    "ref": ref,
                                    "formula": formula_text,
                                }
                            else:
                                master = formula_masters.get(shared_id)
                                formula_text = (
                                    translate_formula(master["formula"], master["ref"], ref)
                                    if master
                                    else ""
                                )
                        if formula_text:
                            cells[ref] = "=" + normalize_formula(formula_text, sheet["name"])
                            cached_value = cell_value(cell, shared_strings)
                            if cached_value != "":
                                cache[ref] = cached_value
                        else:
                            cells[ref] = cell_value(cell, shared_strings)
                    else:
                        value = cell_value(cell, shared_strings)
                        if value != "":
                            cells[ref] = value

            model["sheets"].append(
                {
                    "name": sheet["name"],
                    "state": sheet["state"],
                    "maxRow": max_row,
                    "maxCol": max_col,
                    "cells": cells,
                    "cache": cache,
                }
            )

        optimize_static_index_match(model)
        return model


def optimize_static_index_match(model):
    sheets_by_name = {sheet["name"]: sheet for sheet in model["sheets"]}
    index_pattern = re.compile(
        r"INDEX\('([^']+)'\!\$([A-Z]+)\$(\d+):\$[A-Z]+\$(\d+),"
        r"MATCH\((\$?[A-Z]+\$?\d+),'([^']+)'\!\$([A-Z]+)\$\3:\$[A-Z]+\$\4,0\)\)"
    )
    index_pattern_plain = re.compile(
        r"INDEX\(([A-Za-zÀ-ÿ0-9_]+)!\$([A-Z]+)\$(\d+):\$[A-Z]+\$(\d+),"
        r"MATCH\((\$?[A-Z]+\$?\d+),([A-Za-zÀ-ÿ0-9_]+)!\$([A-Z]+)\$\3:\$[A-Z]+\$\4,0\)\)"
    )

    def static_value(sheet, ref):
        clean_ref = ref.replace("$", "")
        value = sheet["cells"].get(clean_ref)
        if isinstance(value, str) and value.startswith("="):
            return sheet.get("cache", {}).get(clean_ref)
        return value

    lookup_cache = {}

    def row_lookup(source_sheet, lookup_col, first_row, last_row):
        key = (source_sheet["name"], lookup_col, first_row, last_row)
        if key not in lookup_cache:
            lookup_cache[key] = {}
            for row in range(first_row, last_row + 1):
                value = static_value(source_sheet, f"{lookup_col}{row}")
                if value not in (None, ""):
                    lookup_cache[key][value] = row
        return lookup_cache[key]

    for sheet in model["sheets"]:
        cells = sheet["cells"]

        def repl(match):
            index_sheet_name, return_col, first_row, last_row, local_ref, match_sheet_name, lookup_col = match.groups()
            if index_sheet_name != match_sheet_name:
                return match.group(0)
            source_sheet = sheets_by_name.get(index_sheet_name)
            if not source_sheet:
                return match.group(0)
            item = static_value(sheet, local_ref)
            row = row_lookup(source_sheet, lookup_col, int(first_row), int(last_row)).get(item)
            if not row:
                return match.group(0)
            return f"'{index_sheet_name}'!${return_col}${row}"

        for ref, value in list(cells.items()):
            if isinstance(value, str) and value.startswith("=") and "INDEX('" in value and "MATCH(" in value:
                cells[ref] = index_pattern.sub(repl, value)
            if isinstance(cells.get(ref), str) and cells[ref].startswith("=") and "INDEX(" in cells[ref] and "MATCH(" in cells[ref]:
                cells[ref] = index_pattern_plain.sub(repl, cells[ref])

    restore_dynamic_origin_budget_formulas(model)
    break_origin_circular_rows(model)
    replace_unsupported_cached_formulas(model)


def numeric_static_value(sheet, ref):
    value = sheet["cells"].get(ref)
    if isinstance(value, str) and value.startswith("="):
        value = sheet.get("cache", {}).get(ref)
    return value if isinstance(value, (int, float)) else None


def median(values):
    ordered = sorted(values)
    if not ordered:
        return None
    middle = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[middle]
    return (ordered[middle - 1] + ordered[middle]) / 2


def restore_dynamic_origin_budget_formulas(model):
    origem = next((sheet for sheet in model["sheets"] if sheet["name"].endswith("(Origem)")), None)
    simulador = next((sheet for sheet in model["sheets"] if sheet["name"] == "SIMULADOR"), None)
    if not origem or not simulador:
        return

    sim_cells = simulador["cells"]
    item_to_budget_row = {}
    for row in range(11, 103):
        item = simulador.get("cache", {}).get(f"E{row}", sim_cells.get(f"E{row}"))
        if item:
            item_to_budget_row[item] = row

    center_to_toggle_row = {}
    for row in range(11, 27):
        center = sim_cells.get(f"B{row}")
        if center:
            center_to_toggle_row[center] = row

    for ref, formula in list(origem["cells"].items()):
        if not (ref.startswith("D") and isinstance(formula, str) and formula.startswith("=")):
            continue
        row = int(ref[1:])
        if f"$O{row}:$ACZ{row}" not in formula and "FILTER(" not in formula.upper():
            continue

        item = origem["cells"].get(f"B{row}")
        budget_row = item_to_budget_row.get(item)
        if not budget_row:
            continue

        center = sim_cells.get(f"F{budget_row}")
        toggle_row = center_to_toggle_row.get(center)
        if not toggle_row:
            continue

        lower = numeric_static_value(origem, f"L{row}")
        upper = numeric_static_value(origem, f"M{row}")
        values = []
        for col in range(col_to_num("O"), col_to_num("ACZ") + 1):
            number = numeric_static_value(origem, f"{num_to_col(col)}{row}")
            if number is None or number <= 0:
                continue
            values.append(number)

        positive_values = values
        filtered_values = [
            value for value in values
            if (lower is None or value >= lower) and (upper is None or value <= upper)
        ]
        min_value = min(positive_values) if positive_values else origem.get("cache", {}).get(ref, 0)
        median_value = median(filtered_values) if filtered_values else origem.get("cache", {}).get(ref, 0)

        origem["cells"][ref] = (
            f'=IF(SIMULADOR!$C$6="MÍNIMO",{min_value:.15g},{median_value:.15g})'
            f'*IF(SIMULADOR!$C${toggle_row}="SIM",1,0)'
            f'*(1+$E{row})'
        )


def break_origin_circular_rows(model):
    origem = next((sheet for sheet in model["sheets"] if sheet["name"].endswith("(Origem)")), None)
    if not origem:
        return

    def col_to_num_from_ref(ref):
        return col_to_num(re.match(r"([A-Z]+)", ref).group(1))

    cells = origem["cells"]
    cache = origem.get("cache", {})
    for ref, formula in list(cells.items()):
        if not (ref.startswith("D") and isinstance(formula, str) and formula.startswith("=")):
            continue
        row = int(ref[1:])
        if f"$O{row}:$ACZ{row}" not in formula:
            continue
        has_back_reference = False
        for other_ref, other_formula in cells.items():
            if not re.match(rf"[A-Z]+{row}$", other_ref):
                continue
            if col_to_num_from_ref(other_ref) < col_to_num("O"):
                continue
            if isinstance(other_formula, str) and other_formula.startswith("=") and "Orçamento (Mensal)" in other_formula:
                has_back_reference = True
                break
        if has_back_reference and ref in cache:
            cells[ref] = cache[ref]


def replace_unsupported_cached_formulas(model):
    unsupported_markers = ("OFFSET(", "FILTER(")
    for sheet in model["sheets"]:
        cache = sheet.get("cache", {})
        for ref, value in list(sheet["cells"].items()):
            if (
                isinstance(value, str)
                and value.startswith("=")
                and any(marker in value.upper() for marker in unsupported_markers)
                and ref in cache
            ):
                sheet["cells"][ref] = cache[ref]


def build_simulador_config(model):
    simulador = next(sheet for sheet in model["sheets"] if sheet["name"] == "SIMULADOR")
    cells = simulador["cells"]
    cache = simulador.get("cache", {})

    def value(ref):
        current = cells.get(ref, "")
        if isinstance(current, str) and current.startswith("="):
            return cache.get(ref, "")
        return current

    general = []
    for row in range(3, 8):
        general.append(
            {
                "label": value(f"B{row}"),
                "cell": f"C{row}",
                "value": value(f"C{row}"),
            }
        )

    cost_centers = []
    for row in range(11, 27):
        label = value(f"B{row}")
        if label:
            cost_centers.append(
                {
                    "label": label,
                    "cell": f"C{row}",
                    "value": value(f"C{row}"),
                }
            )

    budget = []
    for row in range(11, 103):
        item = value(f"E{row}")
        if item:
            budget.append(
                {
                    "item": item,
                    "center": value(f"F{row}"),
                    "cell": f"G{row}",
                    "value": value(f"G{row}") or 0,
                }
            )

    return {
        "source": model["source"],
        "results": {
            "annual": value("G3"),
            "monthly": value("G4"),
        },
        "general": general,
        "costCenters": cost_centers,
        "budget": budget,
    }


def main():
    workbook = next(name for name in os.listdir(".") if name.lower().endswith(".xlsm"))
    model = extract_workbook(workbook)
    os.makedirs("assets", exist_ok=True)
    with open(os.path.join("assets", "workbook-model.js"), "w", encoding="utf-8") as fh:
        fh.write("window.WORKBOOK_MODEL = ")
        json.dump(model, fh, ensure_ascii=False, separators=(",", ":"))
        fh.write(";\n")
    with open(os.path.join("assets", "simulador-config.js"), "w", encoding="utf-8") as fh:
        fh.write("window.SIMULADOR_CONFIG = ")
        json.dump(build_simulador_config(model), fh, ensure_ascii=False, separators=(",", ":"))
        fh.write(";\n")
    print(f"Generated assets/workbook-model.js and assets/simulador-config.js from {workbook}")


if __name__ == "__main__":
    main()
