#!/usr/bin/env python3
"""Читает Поздравление.xlsx и выводит COLLEAGUES_TABLE для вставки в index.html.
   Ожидает колонки: id (A), Почта (B), Имя (C), Текст поздравления (D).
   Фото: photos/{id}.png
"""
import zipfile
import xml.etree.ElementTree as ET
import re
import json
import os

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}

def get_shared_strings(zip_path):
    with zipfile.ZipFile(zip_path, "r") as z:
        with z.open("xl/sharedStrings.xml") as f:
            tree = ET.parse(f)
    root = tree.getroot()
    strings = []
    for si in root.findall(".//main:si", NS):
        parts = []
        for t in si.findall(".//main:t", NS):
            if t.text:
                parts.append(t.text)
        for r in si.findall(".//main:r", NS):
            t = r.find("main:t", NS)
            if t is not None and t.text:
                parts.append(t.text)
        strings.append("".join(parts) if parts else "")
    return strings

def cell_ref_to_col(cell_ref):
    """A1 -> 0, B1 -> 1, AA1 -> 26"""
    col = re.match(r"^([A-Z]+)", cell_ref, re.I)
    if not col:
        return 0
    s = col.group(1).upper()
    n = 0
    for c in s:
        n = n * 26 + (ord(c) - ord("A") + 1)
    return n - 1

def get_sheet_rows(zip_path, sheet_path="xl/worksheets/sheet2.xml"):
    with zipfile.ZipFile(zip_path, "r") as z:
        with z.open(sheet_path) as f:
            tree = ET.parse(f)
    root = tree.getroot()
    rows = {}
    for row in root.findall(".//main:row", NS):
        r = int(row.get("r", 0))
        rows[r] = {}
        for c in row.findall("main:c", NS):
            ref = c.get("r", "")
            col = cell_ref_to_col(ref)
            val_el = c.find("main:v", NS)
            val = val_el.text if val_el is not None else ""
            is_str = c.get("t") == "s"
            rows[r][col] = (val, is_str)
    return rows

def main():
    # Предпочтительно читаем из проекта: march8/data/Поздравление.xlsx
    xlsx_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "data", "Поздравление.xlsx"))
    if not os.path.isfile(xlsx_path):
        # fallback (если запускают локально и файл лежит в Downloads)
        xlsx_path = "/Users/chernykhvitaly/Downloads/Поздравление (1).xlsx"
    if not os.path.isfile(xlsx_path):
        xlsx_path = "/Users/chernykhvitaly/Downloads/Поздравление.xlsx"
    if not os.path.isfile(xlsx_path):
        print("Файл Поздравление.xlsx не найден в data/ или ~/Downloads")
        return 1

    strings = get_shared_strings(xlsx_path)
    rows_data = get_sheet_rows(xlsx_path)

    # Строка 1 — заголовки (id, Почта, Имя, Текст поздравления), данные с строки 2
    col_id, col_email, col_name, col_text = 0, 1, 2, 3
    header_row = 1

    out = []
    for r in sorted(rows_data.keys()):
        if r <= header_row:
            continue
        row = rows_data[r]
        def get_cell(col):
            if col not in row:
                return ""
            v, is_str = row[col]
            if is_str and v.isdigit():
                return strings[int(v)].strip()
            return (v or "").strip()

        id_val = get_cell(0)
        email = get_cell(1)
        name = get_cell(2)
        text = get_cell(3)

        if not email or "@" not in email:
            continue
        if not name:
            name = "Коллега"

        photo_url = f"photos/{int(float(id_val))}.png" if id_val and id_val.replace(".", "").isdigit() else ""
        # Экранируем для JS-строки
        def esc(s):
            return s.replace("\\", "\\\\").replace("\n", "\\n").replace("\r", "").replace('"', '\\"')
        out.append({
            "email": email,
            "name": name,
            "text": text,
            "photoUrl": photo_url,
        })

    # Вывод как JavaScript массив
    lines = []
    for o in out:
        lines.append(f'      {{ email: "{esc(o["email"])}", name: "{esc(o["name"])}", text: "{esc(o["text"])}", photoUrl: "{o["photoUrl"]}" }}')
    print("const COLLEAGUES_TABLE = [")
    print(",\n".join(lines))
    print("    ];")
    return 0

if __name__ == "__main__":
    exit(main())
