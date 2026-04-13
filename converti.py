#!/usr/bin/env python3
"""
Converte PianificazioneAule.xls → orario.json
Eseguito automaticamente da GitHub Actions ad ogni push del file XLS.
"""

import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path

XLS_PATH  = Path("PianificazioneAule.xls")
JSON_PATH = Path("orario.json")

LESSON_STYLES = {"1", "2", "4", "5"}

# ── HELPERS ────────────────────────────────────────────────

def cell_text(cell, ns):
    data = cell.find("ss:Data", ns)
    if data is None:
        return None
    texts = []
    def recurse(el):
        if el.text:
            texts.append(el.text)
        for child in el:
            recurse(child)
            if child.tail:
                texts.append(child.tail)
    recurse(data)
    result = "".join(texts).strip()
    return result if result else None

def extract_code(raw):
    m = re.match(r"^(\d[A-Z/]+)", raw)
    if m:
        return m.group(1).rstrip("_")
    parts = raw.split("_")
    if parts[0] in ("MI", "FAMI", "XXXX"):
        return "_".join(parts[:3])
    return parts[0]

def parse_cell(raw, file_hour):
    lines = [l.strip() for l in raw.split("\n") if l.strip()]
    if not lines:
        return None

    start_time = class_code = course = uf = teacher = None
    fl = lines[0]
    tm = re.match(r"^(\d{2}:\d{2})\s+(.+)$", fl)
    if tm:
        start_time = tm.group(1)
        rest = tm.group(2).strip()
        cm = re.match(r"^([A-Z0-9/_\-\.]+(?:/[\d_]+)?)\s*", rest)
        class_code = extract_code(cm.group(1)) if cm else rest.split()[0]
    else:
        cm = re.match(r"^([A-Z0-9/_\-\.]+(?:/[\d_]+)?)", fl)
        class_code = extract_code(cm.group(1)) if cm else fl[:20]

    for l in lines[1:]:
        if l.startswith("UF:"):
            uf = l[3:].strip()
        elif l.startswith("DOC:"):
            teacher = l[4:].strip()
        elif not course and not re.match(r"^[A-Z0-9/_\-]+$", l):
            course = l

    if not teacher and not uf:
        return None

    return dict(startTime=start_time, classCode=class_code,
                course=course, uf=uf, teacher=teacher)

# ── MAIN ───────────────────────────────────────────────────

def convert():
    if not XLS_PATH.exists():
        raise FileNotFoundError(f"File non trovato: {XLS_PATH}")

    tree = ET.parse(XLS_PATH)
    root = tree.getroot()
    ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}

    ws = root.findall("ss:Worksheet", ns)[0]
    sheet_name = ws.get("{urn:schemas-microsoft-com:office:spreadsheet}Name", "")
    table = ws.find("ss:Table", ns)
    rows = table.findall("ss:Row", ns)

    rooms = {}
    raw_lessons = []
    current_hour = None

    for ri, row in enumerate(rows):
        cells = row.findall("ss:Cell", ns)
        col = 1
        for cell in cells:
            idx = cell.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
            if idx:
                col = int(idx)
            style = cell.get("{urn:schemas-microsoft-com:office:spreadsheet}StyleID", "")
            text = cell_text(cell, ns)

            if ri == 3 and text:
                rooms[col] = text
            if ri >= 4:
                if style == "s02" and text:
                    current_hour = text
                if style in LESSON_STYLES and text:
                    parsed = parse_cell(text, current_hour)
                    if parsed:
                        parsed["room"] = rooms.get(col, f"col{col}")
                        parsed["fileHour"] = current_hour
                        raw_lessons.append(parsed)
            col += 1

    payload = {
        "title": sheet_name,
        "rooms": {str(k): v for k, v in rooms.items()},
        "rawLessons": raw_lessons,
        "ts": None  # filled by Actions
    }

    # Add timestamp
    from datetime import datetime, timezone
    payload["ts"] = int(datetime.now(timezone.utc).timestamp() * 1000)

    JSON_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✓ Convertite {len(raw_lessons)} lezioni → {JSON_PATH}")

if __name__ == "__main__":
    convert()
