from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "30天重建追蹤.xlsx"
SHEET_NAME = "動作基準"
HEADERS = [
    "動作",
    "建議下次重量（kg）",
    "最近訓練日期",
    "最近使用重量（kg）",
    "最近次數",
    "最近DBR",
    "備註",
]


def norm(value):
    return str(value or "").strip()


def date_key(value):
    if value in (None, ""):
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip().replace("/", "-").replace(".", "-")
    try:
        return datetime.fromisoformat(text[:10]).strftime("%Y-%m-%d")
    except ValueError:
        return text


def as_number(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        text = str(value).strip()
        return float(text) if text else None
    except (TypeError, ValueError):
        return None


def tidy_number(value):
    if value is None:
        return ""
    value = float(value)
    return int(value) if value.is_integer() else value


def find_workout_sheet(wb):
    for name in wb.sheetnames:
        if "重訓" in name:
            return wb[name]
    raise RuntimeError("找不到重訓工作表")


def detect_headers(ws):
    for row in range(1, min(ws.max_row, 12) + 1):
        values = {col: norm(ws.cell(row, col).value) for col in range(1, ws.max_column + 1)}
        mapping = {}
        for col, value in values.items():
            upper = value.upper()
            if "日期" in value:
                mapping["date"] = col
            elif "動作" in value:
                mapping["exercise"] = col
            elif "重量" in value:
                mapping["weight"] = col
            elif "次數" in value:
                mapping["reps"] = col
            elif upper == "DBR":
                mapping["dbr"] = col
            elif upper == "RPE":
                mapping["dbr"] = col
            elif "組次" in value or "組數" in value:
                mapping["set"] = col
        if {"date", "exercise", "weight", "reps", "dbr"}.issubset(mapping):
            return row, mapping
    raise RuntimeError("無法辨識重訓紀錄欄位")


def choose_suggested_weight(rows):
    positive = [r for r in rows if r["weight"] is not None and r["weight"] > 0]
    if not positive:
        return 0.0

    # 優先選最近一次訓練中，至少還留 1 下（DBR >= 1）的最高重量。
    # 若當天所有已知 DBR 都是 0，改採當天較低重量，避免把力竭重量直接當下次起始重量。
    known_dbr = [r for r in positive if r["dbr"] is not None]
    eligible = [r for r in known_dbr if r["dbr"] >= 1]
    if eligible:
        return max(r["weight"] for r in eligible)
    if known_dbr:
        return min(r["weight"] for r in positive)

    # 沒有 DBR 資訊時，只能保守記住最近使用過的最高重量。
    return max(r["weight"] for r in positive)


def main():
    wb = load_workbook(WORKBOOK)
    ws = find_workout_sheet(wb)
    header_row, mapping = detect_headers(ws)

    by_exercise = defaultdict(list)
    for row in range(header_row + 1, ws.max_row + 1):
        exercise = norm(ws.cell(row, mapping["exercise"]).value)
        dt = date_key(ws.cell(row, mapping["date"]).value)
        if not exercise or not dt:
            continue
        by_exercise[exercise].append({
            "row": row,
            "date": dt,
            "set": as_number(ws.cell(row, mapping.get("set", 0)).value) if mapping.get("set") else None,
            "weight": as_number(ws.cell(row, mapping["weight"]).value),
            "reps": as_number(ws.cell(row, mapping["reps"]).value),
            "dbr": as_number(ws.cell(row, mapping["dbr"]).value),
        })

    if SHEET_NAME in wb.sheetnames:
        base = wb[SHEET_NAME]
        wb.remove(base)
    base = wb.create_sheet(SHEET_NAME)

    for col, header in enumerate(HEADERS, 1):
        cell = base.cell(1, col, header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.alignment = Alignment(horizontal="center")

    out_row = 2
    for exercise in sorted(by_exercise):
        rows = by_exercise[exercise]
        latest_date = max(r["date"] for r in rows)
        latest = [r for r in rows if r["date"] == latest_date]
        latest.sort(key=lambda r: ((r["set"] if r["set"] is not None else 9999), r["row"]))

        suggested = choose_suggested_weight(latest)

        # 顯示與建議重量最接近的最後一組；自重則顯示當天最後一組。
        candidates = [r for r in latest if r["weight"] == suggested]
        recent = candidates[-1] if candidates else latest[-1]

        base.cell(out_row, 1, exercise)
        base.cell(out_row, 2, tidy_number(suggested))
        base.cell(out_row, 3, latest_date)
        base.cell(out_row, 4, tidy_number(recent["weight"]))
        base.cell(out_row, 5, tidy_number(recent["reps"]))
        base.cell(out_row, 6, tidy_number(recent["dbr"]))
        base.cell(
            out_row,
            7,
            "自動基準：以最近一次訓練為準；優先採 DBR≥1 的最高重量；若全部已知組皆 DBR0，採當日較低重量。",
        )
        out_row += 1

    widths = {1: 22, 2: 20, 3: 15, 4: 20, 5: 12, 6: 12, 7: 70}
    for col, width in widths.items():
        base.column_dimensions[base.cell(1, col).column_letter].width = width
    base.freeze_panes = "A2"
    base.auto_filter.ref = f"A1:G{max(1, out_row - 1)}"

    wb.save(WORKBOOK)
    print(f"Updated {SHEET_NAME}: {len(by_exercise)} exercises")


if __name__ == "__main__":
    main()
