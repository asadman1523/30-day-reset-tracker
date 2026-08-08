from __future__ import annotations

import copy
import json
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "30天重建追蹤.xlsx"
PENDING = ROOT / "pending_updates.json"
RESULT = ROOT / "tracker_update_result.json"

HEADER_ALIASES = {
    "date": ["日期", "date"],
    "meal": ["餐別", "餐次"],
    "food": ["餐點", "內容", "食物"],
    "source": ["來源"],
    "protein": ["蛋白質"],
    "calories": ["熱量", "卡路里"],
    "note": ["備註"],
}


def norm(value):
    return str(value or "").strip().lower()


def find_sheet(wb, keyword: str):
    for name in wb.sheetnames:
        if keyword in name:
            return wb[name]
    raise RuntimeError(f"找不到包含「{keyword}」的工作表")


def locate_headers(ws):
    for row in range(1, min(ws.max_row, 12) + 1):
        values = {col: norm(ws.cell(row, col).value) for col in range(1, ws.max_column + 1)}
        mapping = {}
        for field, aliases in HEADER_ALIASES.items():
            for col, value in values.items():
                if any(alias.lower() == value or alias.lower() in value for alias in aliases):
                    mapping[field] = col
                    break
        if len(mapping) >= 3 and "date" in mapping:
            return row, mapping
    raise RuntimeError(f"無法辨識工作表 {ws.title} 的欄位標題")


def date_key(value):
    if value is None or value == "":
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip().replace("/", "-").replace(".", "-")
    try:
        return datetime.fromisoformat(text[:10]).strftime("%Y-%m-%d")
    except ValueError:
        return text


def copy_row_style(ws, src_row: int, dst_row: int):
    if src_row < 1:
        return
    for col in range(1, ws.max_column + 1):
        src = ws.cell(src_row, col)
        dst = ws.cell(dst_row, col)
        if src.has_style:
            dst._style = copy.copy(src._style)
        dst.number_format = src.number_format
        dst.alignment = copy.copy(src.alignment)
        dst.fill = copy.copy(src.fill)
        dst.font = copy.copy(src.font)
        dst.border = copy.copy(src.border)
        dst.protection = copy.copy(src.protection)


def food_exists(ws, header_row, mapping, item):
    target_date = item["date"]
    target_meal = norm(item.get("meal"))
    target_food = norm(item.get("food"))
    for row in range(header_row + 1, ws.max_row + 1):
        d = date_key(ws.cell(row, mapping["date"]).value)
        m = norm(ws.cell(row, mapping["meal"]).value) if mapping.get("meal") else ""
        f = norm(ws.cell(row, mapping["food"]).value) if mapping.get("food") else ""
        if d == target_date and m == target_meal and f == target_food:
            return row
    return None


def apply_food(ws, item):
    header_row, mapping = locate_headers(ws)
    existing = food_exists(ws, header_row, mapping, item)
    if existing:
        return {"status": "already_present", "row": existing, "item": item}

    row = ws.max_row + 1
    copy_row_style(ws, max(header_row + 1, ws.max_row), row)
    values = {
        "date": datetime.strptime(item["date"], "%Y-%m-%d").date(),
        "meal": item.get("meal", ""),
        "food": item.get("food", ""),
        "source": item.get("source", ""),
        "protein": item.get("protein", ""),
        "calories": item.get("calories", ""),
        "note": item.get("note", ""),
    }
    for field, value in values.items():
        col = mapping.get(field)
        if col:
            ws.cell(row, col).value = value
    return {"status": "written", "row": row, "item": item}


def verify_food(ws, result):
    item = result["item"]
    header_row, mapping = locate_headers(ws)
    row = food_exists(ws, header_row, mapping, item)
    if not row:
        raise RuntimeError(f"驗證失敗：找不到已寫入餐點 {item}")
    return {"verified": True, "row": row, "date": item["date"], "meal": item.get("meal"), "food": item.get("food")}


def main():
    payload = json.loads(PENDING.read_text(encoding="utf-8"))
    updates = payload.get("updates", [])
    if not updates:
        print("No pending updates.")
        return

    wb = load_workbook(WORKBOOK)
    results = []
    for item in updates:
        if item.get("type") == "food":
            results.append(apply_food(find_sheet(wb, "飲食"), item))
        else:
            raise RuntimeError(f"尚未支援的 update type: {item.get('type')}")
    wb.save(WORKBOOK)

    verify_wb = load_workbook(WORKBOOK, data_only=False)
    verifications = []
    for result in results:
        if result["item"].get("type") == "food":
            verifications.append(verify_food(find_sheet(verify_wb, "飲食"), result))

    RESULT.write_text(json.dumps({"results": results, "verification": verifications}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    PENDING.write_text(json.dumps({"updates": []}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"results": results, "verification": verifications}, ensure_ascii=False))


if __name__ == "__main__":
    main()
