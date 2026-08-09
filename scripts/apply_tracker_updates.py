from __future__ import annotations

import copy
import json
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

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
    "carbs": ["碳水化合物", "碳水", "carbohydrate", "carbs"],
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


def copy_cell_style(src, dst):
    if src.has_style:
        dst._style = copy.copy(src._style)
    dst.number_format = src.number_format
    dst.alignment = copy.copy(src.alignment)
    dst.fill = copy.copy(src.fill)
    dst.font = copy.copy(src.font)
    dst.border = copy.copy(src.border)
    dst.protection = copy.copy(src.protection)


def copy_row_style(ws, src_row: int, dst_row: int):
    if src_row < 1:
        return
    for col in range(1, ws.max_column + 1):
        copy_cell_style(ws.cell(src_row, col), ws.cell(dst_row, col))


def ensure_carbs_column(ws):
    header_row, mapping = locate_headers(ws)
    if "carbs" in mapping:
        return False, mapping["carbs"]

    new_col = ws.max_column + 1
    style_col = mapping.get("protein") or mapping.get("calories") or max(1, new_col - 1)
    ws.cell(header_row, new_col).value = "碳水（g）"
    copy_cell_style(ws.cell(header_row, style_col), ws.cell(header_row, new_col))
    for row in range(header_row + 1, ws.max_row + 1):
        copy_cell_style(ws.cell(row, style_col), ws.cell(row, new_col))
    ws.column_dimensions[get_column_letter(new_col)].width = 12
    return True, new_col


def append_note(existing, extra):
    existing_text = str(existing or "").strip()
    extra_text = str(extra or "").strip()
    if not extra_text or extra_text in existing_text:
        return existing_text
    return f"{existing_text} {extra_text}".strip()


def food_exists(ws, header_row, mapping, item):
    target_date = item["date"]
    target_meal = norm(item.get("meal"))
    target_food = norm(item.get("food"))
    for row in range(header_row + 1, ws.max_row + 1):
        date = date_key(ws.cell(row, mapping["date"]).value)
        meal = norm(ws.cell(row, mapping["meal"]).value) if mapping.get("meal") else ""
        food = norm(ws.cell(row, mapping["food"]).value) if mapping.get("food") else ""
        if date == target_date and meal == target_meal and food == target_food:
            return row
    return None


def apply_food(ws, item):
    header_row, mapping = locate_headers(ws)
    existing = food_exists(ws, header_row, mapping, item)

    if existing and item.get("update_existing"):
        changed_fields = []
        for field in ("source", "protein", "carbs", "calories", "note"):
            if field not in item or not mapping.get(field):
                continue
            ws.cell(existing, mapping[field]).value = item[field]
            changed_fields.append(field)
        if item.get("note_append") and mapping.get("note"):
            cell = ws.cell(existing, mapping["note"])
            cell.value = append_note(cell.value, item["note_append"])
            changed_fields.append("note_append")
        return {"status": "updated", "row": existing, "fields": changed_fields, "item": item}

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
        "carbs": item.get("carbs", ""),
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
        raise RuntimeError(f"驗證失敗：找不到餐點 {item}")

    fields_to_check = ("protein", "carbs", "calories")
    for field in fields_to_check:
        if field not in item or not mapping.get(field):
            continue
        expected = item[field]
        if expected in (None, ""):
            continue
        actual = ws.cell(row, mapping[field]).value
        if str(actual) != str(expected):
            raise RuntimeError(f"驗證失敗：{field} 預期 {expected}，實際 {actual}")

    if item.get("note_append") and mapping.get("note"):
        actual_note = str(ws.cell(row, mapping["note"]).value or "")
        if item["note_append"] not in actual_note:
            raise RuntimeError("驗證失敗：備註修正未寫入")

    return {
        "verified": True,
        "row": row,
        "date": item["date"],
        "meal": item.get("meal"),
        "food": item.get("food"),
        "status": result.get("status"),
    }


def main():
    payload = json.loads(PENDING.read_text(encoding="utf-8"))
    updates = payload.get("updates", [])

    wb = load_workbook(WORKBOOK)
    food_ws = find_sheet(wb, "飲食")
    added_column, carbs_col = ensure_carbs_column(food_ws)

    results = []
    for item in updates:
        if item.get("type") == "food":
            results.append(apply_food(food_ws, item))
        else:
            raise RuntimeError(f"尚未支援的 update type: {item.get('type')}")

    changed = added_column or any(r.get("status") in {"written", "updated"} for r in results)
    if changed:
        wb.save(WORKBOOK)

    verify_wb = load_workbook(WORKBOOK, data_only=False)
    verify_food_ws = find_sheet(verify_wb, "飲食")
    _, verify_mapping = locate_headers(verify_food_ws)
    if "carbs" not in verify_mapping:
        raise RuntimeError("驗證失敗：飲食紀錄沒有碳水欄位")

    verifications = [verify_food(verify_food_ws, result) for result in results]
    migration = {
        "carbs_column_added": added_column,
        "carbs_column": carbs_col,
    }
    output = {"migration": migration, "results": results, "verification": verifications}
    RESULT.write_text(json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if updates:
        PENDING.write_text(json.dumps({"updates": []}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(output, ensure_ascii=False))


if __name__ == "__main__":
    main()
