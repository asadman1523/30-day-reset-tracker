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

# 既有飲食列建立時尚未有碳水欄位。以下數值沿用各列既有份量與備註假設，
# 只補空白欄位，不覆蓋後續人工修正值。
LEGACY_NUTRITION = {
    ("2026-08-03", "早餐", "培根蛋吐司、雞塊"): {
        "carbs": 65,
        "note": "碳水估算：吐司與雞塊裹粉合計約 65 g。",
    },
    ("2026-08-03", "午餐", "未進食"): {"carbs": 0},
    ("2026-08-03", "晚餐", "牛排、雞排、酥皮濃湯"): {
        "carbs": 45,
        "note": "碳水估算：主要來自酥皮濃湯，未另計醬料。",
    },
    ("2026-08-04", "早餐", "地瓜、海鹽水煮蛋、雞胸肉"): {
        "carbs": 35,
        "note": "碳水估算：主要來自地瓜。",
    },
    ("2026-08-04", "午餐", "大麥克、無糖可樂、小杯玉米湯、中薯"): {
        "carbs": 120,
        "note": "碳水估算：大麥克、中薯與小杯玉米湯合計約 120 g。",
    },
    ("2026-08-04", "晚餐", "雙倍雞肉飯，飯一半"): {
        "carbs": 55,
        "note": "碳水估算：主要來自半份白飯。",
    },
    ("2026-08-05", "早餐", "御選肉鬆飯糰、雙蔬鮪魚飯糰"): {
        "carbs": 90,
        "note": "碳水估算：兩顆飯糰合計約 90 g。",
    },
    ("2026-08-05", "午餐", "未進食"): {"carbs": 0},
    ("2026-08-05", "晚餐", "松阪豬、嫩煎豆腐、豆芽菜、高麗菜、豆漿"): {
        "carbs": 28,
        "note": "碳水估算：主要來自豆漿與蔬菜，未另計醬汁。",
    },
    ("2026-08-06", "早餐", "豆漿、地瓜、雞胸肉"): {
        "carbs": 48,
        "note": "碳水估算：地瓜約 35 g、豆漿約 13 g。",
    },
    ("2026-08-06", "午餐", "乳清蛋白粉 40g"): {
        "carbs": 4,
        "note": "碳水估算：未提供品牌，按一般乳清粉估算。",
    },
    ("2026-08-06", "晚餐", "炸雞排、炸杏鮑菇、炸魷魚"): {
        "carbs": 70,
        "note": "碳水估算：主要來自裹粉。",
    },
    ("2026-08-07", "早餐", "地瓜、雞胸肉"): {
        "carbs": 35,
        "note": "碳水估算：主要來自地瓜。",
    },
    ("2026-08-07", "午餐", "乳清蛋白粉 40g"): {
        "carbs": 4,
        "note": "碳水估算：未提供品牌，按一般乳清粉估算。",
    },
    ("2026-08-07", "晚餐", "肉片約 15 片、蔬菜"): {
        "carbs": 15,
        "note": "碳水估算：主要來自蔬菜與湯底，未另計醬料。",
    },
    ("2026-08-08", "早餐", "地瓜、雞胸肉"): {
        "carbs": 35,
        "note": "碳水估算：主要來自地瓜。",
    },
    ("2026-08-08", "午餐", "乳清蛋白粉 60g、豆漿"): {
        "carbs": 19,
        "note": "碳水估算：乳清約 6 g、豆漿約 13 g。",
    },
    ("2026-08-08", "晚餐", "昆布鍋、12 盎司增肌減脂肉片、活鮑魚 3 個、白飯 1 碗"): {
        "carbs": 70,
        "note": "碳水估算：白飯約 60 g，加菜盤與湯底約 10 g。",
    },
    ("2026-08-08", "晚餐", "40g 乳清 + 1 杯豆漿"): {
        "protein": 43,
        "calories": 304,
        "carbs": 17,
        "note": "估算：沿用既有假設，40 g 乳清約蛋白質 30 g、160 kcal、碳水 4 g；豆漿約蛋白質 13 g、144 kcal、碳水 13 g。",
    },
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


def backfill_legacy_nutrition(ws):
    header_row, mapping = locate_headers(ws)
    changed_rows = []
    for row in range(header_row + 1, ws.max_row + 1):
        key = (
            date_key(ws.cell(row, mapping["date"]).value),
            str(ws.cell(row, mapping["meal"]).value or "").strip(),
            str(ws.cell(row, mapping["food"]).value or "").strip(),
        )
        values = LEGACY_NUTRITION.get(key)
        if not values:
            continue

        changed_fields = []
        for field in ("protein", "calories", "carbs"):
            col = mapping.get(field)
            if not col or field not in values:
                continue
            cell = ws.cell(row, col)
            if cell.value in (None, ""):
                cell.value = values[field]
                changed_fields.append(field)

        if changed_fields and mapping.get("note") and values.get("note"):
            note_cell = ws.cell(row, mapping["note"])
            note_cell.value = append_note(note_cell.value, values["note"])

        if changed_fields:
            changed_rows.append({"row": row, "key": key, "fields": changed_fields})
    return changed_rows


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
        raise RuntimeError(f"驗證失敗：找不到已寫入餐點 {item}")

    for field in ("protein", "carbs", "calories"):
        expected = item.get(field, "")
        if expected in (None, "") or not mapping.get(field):
            continue
        actual = ws.cell(row, mapping[field]).value
        if str(actual) != str(expected):
            raise RuntimeError(f"驗證失敗：{field} 預期 {expected}，實際 {actual}")

    return {
        "verified": True,
        "row": row,
        "date": item["date"],
        "meal": item.get("meal"),
        "food": item.get("food"),
    }


def main():
    payload = json.loads(PENDING.read_text(encoding="utf-8"))
    updates = payload.get("updates", [])

    wb = load_workbook(WORKBOOK)
    food_ws = find_sheet(wb, "飲食")
    added_column, carbs_col = ensure_carbs_column(food_ws)
    backfilled = backfill_legacy_nutrition(food_ws)

    results = []
    for item in updates:
        if item.get("type") == "food":
            results.append(apply_food(food_ws, item))
        else:
            raise RuntimeError(f"尚未支援的 update type: {item.get('type')}")

    changed = added_column or bool(backfilled) or bool(results)
    if changed:
        wb.save(WORKBOOK)

    verify_wb = load_workbook(WORKBOOK, data_only=False)
    verify_food_ws = find_sheet(verify_wb, "飲食")
    _, verify_mapping = locate_headers(verify_food_ws)
    if "carbs" not in verify_mapping:
        raise RuntimeError("驗證失敗：飲食紀錄沒有碳水欄位")

    verifications = []
    for result in results:
        if result["item"].get("type") == "food":
            verifications.append(verify_food(verify_food_ws, result))

    migration = {
        "carbs_column_added": added_column,
        "carbs_column": carbs_col,
        "backfilled_rows": backfilled,
    }
    output = {"migration": migration, "results": results, "verification": verifications}
    RESULT.write_text(json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if updates:
        PENDING.write_text(json.dumps({"updates": []}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(output, ensure_ascii=False))


if __name__ == "__main__":
    main()
