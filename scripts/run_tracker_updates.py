import json
import traceback
from pathlib import Path

import apply_tracker_updates as tracker


PROCESSED_IDS = Path("processed_operation_ids.json")
MAX_PROCESSED_IDS = 1000


def _find_food_row(ws, header_row, mapping, item):
    match_food = item.get("match_food")
    if not match_food:
        return tracker.food_exists(ws, header_row, mapping, item)
    match_item = dict(item)
    match_item["food"] = match_food
    return tracker.food_exists(ws, header_row, mapping, match_item)


def apply_food(ws, item):
    header_row, mapping = tracker.locate_headers(ws)
    existing = _find_food_row(ws, header_row, mapping, item)

    if existing and item.get("update_existing"):
        changed_fields = []
        for field in ("food", "source", "protein", "carbs", "calories", "note"):
            if field not in item or not mapping.get(field):
                continue
            if not tracker.scalar_equal(ws.cell(existing, mapping[field]).value, item[field]):
                ws.cell(existing, mapping[field]).value = item[field]
                changed_fields.append(field)
        if item.get("note_append") and mapping.get("note"):
            cell = ws.cell(existing, mapping["note"])
            updated = tracker.append_note(cell.value, item["note_append"])
            if updated != str(cell.value or ""):
                cell.value = updated
                changed_fields.append("note_append")
        return {
            "status": "updated" if changed_fields else "already_present",
            "row": existing,
            "fields": changed_fields,
            "item": item,
        }

    if existing:
        return {"status": "already_present", "row": existing, "item": item}

    row = ws.max_row + 1
    tracker.copy_row_style(ws, max(header_row + 1, ws.max_row), row)
    note = item.get("note", "")
    if item.get("note_append"):
        note = tracker.append_note(note, item["note_append"])
    values = {
        "date": tracker.datetime.strptime(item["date"], "%Y-%m-%d").date(),
        "meal": item.get("meal", ""),
        "food": item.get("food", ""),
        "source": item.get("source", ""),
        "protein": item.get("protein", ""),
        "carbs": item.get("carbs", ""),
        "calories": item.get("calories", ""),
        "note": note,
    }
    for field, value in values.items():
        if mapping.get(field):
            ws.cell(row, mapping[field]).value = value
    return {"status": "written", "row": row, "item": item}


def verify_food(ws, result):
    item = result["item"]
    header_row, mapping = tracker.locate_headers(ws)
    row = result.get("row")
    if not row or row <= header_row or row > ws.max_row:
        row = tracker.food_exists(ws, header_row, mapping, item)
    if not row:
        raise RuntimeError(f"驗證失敗：找不到餐點 {item}")

    if tracker.date_key(ws.cell(row, mapping["date"]).value) != item["date"]:
        raise RuntimeError("驗證失敗：餐點日期不一致")
    if mapping.get("meal") and tracker.norm(ws.cell(row, mapping["meal"]).value) != tracker.norm(item.get("meal")):
        raise RuntimeError("驗證失敗：餐別不一致")
    if mapping.get("food") and tracker.norm(ws.cell(row, mapping["food"]).value) != tracker.norm(item.get("food")):
        raise RuntimeError("驗證失敗：餐點名稱不一致")

    for field in ("source", "protein", "carbs", "calories"):
        if field not in item or not mapping.get(field) or item[field] in (None, ""):
            continue
        actual = ws.cell(row, mapping[field]).value
        if not tracker.scalar_equal(actual, item[field]):
            raise RuntimeError(f"驗證失敗：{field} 預期 {item[field]}，實際 {actual}")

    if "note" in item and mapping.get("note"):
        actual_note = str(ws.cell(row, mapping["note"]).value or "")
        if actual_note != str(item.get("note") or ""):
            raise RuntimeError("驗證失敗：備註不一致")
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


def _load_processed_ids():
    if not PROCESSED_IDS.exists():
        return []
    try:
        payload = json.loads(PROCESSED_IDS.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return []
    values = payload.get("operation_ids", [])
    return [str(value) for value in values if value]


def _write_processed_ids(values):
    deduped = list(dict.fromkeys(values))[-MAX_PROCESSED_IDS:]
    PROCESSED_IDS.write_text(
        json.dumps({"operation_ids": deduped}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def main():
    original_pending = tracker.PENDING.read_text(encoding="utf-8")
    processed_ids = _load_processed_ids()
    processed_set = set(processed_ids)
    skipped_ids = []

    try:
        payload = json.loads(original_pending)
        updates = payload.get("updates", [])
        fresh_updates = []
        for item in updates:
            operation_id = item.get("operation_id")
            if operation_id and str(operation_id) in processed_set:
                skipped_ids.append(str(operation_id))
                continue
            fresh_updates.append(item)

        if len(fresh_updates) != len(updates):
            tracker.PENDING.write_text(
                json.dumps({"updates": fresh_updates}, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )

        tracker.apply_food = apply_food
        tracker.verify_food = verify_food
        tracker.main()

        new_ids = [
            str(item["operation_id"])
            for item in fresh_updates
            if item.get("operation_id")
        ]
        if new_ids:
            _write_processed_ids(processed_ids + new_ids)

        if tracker.RESULT.exists():
            result_payload = json.loads(tracker.RESULT.read_text(encoding="utf-8"))
            result_payload["deduplicated_operation_ids"] = skipped_ids
            tracker.RESULT.write_text(
                json.dumps(result_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
    except Exception as exc:
        tracker.PENDING.write_text(original_pending, encoding="utf-8")
        failure = {
            "status": "failed",
            "error": str(exc),
            "pending_preserved": True,
            "traceback": traceback.format_exc(limit=8),
        }
        tracker.RESULT.write_text(
            json.dumps(failure, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        raise


if __name__ == "__main__":
    main()
