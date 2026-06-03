from .utils import sanitize_filename_component, strip_same_extension


def create_initial_rename_state(file_inventory):
    state = {}
    for item in file_inventory:
        state[item["original_name"]] = {
            "selected": False,
            "new_stem": sanitize_filename_component(item["stem"]),
            "keep_extension": True,
            "ai_suggested_stem": "",
            "ai_reason": "",
            "ai_confidence": "",
            "ai_review": "",
            "ai_review_risk_level": "",
            "ai_review_comment": "",
            "use_ai_suggestion": False,
            "status": "待编辑",
        }
    return state


def apply_paste_name_list(file_inventory, rename_state, text):
    lines = str(text or "").splitlines()
    updated = skipped_empty = extra = 0
    for index, raw in enumerate(lines):
        if index >= len(file_inventory):
            extra += 1
            continue
        if not raw.strip():
            skipped_empty += 1
            continue
        item = file_inventory[index]
        rename_state[item["original_name"]]["new_stem"] = strip_same_extension(raw, item["extension"])
        rename_state[item["original_name"]]["status"] = "待编辑"
        updated += 1
    return {"updated": updated, "skipped_empty": skipped_empty, "extra": extra}


def clear_selection(rename_state):
    for row in rename_state.values():
        row["selected"] = False
        row["use_ai_suggestion"] = False
