import pandas as pd

from universal_rename_tool.history import append_history, load_history, suggest_from_history


def test_append_and_load_history(tmp_path):
    path = tmp_path / "history.xlsx"
    plan = pd.DataFrame([{"原文件名": "13_old.mp3", "输出文件名": "13_new.mp3", "原扩展名": ".mp3", "清洗后的新文件名主体": "13_new"}])
    append_history(path, plan)
    loaded = load_history(path)
    assert len(loaded) == 1


def test_exact_history_suggestion():
    inventory = [{"original_name": "a.mp3"}]
    history = pd.DataFrame([{"原文件名": "a.mp3", "清洗后的新文件名主体": "new_a"}])
    assert suggest_from_history(inventory, history)["a.mp3"] == "new_a"


def test_prefix_history_suggestion():
    inventory = [{"original_name": "13_any.mp3"}]
    history = pd.DataFrame([{"原文件名": "x.mp3", "清洗后的新文件名主体": "13_m_up", "original_prefix_number": "13"}])
    assert suggest_from_history(inventory, history)["13_any.mp3"] == "13_m_up"
