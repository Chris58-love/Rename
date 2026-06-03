from pathlib import Path

from universal_rename_tool.rule_dictionary import build_rule_plan, validate_rule_plan


def test_rule_hit(tmp_path):
    file_path = tmp_path / "13_any.mp3"
    file_path.write_text("x")
    plan = build_rule_plan([file_path], {"13": "13_m_up"})
    assert bool(plan.iloc[0]["是否命中规则"]) is True
    assert plan.iloc[0]["输出文件名"] == "13_m_up.mp3"


def test_rule_miss_uses_cleaned_stem(tmp_path):
    file_path = tmp_path / "99 bad?.mp3"
    file_path.write_text("x")
    plan = build_rule_plan([file_path], {"13": "13_m_up"})
    assert bool(plan.iloc[0]["是否命中规则"]) is False
    assert plan.iloc[0]["输出文件名"] == "99 bad_.mp3"


def test_rule_collision(tmp_path):
    a = tmp_path / "13_a.mp3"
    b = tmp_path / "13_b.mp3"
    a.write_text("a")
    b.write_text("b")
    plan = build_rule_plan([a, b], {"13": "same"})
    assert not validate_rule_plan(plan)["ok"]
