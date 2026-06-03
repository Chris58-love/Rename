from universal_rename_tool.validation import validate_rename_plan


def test_final_name_collision(plan_df_factory):
    plan = plan_df_factory(["A.txt", "a.TXT"])
    result = validate_rename_plan(plan)
    assert not result["ok"]


def test_case_insensitive_collision(plan_df_factory):
    plan = plan_df_factory(["Demo.mp3", "demo.MP3"])
    result = validate_rename_plan(plan)
    assert "最终文件名冲突" in set(result["errors"]["问题类型"])


def test_path_safety(plan_df_factory):
    plan = plan_df_factory(["../bad.txt"])
    result = validate_rename_plan(plan)
    assert not result["ok"]
    assert "路径不安全" in set(result["errors"]["问题类型"])
