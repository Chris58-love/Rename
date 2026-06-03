from universal_rename_tool.rename_state import create_initial_rename_state
from universal_rename_tool.utils import sanitize_filename_component, split_filename, strip_same_extension


def test_illegal_characters_are_cleaned():
    assert sanitize_filename_component('a/b:c*?"<>|') == "a_b_c______"


def test_windows_reserved_name_gets_suffix():
    assert sanitize_filename_component("CON") == "CON_"


def test_no_extension_split():
    assert split_filename("README") == ("README", "", False)


def test_user_filled_same_extension_is_removed():
    assert strip_same_extension("report.xlsx", ".xlsx") == "report"


def test_initial_state_uses_clean_stem():
    inventory = [{"original_name": "a/b.txt", "stem": "a/b", "extension": ".txt"}]
    state = create_initial_rename_state(inventory)
    assert state["a/b.txt"]["new_stem"] == "a_b"
