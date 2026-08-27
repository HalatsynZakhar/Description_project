from mono_processor import load_prompt


def test_prompt_normalizes_ordinary_caps_from_source_description_for_title() -> None:
    prompt = load_prompt()
    assert "CAPS у звичайних словах джерела" in prompt
    assert "оформлення, а не правильний регістр" in prompt


def test_prompt_explicitly_forbids_all_caps_in_ordinary_title_words() -> None:
    assert "повністю CAPS LOCK" in load_prompt()


def test_runtime_prompt_includes_restructured_mono_rules() -> None:
    prompt = load_prompt()
    assert "# Правила MONO для обробки картки товару" in prompt
    assert "Дозволена HTML-розмітка" in prompt
