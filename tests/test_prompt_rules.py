from mono_processor import load_prompt


def test_prompt_normalizes_ordinary_caps_from_source_description_for_title() -> None:
    prompt = load_prompt()
    assert "вважай це лише оформленням" in prompt
    assert "приведи їх до природного регістру" in prompt


def test_prompt_explicitly_forbids_all_caps_in_ordinary_title_words() -> None:
    assert "СУВОРА ЗАБОРОНА CAPS" in load_prompt()
