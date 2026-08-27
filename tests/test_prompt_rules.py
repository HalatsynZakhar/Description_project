from mono_processor import load_prompt


def test_prompt_normalizes_ordinary_caps_from_source_description_for_title() -> None:
    prompt = load_prompt()
    assert "вважай це лише оформленням" in prompt
    assert "приведи їх до природного регістру" in prompt
