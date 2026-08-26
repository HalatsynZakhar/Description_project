"""Швидка перевірка всіх увімкнених ключів Gemini з keys.json."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from google import genai
from google.genai import types

from mono_processor import HTTP_TIMEOUT_MS, KEYS_PATH, MODEL_NAME, classify_api_error


def configured_keys(path: Path) -> list[tuple[str, str]]:
    try:
        data: dict[str, Any] = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        print(f"Не знайдено {path.name}. Створіть його за прикладом keys.example.json.")
        return []
    except (OSError, json.JSONDecodeError) as error:
        print(f"Не вдалося прочитати {path.name}: {error}")
        return []

    result: list[tuple[str, str]] = []
    for number, item in enumerate(data.get("keys", []), start=1):
        if not isinstance(item, dict) or item.get("enabled", True) is False:
            continue
        name = str(item.get("name", f"key-{number}")).strip()
        api_key = str(item.get("api_key", "")).strip()
        if name and api_key and not api_key.startswith("ВСТАВТЕ_"):
            result.append((name, api_key))
    return result


def main() -> int:
    keys = configured_keys(KEYS_PATH)
    if not keys:
        print("Активних ключів для перевірки немає.")
        return 1

    print(f"Модель: {MODEL_NAME}. Перевіряю ключів: {len(keys)}.\n")
    successful = 0
    for name, api_key in keys:
        print(f"Перевірка «{name}»… ", end="", flush=True)
        try:
            client = genai.Client(
                api_key=api_key,
                http_options=types.HttpOptions(
                    timeout=HTTP_TIMEOUT_MS,
                    retryOptions=types.HttpRetryOptions(attempts=1),
                ),
            )
            response = client.models.generate_content(
                model=MODEL_NAME,
                contents="Відповідай рівно одним словом українською: Привіт",
            )
            answer = (getattr(response, "text", "") or "").strip().replace("\n", " ")
            if not answer:
                raise RuntimeError("API повернув порожню відповідь")
        except Exception as error:  # SDK має різні класи винятків для різних HTTP-помилок.
            category, _ = classify_api_error(error)
            print(f"ПОМИЛКА [{category}]: {type(error).__name__}: {error}")
            continue

        successful += 1
        print(f"OK — {answer[:100]}")

    print(f"\nПрацює ключів: {successful}/{len(keys)}")
    return 0 if successful == len(keys) else 1


if __name__ == "__main__":
    sys.exit(main())
