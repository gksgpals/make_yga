from __future__ import annotations

import importlib
import json
import os
import uuid
from pathlib import Path
from typing import Any

try:
    from openai import OpenAI
except ImportError:  # pragma: no cover - optional dependency fallback
    OpenAI = None  # type: ignore[assignment]

_runtime_logging: Any = importlib.import_module("runtime_logging")
ai_call_capture_path = _runtime_logging.ai_call_capture_path
ai_raw_log_path = _runtime_logging.ai_raw_log_path
append_json_log = _runtime_logging.append_json_log
get_logger = _runtime_logging.get_logger
write_text_log = _runtime_logging.write_text_log

LOGGER = get_logger("ai_parser")

OPENAI_API_KEY_ENV = "OPENAI_API_KEY"
AI_ENABLE_ENV = "YGA_ENABLE_AI_PARSER"
AI_MODEL_ENV = "YGA_AI_MODEL"
DEFAULT_AI_MODEL = "gpt-4.1-mini"
FALSEY_VALUES = {"0", "false", "no", "off"}

PROBLEM_TYPE_OPTIONS: list[tuple[int, str, str]] = [
    (0, "unknown", "Does not clearly match any registered type. Use this when a new type should be added."),
    (1, "multiple_choice", "Ordinary multiple-choice question with prompt and choices separated."),
    (2, "blank", "Fill-in-the-blank question."),
    (3, "grammar", "Grammar or usage question asking which expression is correct or incorrect."),
    (4, "vocab", "Vocabulary, word meaning, or word usage question."),
    (5, "title_topic_gist", "Title, topic, main idea, summary, or purpose question."),
    (6, "unrelated_sentence", "Question asking which numbered sentence is unrelated to the flow; numbered lines stay inside the passage body."),
    (7, "sentence_insertion", "Question asking where a given sentence should be inserted; passage contains markers like ( ① )."),
    (8, "sentence_ordering", "Question asking for the correct order of passage blocks such as (A), (B), (C)."),
    (9, "reference", "Question asking what an underlined word or expression refers to."),
]
PROBLEM_TYPE_CODE_TO_NAME = {code: name for code, name, _description in PROBLEM_TYPE_OPTIONS}
PROBLEM_TYPE_CODE_TO_DESCRIPTION = {
    code: description for code, _name, description in PROBLEM_TYPE_OPTIONS
}


def problem_type_options_text() -> str:
    return "\n".join(
        f"- {code} = {name}: {description}" for code, name, description in PROBLEM_TYPE_OPTIONS
    )


AI_CLASSIFY_INSTRUCTIONS = f"""
You classify one already-parsed workbook problem into exactly one registered problem type code.

Return JSON only. Do not rewrite, segment, summarize, or explain the problem text.
Choose exactly one code from the following list:
{problem_type_options_text()}

Rules:
- The input is already one problem. Do not try to split it further.
- For unrelated sentence problems, choose 6 when the numbered sentences belong inside the passage flow itself.
- For sentence insertion problems, choose 7 when the passage contains insertion markers like ( ① ).
- For sentence ordering problems, choose 8 when the choices represent the order of passage blocks such as (A), (B), (C).
- For reference questions, choose 9 when the prompt asks what an underlined word/expression or marked item refers to.
- If the problem does not clearly fit any registered type, return 0.
""".strip()

AI_CLASSIFY_SCHEMA: dict[str, object] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "problem_type_code": {
            "type": "integer",
            "enum": [code for code, _name, _description in PROBLEM_TYPE_OPTIONS],
        },
        "problem_type_name": {"type": "string"},
        "reason": {"type": "string"},
    },
    "required": ["problem_type_code", "problem_type_name", "reason"],
}


def load_env_value(name: str, base_dir: Path | None = None) -> str | None:
    env_value = os.getenv(name)
    if env_value:
        return env_value.strip()

    search_root = base_dir or Path(__file__).resolve().parent
    env_path = search_root / ".env"
    if not env_path.exists():
        return None

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        if key.strip() != name:
            continue
        return value.strip().strip("\"").strip("'")
    return None


def ai_parser_enabled(base_dir: Path | None = None) -> bool:
    flag = load_env_value(AI_ENABLE_ENV, base_dir)
    if flag is not None:
        return flag.strip().lower() not in FALSEY_VALUES
    return bool(load_env_value(OPENAI_API_KEY_ENV, base_dir))


def ai_parser_model(base_dir: Path | None = None) -> str:
    return load_env_value(AI_MODEL_ENV, base_dir) or DEFAULT_AI_MODEL


def response_raw_json(response: Any) -> str:
    model_dump_json = getattr(response, "model_dump_json", None)
    if callable(model_dump_json):
        try:
            raw_json = model_dump_json(indent=None)
            if isinstance(raw_json, str):
                return raw_json
        except Exception:
            pass

    model_dump = getattr(response, "model_dump", None)
    if callable(model_dump):
        try:
            return json.dumps(model_dump(mode="json"), ensure_ascii=False)
        except TypeError:
            try:
                return json.dumps(model_dump(), ensure_ascii=False)
            except Exception:
                pass
        except Exception:
            pass

    to_dict = getattr(response, "to_dict", None)
    if callable(to_dict):
        try:
            return json.dumps(to_dict(), ensure_ascii=False)
        except Exception:
            pass

    return json.dumps({"repr": repr(response)}, ensure_ascii=False)


def problem_type_name_from_code(problem_type_code: int) -> str:
    return PROBLEM_TYPE_CODE_TO_NAME.get(problem_type_code, PROBLEM_TYPE_CODE_TO_NAME[0])


def write_ai_call_capture(path: Path, *, input_text: str, output_text: str) -> None:
    write_text_log(
        path,
        "\n".join(
            [
                "[input]",
                input_text,
                "",
                "[output]",
                output_text,
            ]
        ),
    )


def classify_problem_type_with_ai(
    raw_text: str,
    *,
    base_dir: Path | None = None,
) -> dict[str, Any] | None:
    if not raw_text.strip() or not ai_parser_enabled(base_dir):
        return None

    if OpenAI is None:
        return None

    api_key = load_env_value(OPENAI_API_KEY_ENV, base_dir)
    if not api_key:
        return None

    model_name = ai_parser_model(base_dir)
    request_id = uuid.uuid4().hex
    raw_log_path = ai_raw_log_path(base_dir)
    call_capture_path = ai_call_capture_path(request_id, base_dir)
    write_ai_call_capture(call_capture_path, input_text=raw_text, output_text="")
    append_json_log(
        raw_log_path,
        {
            "event": "ai_type_request",
            "request_id": request_id,
            "model": model_name,
            "instructions": AI_CLASSIFY_INSTRUCTIONS,
            "input_text": raw_text,
        },
    )

    try:
        client = OpenAI(api_key=api_key)
        response = client.responses.create(
            model=model_name,
            input=raw_text,
            instructions=AI_CLASSIFY_INSTRUCTIONS,
            max_output_tokens=400,
            temperature=0,
            store=True,
            text={
                "format": {
                    "type": "json_schema",
                    "name": "problem_type_code",
                    "schema": AI_CLASSIFY_SCHEMA,
                    "strict": True,
                }
            },
        )
    except Exception as exc:  # pragma: no cover - network/API failure fallback
        append_json_log(
            raw_log_path,
            {
                "event": "ai_type_error",
                "request_id": request_id,
                "model": model_name,
                "input_text": raw_text,
                "error": str(exc),
            },
        )
        LOGGER.warning("AI type classification failed; falling back to rules parser: %s", exc)
        return None

    write_ai_call_capture(call_capture_path, input_text=raw_text, output_text=response.output_text or "")
    append_json_log(
        raw_log_path,
        {
            "event": "ai_type_response",
            "request_id": request_id,
            "model": model_name,
            "response_id": str(getattr(response, "id", "")),
            "output_text": response.output_text or "",
            "response_raw_json": response_raw_json(response),
        },
    )

    try:
        payload = json.loads(response.output_text or "")
    except json.JSONDecodeError as exc:  # pragma: no cover - schema should prevent this
        LOGGER.warning("AI type classifier returned invalid JSON; falling back to rules parser: %s", exc)
        return None

    problem_type_code = payload.get("problem_type_code")
    if not isinstance(problem_type_code, int):
        return None
    payload["problem_type_name"] = problem_type_name_from_code(problem_type_code)
    payload["reason"] = str(payload.get("reason", "")).strip()
    return payload
