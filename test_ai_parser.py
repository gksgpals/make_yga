import json
import os
import tempfile
import unittest
from pathlib import Path
from typing import Optional
from unittest.mock import patch

import ai_parser


class _FakeResponse:
    id = "resp_test_123"
    output_text = '{"problem_type_code":6,"problem_type_name":"unrelated_sentence","reason":"numbered sentences stay inside the passage flow"}'

    def model_dump_json(self, indent=None) -> str:
        return json.dumps(
            {
                "id": self.id,
                "output_text": self.output_text,
                "type": "response",
            },
            ensure_ascii=False,
        )


class _FakeResponses:
    last_kwargs: Optional[dict[str, object]] = None

    def create(self, **kwargs):
        self.last_kwargs = kwargs
        return _FakeResponse()


class _FakeClient:
    def __init__(self, api_key: str) -> None:
        self.api_key = api_key
        self.responses = _FakeResponses()


class AiParserLoggingTest(unittest.TestCase):
    def test_ai_parser_logs_raw_request_and_response(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            env_path = temp_root / ".env"
            env_path.write_text(
                "OPENAI_API_KEY=test-key\nYGA_ENABLE_AI_PARSER=1\n",
                encoding="utf-8",
            )

            with patch.dict(
                os.environ,
                {
                    "YGA_ENABLE_AI_PARSER": "1",
                    "OPENAI_API_KEY": "test-key",
                },
                clear=False,
            ), patch.object(ai_parser, "OpenAI", _FakeClient):
                result = ai_parser.classify_problem_type_with_ai("raw input sample", base_dir=temp_root)

            self.assertIsNotNone(result)
            self.assertEqual(result["problem_type_code"], 6)
            self.assertEqual(result["problem_type_name"], "unrelated_sentence")

            log_path = temp_root / "logs" / "ai_parser_raw.jsonl"
            self.assertTrue(log_path.exists())
            call_capture_dir = temp_root / "logs" / "ai_calls"
            self.assertTrue(call_capture_dir.exists())

            entries = [
                json.loads(line)
                for line in log_path.read_text(encoding="utf-8").splitlines()
                if line.strip()
            ]
            self.assertEqual([entry["event"] for entry in entries], ["ai_type_request", "ai_type_response"])
            self.assertEqual(entries[0]["input_text"], "raw input sample")
            self.assertEqual(entries[0]["model"], "gpt-4.1-mini")
            self.assertIn('"problem_type_code":6', entries[1]["output_text"])
            self.assertIn('"type": "response"', entries[1]["response_raw_json"])
            self.assertIn('"id": "resp_test_123"', entries[1]["response_raw_json"])

            capture_files = sorted(call_capture_dir.glob("*.txt"))
            self.assertEqual(len(capture_files), 1)
            capture_text = capture_files[0].read_text(encoding="utf-8")
            self.assertIn("[input]\nraw input sample", capture_text)
            self.assertIn("[output]", capture_text)
            self.assertIn('"problem_type_code":6', capture_text)

    def test_ai_parser_enables_platform_storage(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            env_path = temp_root / ".env"
            env_path.write_text(
                "OPENAI_API_KEY=test-key\nYGA_ENABLE_AI_PARSER=1\n",
                encoding="utf-8",
            )

            fake_responses = _FakeResponses()

            class _CapturingClient:
                def __init__(self, api_key: str) -> None:
                    self.api_key = api_key
                    self.responses = fake_responses

            with patch.dict(
                os.environ,
                {
                    "YGA_ENABLE_AI_PARSER": "1",
                    "OPENAI_API_KEY": "test-key",
                },
                clear=False,
            ), patch.object(ai_parser, "OpenAI", _CapturingClient):
                ai_parser.classify_problem_type_with_ai("raw input sample", base_dir=temp_root)

            assert fake_responses.last_kwargs is not None
            self.assertTrue(fake_responses.last_kwargs["store"])


if __name__ == "__main__":
    unittest.main()
