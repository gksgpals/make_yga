import os
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import app


class AppBootstrapTest(unittest.TestCase):
    def test_direct_python_run_bootstraps_streamlit(self) -> None:
        with patch.object(app.st.runtime, "exists", return_value=False), patch.object(
            app.os,
            "execvp",
        ) as execvp_mock:
            app.ensure_streamlit_runtime()

        execvp_mock.assert_called_once_with(
            sys.executable,
            [
                sys.executable,
                "-m",
                "streamlit",
                "run",
                str(Path(app.__file__).resolve()),
            ],
        )

    def test_bootstrap_is_skipped_inside_streamlit_runtime(self) -> None:
        with patch.object(app.st.runtime, "exists", return_value=True), patch.object(
            app.os,
            "execvp",
        ) as execvp_mock:
            app.ensure_streamlit_runtime()

        execvp_mock.assert_not_called()

    def test_bootstrap_can_be_disabled_for_non_web_runs(self) -> None:
        with patch.dict(os.environ, {app.DIRECT_RUN_BOOTSTRAP_BYPASS_ENV: "1"}), patch.object(
            app.st.runtime,
            "exists",
            return_value=False,
        ), patch.object(app.os, "execvp") as execvp_mock:
            app.ensure_streamlit_runtime()

        execvp_mock.assert_not_called()


class AppInputFlowTest(unittest.TestCase):
    def test_normalize_input_text_trims_whitespace(self) -> None:
        self.assertEqual(app.normalize_input_text("  hello\n"), "hello")

    def test_raw_text_has_pending_changes_uses_normalized_text(self) -> None:
        self.assertFalse(app.raw_text_has_pending_changes("hello  ", " hello"))
        self.assertTrue(app.raw_text_has_pending_changes("hello world", "hello"))

    def test_parse_input_payload_exposes_parser_status(self) -> None:
        fake_problem = app.Problem([], ["body"], ["question"], [], "01")
        fake_parse_result = SimpleNamespace(
            problems=[fake_problem],
            parser_name="hybrid-ai",
            ai_attempted=True,
            ai_used=True,
            reason="ai_type_primary",
        )
        with patch.object(app, "parse_raw_details", return_value=fake_parse_result), patch.object(
            app,
            "paginate_problems",
            return_value=[object(), object()],
        ):
            payload = app.parse_input_payload("sample text")

        self.assertIsNotNone(payload)
        assert payload is not None
        self.assertEqual(payload.parser_name, "hybrid-ai")
        self.assertTrue(payload.ai_attempted)
        self.assertTrue(payload.ai_used)
        self.assertEqual(payload.parse_reason, "ai_type_primary")

    def test_generated_output_is_stale_when_output_signature_changes(self) -> None:
        with patch.object(app, "current_header_title", return_value=""), patch.object(
            app,
            "current_content_font_size",
            return_value=app.DEFAULT_CONTENT_FONT_SIZE_PT,
        ), patch.object(app, "current_output_signature", return_value="new-signature"):
            app.st.session_state["last_generated_text"] = "sample text"
            app.st.session_state["last_generated_header"] = ""
            app.st.session_state["last_generated_font_size_pt"] = app.DEFAULT_CONTENT_FONT_SIZE_PT
            app.st.session_state["last_generated_output_signature"] = "old-signature"

            self.assertTrue(app.generated_output_is_stale("sample text"))


if __name__ == "__main__":
    unittest.main()
