import os
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
from xml.etree import ElementTree as ET

os.environ.setdefault("YGA_ENABLE_AI_PARSER", "0")

import slide_formatter
from slide_formatter import (
    NSMAP,
    Problem,
    build_presentation_files,
    build_slide_xml,
    find_pptx_template,
    paginate_problem,
    paginate_problems,
    parse_raw,
    parse_raw_details,
)


class SlideFormatterRulesTest(unittest.TestCase):
    def test_choice_spacing_is_normalized(self) -> None:
        raw = """문제. 어법상 틀린 것은?
①apologize
② provide
③study
"""
        problems = parse_raw(raw)
        self.assertEqual(len(problems), 1)
        self.assertEqual(
            problems[0].choice_lines,
            ["① apologize", "② provide", "③ study"],
        )

    def test_note_definition_moves_to_passage_end(self) -> None:
        raw = """[PASSAGE]
This is a body line. * term: 의미
Another body line.
[QUESTION]
What is correct?
[A] one
"""
        problems = parse_raw(raw)
        self.assertEqual(len(problems), 1)
        body = problems[0].body_lines
        self.assertIn("This is a body line. Another body line.", body)
        self.assertEqual(body[-1], "* term: 의미")

    def test_starred_word_at_line_start_stays_in_body_text(self) -> None:
        raw = """(B)
Elaine could no longer hold it in. She gently put the cat in a
*kennel and brought it to the back isolation area. Then she went
to the back room.
* kennel: 이동 장, 이동식 케이지
"""
        problems = parse_raw(raw)
        self.assertEqual(len(problems), 1)
        self.assertEqual(
            problems[0].body_lines,
            [
                "(B)",
                "Elaine could no longer hold it in. She gently put the cat in a *kennel and brought it to the back isolation area. Then she went to the back room.",
                "",
                "* kennel: 이동 장, 이동식 케이지",
            ],
        )

    def test_note_font_size_stays_fixed_when_content_font_changes(self) -> None:
        original_font_size = slide_formatter.FONT_SIZE_PT
        try:
            slide_formatter.set_content_font_size(52)
            self.assertEqual(slide_formatter.FONT_SIZE_PT, 52)
            self.assertEqual(
                slide_formatter.NOTE_FONT_PT,
                slide_formatter.DEFAULT_NOTE_FONT_SIZE_PT,
            )
        finally:
            slide_formatter.set_content_font_size(original_font_size)

    def test_body_soft_wrap_lines_are_reflowed(self) -> None:
        raw = """Passage first line broken
across the next line and
continues here.
"""
        problems = parse_raw(raw)
        self.assertEqual(len(problems), 1)
        self.assertEqual(
            problems[0].body_lines,
            ["Passage first line broken across the next line and continues here."],
        )

    def test_choice_soft_wrap_lines_are_reflowed(self) -> None:
        raw = """문제. 옳은 것은?
① first choice is broken
across the next line
② second choice
"""
        problems = parse_raw(raw)
        self.assertEqual(len(problems), 1)
        self.assertEqual(
            problems[0].choice_lines,
            [
                "① first choice is broken across the next line",
                "② second choice",
            ],
        )

    def test_unrelated_sentence_soft_wrap_keeps_embedded_choices_separate(self) -> None:
        raw = """01. 다음 글에서 전체 흐름과 관계 없는 문장은?
Exercise emotions affect participation
and performance.
① Anger cues can increase
strike intensity.
② Positive emotions can support yoga focus.
"""
        problems = parse_raw(raw)
        self.assertEqual(len(problems), 1)
        self.assertEqual(problems[0].choice_lines, [])
        self.assertEqual(
            problems[0].body_lines,
            [
                "01. 다음 글에서 전체 흐름과 관계 없는 문장은?",
                "Exercise emotions affect participation and performance. ① Anger cues can increase strike intensity. ② Positive emotions can support yoga focus.",
            ],
        )

    def test_unrelated_sentence_type_uses_question_first_layout(self) -> None:
        raw = """01. 다음 글에서 전체 흐름과 관계 없는 문장은?
Exercise emotions affect participation and performance.
① Anger cues can increase strike intensity.
② Positive emotions can support yoga focus.
"""
        problems = parse_raw(raw)
        self.assertEqual(len(problems), 1)
        self.assertEqual(problems[0].question_lines, [])
        self.assertEqual(problems[0].choice_lines, [])
        self.assertTrue(problems[0].body_lines[0].startswith("01."))
        self.assertIn("① Anger cues can increase strike intensity.", problems[0].body_lines[1])

    def test_long_reading_multi_question_is_split_into_separate_problems(self) -> None:
        raw = """A short passage line.
01. Which is correct?
①aa
②bb
02. Which is not correct?
①cc
②dd
"""
        problems = parse_raw(raw)
        self.assertEqual(len(problems), 2)
        self.assertEqual(problems[0].body_lines, ["A short passage line."])
        self.assertEqual(problems[0].question_lines, ["01. Which is correct?"])
        self.assertEqual(problems[0].choice_lines, ["① aa", "② bb"])
        self.assertEqual(problems[1].question_lines, ["02. Which is not correct?"])
        self.assertEqual(problems[1].choice_lines, ["① cc", "② dd"])

    def test_paginate_problems_assigns_sequential_numbers_ignoring_input_numbers(self) -> None:
        raw = """01. First question?
① one
② two


01. Second question?
① three
② four
"""
        problems = parse_raw(raw)

        paginated = paginate_problems(problems)

        self.assertEqual([problem.problem_number for problem in paginated], ["1", "2"])
        self.assertEqual(paginated[0].question_lines[0], "1. First question?")
        self.assertEqual(paginated[1].question_lines[0], "2. Second question?")

    def test_paginate_problems_replaces_question_first_body_prompt_number(self) -> None:
        raw = """01. 다음 글에서 전체 흐름과 관계 없는 문장은?
Body sentence one.
① choice one.
② choice two.


01. 글의 흐름으로 보아, 주어진 문장이 들어가기에 가장 적절한 곳은?
Passage line.
( ① ) marker one.
"""
        problems = parse_raw(raw)

        paginated = paginate_problems(problems)

        rendered_text = "\n".join(
            line
            for problem in paginated
            for line in (problem.body_lines + problem.question_lines)
        )
        self.assertIn("1. 다음 글에서 전체 흐름과 관계 없는 문장은?", rendered_text)
        self.assertIn("2. 글의 흐름으로 보아, 주어진 문장이 들어가기에 가장 적절한 곳은?", rendered_text)

    def test_unrelated_sentence_type_can_use_ai_classification(self) -> None:
        raw = """01. 다음 글에서 전체 흐름과 관계 없는 문장은?
Exercise emotions affect participation and performance.
① Anger cues can increase strike intensity.
② Positive emotions can support yoga focus.
"""
        ai_result = {
            "problem_type_code": slide_formatter.PROBLEM_TYPE_UNRELATED_SENTENCE,
            "problem_type_name": "unrelated_sentence",
            "reason": "numbered sentences are part of the passage flow",
        }

        with patch.dict(os.environ, {"YGA_ENABLE_AI_PARSER": "1"}), patch.object(
            slide_formatter,
            "classify_problem_type_with_ai",
            return_value=ai_result,
        ):
            parse_result = parse_raw_details(raw)

        problems = parse_result.problems
        self.assertEqual(len(problems), 1)
        self.assertEqual(parse_result.parser_name, "hybrid-ai")
        self.assertTrue(parse_result.ai_attempted)
        self.assertTrue(parse_result.ai_used)
        self.assertEqual(parse_result.reason, "ai_type_primary")
        self.assertEqual(problems[0].question_lines, [])
        self.assertEqual(problems[0].choice_lines, [])
        self.assertEqual(
            problems[0].body_lines,
            [
                "01. 다음 글에서 전체 흐름과 관계 없는 문장은?",
                "Exercise emotions affect participation and performance. ① Anger cues can increase strike intensity. ② Positive emotions can support yoga focus.",
            ],
        )

    def test_ai_classification_runs_per_problem(self) -> None:
        raw = """01. What is correct?
① one
② two


02. What is not correct?
① three
② four
"""
        ai_result = {
            "problem_type_code": slide_formatter.PROBLEM_TYPE_MULTIPLE_CHOICE,
            "problem_type_name": "multiple_choice",
            "reason": "ordinary multiple-choice question",
        }

        with patch.dict(os.environ, {"YGA_ENABLE_AI_PARSER": "1"}), patch.object(
            slide_formatter,
            "classify_problem_type_with_ai",
            side_effect=[ai_result, ai_result],
        ) as ai_mock:
            parse_result = parse_raw_details(raw)

        self.assertEqual(ai_mock.call_count, 2)
        self.assertIn("01. What is correct?", ai_mock.call_args_list[0].args[0])
        self.assertNotIn("02. What is not correct?", ai_mock.call_args_list[0].args[0])
        self.assertIn("02. What is not correct?", ai_mock.call_args_list[1].args[0])
        self.assertEqual(parse_result.reason, "ai_type_primary")
        self.assertEqual(len(parse_result.problems), 2)

    def test_sentence_insertion_ai_classification_keeps_markers_in_body(self) -> None:
        raw = """01. 글의 흐름으로 보아, 주어진 문장이 들어가기에 가장 적절한 곳은?
For many people though, we spend far too much time on the other end of the scale.
( ① ) These efforts are hard to replicate.
( ② ) This style of working is not cognitively demanding.
"""
        ai_result = {
            "problem_type_code": slide_formatter.PROBLEM_TYPE_SENTENCE_INSERTION,
            "problem_type_name": "sentence_insertion",
            "reason": "insertion markers are embedded in the passage flow",
        }

        with patch.dict(os.environ, {"YGA_ENABLE_AI_PARSER": "1"}), patch.object(
            slide_formatter,
            "classify_problem_type_with_ai",
            return_value=ai_result,
        ):
            parse_result = parse_raw_details(raw)

        self.assertEqual(parse_result.reason, "ai_type_primary")
        self.assertEqual(parse_result.problems[0].question_lines, [])
        self.assertEqual(parse_result.problems[0].choice_lines, [])
        self.assertEqual(
            parse_result.problems[0].body_lines,
            [
                "01. 글의 흐름으로 보아, 주어진 문장이 들어가기에 가장 적절한 곳은?",
                "For many people though, we spend far too much time on the other end of the scale. ( ① ) These efforts are hard to replicate. ( ② ) This style of working is not cognitively demanding.",
            ],
        )

    def test_blank_problem_ai_classification_keeps_choices_separate(self) -> None:
        raw = """01. 다음 빈칸에 들어갈 말로 가장 적절한 것은?
Deep work is the ability to focus without distraction.
① focus
② noise
"""
        ai_result = {
            "problem_type_code": slide_formatter.PROBLEM_TYPE_BLANK,
            "problem_type_name": "blank",
            "reason": "fill-in-the-blank question",
        }

        with patch.dict(os.environ, {"YGA_ENABLE_AI_PARSER": "1"}), patch.object(
            slide_formatter,
            "classify_problem_type_with_ai",
            return_value=ai_result,
        ):
            parse_result = parse_raw_details(raw)

        self.assertEqual(parse_result.problems[0].question_lines, ["01. 다음 빈칸에 들어갈 말로 가장 적절한 것은?"])
        self.assertEqual(parse_result.problems[0].body_lines, ["Deep work is the ability to focus without distraction."])
        self.assertEqual(parse_result.problems[0].choice_lines, ["① focus", "② noise"])

    def test_shared_passage_rule_split_and_ai_classification_keep_body_on_later_questions(self) -> None:
        raw = """[01 ~02] 다음 글을 읽고, 물음에 답하시오.
(A)
Shared passage line.

01. 첫 번째 질문은?
① one
② two

02. 두 번째 질문은?
① three
② four
"""
        ai_result = {
            "problem_type_code": slide_formatter.PROBLEM_TYPE_MULTIPLE_CHOICE,
            "problem_type_name": "multiple_choice",
            "reason": "ordinary shared-passage multiple-choice question",
        }

        with patch.dict(os.environ, {"YGA_ENABLE_AI_PARSER": "1"}), patch.object(
            slide_formatter,
            "classify_problem_type_with_ai",
            side_effect=[ai_result, ai_result],
        ):
            parse_result = parse_raw_details(raw)

        self.assertEqual(len(parse_result.problems), 2)
        self.assertEqual(parse_result.reason, "ai_type_primary")
        self.assertEqual(parse_result.problems[0].body_lines, ["(A)", "Shared passage line."])
        self.assertEqual(parse_result.problems[1].body_lines, ["(A)", "Shared passage line."])
        self.assertEqual(
            parse_result.problems[1].choice_lines,
            ["① three", "② four"],
        )

    def test_find_pptx_template_falls_back_to_any_pptx_in_project(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            fallback_pptx = temp_root / "fallback_sample.pptx"
            fallback_pptx.write_bytes(b"placeholder")

            fake_module_path = temp_root / "slide_formatter.py"
            with patch.object(slide_formatter, "__file__", str(fake_module_path)):
                resolved = find_pptx_template()

        self.assertEqual(resolved.resolve(), fallback_pptx.resolve())

    def test_build_presentation_files_uses_fallback_project_pptx(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_template = Path("test_input_latest.pptx")
            fallback_pptx = temp_root / "fallback_sample.pptx"
            shutil.copyfile(source_template, fallback_pptx)

            fake_module_path = temp_root / "slide_formatter.py"
            out_path = temp_root / "generated_output.pptx"
            problems = parse_raw("문제. What is correct?\n① one\n② two\n")

            with patch.object(slide_formatter, "__file__", str(fake_module_path)):
                built_pptx, built_pdf = build_presentation_files(problems, out_path, None)

            self.assertEqual(built_pptx, out_path)
            self.assertIsNone(built_pdf)
            self.assertTrue(out_path.exists())
            with zipfile.ZipFile(out_path) as zf:
                names = set(zf.namelist())
            self.assertIn("ppt/slides/slide1.xml", names)
            self.assertIn("ppt/presentation.xml", names)

    def test_slide_xml_uses_justified_alignment_for_content(self) -> None:
        slide = build_slide_xml(
            Problem(
                header_lines=["1-1 테스트"],
                body_lines=[
                    "This is a long body line that should use justified alignment in the generated slide XML."
                ],
                question_lines=[],
                choice_lines=[],
                problem_number="1",
            ),
            1,
        )
        root = slide.getroot()
        paragraph_alignments = [
            node.get("algn")
            for node in root.findall(".//a:p/a:pPr", NSMAP)
            if node.get("algn")
        ]
        self.assertIn("just", paragraph_alignments)

    def test_paginate_problem_packs_question_and_choices_into_last_body_page(self) -> None:
        problem = Problem(
            header_lines=["1-1 테스트"],
            body_lines=[f"body line {idx}" for idx in range(1, 9)],
            question_lines=["question line"],
            choice_lines=["① one", "② two"],
            problem_number="1",
        )

        pages = paginate_problem(problem)

        self.assertEqual(len(pages), 1)
        self.assertEqual(len(pages[0].body_lines), 8)
        self.assertEqual(pages[0].question_lines, ["question line"])
        self.assertEqual(pages[0].choice_lines, ["① one", "② two"])

    def test_build_slide_xml_renders_mixed_body_question_choice_shapes(self) -> None:
        slide = build_slide_xml(
            Problem(
                header_lines=["1-1 테스트"],
                body_lines=["body line 1", "body line 2"],
                question_lines=["question line"],
                choice_lines=["① one", "② two"],
                problem_number="1",
            ),
            1,
        )
        root = slide.getroot()
        shape_names = [
            node.get("name")
            for node in root.findall(".//p:cNvPr", NSMAP)
            if node.get("name")
        ]
        self.assertIn("body", shape_names)
        self.assertIn("question", shape_names)
        self.assertIn("choice", shape_names)

    def test_paginate_problem_keeps_choices_after_question_spillover(self) -> None:
        problem = Problem(
            header_lines=["1-1 테스트"],
            body_lines=[],
            question_lines=[
                f"question line {idx}"
                for idx in range(1, slide_formatter.BODY_MAX_LINES + 2)
            ],
            choice_lines=["① one", "② two"],
            problem_number="1",
        )

        pages = paginate_problem(problem)

        self.assertEqual(len(pages), 2)
        self.assertEqual(len(pages[0].question_lines), slide_formatter.BODY_MAX_LINES)
        self.assertEqual(pages[0].choice_lines, [])
        self.assertEqual(pages[1].question_lines, [f"question line {slide_formatter.BODY_MAX_LINES + 1}"])
        self.assertEqual(pages[1].choice_lines, ["① one", "② two"])

    def test_paginate_problem_splits_body_line_to_fill_remaining_space(self) -> None:
        prefix_lines = [
            f"body line {idx}"
            for idx in range(1, (slide_formatter.BODY_MAX_LINES - 2) + 1)
        ]
        long_line = "x" * (slide_formatter.BODY_CHARS_PER_LINE * 3)
        problem = Problem(
            header_lines=["1-1 테스트"],
            body_lines=prefix_lines + [long_line],
            question_lines=[],
            choice_lines=[],
            problem_number="1",
        )

        pages = paginate_problem(problem)

        self.assertEqual(len(pages), 2)
        self.assertEqual(len(pages[0].body_lines), len(prefix_lines) + 1)
        self.assertGreaterEqual(
            slide_formatter.estimate_visual_lines(
                pages[0].body_lines[-1],
                slide_formatter.BODY_CHARS_PER_LINE,
            ),
            1,
        )
        self.assertGreaterEqual(
            slide_formatter.estimate_visual_lines(
                pages[1].body_lines[0],
                slide_formatter.BODY_CHARS_PER_LINE,
            ),
            1,
        )

    def test_paginate_problem_splits_boundary_paragraph_before_overflow(self) -> None:
        raw = """(A)
The manager of the Humane Society was out of town for two weeks, and Elaine was placed in charge. Since she had taken on this responsibility before, she was familiar with the challenges she might encounter — *impounding **stray animals, reuniting pets with their owners, informing people their pet hadn’t been found, accepting animals from those who could no longer care for them, or turning animals away when the shelter was full. Though she found the work demanding, (a) she handled these situations with confidence and calm.

(B)
Elaine could no longer hold it in. She gently put the cat in a *kennel and brought it to the back isolation area. Then she went to the back room and called the absent manager, who encouraged her to let it all out. Elaine shared what happened and released (b) her emotions while the manager listened. She felt much better afterward and was relieved to have
"""
        problems = parse_raw(raw)

        self.assertEqual(len(problems), 1)
        pages = paginate_problem(problems[0])

        self.assertEqual(len(pages), 2)
        self.assertEqual(pages[0].body_lines[:4], ["(A)", problems[0].body_lines[1], "", "(B)"])
        self.assertTrue(pages[1].body_lines)
        self.assertLessEqual(
            sum(
                slide_formatter.estimate_visual_lines(line, slide_formatter.BODY_CHARS_PER_LINE)
                for line in pages[0].body_lines
            ),
            slide_formatter.BODY_MAX_LINES - slide_formatter.BODY_VISUAL_LINE_SAFETY_MARGIN,
        )

    def test_paginate_problems_keeps_same_number_for_split_pages(self) -> None:
        repeated_line = (
            "This is a long body line that should keep wrapping across pages when repeated. "
            * 10
        ).strip()
        raw = f"""01. First question?
{repeated_line}
{repeated_line}
{repeated_line}
{repeated_line}
"""
        problems = parse_raw(raw)

        paginated = paginate_problems(problems)

        self.assertGreaterEqual(len(paginated), 2)
        self.assertTrue(all(problem.problem_number == "1" for problem in paginated))


if __name__ == "__main__":
    unittest.main()
