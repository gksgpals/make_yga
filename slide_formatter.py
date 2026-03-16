#!/usr/bin/env python3
"""Format English problem text into PPT-ready slide blocks.

Usage:
  python3 slide_formatter.py input.txt
  python3 slide_formatter.py input.txt -o output.txt
  python3 slide_formatter.py input.txt --pptx class_slides.pptx --pdf class_slides.pdf
  cat input.txt | python3 slide_formatter.py
"""

from __future__ import annotations

import argparse
import importlib
import math
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, List, Tuple
from xml.etree import ElementTree as ET

_ai_parser: Any = importlib.import_module("ai_parser")
ai_parser_enabled = _ai_parser.ai_parser_enabled
classify_problem_type_with_ai = _ai_parser.classify_problem_type_with_ai
_runtime_logging: Any = importlib.import_module("runtime_logging")
get_logger = _runtime_logging.get_logger

DEFAULT_FONT_SIZE_PT = 40
DEFAULT_NOTE_FONT_SIZE_PT = 30
DEFAULT_LINE_SPACING_PERCENT = 165


def build_style_block(font_size_pt: int, line_spacing_percent: int) -> str:
    return (
        "[스타일]\n"
        "배경: 검정색(#000000)\n"
        "글자색: 흰색(#FFFFFF)\n"
        "폰트: 굴림\n"
        f"글자크기: {font_size_pt}pt\n"
        f"줄간격: {line_spacing_percent}%"
    )

FONT_NAME = "굴림"
HEADER_FONT_NAME = "굴림"
FONT_SIZE_PT = DEFAULT_FONT_SIZE_PT
LINE_SPACING_PERCENT = DEFAULT_LINE_SPACING_PERCENT

LOGGER = get_logger("slide_formatter")

SLIDE_WIDTH = 1920
CONTENT_SIDE_MARGIN = 24
QUESTION_SIDE_MARGIN = 28
TEXTBOX_HORIZONTAL_INSET = "0"
TEXTBOX_VERTICAL_INSET = "25400"
CONTENT_TEXT_ALIGN = "just"
CONTENT_FLOW_TOP = 108
CONTENT_FLOW_GAP = 18
BODY_VISUAL_LINE_SAFETY_MARGIN = 1

TOP_LEFT_BOX = (14, 8, 220, 56)
TOP_RIGHT_BOX = (1180, 8, 720, 56)
TOP_RED_LINE_BOX = (0, 66, SLIDE_WIDTH, 8)
BODY_BOX = (CONTENT_SIDE_MARGIN, 108, SLIDE_WIDTH - (CONTENT_SIDE_MARGIN * 2), 920)
BODY_ONLY_BOX = BODY_BOX
QUESTION_BOX = (QUESTION_SIDE_MARGIN, 172, SLIDE_WIDTH - (QUESTION_SIDE_MARGIN * 2), 145)
CHOICE_BOX = (QUESTION_SIDE_MARGIN, 320, SLIDE_WIDTH - (QUESTION_SIDE_MARGIN * 2), 680)

TOP_LEFT_FONT_PT = 35
TOP_RIGHT_FONT_PT = 35
NOTE_FONT_PT = DEFAULT_NOTE_FONT_SIZE_PT
TOP_RIGHT_LABEL = "YGA 2026 KO Reading"
PREFERRED_TEMPLATE_NAMES = (
    "pptx_template.pptx",
    "my_class.pptx",
    "test_input_latest.pptx",
)


def calc_max_lines(box_height: int) -> int:
    line_height = FONT_SIZE_PT * (LINE_SPACING_PERCENT / 100.0)
    # Use most of the box while still reserving a small safety margin.
    return max(1, int((box_height * 0.99) // line_height))


def calc_chars_per_line(box_width: int) -> int:
    # Rough width estimate for mixed KR/EN text at current font size.
    # Bold Gulim runs slightly wider than the previous estimate on boundary cases.
    avg_char_width = FONT_SIZE_PT * 0.525
    return max(16, int((box_width * 0.995) // avg_char_width))


STYLE_BLOCK = build_style_block(FONT_SIZE_PT, LINE_SPACING_PERCENT)
BODY_MAX_LINES = calc_max_lines(BODY_BOX[3])
BODY_ONLY_MAX_LINES = calc_max_lines(BODY_ONLY_BOX[3])
QUESTION_MAX_LINES = calc_max_lines(QUESTION_BOX[3])
CHOICE_MAX_LINES = calc_max_lines(CHOICE_BOX[3])

BODY_CHARS_PER_LINE = calc_chars_per_line(BODY_BOX[2])
BODY_ONLY_CHARS_PER_LINE = calc_chars_per_line(BODY_ONLY_BOX[2])
QUESTION_CHARS_PER_LINE = calc_chars_per_line(QUESTION_BOX[2])
CHOICE_CHARS_PER_LINE = calc_chars_per_line(CHOICE_BOX[2])


def set_content_font_size(font_size_pt: int) -> None:
    global FONT_SIZE_PT, STYLE_BLOCK
    global BODY_MAX_LINES, BODY_ONLY_MAX_LINES, QUESTION_MAX_LINES, CHOICE_MAX_LINES
    global BODY_CHARS_PER_LINE, BODY_ONLY_CHARS_PER_LINE, QUESTION_CHARS_PER_LINE, CHOICE_CHARS_PER_LINE

    FONT_SIZE_PT = max(16, int(font_size_pt))
    STYLE_BLOCK = build_style_block(FONT_SIZE_PT, LINE_SPACING_PERCENT)
    BODY_MAX_LINES = calc_max_lines(BODY_BOX[3])
    BODY_ONLY_MAX_LINES = calc_max_lines(BODY_ONLY_BOX[3])
    QUESTION_MAX_LINES = calc_max_lines(QUESTION_BOX[3])
    CHOICE_MAX_LINES = calc_max_lines(CHOICE_BOX[3])
    BODY_CHARS_PER_LINE = calc_chars_per_line(BODY_BOX[2])
    BODY_ONLY_CHARS_PER_LINE = calc_chars_per_line(BODY_ONLY_BOX[2])
    QUESTION_CHARS_PER_LINE = calc_chars_per_line(QUESTION_BOX[2])
    CHOICE_CHARS_PER_LINE = calc_chars_per_line(CHOICE_BOX[2])

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NSMAP = {"a": NS_A, "p": NS_P}

EMU_PER_PT = 12700
UNDERLINE_OPEN = "{{__U__}}"
UNDERLINE_CLOSE = "{{__/U__}}"

HEADER_START_RE = re.compile(r"^\s*[0-9]+\s*[-.]\s*\S.*$")
QUESTION_MARKER_RE = re.compile(
    r"^(문제\.?|Q\.?|Question\.?|다음\s+글|물음|다음\s+중|주어진\s+글)",
    re.IGNORECASE,
)
QUESTION_NUMBERED_RE = re.compile(r"^\s*(\d+)\.\s*(.+)$")
PROBLEM_NUMBER_PREFIX_RE = re.compile(
    r"^(?P<lead>\s*)(?P<number>\d+)(?P<sep>\s*[.)-]?\s*)(?P<tail>.+)$"
)
LABELED_LINE_RE = re.compile(r"^\s*\[([^\]]+)\]\s*(.*)$")
COLON_LABEL_RE = re.compile(r"^\s*([A-Za-z가-힣]+)\s*[:：]\s*(.*)$")
CHOICE_LINE_RE = re.compile(
    r"^(?:(?P<circled>[①②③④⑤])|(?P<num>[1-5])[.)]|(?P<alpha>[A-Ea-e])[.)])\s*(?P<text>.*)$"
)
EMBEDDED_CHOICE_LINE_RE = re.compile(r"^\(\s*(?:[①②③④⑤]|[1-5]|[A-Ea-e])\s*\)\s*.*$")
CHOICE_PREFIX_LINE_RE = re.compile(
    r"^(?P<lead>\s*)(?P<prefix>(?:[①②③④⑤]|[1-5][.)]|[A-Ea-e][.)]))\s*(?P<text>.*)$"
)
NOTE_DEFINITION_MARKER_RE = re.compile(r"\*{1,3}\s*[^*:\n]{1,80}\s*[:：]")

CHOICE_NUMBER_TO_CIRCLED = {
    "1": "①",
    "2": "②",
    "3": "③",
    "4": "④",
    "5": "⑤",
}
CIRCLED_CHOICES = tuple(CHOICE_NUMBER_TO_CIRCLED.values())
META_NOISE_LINE_RE = re.compile(
    r"^\s*(?:"
    r"(?:정답|답)\s*[:：]?"
    r"|(?:해설|풀이|출처|자료|저작권)\s*[:：]?"
    r"|source\s*[:：]?"
    r"|copyright\s*[:：]?"
    r"|page\s*\d+"
    r"|p\.?\s*\d+"
    r"|YGA\s*\d{2,4}\s*KO\s*Reading"
    r"|[0-9]+\s*번"
    r")\s*$",
    re.IGNORECASE,
)

HEADER_LABELS = {
    "unit",
    "title",
    "header",
    "lesson",
    "part",
    "chapter",
    "상단표기",
    "제목",
    "단원",
}
BODY_LABELS = {"passage", "body", "text", "본문", "본문단", "지문"}
QUESTION_LABELS = {"q", "question", "prompt", "문제", "문제단", "질문"}
CHOICE_LABELS = {"choice", "choices", "option", "options", "선지", "선지단", "보기"}
ANSWER_LABELS = {"answer", "정답"}
EXPLANATION_LABELS = {"explanation", "explain", "해설", "풀이"}
PROBLEM_START_LABELS = {"problem", "slide", "슬라이드", "문항"}
QUESTION_HINT_KEYWORDS = {
    "밑줄",
    "빈칸",
    "어법",
    "문법",
    "제목",
    "주제",
    "요지",
    "적절",
    "옳은",
    "틀린",
    "의미",
    "고르시오",
    "고를",
}
QUESTION_FIRST_LAYOUT_KEYWORDS = {
    "무관한 문장",
    "관계 없는 문장",
    "흐름과 관계 없는 문장",
    "문단 속에",
    "문장 넣기",
}
PROBLEM_TYPE_UNKNOWN = 0
PROBLEM_TYPE_MULTIPLE_CHOICE = 1
PROBLEM_TYPE_BLANK = 2
PROBLEM_TYPE_GRAMMAR = 3
PROBLEM_TYPE_VOCAB = 4
PROBLEM_TYPE_TITLE_TOPIC_GIST = 5
PROBLEM_TYPE_UNRELATED_SENTENCE = 6
PROBLEM_TYPE_SENTENCE_INSERTION = 7
PROBLEM_TYPE_SENTENCE_ORDERING = 8
PROBLEM_TYPE_REFERENCE = 9
QUESTION_FIRST_TYPE_CODES = {
    PROBLEM_TYPE_UNRELATED_SENTENCE,
    PROBLEM_TYPE_SENTENCE_INSERTION,
}

@dataclass
class Problem:
    header_lines: List[str]
    body_lines: List[str]
    question_lines: List[str]
    choice_lines: List[str]
    problem_number: str = ""


@dataclass(frozen=True)
class ParseResult:
    problems: List[Problem]
    parser_name: str
    ai_attempted: bool
    ai_used: bool
    reason: str


# -------------------------------
# Parsing and normalization
# -------------------------------
def normalize_lines(text: str) -> List[str]:
    text = sanitize_xml_text(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return text.split("\n")


def sanitize_xml_text(text: str) -> str:
    out: List[str] = []
    for ch in text:
        cp = ord(ch)
        if cp in (0x9, 0xA, 0xD):
            out.append(ch)
            continue
        if 0x20 <= cp <= 0xD7FF or 0xE000 <= cp <= 0xFFFD or 0x10000 <= cp <= 0x10FFFF:
            out.append(ch)
    return "".join(out)


def split_problems(lines: List[str]) -> List[List[str]]:
    # Prefer explicit large blank-line separators when present.
    raw = "\n".join(lines)
    pieces = [p.strip("\n") for p in re.split(r"\n{3,}", raw) if p.strip()]
    if len(pieces) > 1:
        return [piece.split("\n") for piece in pieces]

    chunks: List[List[str]] = []
    current: List[str] = []
    current_has_header = False

    for line in lines:
        stripped = line.strip()
        is_header = is_header_line(stripped)
        if is_header:
            # Keep leading passage text together with the first detected header.
            if current and current_has_header:
                if any(token.strip() for token in current):
                    chunks.append(current)
                current = [line]
                current_has_header = True
                continue
            current_has_header = True
        current.append(line)

    if any(token.strip() for token in current):
        chunks.append(current)

    return chunks


def is_header_line(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return False
    if re.match(r"^\[\s*\d+\s*~\s*\d+\s*\]", stripped):
        return True
    if HEADER_START_RE.match(stripped):
        if looks_like_question_prompt(stripped):
            return False
        return True

    lowered = stripped.lower()
    english_prefixes = ("part", "lesson", "unit", "chapter")
    if any(lowered.startswith(prefix) for prefix in english_prefixes):
        return True

    korean_prefixes = ("유형", "기출", "수능", "독해", "문제집", "예제", "단원", "제목")
    return any(stripped.startswith(prefix) for prefix in korean_prefixes)


def normalize_label_key(label: str) -> str:
    return re.sub(r"[\s_\-]+", "", label.strip().lower())


def safe_slice_text(value: str, start: int, end: int | None = None) -> str:
    length = len(value)

    start_idx = start
    if start_idx < 0:
        start_idx = max(length + start_idx, 0)
    if start_idx > length:
        start_idx = length

    if end is None:
        end_idx = length
    else:
        end_idx = end
        if end_idx < 0:
            end_idx = max(length + end_idx, 0)
        if end_idx > length:
            end_idx = length

    if end_idx < start_idx:
        return ""

    chars: List[str] = []
    idx = start_idx
    while idx < end_idx:
        chars.append(value[idx])
        idx += 1
    return "".join(chars)


def copy_line_range(values: List[str], start: int, end: int) -> List[str]:
    copied: List[str] = []
    for idx, value in enumerate(values):
        if idx < start:
            continue
        if idx >= end:
            break
        copied.append(value)
    return copied


def remove_index(values: List[str], index_to_skip: int) -> List[str]:
    out: List[str] = []
    for idx, value in enumerate(values):
        if idx == index_to_skip:
            continue
        out.append(value)
    return out


def is_meta_noise_line(line: str) -> bool:
    return META_NOISE_LINE_RE.match(line.strip()) is not None


def looks_like_question_prompt(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return False

    if QUESTION_MARKER_RE.match(stripped):
        return True

    numbered = QUESTION_NUMBERED_RE.match(stripped)
    if numbered is None:
        return False

    tail = numbered.group(2)
    if "?" in tail:
        return True
    return any(keyword in tail for keyword in QUESTION_HINT_KEYWORDS)


def looks_like_choice_line(line: str) -> bool:
    return CHOICE_LINE_RE.match(line.strip()) is not None


def looks_like_embedded_choice_line(line: str) -> bool:
    return EMBEDDED_CHOICE_LINE_RE.match(line.strip()) is not None


def is_choice_group_start(line: str) -> bool:
    return looks_like_choice_line(line) or looks_like_embedded_choice_line(line)


def normalize_choice_spacing(line: str) -> str:
    stripped = line.rstrip()
    match = CHOICE_PREFIX_LINE_RE.match(stripped)
    if match is None:
        return stripped

    lead = match.group("lead") or ""
    prefix = match.group("prefix") or ""
    text = (match.group("text") or "").strip()
    if text:
        return f"{lead}{prefix} {text}"
    return f"{lead}{prefix}"


def normalize_choice_lines(lines: List[str]) -> List[str]:
    normalized: List[str] = []
    for line in lines:
        stripped = line.strip()
        if stripped and looks_like_choice_line(stripped):
            normalized.append(normalize_choice_spacing(line))
        else:
            normalized.append(line.rstrip())
    return normalized


def has_multiple_choice_markers(line: str) -> bool:
    markers = re.findall(r"[①②③④⑤]|[1-5][.)]|[A-Ea-e][.)]", line)
    return len(markers) >= 2


def should_keep_choice_like_line_in_question(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return False
    if "?" in stripped or "？" in stripped:
        return True
    if has_multiple_choice_markers(stripped):
        return True
    return False


def split_underlined_segments(line: str) -> List[tuple[str, bool]]:
    if UNDERLINE_OPEN not in line and UNDERLINE_CLOSE not in line:
        return [(line, False)]

    tokens = re.split(
        f"({re.escape(UNDERLINE_OPEN)}|{re.escape(UNDERLINE_CLOSE)})",
        line,
    )
    segments: List[tuple[str, bool]] = []
    underline = False
    for token in tokens:
        if token == UNDERLINE_OPEN:
            underline = True
            continue
        if token == UNDERLINE_CLOSE:
            underline = False
            continue
        if token:
            segments.append((token, underline))
    return segments if segments else [("", False)]


def option_label_to_prefix(label_key: str) -> str | None:
    if label_key in CHOICE_NUMBER_TO_CIRCLED:
        return CHOICE_NUMBER_TO_CIRCLED[label_key]
    if label_key in {"a", "b", "c", "d", "e"}:
        return f"{label_key.upper()}."
    if label_key in CIRCLED_CHOICES:
        return label_key
    return None


def resolve_label_type(label_key: str) -> str | None:
    if label_key in HEADER_LABELS:
        return "header"
    if label_key in BODY_LABELS:
        return "body"
    if label_key in QUESTION_LABELS:
        return "question"
    if label_key in CHOICE_LABELS:
        return "choice"
    if label_key in ANSWER_LABELS:
        return "answer"
    if label_key in EXPLANATION_LABELS:
        return "explanation"
    if label_key in PROBLEM_START_LABELS:
        return "problem"
    if option_label_to_prefix(label_key):
        return "option"
    return None


def empty_sections() -> dict[str, List[str]]:
    return {
        "header": [],
        "body": [],
        "question": [],
        "choice": [],
        "answer": [],
        "explanation": [],
    }


def has_section_content(sections: dict[str, List[str]]) -> bool:
    for key in ("body", "question", "choice"):
        if any(line.strip() for line in sections.get(key, [])):
            return True
    return False


def add_line_to_sections(
    sections: dict[str, List[str]], section: str, line: str, preserve_blank: bool = True
) -> None:
    if section not in sections:
        return

    line_rstrip = line.rstrip()
    if not line_rstrip.strip():
        if not preserve_blank:
            return
        if sections[section] and sections[section][-1] != "":
            sections[section].append("")
        return

    sections[section].append(line_rstrip)


def trim_trailing_blank_lines(lines: List[str]) -> List[str]:
    trimmed = list(lines)
    while trimmed and not trimmed[-1].strip():
        trimmed.pop()
    return trimmed


def append_text_line(base: str, addition: str) -> str:
    base_text = base.rstrip()
    addition_text = addition.strip()
    if not base_text:
        return addition_text
    if not addition_text:
        return base_text
    return f"{base_text} {addition_text}"


def is_inline_break_marker(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return False
    if NOTE_DEFINITION_MARKER_RE.search(stripped):
        return True
    if re.fullmatch(r"\([A-Z]\)", stripped):
        return True
    if is_header_line(stripped):
        return True
    lowered = stripped.lower()
    if lowered.startswith(("dear ", "best regards", "warm regards", "kind regards", "sincerely", "regards")):
        return True
    if stripped.endswith(",") and len(stripped) <= 40:
        return True
    return False


def reflow_body_lines(lines: List[str]) -> List[str]:
    reflowed: List[str] = []
    current = ""

    def flush() -> None:
        nonlocal current
        if current:
            reflowed.append(current)
        current = ""

    for raw_line in trim_trailing_blank_lines(lines):
        stripped = raw_line.strip()
        if not stripped:
            flush()
            if reflowed and reflowed[-1] != "":
                reflowed.append("")
            continue

        if is_inline_break_marker(stripped):
            flush()
            reflowed.append(stripped)
            continue

        if is_choice_group_start(stripped):
            normalized_choice = normalize_choice_spacing(raw_line)
            if current:
                current = append_text_line(current, normalized_choice)
            else:
                current = normalized_choice
            continue

        if current and current.rstrip().endswith(("?", "？")):
            flush()
            current = stripped
            continue

        if current:
            current = append_text_line(current, stripped)
        else:
            current = stripped

    flush()
    return trim_trailing_blank_lines(reflowed)


def reflow_question_lines(lines: List[str]) -> List[str]:
    reflowed: List[str] = []
    current = ""

    def flush() -> None:
        nonlocal current
        if current:
            reflowed.append(current)
        current = ""

    for raw_line in trim_trailing_blank_lines(lines):
        stripped = raw_line.strip()
        if not stripped:
            flush()
            if reflowed and reflowed[-1] != "":
                reflowed.append("")
            continue

        if current and looks_like_question_prompt(current) and current.endswith(("?", "？")) and not is_choice_group_start(stripped):
            flush()
            current = stripped
            continue

        if looks_like_question_prompt(stripped) or is_choice_group_start(stripped) or is_inline_break_marker(stripped):
            flush()
            current = normalize_choice_spacing(raw_line) if is_choice_group_start(stripped) else stripped
            continue

        if current:
            current = append_text_line(current, stripped)
        else:
            current = stripped

    flush()
    return trim_trailing_blank_lines(reflowed)


def reflow_choice_lines(lines: List[str]) -> List[str]:
    reflowed: List[str] = []
    current = ""

    def flush() -> None:
        nonlocal current
        if current:
            reflowed.append(normalize_choice_spacing(current))
        current = ""

    for raw_line in trim_trailing_blank_lines(lines):
        stripped = raw_line.strip()
        if not stripped:
            flush()
            continue

        if is_choice_group_start(stripped):
            flush()
            current = normalize_choice_spacing(raw_line)
            continue

        if current:
            current = append_text_line(current, stripped)
        else:
            current = stripped

    flush()
    return trim_trailing_blank_lines(reflowed)


def split_body_and_note_line(line: str) -> tuple[str, str | None]:
    stripped = line.rstrip()
    if not stripped.strip():
        return "", None

    compact = stripped.strip()
    if compact.startswith("*") and NOTE_DEFINITION_MARKER_RE.search(compact):
        return "", compact

    for match in NOTE_DEFINITION_MARKER_RE.finditer(stripped):
        marker_idx = match.start()
        prefix = safe_slice_text(stripped, 0, marker_idx)
        if not prefix.strip():
            continue
        # Inline glossary notes can be attached after a single space:
        # "... sentence. * term: meaning"
        note_tail = safe_slice_text(stripped, marker_idx, None)
        return prefix.rstrip(), note_tail.strip()

    return stripped, None


def move_body_notes_to_end(lines: List[str]) -> List[str]:
    main_lines: List[str] = []
    note_lines: List[str] = []

    for raw in lines:
        if not raw.strip():
            if main_lines and main_lines[-1] != "":
                main_lines.append("")
            continue

        body_part, note_part = split_body_and_note_line(raw)
        if body_part.strip():
            main_lines.append(body_part)
        if note_part:
            note_lines.append(note_part)

    ordered_main = trim_trailing_blank_lines(main_lines)
    if note_lines:
        if ordered_main and ordered_main[-1] != "":
            ordered_main.append("")
        ordered_main.extend(trim_trailing_blank_lines(note_lines))
    return trim_trailing_blank_lines(ordered_main)


def has_embedded_passage_after_prompt(question_lines: List[str]) -> bool:
    if len(question_lines) <= 1:
        return False

    candidates: List[str] = []
    for idx, line in enumerate(question_lines):
        if idx == 0:
            continue
        stripped = line.strip()
        if not stripped:
            continue
        if looks_like_question_prompt(stripped):
            continue
        if looks_like_choice_line(stripped):
            continue
        candidates.append(stripped)

    if len(candidates) >= 2:
        return True
    if candidates and len(candidates[0]) >= 60:
        return True
    return False


def normalize_multi_question_flow_lines(lines: List[str]) -> List[str]:
    normalized: List[str] = []

    for line in lines:
        stripped = line.strip()
        if not stripped:
            if normalized and normalized[-1] != "":
                normalized.append("")
            continue

        if looks_like_question_prompt(stripped):
            prev_non_empty = ""
            for prev in reversed(normalized):
                if prev.strip():
                    prev_non_empty = prev
                    break
            if prev_non_empty and not looks_like_question_prompt(prev_non_empty.strip()):
                if normalized and normalized[-1] != "":
                    normalized.append("")

        normalized.append(line)

    return trim_trailing_blank_lines(normalized)


def apply_problem_layout_rules(problem: Problem, qa_flow_lines: List[str] | None = None) -> Problem:
    question_text = "\n".join(problem.question_lines)
    prompt_count = 0
    if qa_flow_lines:
        prompt_count = sum(
            1 for line in qa_flow_lines if line.strip() and looks_like_question_prompt(line.strip())
        )

    if prompt_count >= 2 and len(problem.choice_lines) > 0:
        normalized_flow_lines = normalize_choice_lines(qa_flow_lines or [])
        return Problem(
            header_lines=problem.header_lines,
            body_lines=problem.body_lines,
            question_lines=normalize_multi_question_flow_lines(normalized_flow_lines),
            choice_lines=[],
            problem_number=problem.problem_number,
        )

    return problem


def normalize_problem(problem: Problem) -> Problem:
    return Problem(
        header_lines=trim_trailing_blank_lines(problem.header_lines),
        body_lines=reflow_body_lines(move_body_notes_to_end(trim_trailing_blank_lines(problem.body_lines))),
        question_lines=reflow_question_lines(problem.question_lines),
        choice_lines=reflow_choice_lines(problem.choice_lines),
        problem_number=problem.problem_number.strip(),
    )


def finalize_problem(
    problem: Problem,
    qa_flow_lines: List[str] | None = None,
) -> Problem:
    normalized_qa_flow_lines: List[str] | None = None
    if qa_flow_lines is not None:
        normalized_qa_flow_lines = trim_trailing_blank_lines(qa_flow_lines)

    normalized = normalize_problem(problem)
    return apply_problem_layout_rules(normalized, qa_flow_lines=normalized_qa_flow_lines)


def question_first_layout_keyword(raw: str) -> str | None:
    normalized_raw = raw.strip()
    if not normalized_raw:
        return None
    for keyword in QUESTION_FIRST_LAYOUT_KEYWORDS:
        if keyword in normalized_raw:
            return keyword
    return None


def coerce_ai_int(problem_data: dict[str, Any], key: str, default: int = 0) -> int:
    value = problem_data.get(key, default)
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    if isinstance(value, str):
        stripped = value.strip()
        if stripped.lstrip("-").isdigit():
            return int(stripped)
    return default


def serialize_problem_for_ai(problem: Problem) -> str:
    lines: List[str] = []

    if problem.header_lines:
        lines.append("[HEADER]")
        lines.extend(problem.header_lines)
        lines.append("")

    if problem.body_lines:
        lines.append("[BODY]")
        lines.extend(problem.body_lines)
        lines.append("")

    if problem.question_lines:
        lines.append("[QUESTION]")
        lines.extend(problem.question_lines)
        lines.append("")

    if problem.choice_lines:
        lines.append("[CHOICE]")
        lines.extend(problem.choice_lines)

    return "\n".join(trim_trailing_blank_lines(lines)).strip()


def merge_problem_into_body(problem: Problem) -> Problem:
    if not problem.question_lines and not problem.choice_lines:
        return problem

    merged_body = list(problem.question_lines)
    merged_body.extend(problem.body_lines)
    if problem.choice_lines:
        merged_body.extend(problem.choice_lines)
    return normalize_problem(
        Problem(
            header_lines=problem.header_lines,
            body_lines=trim_trailing_blank_lines(merged_body),
            question_lines=[],
            choice_lines=[],
            problem_number=problem.problem_number,
        )
    )


def line_looks_like_passage(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return False
    if is_inline_break_marker(stripped) or is_choice_group_start(stripped):
        return True

    ascii_alpha_count = sum(1 for ch in stripped if ("a" <= ch.lower() <= "z"))
    hangul_count = sum(1 for ch in stripped if "가" <= ch <= "힣")
    if ascii_alpha_count >= 20 and ascii_alpha_count >= hangul_count:
        return True
    return False


def separate_prompt_and_passage(problem: Problem) -> Problem:
    if not problem.question_lines:
        return problem

    prompt_lines: List[str] = []
    passage_lines = list(problem.body_lines)
    moved_to_passage = False

    for line in problem.question_lines:
        stripped = line.strip()
        if not stripped:
            target = passage_lines if moved_to_passage else prompt_lines
            if target and target[-1] != "":
                target.append("")
            continue

        if prompt_lines and line_looks_like_passage(line):
            moved_to_passage = True

        if moved_to_passage:
            passage_lines.append(line.rstrip())
        else:
            prompt_lines.append(line.rstrip())

    return normalize_problem(
        Problem(
            header_lines=list(problem.header_lines),
            body_lines=trim_trailing_blank_lines(passage_lines),
            question_lines=trim_trailing_blank_lines(prompt_lines),
            choice_lines=list(problem.choice_lines),
            problem_number=problem.problem_number,
        )
    )


def infer_problem_type_code(problem: Problem, raw_text: str) -> int:
    question_text = "\n".join(problem.question_lines).lower()
    body_text = "\n".join(problem.body_lines)
    choice_text = "\n".join(problem.choice_lines)
    combined_text = f"{question_text}\n{body_text}\n{choice_text}"

    keyword = question_first_layout_keyword(raw_text)
    if keyword is not None:
        if any(marker in combined_text for marker in ("( ① )", "(①)", "문장 넣기", "들어가기에")):
            return PROBLEM_TYPE_SENTENCE_INSERTION
        return PROBLEM_TYPE_UNRELATED_SENTENCE

    if "빈칸" in question_text:
        return PROBLEM_TYPE_BLANK
    if "어법" in question_text or "문법" in question_text:
        return PROBLEM_TYPE_GRAMMAR
    if "어휘" in question_text or "낱말" in question_text or "vocab" in question_text:
        return PROBLEM_TYPE_VOCAB
    if any(token in question_text for token in ("제목", "주제", "요지", "목적", "summary", "gist", "topic", "title")):
        return PROBLEM_TYPE_TITLE_TOPIC_GIST
    if any(token in question_text for token in ("순서", "배열")):
        return PROBLEM_TYPE_SENTENCE_ORDERING
    if "가리키는 대상" in question_text or "가리키는" in question_text or "refer" in question_text:
        return PROBLEM_TYPE_REFERENCE
    return PROBLEM_TYPE_MULTIPLE_CHOICE


def apply_problem_type_code(problem: Problem, problem_type_code: int) -> Problem:
    if problem_type_code in QUESTION_FIRST_TYPE_CODES:
        return merge_problem_into_body(problem)
    if problem_type_code != PROBLEM_TYPE_UNKNOWN:
        return separate_prompt_and_passage(problem)
    return problem


def split_question_choice_group(group_lines: List[str]) -> tuple[List[str], List[str]]:
    question_lines: List[str] = []
    choice_lines: List[str] = []
    in_choice_block = False

    for line in group_lines:
        stripped = line.strip()
        if not stripped:
            target = choice_lines if in_choice_block else question_lines
            if target and target[-1] != "":
                target.append("")
            continue

        if (
            not in_choice_block
            and looks_like_choice_line(stripped)
            and not looks_like_question_prompt(stripped)
        ):
            in_choice_block = True

        if in_choice_block:
            choice_lines.append(normalize_choice_spacing(line))
        else:
            question_lines.append(line.rstrip())

    return trim_trailing_blank_lines(question_lines), trim_trailing_blank_lines(choice_lines)


def split_multi_question_problem(problem: Problem) -> List[Problem]:
    prompt_indexes = [
        idx
        for idx, line in enumerate(problem.question_lines)
        if line.strip() and looks_like_question_prompt(line.strip())
    ]
    if len(prompt_indexes) <= 1:
        return [problem]

    split_problems: List[Problem] = []
    for idx, start in enumerate(prompt_indexes):
        end = prompt_indexes[idx + 1] if idx + 1 < len(prompt_indexes) else len(problem.question_lines)
        group_lines = trim_trailing_blank_lines(copy_line_range(problem.question_lines, start, end))
        if not any(line.strip() for line in group_lines):
            continue

        group_question_lines, group_choice_lines = split_question_choice_group(group_lines)
        grouped_problem = finalize_problem(
            Problem(
                header_lines=list(problem.header_lines),
                body_lines=list(problem.body_lines),
                question_lines=group_question_lines,
                choice_lines=group_choice_lines,
                problem_number="",
            ),
            qa_flow_lines=group_lines,
        )
        split_problems.append(
            Problem(
                header_lines=list(grouped_problem.header_lines),
                body_lines=list(grouped_problem.body_lines),
                question_lines=list(grouped_problem.question_lines),
                choice_lines=list(grouped_problem.choice_lines),
                problem_number=extract_problem_number(grouped_problem) or problem.problem_number,
            )
        )

    return split_problems or [problem]


def parse_rule_chunk(lines: List[str]) -> List[Problem]:
    labeled = parse_labeled_problems(lines)
    base_problems = labeled if labeled else [parse_problem(lines)]
    expanded: List[Problem] = []
    for problem in base_problems:
        expanded.extend(split_multi_question_problem(problem))
    return expanded


def parse_raw_details(raw: str) -> ParseResult:
    lines = normalize_lines(raw)
    if not any(line.strip() for line in lines):
        result = ParseResult(
            problems=[],
            parser_name="rules",
            ai_attempted=False,
            ai_used=False,
            reason="empty_input",
        )
        LOGGER.info(
            "parse_completed parser=%s ai_attempted=%s ai_used=%s reason=%s problems=%d input_chars=%d",
            result.parser_name,
            result.ai_attempted,
            result.ai_used,
            result.reason,
            len(result.problems),
            len(raw),
        )
        return result

    chunks = [chunk for chunk in split_problems(lines) if any(line.strip() for line in chunk)]
    base_dir = Path(__file__).resolve().parent
    ai_enabled = ai_parser_enabled(base_dir)
    final_problems: List[Problem] = []
    ai_used_problem_markers: List[None] = []
    rule_used_problem_markers: List[None] = []
    total_problem_candidates = 0

    for chunk in chunks:
        chunk_problems = parse_rule_chunk(chunk)
        total_problem_candidates += len(chunk_problems)

        for problem in chunk_problems:
            classification_payload: dict[str, Any] | None = None
            if ai_enabled:
                classification_payload = classify_problem_type_with_ai(
                    serialize_problem_for_ai(problem),
                    base_dir=base_dir,
                )

            if classification_payload is None:
                problem_type_code = infer_problem_type_code(problem, serialize_problem_for_ai(problem))
                normalized_problem = apply_problem_type_code(problem, problem_type_code)
                rule_used_problem_markers.append(None)
            else:
                problem_type_code = coerce_ai_int(classification_payload, "problem_type_code", PROBLEM_TYPE_UNKNOWN)
                if problem_type_code == PROBLEM_TYPE_UNKNOWN:
                    normalized_problem = problem
                    rule_used_problem_markers.append(None)
                else:
                    normalized_problem = apply_problem_type_code(problem, problem_type_code)
                    ai_used_problem_markers.append(None)

            final_problems.append(normalized_problem)

    ai_used_problems = len(ai_used_problem_markers)
    rule_used_problems = len(rule_used_problem_markers)
    ai_attempted = ai_enabled and total_problem_candidates > 0
    ai_used = ai_used_problems > 0
    if ai_enabled:
        if ai_used_problems == total_problem_candidates and total_problem_candidates > 0:
            reason = "ai_type_primary"
        elif ai_used_problems > 0:
            reason = f"ai_type_mixed:{ai_used_problems}/{total_problem_candidates}"
        else:
            reason = "ai_type_unknown"
    else:
        reason = "rules_only"

    result = ParseResult(
        problems=final_problems,
        parser_name="hybrid-ai" if ai_used else "rules",
        ai_attempted=ai_attempted,
        ai_used=ai_used,
        reason=reason,
    )
    LOGGER.info(
        "parse_completed parser=%s ai_attempted=%s ai_used=%s reason=%s problems=%d chunks=%d ai_problems=%d rule_problems=%d input_chars=%d",
        result.parser_name,
        result.ai_attempted,
        result.ai_used,
        result.reason,
        len(result.problems),
        len(chunks),
        ai_used_problems,
        rule_used_problems,
        len(raw),
    )
    return result


def build_problem_from_sections(sections: dict[str, List[str]]) -> Problem | None:
    if not has_section_content(sections):
        return None

    return finalize_problem(
        Problem(
            header_lines=list(sections["header"]),
            body_lines=list(sections["body"]),
            question_lines=list(sections["question"]),
            choice_lines=list(sections["choice"]),
            problem_number="",
        )
    )


def parse_labeled_problems(lines: List[str]) -> List[Problem]:
    sections = empty_sections()
    problems: List[Problem] = []
    current_section: str = "body"
    seen_start_labels: set[str] = set()
    found_known_label = False

    def flush_current() -> None:
        nonlocal sections, current_section, seen_start_labels
        built = build_problem_from_sections(sections)
        if built is not None:
            problems.append(built)
        sections = empty_sections()
        current_section = "body"
        seen_start_labels = set()

    for raw_line in lines:
        line = raw_line.rstrip()
        stripped = line.strip()
        if stripped and is_meta_noise_line(stripped):
            continue

        label_key = ""
        inline = ""
        label_match = LABELED_LINE_RE.match(stripped)
        if label_match:
            label_key = normalize_label_key(label_match.group(1))
            inline = label_match.group(2).strip()
        else:
            colon_match = COLON_LABEL_RE.match(stripped)
            if colon_match:
                maybe_key = normalize_label_key(colon_match.group(1))
                if resolve_label_type(maybe_key) is not None:
                    label_key = maybe_key
                    inline = colon_match.group(2).strip()

        if label_key:
            label_type = resolve_label_type(label_key)

            if label_type is None:
                # Unknown bracket label is treated as body text.
                if stripped:
                    add_line_to_sections(sections, "body", stripped)
                continue

            found_known_label = True

            if label_type == "body":
                has_question_lines = any(line.strip() for line in sections["question"])
                has_choice_lines = any(line.strip() for line in sections["choice"])
                if has_section_content(sections) and (has_question_lines or has_choice_lines):
                    flush_current()
            elif label_type == "question":
                has_choice_lines = any(line.strip() for line in sections["choice"])
                if has_section_content(sections) and has_choice_lines:
                    flush_current()

            if label_type in {"problem", "header"}:
                if label_key in seen_start_labels and has_section_content(sections):
                    flush_current()
                seen_start_labels.add(label_key)

            if label_type == "problem":
                current_section = "header"
                if inline:
                    add_line_to_sections(sections, "header", inline)
                continue

            if label_type in {"answer", "explanation"}:
                # Enforce body/question/choice-only output rule.
                continue

            if label_type == "option":
                current_section = "choice"
                option_prefix = option_label_to_prefix(label_key)
                if option_prefix:
                    if inline:
                        add_line_to_sections(
                            sections,
                            "choice",
                            f"{option_prefix} {inline}",
                            preserve_blank=False,
                        )
                    else:
                        add_line_to_sections(
                            sections,
                            "choice",
                            option_prefix,
                            preserve_blank=False,
                        )
                continue

            current_section = label_type
            if inline:
                add_line_to_sections(
                    sections,
                    current_section,
                    inline,
                    preserve_blank=current_section != "choice",
                )
            continue

        if not stripped:
            if current_section in {"body", "question"}:
                add_line_to_sections(sections, current_section, "")
            continue

        if looks_like_choice_line(stripped):
            if current_section == "body":
                add_line_to_sections(sections, "body", line)
                continue

            if current_section == "question":
                if should_keep_choice_like_line_in_question(stripped):
                    add_line_to_sections(sections, "question", line)
                    continue
                add_line_to_sections(sections, "choice", line, preserve_blank=False)
                continue

            current_section = "choice"
            add_line_to_sections(sections, "choice", line, preserve_blank=False)
            continue

        if current_section == "choice" and sections["choice"]:
            if is_meta_noise_line(stripped):
                continue
            add_line_to_sections(sections, "choice", line, preserve_blank=False)
            continue

        add_line_to_sections(sections, current_section, line)

    flush_current()

    if not found_known_label:
        return []
    return problems


# -------------------------------
# Freeform parser fallback
# -------------------------------
def parse_problem(raw_lines: List[str]) -> Problem:
    lines = [ln for ln in raw_lines]

    header_lines: List[str] = []
    body_lines: List[str] = []
    question_lines: List[str] = []
    choice_lines: List[str] = []

    idx = 0
    while idx < len(lines):
        stripped = lines[idx].strip()
        if not stripped:
            idx += 1
            continue
        if is_header_line(stripped):
            header_lines.append(stripped)
            idx += 1
            continue
        break

    content: List[str] = []
    for pos, line in enumerate(lines):
        if pos >= idx:
            content.append(line)

    if not header_lines:
        embedded_header_idx = next(
            (
                i
                for i, line in enumerate(content)
                if HEADER_START_RE.match(line.strip())
                and not looks_like_question_prompt(line.strip())
            ),
            None,
        )
        if embedded_header_idx is not None:
            header_lines.append(content[embedded_header_idx].strip())
            content = remove_index(content, embedded_header_idx)

    qa_flow_lines: List[str] = []
    in_qa_flow = False

    for line in content:
        stripped = line.strip()

        if not stripped:
            target = qa_flow_lines if in_qa_flow else body_lines
            if target and target[-1] != "":
                target.append("")
            continue

        if is_meta_noise_line(stripped):
            continue

        if looks_like_question_prompt(stripped):
            in_qa_flow = True
            qa_flow_lines.append(line.rstrip())
            continue

        if in_qa_flow:
            qa_flow_lines.append(line.rstrip())
        else:
            body_lines.append(line.rstrip())

    in_choice_block = False
    for line in qa_flow_lines:
        stripped = line.strip()
        if not stripped:
            target = choice_lines if in_choice_block else question_lines
            if target and target[-1] != "":
                target.append("")
            continue

        if (
            not in_choice_block
            and looks_like_choice_line(stripped)
            and not looks_like_question_prompt(stripped)
        ):
            in_choice_block = True

        if in_choice_block:
            choice_lines.append(normalize_choice_spacing(line))
        else:
            question_lines.append(line.rstrip())

    return finalize_problem(
        Problem(
            header_lines=header_lines,
            body_lines=body_lines,
            question_lines=question_lines,
            choice_lines=choice_lines,
            problem_number="",
        ),
        qa_flow_lines=qa_flow_lines,
    )


def render_problem(problem: Problem, index: int) -> str:
    lines: List[str] = []
    lines.append(f"[슬라이드 {index}]")
    lines.append("")
    lines.append(STYLE_BLOCK)
    lines.append("")

    lines.append("[상단표기]")
    if problem.header_lines:
        lines.extend(problem.header_lines)
    lines.append("")

    lines.append("[본문단]")
    if problem.body_lines:
        lines.extend(problem.body_lines)
    lines.append("")

    lines.append("[문제단]")
    if problem.question_lines:
        lines.extend(problem.question_lines)
    lines.append("")

    lines.append("[선지단]")
    if problem.choice_lines:
        lines.extend(problem.choice_lines)

    return "\n".join(lines).rstrip()


def parse_raw(raw: str) -> List[Problem]:
    """Parse raw input text into normalized problem records."""
    return parse_raw_details(raw).problems


# -------------------------------
# Pagination and text rendering
# -------------------------------
def normalize_content_lines(lines: List[str]) -> List[str]:
    out: List[str] = []
    for line in lines:
        stripped = line.rstrip()
        if not stripped:
            if out and out[-1] != "":
                out.append("")
            continue
        out.append(stripped)
    while out and out[-1] == "":
        out.pop()
    return out


def estimate_visual_lines(line: str, chars_per_line: int) -> int:
    if not line:
        return 1
    return max(1, int(math.ceil(len(line) / chars_per_line)))


def add_visual_lines(current_count: int, needed_count: int) -> int:
    return current_count + needed_count


def visual_line_markers(count: int) -> List[None]:
    markers: List[None] = []
    for _ in range(count):
        markers.append(None)
    return markers


def append_visual_line_markers(markers: List[None], count: int) -> List[None]:
    updated: List[None] = list(markers)
    for _ in range(count):
        updated.append(None)
    return updated


def remaining_visual_line_capacity(max_lines: int, used_lines: int) -> int:
    remaining_markers: List[None] = []
    used_markers: List[None] = []
    for _ in range(max_lines):
        if len(used_markers) < used_lines:
            used_markers.append(None)
        else:
            remaining_markers.append(None)
    return len(remaining_markers)


def chunk_visual_line_count(chunk: List[str], chars_per_line: int) -> int:
    total = 0
    for line in chunk:
        total += estimate_visual_lines(line, chars_per_line)
    return total


def has_content_lines(lines: List[str]) -> bool:
    return any(line.strip() for line in lines)


def vertical_inset_points() -> float:
    return int(TEXTBOX_VERTICAL_INSET) / EMU_PER_PT


def line_height_points() -> float:
    return FONT_SIZE_PT * (LINE_SPACING_PERCENT / 100.0)


def max_visual_lines_for_height(height: int) -> int:
    usable_height = max(0.0, float(height) - (vertical_inset_points() * 2.0))
    if usable_height <= 0:
        return 0
    return max(0, int(math.floor(usable_height / line_height_points())))


def take_first_chunk(
    lines: List[str],
    max_lines: int,
    chars_per_line: int,
    *,
    allow_partial_line_split: bool = False,
) -> tuple[List[str], List[str]]:
    if max_lines <= 0 or not has_content_lines(lines):
        return [], list(lines)

    chunks = chunk_lines_for_box(
        lines,
        max_lines,
        chars_per_line,
        allow_partial_line_split=allow_partial_line_split,
    )
    if not chunks:
        return [], []

    first_chunk = list(chunks[0])
    remaining: List[str] = []
    for chunk_index in range(1, len(chunks)):
        remaining.extend(chunks[chunk_index])
    return first_chunk, remaining


def take_first_chunk_by_height(
    lines: List[str],
    available_height: int,
    chars_per_line: int,
    *,
    allow_partial_line_split: bool = False,
    reserved_visual_lines: int = 0,
) -> tuple[List[str], List[str]]:
    if available_height <= 0 or not has_content_lines(lines):
        return [], list(lines)

    max_lines = max_visual_lines_for_height(available_height)
    if reserved_visual_lines > 0:
        max_lines = max(0, max_lines - reserved_visual_lines)
    if max_lines <= 0:
        return [], list(lines)
    return take_first_chunk(
        lines,
        max_lines,
        chars_per_line,
        allow_partial_line_split=allow_partial_line_split,
    )


def content_flow_bounds(has_body: bool) -> tuple[int, int]:
    if has_body:
        return CONTENT_FLOW_TOP, content_box_bottom(BODY_ONLY_BOX)
    return CONTENT_FLOW_TOP, content_box_bottom(BODY_ONLY_BOX)


def soften_chunk_boundaries(
    chunks: List[List[str]], max_lines: int, chars_per_line: int
) -> List[List[str]]:
    if len(chunks) < 2:
        return chunks

    softened: List[List[str]] = [list(chunks[0])]
    for next_idx in range(1, len(chunks)):
        next_chunk = chunks[next_idx]
        prev_chunk = softened[-1]
        prev_lines = chunk_visual_line_count(prev_chunk, chars_per_line)
        next_lines = chunk_visual_line_count(next_chunk, chars_per_line)

        # Avoid tiny orphan chunks such as a salutation line split from the body.
        if (prev_lines <= 1 or next_lines <= 1) and (prev_lines + next_lines <= max_lines):
            softened[-1] = prev_chunk + list(next_chunk)
        else:
            softened.append(list(next_chunk))

    return softened


def split_very_long_line(line: str, chars_per_line: int, max_lines: int) -> List[str]:
    # Split only at page-sized chunks, not per visual line, so rendered text wraps naturally.
    max_chars_per_segment = max(chars_per_line, chars_per_line * max_lines)
    if len(line) <= max_chars_per_segment:
        return [line]

    chunks: List[str] = []
    remaining: str = line
    while len(remaining) > 0:
        if len(remaining) <= max_chars_per_segment:
            chunks.append(remaining)
            break

        cut_idx: int = remaining.rfind(" ", 0, max_chars_per_segment + 1)
        if cut_idx <= 0:
            cut_idx = max_chars_per_segment

        part = safe_slice_text(remaining, 0, cut_idx).rstrip()
        if part:
            chunks.append(part)
        remaining = safe_slice_text(remaining, cut_idx, None).lstrip()

    return chunks if chunks else [line]


def split_line_for_available_lines(
    line: str, available_lines: int, chars_per_line: int
) -> tuple[str, str] | None:
    if available_lines <= 0:
        return None

    max_chars = max(chars_per_line, chars_per_line * available_lines)
    if len(line) <= max_chars:
        return None

    cut_idx = line.rfind(" ", 0, max_chars + 1)
    if cut_idx <= 0:
        cut_idx = max_chars

    prefix = safe_slice_text(line, 0, cut_idx).rstrip()
    remainder = safe_slice_text(line, cut_idx, None).lstrip()
    if not prefix or not remainder:
        return None
    return prefix, remainder


def chunk_lines_for_box(
    lines: List[str],
    max_lines: int,
    chars_per_line: int,
    *,
    allow_partial_line_split: bool = False,
) -> List[List[str]]:
    cleaned = normalize_content_lines(lines)
    if not cleaned:
        return [[]]

    chunks: List[List[str]] = []
    current: List[str] = []
    current_visual_line_markers: List[None] = []

    for raw_line in cleaned:
        segments: List[str] = [raw_line]
        if raw_line and estimate_visual_lines(raw_line, chars_per_line) > max_lines:
            segments = split_very_long_line(raw_line, chars_per_line, max_lines)

        segment_index = 0
        while segment_index < len(segments):
            segment = segments[segment_index]
            line_text: str = segment
            if line_text == "" and not current:
                segment_index += 1
                continue

            needed_visual_lines: int = estimate_visual_lines(line_text, chars_per_line)
            projected_visual_lines: int = add_visual_lines(
                len(current_visual_line_markers), needed_visual_lines
            )

            if current and projected_visual_lines > max_lines:
                if allow_partial_line_split:
                    remaining_lines = remaining_visual_line_capacity(
                        max_lines,
                        len(current_visual_line_markers),
                    )
                    split_result = split_line_for_available_lines(
                        line_text,
                        remaining_lines,
                        chars_per_line,
                    )
                    if split_result is not None:
                        prefix, remainder = split_result
                        current = current + [prefix]
                        current_visual_line_markers = append_visual_line_markers(
                            current_visual_line_markers,
                            estimate_visual_lines(prefix, chars_per_line),
                        )
                        current = trim_trailing_blank_lines(current)
                        if current:
                            chunks.append(current)
                        current = []
                        reset_markers: List[None] = []
                        current_visual_line_markers = reset_markers
                        segments[segment_index] = remainder
                        continue

                current = trim_trailing_blank_lines(current)
                if current:
                    chunks.append(current)
                next_current: List[str] = []
                current = next_current
                next_markers: List[None] = []
                current_visual_line_markers = next_markers

            current = current + [line_text]
            current_visual_line_markers = append_visual_line_markers(
                current_visual_line_markers,
                needed_visual_lines,
            )
            segment_index += 1

        if raw_line == "" and not current:
            continue

    current = trim_trailing_blank_lines(current)
    if current:
        chunks.append(current)

    if not chunks:
        empty_chunk: List[str] = []
        return [empty_chunk]
    return soften_chunk_boundaries(chunks, max_lines, chars_per_line)


def paginate_problem(problem: Problem) -> List[Problem]:
    body_remaining = list(problem.body_lines)
    question_remaining = list(problem.question_lines)
    choice_remaining = list(problem.choice_lines)
    pages: List[Problem] = []

    while (
        has_content_lines(body_remaining)
        or has_content_lines(question_remaining)
        or has_content_lines(choice_remaining)
    ):
        body_source: List[str] = []
        question_source: List[str] = []
        choice_source: List[str] = []
        has_body_for_page = has_content_lines(body_remaining)
        flow_top, flow_bottom = content_flow_bounds(has_body_for_page)
        remaining_height = flow_bottom - flow_top
        has_page_content = False

        if has_body_for_page:
            body_source, body_remaining = take_first_chunk_by_height(
                body_remaining,
                remaining_height,
                BODY_CHARS_PER_LINE,
                allow_partial_line_split=True,
                reserved_visual_lines=BODY_VISUAL_LINE_SAFETY_MARGIN,
            )
            if has_content_lines(body_source):
                remaining_height -= required_text_box_height(
                    chunk_visual_line_count(body_source, BODY_CHARS_PER_LINE)
                )
                has_page_content = True
            if has_content_lines(body_remaining):
                pages.append(
                    Problem(
                        header_lines=list(problem.header_lines),
                        body_lines=list(body_source),
                        question_lines=[],
                        choice_lines=[],
                        problem_number=problem.problem_number,
                    )
                )
                continue

        if has_content_lines(question_remaining):
            question_space = remaining_height - (CONTENT_FLOW_GAP if has_page_content else 0)
            question_source, question_remaining = take_first_chunk_by_height(
                question_remaining, question_space, QUESTION_CHARS_PER_LINE
            )
            if has_content_lines(question_source):
                if has_page_content:
                    remaining_height -= CONTENT_FLOW_GAP
                remaining_height -= required_text_box_height(
                    chunk_visual_line_count(question_source, QUESTION_CHARS_PER_LINE)
                )
                has_page_content = True

        if has_content_lines(choice_remaining) and not has_content_lines(question_remaining):
            choice_space = remaining_height - (CONTENT_FLOW_GAP if has_page_content else 0)
            choice_source, choice_remaining = take_first_chunk_by_height(
                choice_remaining, choice_space, CHOICE_CHARS_PER_LINE
            )
            if has_content_lines(choice_source):
                if has_page_content:
                    remaining_height -= CONTENT_FLOW_GAP
                remaining_height -= required_text_box_height(
                    chunk_visual_line_count(choice_source, CHOICE_CHARS_PER_LINE)
                )
                has_page_content = True

        page = Problem(
            header_lines=list(problem.header_lines),
            body_lines=list(body_source),
            question_lines=list(question_source),
            choice_lines=list(choice_source),
            problem_number=problem.problem_number,
        )
        if (
            not has_content_lines(page.body_lines)
            and not has_content_lines(page.question_lines)
            and not has_content_lines(page.choice_lines)
        ):
            break
        pages.append(page)

    return pages if pages else [problem]


def paginate_problems(problems: List[Problem]) -> List[Problem]:
    paginated: List[Problem] = []
    for idx, problem in enumerate(problems, start=1):
        base_number = str(idx)
        normalized = apply_display_number_to_problem(problem, base_number)
        paginated.extend(paginate_problem(normalized))
    return paginated


def format_text(raw: str) -> str:
    """Render parsed problems as labeled, human-readable slide text."""
    parsed = paginate_problems(parse_raw(raw))
    if not parsed:
        return ""

    rendered = [render_problem(problem, idx + 1) for idx, problem in enumerate(parsed)]
    return "\n\n".join(rendered).rstrip() + "\n"


# -------------------------------
# PPTX build helpers
# -------------------------------
def template_candidate_paths(base_dir: Path) -> List[Path]:
    preferred = [base_dir / name for name in PREFERRED_TEMPLATE_NAMES]
    preferred_names = {path.name for path in preferred}
    fallback = sorted(
        candidate
        for candidate in base_dir.glob("*.pptx")
        if candidate.name not in preferred_names
    )
    return preferred + fallback


def find_pptx_template() -> Path:
    base_dir = Path(__file__).resolve().parent
    for candidate in template_candidate_paths(base_dir):
        if candidate.exists():
            return candidate
    raise RuntimeError(
        "사용 가능한 PPTX 기준 파일을 찾을 수 없습니다. "
        "프로젝트 폴더에 .pptx 파일을 하나 두고 다시 실행하세요."
    )


def extract_problem_number(problem: Problem) -> str:
    if problem.problem_number.strip():
        return problem.problem_number.strip()

    number_re = re.compile(r"^\s*(\d+)\s*[.)-]?")
    candidates = problem.header_lines + problem.question_lines + problem.body_lines
    for line in candidates:
        match = number_re.match(line.strip())
        if match:
            return match.group(1)
    return ""


def replace_problem_number_prefix(line: str, display_number: str) -> str:
    match = PROBLEM_NUMBER_PREFIX_RE.match(line)
    if match is None:
        return line

    lead = match.group("lead") or ""
    sep = match.group("sep") or ""
    tail = match.group("tail") or ""
    return f"{lead}{display_number}{sep}{tail}"


def replace_first_problem_prompt_number(lines: List[str], display_number: str) -> List[str]:
    replaced: List[str] = []
    updated = False
    for line in lines:
        if not updated and line.strip() and looks_like_question_prompt(line.strip()):
            replaced.append(replace_problem_number_prefix(line, display_number))
            updated = True
            continue
        replaced.append(line)
    return replaced


def apply_display_number_to_problem(problem: Problem, display_number: str) -> Problem:
    return Problem(
        header_lines=list(problem.header_lines),
        body_lines=replace_first_problem_prompt_number(problem.body_lines, display_number),
        question_lines=replace_first_problem_prompt_number(problem.question_lines, display_number),
        choice_lines=list(problem.choice_lines),
        problem_number=display_number,
    )


def add_text_paragraphs(
    tx_body: ET.Element,
    text: str,
    *,
    font_size_pt: int = FONT_SIZE_PT,
    bold: bool = True,
    align: str = "l",
    font_name: str = FONT_NAME,
    align_note_lines_right: bool = False,
    note_font_size_pt: int | None = None,
    note_font_name: str | None = None,
) -> None:
    safe_text = sanitize_xml_text(text)
    lines = safe_text.splitlines() if safe_text else [""]
    if not lines:
        lines = [""]

    for line in lines:
        stripped = line.strip()
        is_note_line = NOTE_DEFINITION_MARKER_RE.search(stripped) is not None
        paragraph_align = "r" if (align_note_lines_right and is_note_line) else align
        paragraph_font_size_pt: int = font_size_pt
        paragraph_font_name: str = font_name
        if is_note_line:
            if isinstance(note_font_size_pt, int):
                paragraph_font_size_pt = note_font_size_pt
            if isinstance(note_font_name, str):
                paragraph_font_name = note_font_name
        paragraph = ET.SubElement(tx_body, qn(NS_A, "p"))
        enforce_paragraph_style(paragraph, align=paragraph_align)
        if line:
            segments = split_underlined_segments(line)
            for segment_text, underlined in segments:
                if not segment_text:
                    continue
                run = ET.SubElement(paragraph, qn(NS_A, "r"))
                run_pr = ET.SubElement(run, qn(NS_A, "rPr"))
                enforce_run_style(
                    run_pr,
                    font_size_pt=paragraph_font_size_pt,
                    bold=bold,
                    underlined=underlined,
                    font_name=paragraph_font_name,
                )
                text_el = ET.SubElement(run, qn(NS_A, "t"))
                has_leading_space = bool(segment_text) and segment_text[0].isspace()
                has_trailing_space = bool(segment_text) and segment_text[len(segment_text) - 1].isspace()
                if has_leading_space or has_trailing_space:
                    text_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                text_el.text = segment_text
        end_para = ET.SubElement(paragraph, qn(NS_A, "endParaRPr"))
        enforce_run_style(
            end_para,
            font_size_pt=paragraph_font_size_pt,
            bold=bold,
            font_name=paragraph_font_name,
        )


def add_text_shape(
    sp_tree: ET.Element,
    shape_id: int,
    shape_name: str,
    text: str,
    box: Tuple[int, int, int, int],
    *,
    font_size_pt: int = FONT_SIZE_PT,
    bold: bool = True,
    align: str = "l",
    font_name: str = FONT_NAME,
    align_note_lines_right: bool = False,
    note_font_size_pt: int | None = None,
    note_font_name: str | None = None,
    l_ins: str = TEXTBOX_HORIZONTAL_INSET,
    r_ins: str | None = TEXTBOX_HORIZONTAL_INSET,
) -> None:
    box_x, box_y, box_w, box_h = box
    shape = ET.SubElement(sp_tree, qn(NS_P, "sp"))

    nv_sp_pr = ET.SubElement(shape, qn(NS_P, "nvSpPr"))
    ET.SubElement(
        nv_sp_pr,
        qn(NS_P, "cNvPr"),
        {"id": str(shape_id), "name": shape_name},
    )
    ET.SubElement(nv_sp_pr, qn(NS_P, "cNvSpPr"), {"txBox": "1"})
    ET.SubElement(nv_sp_pr, qn(NS_P, "nvPr"))

    sp_pr = ET.SubElement(shape, qn(NS_P, "spPr"))
    xfrm = ET.SubElement(sp_pr, qn(NS_A, "xfrm"))
    ET.SubElement(xfrm, qn(NS_A, "off"), {"x": to_emu(box_x), "y": to_emu(box_y)})
    ET.SubElement(xfrm, qn(NS_A, "ext"), {"cx": to_emu(box_w), "cy": to_emu(box_h)})
    prst_geom = ET.SubElement(sp_pr, qn(NS_A, "prstGeom"), {"prst": "rect"})
    ET.SubElement(prst_geom, qn(NS_A, "avLst"))
    line_style = ET.SubElement(sp_pr, qn(NS_A, "ln"))
    ET.SubElement(line_style, qn(NS_A, "noFill"))

    tx_body = ET.SubElement(shape, qn(NS_P, "txBody"))
    resolved_r_ins = l_ins if r_ins is None else r_ins
    body_pr = ET.SubElement(
        tx_body,
        qn(NS_A, "bodyPr"),
        {
            "anchor": "t",
            "lIns": l_ins,
            "tIns": TEXTBOX_VERTICAL_INSET,
            "rIns": resolved_r_ins,
            "bIns": TEXTBOX_VERTICAL_INSET,
            "wrap": "square",
        },
    )
    ET.SubElement(body_pr, qn(NS_A, "noAutofit"))
    ET.SubElement(tx_body, qn(NS_A, "lstStyle"))
    add_text_paragraphs(
        tx_body,
        text,
        font_size_pt=font_size_pt,
        bold=bold,
        align=align,
        font_name=font_name,
        align_note_lines_right=align_note_lines_right,
        note_font_size_pt=note_font_size_pt,
        note_font_name=note_font_name,
    )


def content_box_bottom(box: Tuple[int, int, int, int]) -> int:
    return box[1] + box[3]


def required_text_box_height(visual_lines: int) -> int:
    vertical_inset_pt = int(TEXTBOX_VERTICAL_INSET) / EMU_PER_PT
    text_height = max(1, visual_lines) * line_height_points()
    return max(1, int(round(text_height + (vertical_inset_pt * 2))))


def build_flow_content_shapes(problem: Problem) -> List[tuple[str, str, Tuple[int, int, int, int], bool]]:
    sections: List[tuple[str, List[str], int, int, bool, int]] = []
    if has_content_lines(problem.body_lines):
        sections.append(
            (
                "body",
                list(problem.body_lines),
                BODY_BOX[0],
                BODY_BOX[2],
                True,
                BODY_CHARS_PER_LINE,
            )
        )
    if has_content_lines(problem.question_lines):
        sections.append(
            (
                "question",
                list(problem.question_lines),
                QUESTION_BOX[0],
                QUESTION_BOX[2],
                False,
                QUESTION_CHARS_PER_LINE,
            )
        )
    if has_content_lines(problem.choice_lines):
        sections.append(
            (
                "choice",
                list(problem.choice_lines),
                CHOICE_BOX[0],
                CHOICE_BOX[2],
                False,
                CHOICE_CHARS_PER_LINE,
            )
        )

    if not sections:
        return []

    if len(sections) == 1:
        name, lines, box_x, box_w, align_notes, _chars_per_line = sections[0]
        single_text = "\n".join(lines).rstrip()
        if name == "body":
            return [(name, single_text, BODY_ONLY_BOX, align_notes)]
        return [(name, single_text, BODY_ONLY_BOX, align_notes)]

    has_body = sections[0][0] == "body"
    flow_top, flow_bottom = content_flow_bounds(has_body)
    total_height = flow_bottom - flow_top

    required_heights: List[int] = []
    for _name, lines, _box_x, _box_w, _align_notes, chars_per_line in sections:
        visual_lines = chunk_visual_line_count(normalize_content_lines(lines), chars_per_line)
        required_heights.append(required_text_box_height(visual_lines))

    total_gap = CONTENT_FLOW_GAP * max(0, len(sections) - 1)
    available_for_sections = max(0, total_height - total_gap)
    total_required = sum(required_heights)

    scaled_heights = list(required_heights)
    if total_required > available_for_sections and total_required > 0:
        ratio = available_for_sections / total_required
        scaled_heights = [max(1, int(math.floor(height * ratio))) for height in required_heights]
        remainder = available_for_sections - sum(scaled_heights)
        index = 0
        while remainder > 0 and scaled_heights:
            scaled_heights[index % len(scaled_heights)] += 1
            remainder -= 1
            index += 1

    shapes: List[tuple[str, str, Tuple[int, int, int, int], bool]] = []
    cursor_y = flow_top
    for idx, section in enumerate(sections):
        name, lines, box_x, box_w, align_notes, _chars_per_line = section
        if idx > 0:
            cursor_y += CONTENT_FLOW_GAP

        if idx == len(sections) - 1:
            box_height = max(1, flow_bottom - cursor_y)
        else:
            box_height = scaled_heights[idx]

        text = "\n".join(lines).rstrip()
        shapes.append((name, text, (box_x, cursor_y, box_w, box_height), align_notes))
        cursor_y += box_height

    return shapes


def add_fill_rect_shape(
    sp_tree: ET.Element,
    shape_id: int,
    shape_name: str,
    box: Tuple[int, int, int, int],
    fill_hex: str,
) -> None:
    box_x, box_y, box_w, box_h = box
    shape = ET.SubElement(sp_tree, qn(NS_P, "sp"))

    nv_sp_pr = ET.SubElement(shape, qn(NS_P, "nvSpPr"))
    ET.SubElement(
        nv_sp_pr,
        qn(NS_P, "cNvPr"),
        {"id": str(shape_id), "name": shape_name},
    )
    ET.SubElement(nv_sp_pr, qn(NS_P, "cNvSpPr"))
    ET.SubElement(nv_sp_pr, qn(NS_P, "nvPr"))

    sp_pr = ET.SubElement(shape, qn(NS_P, "spPr"))
    xfrm = ET.SubElement(sp_pr, qn(NS_A, "xfrm"))
    ET.SubElement(xfrm, qn(NS_A, "off"), {"x": to_emu(box_x), "y": to_emu(box_y)})
    ET.SubElement(xfrm, qn(NS_A, "ext"), {"cx": to_emu(box_w), "cy": to_emu(box_h)})
    prst_geom = ET.SubElement(sp_pr, qn(NS_A, "prstGeom"), {"prst": "rect"})
    ET.SubElement(prst_geom, qn(NS_A, "avLst"))

    solid_fill = ET.SubElement(sp_pr, qn(NS_A, "solidFill"))
    ET.SubElement(solid_fill, qn(NS_A, "srgbClr"), {"val": fill_hex})
    line_style = ET.SubElement(sp_pr, qn(NS_A, "ln"))
    ET.SubElement(line_style, qn(NS_A, "noFill"))


def build_slide_xml(
    problem: Problem,
    page_fallback_number: int,
    top_right_label: str | None = TOP_RIGHT_LABEL,
) -> ET.ElementTree:
    ET.register_namespace("a", NS_A)
    ET.register_namespace("p", NS_P)
    ET.register_namespace("r", NS_R)
    resolved_top_right_label = TOP_RIGHT_LABEL if top_right_label is None else top_right_label.strip()

    number_text = problem.problem_number.strip() or str(page_fallback_number)
    top_left_text = f"{number_text}번"

    content_shapes = build_flow_content_shapes(problem)

    root = ET.Element(qn(NS_P, "sld"), {"showMasterSp": "1", "showMasterPhAnim": "1"})
    c_sld = ET.SubElement(root, qn(NS_P, "cSld"))
    bg = ET.SubElement(c_sld, qn(NS_P, "bg"))
    bg_pr = ET.SubElement(bg, qn(NS_P, "bgPr"))
    solid_fill = ET.SubElement(bg_pr, qn(NS_A, "solidFill"))
    ET.SubElement(solid_fill, qn(NS_A, "srgbClr"), {"val": "000000"})

    sp_tree = ET.SubElement(c_sld, qn(NS_P, "spTree"))
    nv_grp_sp_pr = ET.SubElement(sp_tree, qn(NS_P, "nvGrpSpPr"))
    ET.SubElement(nv_grp_sp_pr, qn(NS_P, "cNvPr"), {"id": "1", "name": ""})
    ET.SubElement(nv_grp_sp_pr, qn(NS_P, "cNvGrpSpPr"))
    ET.SubElement(nv_grp_sp_pr, qn(NS_P, "nvPr"))
    grp_sp_pr = ET.SubElement(sp_tree, qn(NS_P, "grpSpPr"))
    grp_xfrm = ET.SubElement(grp_sp_pr, qn(NS_A, "xfrm"))
    ET.SubElement(grp_xfrm, qn(NS_A, "off"), {"x": "0", "y": "0"})
    ET.SubElement(grp_xfrm, qn(NS_A, "ext"), {"cx": "0", "cy": "0"})
    ET.SubElement(grp_xfrm, qn(NS_A, "chOff"), {"x": "0", "y": "0"})
    ET.SubElement(grp_xfrm, qn(NS_A, "chExt"), {"cx": "0", "cy": "0"})

    add_text_shape(
        sp_tree,
        2,
        "top-left",
        top_left_text,
        TOP_LEFT_BOX,
        font_size_pt=TOP_LEFT_FONT_PT,
        bold=True,
        align="l",
        font_name=HEADER_FONT_NAME,
    )
    add_text_shape(
        sp_tree,
        3,
        "top-right",
        resolved_top_right_label,
        TOP_RIGHT_BOX,
        font_size_pt=TOP_RIGHT_FONT_PT,
        bold=True,
        align="r",
        font_name=HEADER_FONT_NAME,
    )
    add_fill_rect_shape(sp_tree, 4, "top-red-line", TOP_RED_LINE_BOX, "FF0000")

    for idx, (shape_name, text, box, align_note_lines_right) in enumerate(content_shapes, start=10):
        add_text_shape(
            sp_tree,
            idx,
            shape_name,
            text,
            box,
            font_size_pt=FONT_SIZE_PT,
            bold=True,
            align=CONTENT_TEXT_ALIGN,
            align_note_lines_right=align_note_lines_right,
            note_font_size_pt=NOTE_FONT_PT,
        )

    clr_map_ovr = ET.SubElement(root, qn(NS_P, "clrMapOvr"))
    ET.SubElement(clr_map_ovr, qn(NS_A, "masterClrMapping"))

    return ET.ElementTree(root)


def build_slide_rels_xml() -> ET.ElementTree:
    root = ET.Element(qn(NS_REL, "Relationships"))
    ET.SubElement(
        root,
        qn(NS_REL, "Relationship"),
        {
            "Id": "rId1",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
            "Target": "../slideLayouts/slideLayout1.xml",
        },
    )
    return ET.ElementTree(root)


def rid_to_number(value: str) -> int:
    match = re.fullmatch(r"rId(\d+)", value)
    if match is None:
        return 0
    return int(match.group(1))


def rewrite_presentation_rels(path: Path, slide_count: int) -> List[str]:
    ET.register_namespace("", NS_REL)
    tree = ET.parse(path)
    root = tree.getroot()

    rel_tag = qn(NS_REL, "Relationship")
    for rel in list(root.findall(rel_tag)):
        rel_type = rel.get("Type", "")
        if rel_type.endswith("/slide"):
            root.remove(rel)

    existing_rids = [rid_to_number(rel.get("Id", "")) for rel in root.findall(rel_tag)]
    start = max(existing_rids) + 1 if existing_rids else 1
    slide_rids: List[str] = []

    for index in range(slide_count):
        rid = f"rId{start + index}"
        slide_rids.append(rid)
        ET.SubElement(
            root,
            rel_tag,
            {
                "Id": rid,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                "Target": f"slides/slide{index + 1}.xml",
            },
        )

    tree.write(path, encoding="UTF-8", xml_declaration=True)
    return slide_rids


def rewrite_presentation_xml(path: Path, slide_rids: List[str]) -> None:
    ET.register_namespace("a", NS_A)
    ET.register_namespace("p", NS_P)
    ET.register_namespace("r", NS_R)

    tree = ET.parse(path)
    root = tree.getroot()

    sld_id_list = root.find("p:sldIdLst", NSMAP)
    if sld_id_list is None:
        sld_id_list = ET.SubElement(root, qn(NS_P, "sldIdLst"))

    for child in list(sld_id_list):
        sld_id_list.remove(child)

    for index, rid in enumerate(slide_rids):
        ET.SubElement(
            sld_id_list,
            qn(NS_P, "sldId"),
            {
                "id": str(256 + index),
                qn(NS_R, "id"): rid,
            },
        )

    tree.write(path, encoding="UTF-8", xml_declaration=True)


def rewrite_content_types(path: Path, slide_count: int) -> None:
    ET.register_namespace("", NS_CT)
    tree = ET.parse(path)
    root = tree.getroot()

    override_tag = qn(NS_CT, "Override")
    for override in list(root.findall(override_tag)):
        part_name = override.get("PartName", "")
        if part_name.startswith("/ppt/slides/slide"):
            root.remove(override)

    for index in range(slide_count):
        ET.SubElement(
            root,
            override_tag,
            {
                "PartName": f"/ppt/slides/slide{index + 1}.xml",
                "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
            },
        )

    tree.write(path, encoding="UTF-8", xml_declaration=True)


def build_pptx_without_keynote(
    problems: List[Problem],
    output_path: Path,
    top_right_label: str | None = TOP_RIGHT_LABEL,
) -> Path:
    template_path = find_pptx_template()
    with tempfile.TemporaryDirectory(prefix="pptx_build_") as temp_dir:
        temp_root = Path(temp_dir)
        with zipfile.ZipFile(template_path, "r") as zf:
            zf.extractall(temp_root)

        slides_dir = temp_root / "ppt" / "slides"
        slides_rels_dir = slides_dir / "_rels"
        slides_dir.mkdir(parents=True, exist_ok=True)
        slides_rels_dir.mkdir(parents=True, exist_ok=True)

        for old_slide in slides_dir.glob("slide*.xml"):
            old_slide.unlink()
        for old_rel in slides_rels_dir.glob("slide*.xml.rels"):
            old_rel.unlink()

        for index, problem in enumerate(problems, start=1):
            slide_tree = build_slide_xml(problem, index, top_right_label=top_right_label)
            slide_tree.write(
                slides_dir / f"slide{index}.xml",
                encoding="UTF-8",
                xml_declaration=True,
            )

            slide_rel_tree = build_slide_rels_xml()
            slide_rel_tree.write(
                slides_rels_dir / f"slide{index}.xml.rels",
                encoding="UTF-8",
                xml_declaration=True,
            )

        presentation_rels_path = temp_root / "ppt" / "_rels" / "presentation.xml.rels"
        presentation_path = temp_root / "ppt" / "presentation.xml"
        content_types_path = temp_root / "[Content_Types].xml"

        slide_rids = rewrite_presentation_rels(presentation_rels_path, len(problems))
        rewrite_presentation_xml(presentation_path, slide_rids)
        rewrite_content_types(content_types_path, len(problems))

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for file_path in sorted(temp_root.rglob("*")):
                if file_path.is_file():
                    arc_name = file_path.relative_to(temp_root).as_posix()
                    zf.write(file_path, arc_name)

    return output_path


def export_pdf_with_soffice(pptx_path: Path, pdf_path: Path) -> None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError(
            "PDF 자동 변환을 위해 LibreOffice(soffice)가 필요합니다. "
            "현재는 PPTX만 생성됩니다."
        )

    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    proc = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(pdf_path.parent),
            str(pptx_path),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    if proc.returncode != 0:
        stderr = proc.stderr.strip() or proc.stdout.strip() or "Unknown LibreOffice error"
        raise RuntimeError(f"PDF 변환 실패: {stderr}")

    converted = pdf_path.parent / f"{pptx_path.stem}.pdf"
    if not converted.exists():
        raise RuntimeError("PDF 변환 결과 파일을 찾을 수 없습니다.")
    if converted.resolve() != pdf_path.resolve():
        converted.replace(pdf_path)


def qn(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}"


def enforce_run_style(
    run_pr: ET.Element,
    *,
    font_size_pt: int = FONT_SIZE_PT,
    bold: bool = True,
    underlined: bool = False,
    font_name: str = FONT_NAME,
) -> None:
    run_pr.set("sz", str(font_size_pt * 100))
    run_pr.set("lang", "ko-KR")
    run_pr.set("b", "1" if bold else "0")
    run_pr.set("u", "sng" if underlined else "none")

    removable = {
        "solidFill",
        "noFill",
        "gradFill",
        "pattFill",
        "blipFill",
        "schemeClr",
        "srgbClr",
        "latin",
        "ea",
        "cs",
    }
    for child in list(run_pr):
        if child.tag.split("}")[-1] in removable:
            run_pr.remove(child)

    solid_fill = ET.SubElement(run_pr, qn(NS_A, "solidFill"))
    ET.SubElement(solid_fill, qn(NS_A, "srgbClr"), {"val": "FFFFFF"})
    ET.SubElement(run_pr, qn(NS_A, "latin"), {"typeface": font_name})
    ET.SubElement(run_pr, qn(NS_A, "ea"), {"typeface": font_name})
    ET.SubElement(run_pr, qn(NS_A, "cs"), {"typeface": font_name})


def enforce_paragraph_style(paragraph: ET.Element, *, align: str = "l") -> None:
    ppr = paragraph.find("a:pPr", NSMAP)
    if ppr is None:
        ppr = ET.Element(qn(NS_A, "pPr"))
        paragraph.insert(0, ppr)
    ppr.set("algn", align)

    for tag in ("lnSpc", "spcBef", "spcAft"):
        existing = ppr.find(f"a:{tag}", NSMAP)
        if existing is not None:
            ppr.remove(existing)

    ln_spc = ET.SubElement(ppr, qn(NS_A, "lnSpc"))
    ET.SubElement(ln_spc, qn(NS_A, "spcPct"), {"val": str(LINE_SPACING_PERCENT * 1000)})

    spc_bef = ET.SubElement(ppr, qn(NS_A, "spcBef"))
    ET.SubElement(spc_bef, qn(NS_A, "spcPts"), {"val": "0"})

    spc_aft = ET.SubElement(ppr, qn(NS_A, "spcAft"))
    ET.SubElement(spc_aft, qn(NS_A, "spcPts"), {"val": "0"})


def to_emu(points: int) -> str:
    return str(points * EMU_PER_PT)


# -------------------------------
# Public build API
# -------------------------------
def build_presentation_files(
    problems: List[Problem],
    pptx_output: Path,
    pdf_output: Path | None,
    top_right_label: str | None = TOP_RIGHT_LABEL,
) -> tuple[Path, Path | None]:
    """Build PPTX (and optional PDF) outputs from parsed problems."""
    paginated = paginate_problems(problems)
    if not paginated:
        raise RuntimeError("생성할 슬라이드가 없습니다.")

    built_pptx = build_pptx_without_keynote(
        paginated,
        pptx_output,
        top_right_label=top_right_label,
    )

    built_pdf: Path | None = None
    if pdf_output:
        export_pdf_with_soffice(built_pptx, pdf_output)
        built_pdf = pdf_output

    return built_pptx, built_pdf


# -------------------------------
# CLI entrypoint
# -------------------------------
def read_input(path: str | None) -> str:
    if path:
        return Path(path).read_text(encoding="utf-8")
    return sys.stdin.read()


def write_output(text: str, path: str | None) -> None:
    if path:
        Path(path).write_text(text, encoding="utf-8")
    else:
        sys.stdout.write(text)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert English problem text into PPT slide layout blocks."
    )
    parser.add_argument("input", nargs="?", help="Input text file path. If omitted, reads stdin.")
    parser.add_argument("-o", "--output", help="Output text file path. If omitted, writes stdout.")
    parser.add_argument(
        "--pptx",
        help="Output PPTX file path (pure Python generation).",
    )
    parser.add_argument(
        "--pdf",
        help="Output PDF file path (requires LibreOffice/soffice).",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    raw = read_input(args.input)
    result = format_text(raw)

    if args.pptx or args.pdf:
        problems = parse_raw(raw)
        if not problems:
            raise SystemExit("입력 텍스트가 비어 있어 PPT/PDF를 생성할 수 없습니다.")

        pptx_path = (
            Path(args.pptx) if args.pptx else Path(args.pdf).with_suffix(".pptx")
        ).expanduser().resolve()
        pdf_path = Path(args.pdf).expanduser().resolve() if args.pdf else None

        try:
            built_pptx, built_pdf = build_presentation_files(problems, pptx_path, pdf_path)
        except RuntimeError as exc:
            raise SystemExit(f"PPT/PDF 생성 실패: {exc}") from exc

        if args.output:
            write_output(result, args.output)
        else:
            sys.stdout.write(
                f"PPT 생성 완료: {built_pptx}\n"
                + (f"PDF 생성 완료: {built_pdf}\n" if built_pdf else "")
            )
        return 0

    write_output(result, args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
