#!/usr/bin/env python3
"""Streamlit UI for generating YGA slide decks from raw workbook text."""

from __future__ import annotations

import importlib
import os
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Optional, cast

st: Any = importlib.import_module("streamlit")
components: Any = importlib.import_module("streamlit.components.v1")
_runtime_logging: Any = importlib.import_module("runtime_logging")
get_logger = _runtime_logging.get_logger

_slide_formatter: Any = importlib.import_module("slide_formatter")
DEFAULT_CONTENT_FONT_SIZE_PT = _slide_formatter.DEFAULT_FONT_SIZE_PT
Problem = _slide_formatter.Problem
build_presentation_files = _slide_formatter.build_presentation_files
paginate_problems = _slide_formatter.paginate_problems
parse_raw_details = _slide_formatter.parse_raw_details
parse_raw = _slide_formatter.parse_raw
set_content_font_size = _slide_formatter.set_content_font_size

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
DEFAULT_DOWNLOAD_NAME = "class_slides.pptx"
TEMP_PPTX_NAME = DEFAULT_DOWNLOAD_NAME
PREVIEW_TEXT_HEIGHT = 360
BUILD_SECTION_LABEL = "BUILD"
APPLY_INPUT_BUTTON_LABEL = "제출"
DOWNLOAD_NAME_INPUT_KEY = "download_file_name_input_v4"
HEADER_TITLE_INPUT_KEY = "header_title_input_v4"
FONT_SIZE_INPUT_KEY = "content_font_size_input_v1"
RAW_TEXT_INPUT_KEY = "raw_text_area_v3"
DIRECT_RUN_BOOTSTRAP_BYPASS_ENV = "YGA_SKIP_STREAMLIT_BOOTSTRAP"
LOGGER = get_logger("app")
UI_THEME_STYLE = """
<style>
  :root {
    --bg: #f6f7f9;
    --surface: #ffffff;
    --ink: #111827;
    --muted: #6b7280;
    --line: #e5e7eb;
    --accent: #1f4d7a;
    --accent-strong: #102a43;
  }

  .stApp {
    background: var(--bg);
    color: var(--ink);
    font-family: "SUIT", "Pretendard", "IBM Plex Sans KR", "Noto Sans KR", sans-serif;
  }

  [data-testid="stAppViewContainer"] {
    width: 100%;
  }

  [data-testid="stMain"] {
    width: 100%;
    align-items: stretch !important;
  }

  [data-testid="stMain"] > div,
  [data-testid="stMainBlockContainer"],
  .block-container {
    padding-top: 1.1rem;
    padding-bottom: 2rem;
    padding-left: 0.45rem;
    padding-right: 0.45rem;
    max-width: 100% !important;
    width: 100% !important;
  }

  [data-testid="stVerticalBlock"],
  [data-testid="stElementContainer"] {
    width: 100% !important;
    max-width: 100% !important;
    align-self: stretch !important;
  }

  @keyframes fadeIn {
    from { opacity: 0; transform: translateY(4px); }
    to { opacity: 1; transform: translateY(0); }
  }

  .hero-shell {
    background: var(--surface);
    border: 1px solid var(--line);
    border-left: 4px solid var(--accent);
    border-radius: 12px;
    padding: 0.9rem 1rem;
    margin-bottom: 0.45rem;
    animation: fadeIn 0.22s ease-out;
  }

  .hero-title {
    margin: 0.2rem 0 0.1rem 0;
    font-size: 1.58rem;
    line-height: 1.26;
    font-weight: 760;
    color: var(--ink);
  }

  .hero-subtitle {
    margin: 0.08rem 0 0 0;
    color: var(--muted);
    font-size: 0.92rem;
  }

  div[data-testid="stMetric"] {
    background: var(--surface);
    border: 1px solid var(--line);
    border-radius: 10px;
    padding: 0.48rem 0.62rem;
  }

  .stTextArea textarea,
  .stTextInput input,
  .stNumberInput input {
    border-radius: 10px !important;
    border: 1px solid var(--line) !important;
    background: #ffffff !important;
  }

  .stButton > button,
  .stDownloadButton > button {
    border: 1px solid var(--accent-strong) !important;
    border-radius: 10px !important;
    background: var(--accent-strong) !important;
    color: #ffffff !important;
    font-weight: 650 !important;
    box-shadow: none;
    transition: filter 0.15s ease;
  }
  .stButton > button:hover,
  .stDownloadButton > button:hover {
    filter: brightness(1.06);
  }

  .stButton > button[kind="primary"] {
    min-height: 3.4rem;
    padding: 0.9rem 1.1rem;
    font-size: 1.05rem;
    font-weight: 760 !important;
  }

  .section-label {
    margin: 0.15rem 0 0.45rem 0;
    color: #445267;
    font-size: 0.82rem;
    font-weight: 700;
    letter-spacing: 0.04em;
  }

  #yga-scroll-progress-track {
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 4px;
    background: rgba(17, 24, 39, 0.08);
    z-index: 9999;
  }

  #yga-scroll-progress {
    width: 0%;
    height: 100%;
    background: linear-gradient(90deg, var(--accent), var(--accent-strong));
    transition: width 0.12s ease-out;
  }

  #yga-scroll-top {
    position: fixed;
    right: 18px;
    bottom: 18px;
    width: 40px;
    height: 40px;
    border-radius: 999px;
    border: 1px solid #d1d5db;
    background: #ffffff;
    color: #111827;
    font-size: 18px;
    line-height: 1;
    cursor: pointer;
    box-shadow: 0 6px 14px rgba(17, 24, 39, 0.12);
    opacity: 0;
    pointer-events: none;
    transform: translateY(10px);
    transition: opacity 0.2s ease, transform 0.2s ease;
    z-index: 9999;
  }

  #yga-scroll-top.show {
    opacity: 1;
    pointer-events: auto;
    transform: translateY(0);
  }

  .scroll-reveal {
    opacity: 0;
    transform: translateY(10px);
    transition: opacity 0.32s ease, transform 0.32s ease;
  }

  .scroll-reveal.is-visible {
    opacity: 1;
    transform: translateY(0);
  }

  [data-testid="stToolbar"] {
    display: none !important;
  }

  @media (max-width: 900px) {
    .hero-title { font-size: 1.42rem; }
    .hero-shell { padding: 0.85rem 0.85rem; }
    [data-testid="stMain"] > div,
    [data-testid="stMainBlockContainer"],
    .block-container {
      padding-left: 0.75rem;
      padding-right: 0.75rem;
    }
  }
</style>
"""
HERO_HTML = """
<section class="hero-shell">
  <p class="hero-subtitle">YGA Slide Formatter</p>
  <h1 class="hero-title">영어교재 PPT 자동 생성기</h1>
</section>
"""
SCROLL_CHROME_HTML = """
<div id="yga-scroll-progress-track">
  <div id="yga-scroll-progress"></div>
</div>
<button id="yga-scroll-top" type="button" aria-label="scroll-to-top">↑</button>
"""
SCROLL_SCRIPT = """
<script>
(function () {
  const parentWindow = window.parent;
  const doc = parentWindow.document;
  if (!doc) return;

  const progress = doc.getElementById("yga-scroll-progress");
  const topBtn = doc.getElementById("yga-scroll-top");
  if (!progress || !topBtn) return;

  function applyLayoutOverrides() {
    const appView = doc.querySelector("[data-testid='stAppViewContainer']");
    const main = doc.querySelector("[data-testid='stMain']");
    const blockContainers = doc.querySelectorAll(
      "[data-testid='stMain'] > div, [data-testid='stMainBlockContainer'], .block-container"
    );
    const fullWidthNodes = doc.querySelectorAll(
      "[data-testid='stVerticalBlock'], [data-testid='stElementContainer']"
    );
    const sidePadding = parentWindow.innerWidth <= 900 ? "0.75rem" : "0.35rem";

    if (appView) {
      appView.dataset.layout = "wide";
      appView.style.width = "100%";
    }

    if (main) {
      main.style.width = "100%";
      main.style.alignItems = "stretch";
    }

    blockContainers.forEach((node) => {
      node.style.width = "100%";
      node.style.maxWidth = "100%";
      node.style.paddingLeft = sidePadding;
      node.style.paddingRight = sidePadding;
    });

    fullWidthNodes.forEach((node) => {
      node.style.width = "100%";
      node.style.maxWidth = "100%";
      node.style.alignSelf = "stretch";
    });
  }

  function updateScrollUi() {
    const de = doc.documentElement;
    const scrollTop = de.scrollTop || doc.body.scrollTop || 0;
    const scrollHeight = Math.max(de.scrollHeight, doc.body.scrollHeight) - de.clientHeight;
    const pct = scrollHeight > 0 ? (scrollTop / scrollHeight) * 100 : 0;
    progress.style.width = pct.toFixed(2) + "%";
    if (scrollTop > 240) {
      topBtn.classList.add("show");
    } else {
      topBtn.classList.remove("show");
    }
  }

  if (!parentWindow.__ygaScrollBound) {
    doc.addEventListener("scroll", updateScrollUi, { passive: true });
    parentWindow.addEventListener("scroll", updateScrollUi, { passive: true });
    parentWindow.addEventListener("resize", applyLayoutOverrides, { passive: true });
    topBtn.addEventListener("click", function () {
      parentWindow.scrollTo({ top: 0, behavior: "smooth" });
    });
    parentWindow.__ygaScrollBound = true;
  }

  const targets = doc.querySelectorAll(".hero-shell, div[data-testid='stMetric']");
  targets.forEach((node) => node.classList.add("scroll-reveal"));

  if (!parentWindow.__ygaRevealObserver) {
    parentWindow.__ygaRevealObserver = new parentWindow.IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            entry.target.classList.add("is-visible");
          }
        });
      },
      { threshold: 0.08 }
    );
  }
  const observer = parentWindow.__ygaRevealObserver;
  targets.forEach((node) => observer.observe(node));

  if (!parentWindow.__ygaLayoutObserver) {
    parentWindow.__ygaLayoutObserver = new parentWindow.MutationObserver(() => {
      applyLayoutOverrides();
    });
    parentWindow.__ygaLayoutObserver.observe(doc.body, { childList: true, subtree: true });
  }

  applyLayoutOverrides();
  updateScrollUi();
})();
</script>
"""
STATE_DEFAULTS: dict[str, Any] = {
    "raw_text": "",
    "submitted_text": "",
    "show_submit_spinner": False,
    "out_pptx": None,
    "download_name": DEFAULT_DOWNLOAD_NAME,
    "header_title": "",
    "content_font_size_pt": DEFAULT_CONTENT_FONT_SIZE_PT,
    "last_generated_text": "",
    "last_generated_header": "",
    "last_generated_font_size_pt": DEFAULT_CONTENT_FONT_SIZE_PT,
    "last_generated_output_signature": "",
    "last_generated_at": "",
}


@dataclass(frozen=True)
class ParsedInput:
    normalized_text: str
    base_problems: list[Problem]
    slide_count: int
    parser_name: str
    ai_attempted: bool
    ai_used: bool
    parse_reason: str


def render_html(markup: str) -> None:
    st.markdown(markup, unsafe_allow_html=True)


def render_section_label(label: str) -> None:
    render_html(f"<p class='section-label'>{label}</p>")


def inject_ui_theme() -> None:
    render_html(UI_THEME_STYLE)


def render_hero() -> None:
    render_html(HERO_HTML)


def inject_scroll_effects() -> None:
    render_html(SCROLL_CHROME_HTML)
    components.html(SCROLL_SCRIPT, height=0)


def create_output(problems: list[Problem], header_title: str, content_font_size_pt: int) -> bytes:
    if not problems:
        raise ValueError("입력 텍스트가 비어 있어 PPT를 생성할 수 없습니다.")

    set_content_font_size(content_font_size_pt)
    with tempfile.TemporaryDirectory(prefix="slide_gui_") as temp_dir:
        pptx_path = Path(temp_dir) / TEMP_PPTX_NAME
        built_pptx, _ = build_presentation_files(
            problems,
            pptx_path,
            None,
            top_right_label=header_title,
        )
        return built_pptx.read_bytes()


def normalize_download_name(raw_name: str) -> str:
    name = raw_name.strip()
    if not name:
        return DEFAULT_DOWNLOAD_NAME
    if name.lower().endswith(".pptx"):
        return name
    return f"{name}.pptx"


def init_state() -> None:
    for key, default in STATE_DEFAULTS.items():
        st.session_state.setdefault(key, default)


def invalidate_generated_pptx() -> None:
    st.session_state["out_pptx"] = None


def normalize_input_text(raw_text: str) -> str:
    return raw_text.strip()


def current_download_name() -> str:
    return str(st.session_state["download_name"])


def current_download_stem() -> str:
    return Path(current_download_name()).stem


def current_header_title() -> str:
    return str(st.session_state["header_title"])


def current_content_font_size() -> int:
    return int(st.session_state.get("content_font_size_pt", DEFAULT_CONTENT_FONT_SIZE_PT))


def resolve_module_path(module: Any) -> Path:
    module_path = cast(Optional[str], getattr(module, "__file__", None))
    if module_path is None:
        module_spec = getattr(module, "__spec__", None)
        module_path = cast(Optional[str], getattr(module_spec, "origin", None))
    if module_path is None:
        raise RuntimeError("Unable to resolve module path.")
    return Path(module_path).resolve()


def current_output_signature() -> str:
    app_mtime = str(Path(__file__).resolve().stat().st_mtime_ns)
    formatter_mtime = str(resolve_module_path(_slide_formatter).stat().st_mtime_ns)
    return f"{app_mtime}:{formatter_mtime}"


def current_submitted_text() -> str:
    return str(st.session_state["submitted_text"])


def update_download_name(raw_name: str) -> None:
    st.session_state["download_name"] = normalize_download_name(raw_name)


def update_header_title(raw_title: str) -> None:
    st.session_state["header_title"] = raw_title.strip()


def update_content_font_size(raw_size: int) -> None:
    size = max(16, int(raw_size))
    if int(st.session_state.get("content_font_size_pt", DEFAULT_CONTENT_FONT_SIZE_PT)) != size:
        st.session_state["content_font_size_pt"] = size
        invalidate_generated_pptx()
    set_content_font_size(size)


def generated_output_is_stale(normalized_text: str) -> bool:
    return (
        st.session_state["last_generated_text"] != normalized_text
        or st.session_state["last_generated_header"] != current_header_title()
        or int(st.session_state.get("last_generated_font_size_pt", DEFAULT_CONTENT_FONT_SIZE_PT))
        != current_content_font_size()
        or st.session_state.get("last_generated_output_signature", "") != current_output_signature()
    )


def raw_text_has_pending_changes(raw_text: str, submitted_text: str) -> bool:
    return normalize_input_text(raw_text) != normalize_input_text(submitted_text)


def apply_raw_text(raw_text: str) -> None:
    st.session_state["submitted_text"] = normalize_input_text(raw_text)
    st.session_state["show_submit_spinner"] = True
    invalidate_generated_pptx()


def render_download_name_input() -> None:
    file_name = st.text_input(
        "다운로드 파일명",
        value=current_download_stem(),
        key=DOWNLOAD_NAME_INPUT_KEY,
        placeholder="예: class_slides",
        help="확장자(.pptx)는 자동 보정됩니다.",
    )
    update_download_name(file_name)


def render_header_title_input() -> None:
    header_title = st.text_input(
        "교재 제목 (우측 머리말)",
        value=current_header_title(),
        key=HEADER_TITLE_INPUT_KEY,
        placeholder="예: YGA 2026 KO Reading",
        help="비워두면 우측 머리말 텍스트를 넣지 않습니다. 좌측 번호는 유지됩니다.",
    )
    update_header_title(header_title)


def render_font_size_input() -> None:
    font_size = st.slider(
        "본문 글자 크기",
        min_value=20,
        max_value=60,
        value=current_content_font_size(),
        step=1,
        key=FONT_SIZE_INPUT_KEY,
        help="PPT 본문/문제/선지 글자 크기입니다.",
    )
    update_content_font_size(font_size)


def render_controls() -> None:
    settings_col1, settings_col2, settings_col3 = st.columns(3)
    with settings_col1:
        render_download_name_input()

    with settings_col2:
        render_header_title_input()

    with settings_col3:
        render_font_size_input()


def render_raw_text_input() -> str:
    raw_text = st.text_area(
        "원문 텍스트",
        value=st.session_state["raw_text"],
        height=PREVIEW_TEXT_HEIGHT,
        key=RAW_TEXT_INPUT_KEY,
    )
    st.session_state["raw_text"] = raw_text

    has_pending_changes = raw_text_has_pending_changes(raw_text, current_submitted_text())
    action_col, hint_col = st.columns([1, 3])
    with action_col:
        if st.button(
            APPLY_INPUT_BUTTON_LABEL,
            use_container_width=True,
            disabled=not has_pending_changes,
        ):
            apply_raw_text(raw_text)

    with hint_col:
        if has_pending_changes and normalize_input_text(raw_text):
            st.caption(f"텍스트 수정 후에는 `{APPLY_INPUT_BUTTON_LABEL}` 버튼을 눌러야 결과가 갱신됩니다.")
        elif current_submitted_text():
            st.caption(f"결과 갱신은 `{APPLY_INPUT_BUTTON_LABEL}` 버튼으로만 진행됩니다.")

    return current_submitted_text()


def parse_input_payload(normalized_text: str) -> Optional[ParsedInput]:
    if not normalized_text:
        st.session_state["show_submit_spinner"] = False
        st.info("텍스트를 입력하세요.")
        return None

    set_content_font_size(current_content_font_size())
    try:
        show_submit_spinner = bool(st.session_state.get("show_submit_spinner", False))
        if show_submit_spinner:
            with st.spinner("입력 분석 중..."):
                parse_result = parse_raw_details(normalized_text)
                base_problems = parse_result.problems
                slide_count = len(paginate_problems(base_problems))
        else:
            parse_result = parse_raw_details(normalized_text)
            base_problems = parse_result.problems
            slide_count = len(paginate_problems(base_problems))
    except Exception as exc:
        st.error(f"입력 파싱 실패: {exc}")
        LOGGER.exception("input_parse_failed input_chars=%d", len(normalized_text))
        return None
    finally:
        st.session_state["show_submit_spinner"] = False

    if not base_problems:
        st.warning("입력에서 문제를 인식하지 못했습니다. 문제/지문/선지 형식을 다시 확인하세요.")
        return None

    LOGGER.info(
        "input_parse_ready parser=%s ai_attempted=%s ai_used=%s reason=%s problems=%d slides=%d input_chars=%d",
        parse_result.parser_name,
        parse_result.ai_attempted,
        parse_result.ai_used,
        parse_result.reason,
        len(base_problems),
        slide_count,
        len(normalized_text),
    )
    return ParsedInput(
        normalized_text=normalized_text,
        base_problems=base_problems,
        slide_count=slide_count,
        parser_name=parse_result.parser_name,
        ai_attempted=parse_result.ai_attempted,
        ai_used=parse_result.ai_used,
        parse_reason=parse_result.reason,
    )


def render_problem_metrics(payload: ParsedInput) -> None:
    metric_col1, metric_col2, metric_col3 = st.columns(3)
    metric_col1.metric("문제 수", len(payload.base_problems))
    metric_col2.metric("예상 슬라이드 수", payload.slide_count)
    metric_col3.metric("입력 글자 수", len(payload.normalized_text))


def store_generated_output(pptx_bytes: bytes, normalized_text: str) -> None:
    st.session_state["out_pptx"] = pptx_bytes
    st.session_state["last_generated_text"] = normalized_text
    st.session_state["last_generated_header"] = current_header_title()
    st.session_state["last_generated_font_size_pt"] = current_content_font_size()
    st.session_state["last_generated_output_signature"] = current_output_signature()
    st.session_state["last_generated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def render_generated_download(normalized_text: str) -> None:
    if st.session_state["out_pptx"] is None:
        return

    if generated_output_is_stale(normalized_text):
        st.warning("입력/제목이 바뀌었습니다. 최신 내용으로 다시 `PPT 생성` 후 다운로드하세요.")

    last_generated_at = st.session_state["last_generated_at"] or "-"
    st.caption(f"마지막 생성 시각: {last_generated_at}")
    st.download_button(
        "PPTX 다운로드",
        data=st.session_state["out_pptx"],
        file_name=current_download_name(),
        mime=PPTX_MIME,
        use_container_width=True,
    )


def handle_ppt_generation(payload: ParsedInput) -> None:
    if not st.button("PPT 생성", type="primary", use_container_width=True):
        return

    LOGGER.info(
        "ppt_generation_started parser=%s ai_used=%s problems=%d slide_estimate=%d",
        payload.parser_name,
        payload.ai_used,
        len(payload.base_problems),
        payload.slide_count,
    )
    try:
        with st.spinner("PPT 생성 중..."):
            pptx_bytes = create_output(
                payload.base_problems,
                current_header_title(),
                current_content_font_size(),
            )
    except Exception as exc:
        st.error(f"생성 실패: {exc}")
        LOGGER.exception(
            "ppt_generation_failed parser=%s ai_used=%s problems=%d",
            payload.parser_name,
            payload.ai_used,
            len(payload.base_problems),
        )
        invalidate_generated_pptx()
    else:
        st.success("생성 완료")
        LOGGER.info(
            "ppt_generation_completed parser=%s ai_used=%s problems=%d output_bytes=%d",
            payload.parser_name,
            payload.ai_used,
            len(payload.base_problems),
            len(pptx_bytes),
        )
        store_generated_output(pptx_bytes, payload.normalized_text)


def render_build_tab(payload: ParsedInput) -> None:
    render_section_label(BUILD_SECTION_LABEL)
    handle_ppt_generation(payload)
    render_generated_download(payload.normalized_text)


def render_results(payload: ParsedInput) -> None:
    render_problem_metrics(payload)
    render_build_tab(payload)


def ensure_streamlit_runtime() -> None:
    if st.runtime.exists() or os.environ.get(DIRECT_RUN_BOOTSTRAP_BYPASS_ENV) == "1":
        return

    script_path = str(Path(__file__).resolve())
    try:
        os.execvp(
            sys.executable,
            [sys.executable, "-m", "streamlit", "run", script_path],
        )
    except OSError as exc:
        raise SystemExit(
            "Streamlit 서버를 시작하지 못했습니다. "
            "`python3 -m streamlit run app.py`로 직접 실행해 보세요."
        ) from exc


def main() -> None:
    st.set_page_config(page_title="YGA Slide Formatter", layout="wide")
    init_state()
    set_content_font_size(current_content_font_size())
    inject_ui_theme()
    inject_scroll_effects()

    render_hero()
    render_controls()

    payload = parse_input_payload(render_raw_text_input())
    if payload is None:
        return

    render_results(payload)


if __name__ == "__main__":
    ensure_streamlit_runtime()
    main()
