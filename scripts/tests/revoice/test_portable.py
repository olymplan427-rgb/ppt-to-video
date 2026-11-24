# -*- coding: utf-8 -*-
"""Windows/macOS 크로스플랫폼 이식성(.env, deck 폴백, 폰트 폴백, 발음 사전) 테스트."""
import os
import pytest


def test_dotenv_loader_respects_existing_env(tmp_path, monkeypatch):
    from revoice import env
    dot = tmp_path / ".env"
    dot.write_text(
        "# comment\n"
        "TEST_PORTABLE_A=from_dotenv\n"
        "TEST_PORTABLE_B=\"quoted value\"\n"
        "TEST_PORTABLE_KEEP=ignored\n",
        encoding="utf-8",
    )
    monkeypatch.delenv("TEST_PORTABLE_A", raising=False)
    monkeypatch.delenv("TEST_PORTABLE_B", raising=False)
    monkeypatch.setenv("TEST_PORTABLE_KEEP", "from_shell")
    env._LOADED_PATHS.discard(str(dot.resolve()))

    loaded = env.load_dotenv(str(tmp_path))
    assert loaded == str(dot.resolve())
    assert os.environ["TEST_PORTABLE_A"] == "from_dotenv"
    assert os.environ["TEST_PORTABLE_B"] == "quoted value"
    assert os.environ["TEST_PORTABLE_KEEP"] == "from_shell"


def test_parse_dialogue_supports_custom_speakers():
    from revoice.revoice_tts import parse_dialogue
    text = "호스트: 반갑습니다.\n게스트: 안녕하세요."
    assert parse_dialogue(text, speakers=["호스트", "게스트"]) == [
        ("호스트", "반갑습니다."),
        ("게스트", "안녕하세요."),
    ]


def test_pronounce_override_applies_extra_map():
    from revoice.revoice_tts import _apply_tts_pronounce
    out = _apply_tts_pronounce("ABC 서비스 채점표", {"ABC": "에이비씨", "채점표": "채쩜표"})
    assert out == "에이비씨 서비스 채쩜표"


def test_export_slides_reuse_backend(tmp_path):
    from revoice import deck
    (tmp_path / "slide_01.png").write_bytes(b"png")
    (tmp_path / "slide_02.png").write_bytes(b"png")
    res = deck.export_slides("dummy.pptx", str(tmp_path), backend="reuse")
    assert [os.path.basename(p) for p in res] == ["slide_01.png", "slide_02.png"]


def test_render_animated_slides_returns_empty_without_powerpoint(monkeypatch, tmp_path):
    from revoice import deck
    monkeypatch.setattr(deck, "has_powerpoint", lambda: False)
    assert deck.render_animated_slides("dummy.pptx", {1: 5.0}, str(tmp_path)) == {}


def test_subtitle_font_fallback_when_bundled_font_missing(monkeypatch, tmp_path):
    from revoice import compose
    monkeypatch.setattr(compose, "_FONTS_DIR", str(tmp_path / "missing_fonts"))
    out = compose.create_scene_srt("테스트 문장입니다.", 3.0, str(tmp_path / "s.srt"))
    body = open(out, encoding="utf-8").read()
    assert "Wanted Sans ExtraBold" not in body
    assert compose._subtitle_font_name() in body
    assert "fontsdir" not in compose.ass_filter(out)
