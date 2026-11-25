# -*- coding: utf-8 -*-
"""TTS 모델 3.8 지원, 모델 이력 우선순위, 캐시 보존 및 종료 모델 감지 테스트."""
import pytest


def test_default_tts_model_is_gemini_38():
    from revoice import revoice_tts
    assert revoice_tts.DEFAULT_MODEL == "gemini-3.8-flash-tts"
    assert revoice_tts.is_structured_tts_model("gemini-3.8-flash-tts") is True
    assert revoice_tts.is_structured_tts_model("gemini-3.8-flash-lite-tts") is True
    assert revoice_tts.is_structured_tts_model("gemini-3.1-flash-tts-preview") is False
    assert revoice_tts.is_structured_tts_model("gemini-2.5-pro-preview-tts") is False


def test_deprecation_status_before_and_after_shutdown():
    from revoice import revoice_tts
    before = revoice_tts.check_model_status("gemini-3.1-flash-tts-preview", today="2026-11-25")
    assert before["status"] == "deprecated"
    assert before["shutdown_date"] == "2026-12-18"
    assert before["replacement"] == "gemini-3.8-flash-tts"

    after = revoice_tts.check_model_status("gemini-3.1-flash-tts-preview", today="2026-12-19")
    assert after["status"] == "retired"

    fresh = revoice_tts.check_model_status("gemini-3.8-flash-tts", today="2027-01-10")
    assert fresh["status"] == "ok"


def test_resolve_tts_model_precedence(monkeypatch):
    from revoice.build import _resolve_tts_model

    class _M:
        tts_model = ""

    m = _M()
    monkeypatch.delenv("GEMINI_TTS_MODEL", raising=False)
    assert _resolve_tts_model(m)[0] == "gemini-3.8-flash-tts"

    monkeypatch.setenv("GEMINI_TTS_MODEL", "gemini-3.8-flash-lite-tts")
    assert _resolve_tts_model(m)[0] == "gemini-3.8-flash-lite-tts"

    # 기존 매니페스트 이력이 .env보다 우선해야 부분 수정 시 음색이 보존된다.
    m.tts_model = "gemini-3.1-flash-tts-preview"
    assert _resolve_tts_model(m)[0] == "gemini-3.1-flash-tts-preview"


def test_voice_key_preserves_legacy_31_and_invalidates_on_38_migration():
    from revoice.build import _voice_key

    class _Sc:
        narration = "진행자: 질문입니다.\n전문가: 답변입니다."

    class _M:
        tts_style = "차분하게: "
        voice = "Charon"
        voices = {"진행자": "Erinome", "전문가": "Alnilam"}
        dialogue_style = "스타일:\n"
        tts_model = ""

    m_none = _M()
    m_31 = _M(); m_31.tts_model = "gemini-3.1-flash-tts-preview"
    m_38 = _M(); m_38.tts_model = "gemini-3.8-flash-tts"

    # 기존 10·11월판(모델 미지정 vs 3.1 명시)은 캐시 키가 100% 동일해야 기존 WAV가 유지된다.
    assert _voice_key(_Sc(), m_none) == _voice_key(_Sc(), m_31)
    # 3.8로 전환하면 키가 달라져 전체 씬이 일괄 재합성된다.
    assert _voice_key(_Sc(), m_38) != _voice_key(_Sc(), m_31)


def test_404_raises_model_unavailable_error():
    from revoice import revoice_tts

    class _NotFound(Exception):
        code = 404

    class _DeadClient:
        def __init__(self):
            self.models = self
        def generate_content(self, **kw):
            raise _NotFound("models/gemini-3.1-flash-tts-preview is not found")

    with pytest.raises(revoice_tts.ModelUnavailableError) as exc:
        revoice_tts._generate(_DeadClient(), model="gemini-3.1-flash-tts-preview")
    assert exc.value.replacement == "gemini-3.8-flash-tts"
