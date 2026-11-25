# -*- coding: utf-8 -*-
"""VO/SLIDE 씬 대본 → Gemini TTS WAV 합성 (1인 나레이션 & 2인 팟캐스트).

- 신규 전체 제작 기본 모델: `gemini-3.8-flash-tts`
- 기존 작업 부분 수정: 매니페스트(`tts_model`)에 기록된 기존 모델을 그대로 유지해 음색 일관성 확보
- 3.8 계열 모델은 다화자 요청 시 `speechMetadata.speaker` 구조화 형식을 자동 적용하고,
  1인 나레이션 시에도 스타일 프롬프트를 낭독하지 않도록 본문만 전송한다.
"""
import base64
import datetime
import json
import os
import re
import struct
import time
import urllib.error
import urllib.request
import wave

from revoice.env import get_api_key as _get_api_key, load_dotenv as _load_dotenv

_load_dotenv()

# 신규 전체 제작 기본 모델 (Google 공식 권장 안정화 모델)
DEFAULT_MODEL = "gemini-3.8-flash-tts"
# 기존 10·11월판 캐시 호환 판별용 레거시 모델
LEGACY_DEFAULT_MODEL = "gemini-3.1-flash-tts-preview"

# 모델 ID는 API의 ListModels로 확인한 값(추측 금지 — 세대마다 표기 규칙이 다르다).
# 일일 쿼터가 `per_model_per_day`라 모델을 바꾸면 쿼터 버킷도 갈린다.
MODEL = os.environ.get("GEMINI_TTS_MODEL", DEFAULT_MODEL)
SR = 24000
_BEAT = re.compile(r"\[BEAT\s+([0-9.]+)s?\]")

# 종료 예정(또는 종료된) 모델 안내표: model_id -> (종료일 YYYY-MM-DD, 권장 대체 모델)
DEPRECATED_MODELS = {
    "gemini-3.1-flash-tts-preview": ("2026-12-18", DEFAULT_MODEL),
}


class ModelUnavailableError(RuntimeError):
    """지정된 TTS 모델이 종료되었거나(404/retired) 사용할 수 없을 때 발생."""

    def __init__(self, model: str, detail: str = "", replacement: str = DEFAULT_MODEL):
        self.model = model
        self.detail = detail
        self.replacement = replacement
        super().__init__(
            f"TTS 모델 '{model}'을(를) 사용할 수 없습니다({detail}). "
            f"기존 씬들과 음색이 섞이지 않도록 매니페스트의 `tts_model`을 "
            f"'{replacement}'로 변경한 뒤 전체 씬을 새로 생성하는 것을 권장합니다."
        )


def is_structured_tts_model(model: str) -> bool:
    """Gemini 3.8+ 계열 TTS 모델 여부 (speechMetadata.speaker 구조화 규격 사용)."""
    m = (model or "").lower()
    if "2.5" in m or "3.1" in m:
        return False
    return "3.8" in m or "gemini-4" in m


def check_model_status(model: str, today: str = None) -> dict:
    """모델의 종료 예정/종료 여부를 반환: status ∈ {'ok', 'deprecated', 'retired'}."""
    info = DEPRECATED_MODELS.get(model or "")
    if not info:
        return {"model": model, "status": "ok", "shutdown_date": "", "replacement": ""}
    shutdown_date, replacement = info
    today_str = today or datetime.date.today().isoformat()
    status = "retired" if today_str > shutdown_date else "deprecated"
    return {
        "model": model,
        "status": status,
        "shutdown_date": shutdown_date,
        "replacement": replacement,
    }


def get_api_key():
    return _get_api_key()


def _client():
    from google import genai
    key = get_api_key()
    if not key:
        raise SystemExit("GEMINI_API_KEY 환경변수를 설정하세요.")
    return genai.Client(api_key=key)


# 일시적 서버 오류 — 재시도하면 대개 통과한다. 4xx 중 429(쿼터)만 재시도 대상.
_RETRY_CODES = (429, 500, 502, 503, 504)


def _is_model_gone_error(code, msg: str) -> bool:
    low = (msg or "").lower()
    if code == 404:
        return True
    return any(tok in low for tok in (
        "is not found for api version",
        "model is not found",
        "has been deprecated",
        "has been discontinued",
        "no longer available",
        "retired",
    ))


def _generate(client, **kw):
    """TTS 호출 + 일시 오류 백오프 재시도 + 모델 종료(404) 감지."""
    delay = 5.0
    model = kw.get("model", MODEL)
    for attempt in range(1, 7):
        try:
            return client.models.generate_content(**kw)
        except Exception as e:
            code = getattr(e, "code", None) or getattr(e, "status_code", None)
            if _is_model_gone_error(code, str(e)):
                raise ModelUnavailableError(model, str(e)) from e
            if code not in _RETRY_CODES or attempt == 6:
                raise
            print(f"[tts] {code} 일시 오류 — {delay:.0f}초 후 재시도 ({attempt}/5)")
            time.sleep(delay)
            delay = min(delay * 2, 120)


def _generate_structured_dialogue_rest(model: str, block: list, voices: dict,
                                       pronounce: dict = None) -> bytes:
    """Gemini 3.8+ 다화자 TTS — `speechMetadata.speaker` 구조화 REST 호출 + 재시도."""
    key = get_api_key()
    if not key:
        raise SystemExit("GEMINI_API_KEY 환경변수를 설정하세요.")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
    parts = [
        {"text": _apply_tts_pronounce(t, pronounce), "speechMetadata": {"speaker": w}}
        for w, t in block
    ]
    payload = {
        "contents": [{"role": "user", "parts": parts}],
        "generationConfig": {
            "responseModalities": ["AUDIO"],
            "speechConfig": {
                "multiSpeakerVoiceConfig": {
                    "speakerVoiceConfigs": [
                        {"speaker": spk, "voiceConfig": {"prebuiltVoiceConfig": {"voiceName": v}}}
                        for spk, v in voices.items()
                    ]
                }
            },
        },
    }
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    delay = 5.0
    for attempt in range(1, 7):
        req = urllib.request.Request(
            url, data=body, headers={"Content-Type": "application/json"}, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=180) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            b64 = data["candidates"][0]["content"]["parts"][0]["inlineData"]["data"]
            return base64.b64decode(b64)
        except urllib.error.HTTPError as e:
            err_text = e.read().decode("utf-8", errors="replace")
            if _is_model_gone_error(e.code, err_text):
                raise ModelUnavailableError(model, f"HTTP {e.code}: {err_text[:200]}") from e
            if e.code not in _RETRY_CODES or attempt == 6:
                raise RuntimeError(f"TTS HTTP {e.code}: {err_text[:300]}") from e
            print(f"[tts] {e.code} 일시 오류 — {delay:.0f}초 후 재시도 ({attempt}/5)")
            time.sleep(delay)
            delay = min(delay * 2, 120)
        except urllib.error.URLError as e:
            if attempt == 6:
                raise
            print(f"[tts] 네트워크 일시 오류 — {delay:.0f}초 후 재시도 ({attempt}/5)")
            time.sleep(delay)
            delay = min(delay * 2, 120)


def _strip_annotations(text: str) -> str:
    # [영상]/⚠[확인] 등 지시문 제거, [BEAT]는 유지(무음 처리)
    text = re.sub(r"[⚠]?\[확인[^\]]*\]", "", text)
    text = re.sub(r"\[영상[^\]]*\]", "", text)
    return text.strip()


def synth_scene(text: str, style: str, voice: str, out_wav: str,
                pronounce: dict = None, model: str = None) -> str:
    from google.genai import types
    mdl = model or MODEL
    client = _client()
    clean = _strip_annotations(text)
    # [BEAT n]를 무음 마커로 분리
    chunks = _BEAT.split(clean)  # [text, secs, text, secs, ...]
    pcm = bytearray()
    structured = is_structured_tts_model(mdl)
    for i, part in enumerate(chunks):
        if i % 2 == 1:  # BEAT seconds
            silence = int(float(part) * SR)
            pcm += struct.pack("<%dh" % silence, *([0] * silence))
            continue
        spoken = _apply_tts_pronounce(part.strip(), pronounce)
        if not spoken:
            continue
        # 3.8+ 모델은 contents 전체를 발화 텍스트로 읽으므로 style 지시문을 붙이지 않는다.
        contents = spoken if structured else (style + spoken)
        resp = _generate(
            client,
            model=mdl,
            contents=contents,
            config=types.GenerateContentConfig(
                response_modalities=["AUDIO"],
                speech_config=types.SpeechConfig(
                    voice_config=types.VoiceConfig(
                        prebuilt_voice_config=types.PrebuiltVoiceConfig(voice_name=voice)
                    )
                ),
            ),
        )
        pcm += resp.candidates[0].content.parts[0].inline_data.data
    os.makedirs(os.path.dirname(out_wav) or ".", exist_ok=True)
    with wave.open(out_wav, "wb") as wf:
        wf.setnchannels(1); wf.setsampwidth(2); wf.setframerate(SR)
        wf.writeframes(bytes(pcm))
    return out_wav


# ------------------------------------------------------- 다화자(팟캐스트) 합성

# 진행자·전문가 2인 대화. Gemini TTS 다화자 모드는 화자 2명이 상한이라 이 배합이 끝이다.
DIALOGUE_VOICES = {"진행자": "Erinome", "전문가": "Alnilam"}

# 2.5 / 3.1 레거시 모델용 화자 일관성 프롬프트 (짧게 유지할 것).
DIALOGUE_STYLE = (
    "프레젠테이션·사업설명회를 대화로 진행하는 팟캐스트입니다. "
    "진행자는 30대 후반 여성 사회자로 또렷하고 침착하게 묻고, "
    "전문가는 40대 남성 실무 책임자로 중저음의 신뢰감 있는 목소리로 답합니다. "
    "단조롭게 낭독하지 말고 숫자와 핵심 주장에 힘을 주며, 실제 대화처럼 자연스럽게 "
    "주고받아 주세요:\n"
)

_SPEAKER = re.compile(r"^(진행자|전문가)\s*[:：]\s*(.+)$")

# 자막 표기는 그대로 유지하되 TTS 발음만 자연스럽게 보정하는 치환표.
_TTS_PRONOUNCE = ()


def _apply_tts_pronounce(text: str, extra: dict = None) -> str:
    for src, dst in _TTS_PRONOUNCE:
        text = text.replace(src, dst)
    if extra:
        for src, dst in extra.items():
            text = text.replace(src, dst)
    return text


def parse_dialogue(text: str, speakers=None) -> list:
    """대본 블록 → [(화자, 대사)] 순서 목록. 화자가 "BEAT"면 값은 무음 초."""
    pat = _SPEAKER
    if speakers:
        escaped = "|".join(re.escape(s) for s in speakers)
        pat = re.compile(rf"^({escaped})\s*[:：]\s*(.+)$")
    turns = []
    for raw in _strip_annotations(text).splitlines():
        line = raw.strip()
        if not line or line[0] in ">#|" or line.startswith("**"):
            continue
        beat = _BEAT.fullmatch(line)
        if beat:
            turns.append(("BEAT", beat.group(1)))
            continue
        mo = pat.match(line)
        if mo:
            turns.append((mo.group(1), mo.group(2).strip()))
    return turns


def _trim_trailing_silence(pcm: bytearray, label: str = "") -> bytearray:
    """끝에 붙은 과도한 무음(3.1 TTS 제로 패딩 등)을 자연스럽게 압축한다."""
    import array
    a = array.array("h")
    a.frombytes(bytes(pcm[:len(pcm) // 2 * 2]))
    if not a:
        return pcm

    win = int(0.05 * SR)
    thresh = 200
    keep_gap = int(0.45 * SR)
    loud = [max(max(a[i:i + win]), -min(a[i:i + win])) >= thresh
            for i in range(0, len(a) - win, win)]
    if not any(loud):
        return pcm

    out = array.array("h")
    run = 0
    for k, is_loud in enumerate(loud):
        seg = a[k * win:(k + 1) * win]
        if is_loud:
            if run:
                out.extend([0] * min(run * win, keep_gap))
                run = 0
            out.extend(seg)
        else:
            run += 1
    out.extend([0] * min(int(0.25 * SR), keep_gap))

    cut = (len(a) - len(out)) / SR
    if cut > 3.0:
        print(f"[tts] {label or '출력'}: 과도한 공백 {cut:.0f}초 압축 "
              f"({len(a)/SR:.0f}초 → {len(out)/SR:.0f}초)")
    return bytearray(out.tobytes())


def synth_dialogue(text: str, style: str = "", out_wav: str = "",
                   voices: dict = None, pronounce: dict = None,
                   model: str = None) -> str:
    """진행자·전문가 대화를 한 WAV로 합성 — 2.5/3.1(SDK) 및 3.8+(speechMetadata) 자동 분기."""
    mdl = model or MODEL
    voices = voices or DIALOGUE_VOICES
    turns = parse_dialogue(text, speakers=list(voices.keys()))
    if not turns or all(w == "BEAT" for w, _ in turns):
        raise ValueError("대사를 찾지 못했습니다 — '진행자:' / '전문가:' 형식인지 확인하세요.")

    structured = is_structured_tts_model(mdl)
    client = None
    cfg = None
    if not structured:
        from google.genai import types
        client = _client()
        speaker_cfg = [
            types.SpeakerVoiceConfig(
                speaker=name,
                voice_config=types.VoiceConfig(
                    prebuilt_voice_config=types.PrebuiltVoiceConfig(voice_name=voice)))
            for name, voice in voices.items()
        ]
        speech = types.SpeechConfig(
            multi_speaker_voice_config=types.MultiSpeakerVoiceConfig(
                speaker_voice_configs=speaker_cfg))
        cfg = types.GenerateContentConfig(
            response_modalities=["AUDIO"], speech_config=speech)

    pcm = bytearray()
    block = []
    for who, content in turns + [("BEAT", "0")]:   # 끝 sentinel — 마지막 블록도 flush
        if who != "BEAT":
            block.append((who, content))
            continue
        if block:
            if structured:
                pcm.extend(_generate_structured_dialogue_rest(mdl, block, voices, pronounce))
            else:
                script = "\n".join(f"{w}: {_apply_tts_pronounce(t, pronounce)}" for w, t in block)
                resp = _generate(
                    client, model=mdl,
                    contents=(style or DIALOGUE_STYLE) + script, config=cfg)
                pcm.extend(resp.candidates[0].content.parts[0].inline_data.data)
            block = []
        secs = float(content)
        if secs > 0:
            n = int(secs * SR)
            pcm.extend(struct.pack("<%dh" % n, *([0] * n)))

    pcm = _trim_trailing_silence(pcm, out_wav)

    spoken_chars = sum(len(t) for w, t in turns if w != "BEAT")
    dur = len(pcm) / 2 / SR
    if spoken_chars and dur > max(60.0, spoken_chars / 3.0):
        raise RuntimeError(
            "TTS 출력이 비정상적으로 깁니다(%.0f초 / 대본 %d자). 모델이 긴 침묵을 "
            "삽입했을 가능성이 큽니다 — 톤 지시에 멈춤 관련 문구가 들어갔는지 확인하세요."
            % (dur, spoken_chars))
    os.makedirs(os.path.dirname(out_wav) or ".", exist_ok=True)
    with wave.open(out_wav, "wb") as wf:
        wf.setnchannels(1); wf.setsampwidth(2); wf.setframerate(SR)
        wf.writeframes(bytes(pcm))
    return out_wav


def revoice_manifest(m, audio_dir: str = "audio") -> dict:
    out = {}
    mdl = getattr(m, "tts_model", "") or MODEL
    for sc in m.scenes:
        if sc.type != "VO":
            continue
        wav = os.path.join(audio_dir, f"{sc.id}.wav")
        synth_scene(sc.narration, m.tts_style, m.voice, wav, model=mdl)
        out[sc.id] = wav
        print(f"[revoice] {sc.id} → {wav}")
    return out
