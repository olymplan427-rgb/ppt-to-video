"""VO 씬 대본 → Charon TTS WAV (API). tts_gemini.py 설정 재사용."""
import os
import re
import wave
import struct

from revoice.env import get_api_key as _get_api_key, load_dotenv as _load_dotenv

# 모델 ID는 API의 ListModels로 확인한 값(추측 금지 — 세대마다 표기 규칙이 다르다).
# 일일 쿼터가 `per_model_per_day`라 모델을 바꾸면 쿼터 버킷도 갈린다.
_load_dotenv()
MODEL = os.environ.get("GEMINI_TTS_MODEL", "gemini-3.1-flash-tts-preview")
SR = 24000
_BEAT = re.compile(r"\[BEAT\s+([0-9.]+)s?\]")


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


def _generate(client, **kw):
    """TTS 호출 + 일시 오류 백오프 재시도.

    한 영상에 수십 번 연속 호출하는데, 그중 한 번이라도 503(과부하)이 오면 렌더가
    통째로 죽는다. 이미 만든 음성은 캐시가 지켜 주지만 사람이 매번 다시 실행해 줘야
    하는 건 파이프라인이 아니다. 실제로 44씬 렌더가 22번째 호출에서 죽어 추가했다.
    """
    import time
    delay = 5.0
    for attempt in range(1, 7):
        try:
            return client.models.generate_content(**kw)
        except Exception as e:
            code = getattr(e, "code", None) or getattr(e, "status_code", None)
            if code not in _RETRY_CODES or attempt == 6:
                raise
            print(f"[tts] {code} 일시 오류 — {delay:.0f}초 후 재시도 ({attempt}/5)")
            time.sleep(delay)
            delay = min(delay * 2, 120)


def _strip_annotations(text: str) -> str:
    # [영상]/⚠[확인] 등 지시문 제거, [BEAT]는 유지(무음 처리)
    text = re.sub(r"[⚠]?\[확인[^\]]*\]", "", text)
    text = re.sub(r"\[영상[^\]]*\]", "", text)
    return text.strip()


def synth_scene(text: str, style: str, voice: str, out_wav: str,
                pronounce: dict = None) -> str:
    from google.genai import types
    client = _client()
    clean = _strip_annotations(text)
    # [BEAT n]를 무음 마커로 분리
    chunks = _BEAT.split(clean)  # [text, secs, text, secs, ...]
    pcm = bytearray()
    for i, part in enumerate(chunks):
        if i % 2 == 1:  # BEAT seconds
            silence = int(float(part) * SR)
            pcm += struct.pack("<%dh" % silence, *([0] * silence))
            continue
        spoken = _apply_tts_pronounce(part.strip(), pronounce)
        if not spoken:
            continue
        resp = _generate(
            client,
            model=MODEL,
            contents=style + spoken,
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
    os.makedirs(os.path.dirname(out_wav), exist_ok=True)
    with wave.open(out_wav, "wb") as wf:
        wf.setnchannels(1); wf.setsampwidth(2); wf.setframerate(SR)
        wf.writeframes(bytes(pcm))
    return out_wav


# ------------------------------------------------------- 다화자(팟캐스트) 합성

# 진행자·전문가 2인 대화. Gemini TTS 다화자 모드는 화자 2명이 상한이라 이 배합이 끝이다.
DIALOGUE_VOICES = {"진행자": "Erinome", "전문가": "Alnilam"}

# 씬마다 API를 따로 호출하므로 같은 보이스라도 톤·속도·음색이 씬 사이에서 흘러간다
# (사용자 지적 2026-08-23: "초반에 유지되다가 목소리가 바뀌고 말하는 스타일도 변한다").
# 시드 파라미터가 없으므로 **인물 설정을 매 호출 똑같이 적어 넣어** 드리프트를 붙잡는다.
# 나이·성별·음역·말버릇까지 고정해야 회차 간 편차가 줄어든다.
# ⚠ 이 프롬프트는 **짧게 유지할 것.** 인물 설정과 읽기 지시를 불릿 목록으로 길게(600자)
# 넣었더니 3.1 TTS가 무너졌다 — 씬마다 43분짜리 버퍼를 반환하고 그 안에 대사를 흩뿌리며,
# 첫 대사를 빠뜨리고 중간 문장을 중복시키고 마지막을 잘라먹었다(2026-08-23 실측).
# 짧은 프롬프트에서는 같은 모델이 정상 동작한다. 화자 일관성에 필요한 최소한만 담는다.
# 멈춤·쉼 관련 지시도 넣지 말 것(모델이 문자 그대로 받아 분 단위 침묵을 만든다).
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
    """대본 블록 → [(화자, 대사)] 순서 목록. 화자가 "BEAT"면 값은 무음 초.

    `> (인터뷰 원음 발췌 …)` 인용 줄은 CLIP 원음이 재생될 자리라 TTS 대상이 아니므로
    버린다. 마크다운 헤더·**화면** 줄도 마찬가지. `[BEAT n]`은 synth_scene과 같은
    규칙으로 무음 마커로 살린다. 화자 표기가 하나도 없으면 빈 목록을 돌려주므로,
    호출부는 이걸로 1인 나레이션인지 대화인지 판별할 수 있다.
    """
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
    """끝에 붙은 완전 무음(값이 정확히 0)을 잘라낸다.

    3.1 TTS가 이따금 실제 발화 뒤에 거대한 제로 패딩을 붙여 돌려준다 — 실측으로
    씬마다 정확히 125,827,244바이트(43.7분)가 나왔고, 앞부분 20초만 진짜 말이고
    나머지는 전부 0이었다. 그대로 두면 씬 하나가 43분짜리가 돼 영상이 망가진다.
    값이 '정확히 0'인 구간만 자르므로 실제 녹음의 조용한 대목은 건드리지 않는다
    (BEAT로 넣은 무음도 끝에 오지 않는 한 보존된다).
    """
    import array
    a = array.array("h")
    a.frombytes(bytes(pcm[:len(pcm) // 2 * 2]))
    if not a:
        return pcm

    win = int(0.05 * SR)                 # 50ms 창
    thresh = 200                         # 32767 기준 약 -44dB. 말소리엔 한참 못 미친다
    keep_gap = int(0.45 * SR)            # 대사 사이에 남길 자연스러운 간격
    loud = [max(max(a[i:i + win]), -min(a[i:i + win])) >= thresh
            for i in range(0, len(a) - win, win)]
    if not any(loud):
        return pcm

    out = array.array("h")
    run = 0                              # 연속된 무음 창 개수
    for k, is_loud in enumerate(loud):
        seg = a[k * win:(k + 1) * win]
        if is_loud:
            if run:                      # 무음 구간이 끝났다 — 짧게 줄여서 이어 붙인다
                out.extend([0] * min(run * win, keep_gap))
                run = 0
            out.extend(seg)
        else:
            run += 1
    out.extend([0] * min(int(0.25 * SR), keep_gap))     # 끝맺음 여유

    cut = (len(a) - len(out)) / SR
    if cut > 3.0:
        print(f"[tts] {label or '출력'}: 과도한 공백 {cut:.0f}초 압축 "
              f"({len(a)/SR:.0f}초 → {len(out)/SR:.0f}초)")
    return bytearray(out.tobytes())


def synth_dialogue(text: str, style: str = "", out_wav: str = "",
                   voices: dict = None, pronounce: dict = None) -> str:
    """진행자·전문가 대화를 한 WAV로 합성 — Gemini MultiSpeakerVoiceConfig.

    화자별로 따로 합성해 이어 붙이면 서로의 호흡을 모른 채 읽어 대화가 아니라 낭독
    두 개가 된다. 한 번의 호출에 두 화자를 같이 넘겨야 주고받는 억양이 나온다.
    [BEAT]를 만나면 거기서 호출을 끊고 무음을 넣은 뒤 다음 블록을 새로 합성한다.
    """
    from google.genai import types
    voices = voices or DIALOGUE_VOICES
    turns = parse_dialogue(text, speakers=list(voices.keys()))
    if not turns or all(w == "BEAT" for w, _ in turns):
        raise ValueError("대사를 찾지 못했습니다 — '진행자:' / '전문가:' 형식인지 확인하세요.")

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
    # temperature는 지정하지 않는다 — 긴 프롬프트와 함께 넣었을 때 출력이 무너졌고,
    # 정상 동작하던 설정에는 이 필드가 없었다. 화자 일관성은 프롬프트로만 잡는다.
    cfg = types.GenerateContentConfig(
        response_modalities=["AUDIO"], speech_config=speech)

    pcm = bytearray()
    block = []
    for who, content in turns + [("BEAT", "0")]:   # 끝 sentinel — 마지막 블록도 flush
        if who != "BEAT":
            block.append((who, content))
            continue
        if block:
            script = "\n".join(f"{w}: {_apply_tts_pronounce(t, pronounce)}" for w, t in block)
            resp = _generate(
                client, model=MODEL,
                contents=(style or DIALOGUE_STYLE) + script, config=cfg)
            pcm.extend(resp.candidates[0].content.parts[0].inline_data.data)
            block = []
        secs = float(content)
        if secs > 0:
            n = int(secs * SR)
            pcm.extend(struct.pack("<%dh" % n, *([0] * n)))

    pcm = _trim_trailing_silence(pcm, out_wav)

    # 마지막 안전장치 — 대본 길이에 견줘 말도 안 되게 긴 결과는 모델 오작동이다.
    # 한글은 초당 5~7자 정도로 읽히므로 글자수÷3초면 아주 넉넉한 상한이다.
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
    for sc in m.scenes:
        if sc.type != "VO":
            continue
        wav = os.path.join(audio_dir, f"{sc.id}.wav")
        synth_scene(sc.narration, m.tts_style, m.voice, wav)
        out[sc.id] = wav
        print(f"[revoice] {sc.id} → {wav}")
    return out
