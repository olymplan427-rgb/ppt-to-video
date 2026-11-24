"""설명회 영상 제작 오케스트레이터.

한 매니페스트로 두 제작 방식을 모두 다룬다:
  - 받은 PPT로 새 영상을 만든다      → SLIDE 씬 위주
  - 기존 설명회 영상을 고쳐 쓴다      → VO/ORIGINAL 씬 위주
둘은 섞을 수 있고, 섞는 게 정상이다(신규 슬라이드 + 기존 인터뷰 원음).

명령:
  python -m revoice.deck  <파일.pptx> [deck]   PPT → 슬라이드 PNG + outline.json
  python -m revoice.build seed                 기존 영상 전사 → 매니페스트 시드
  python -m revoice.build render --dry-run     무엇이 다시 만들어질지만 보기
  python -m revoice.build render               변경분만 재빌드 → 최종 mp4
  python -m revoice.build render --only SC03   그 씬만 다시 만들고 최종 재결합
  python -m revoice.build render --all         캐시 무시하고 전부 재빌드
  python -m revoice.build render --static-slides  SLIDE 씬을 PPT 애니메이션 없이 정지+푸시인으로
  python -m revoice.build render --manifest manifest_client.json   다른 매니페스트로 빌드
  python -m revoice.build trust-audio          기존 나레이션 WAV를 현행으로 인정

여러 매니페스트가 한 프로젝트에 병존할 수 있다(기존 리컷용 manifest.json + 신규 PPT용
manifest_client.json 등). list.txt·베이스 영상·BGM 파일명은 manifest의 output 파일명에서
파생되므로 서로 덮어쓰지 않는다 — 단, output 값을 서로 다르게 지정해야 한다.

증분 빌드가 핵심이다. 검수는 "씬 하나 고쳐줘"의 반복인데, 예전 구조는 그때마다
VO 전체를 TTS 재호출하고 BGM까지 새로 생성해(=영상 전체 배경음악이 바뀜) 한 번에
10분 넘게 걸렸다. 지금은 입력이 바뀐 씬만 다시 만든다.

SLIDE 씬은 기본으로 PPT 자체 애니메이션(요소 순차 페이드인, PowerPoint COM으로 렌더)을
배경으로 쓴다 — 정지 이미지에 밋밋한 푸시인만 얹던 이전 방식 대신, HTML로 다시 그리지
않고도 브랜드 룩 그대로 모션을 얻는다. deck/outline.json이 없거나 --static-slides를
주면 이전 방식(정지+푸시인)으로 대체된다.
"""
import os
import re
import sys
import json
import hashlib
import subprocess

from revoice.manifest import load_manifest, save_manifest
from revoice.timecode import parse_tc
from revoice.timing import reconcile
from revoice import compose, deck, download, env, transcribe, revoice_tts, lyria_bgm

env.load_dotenv()

MANIFEST = "manifest.json"
BUILD = "build"
AUDIO = "audio"
CACHE = os.path.join(BUILD, ".cache.json")

_SLIDE_NO_RE = re.compile(r"slide_(\d+)\.png$")

# 캐시는 입력 데이터만 해시하고 렌더링 코드 자체는 보지 않는다 — _render_scene의 조립
# 로직을 바꿀 때마다(이번엔 file= 경로에 리드인 적용 누락을 고침) 이 값을 올려야 예전
# 캐시가 "안 바뀐 것"으로 오판돼 낡은 세그먼트를 계속 재사용하는 사고를 막는다.
_CODEGEN_VERSION = "9"   # 9: no-anim 씬은 푸시인 제거(zoompan 1px 떨림) — 완전 고정 화면

# 애니메이션 캐시 전용 버전 — 체크포인트 계산 알고리즘(_sentence_checkpoints)이 바뀌면
# narration·duration이 그대로여도 결과가 달라지므로 올려야 한다.
_ANIM_CODEGEN_VERSION = "4"   # 4: 헤더(제목·로고)는 도입부 일괄, 본문만 대사에 배분

# 확정되지 않은 수치·문구 표식. 1차 제작 때 "⚠확인"이 붙은 채로 최종 프레임까지
# 흘러가 화면에 그대로 박힌 사고가 있었다. 렌더 전에 차단한다.
GATE_MARKERS = ("[확인", "⚠", "(대본 확인 필요)", "TBD")


def _probe_dur(path: str) -> float:
    out = subprocess.run(
        ["ffprobe", "-v", "error", "-show_entries", "format=duration",
         "-of", "csv=p=0", path], capture_output=True, text=True, check=True)
    return float(out.stdout.strip())


# ---------------------------------------------------------------- 캐시

def _load_cache() -> dict:
    try:
        with open(CACHE, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def _save_cache(c: dict) -> None:
    os.makedirs(BUILD, exist_ok=True)
    with open(CACHE, "w", encoding="utf-8") as f:
        json.dump(c, f, ensure_ascii=False, indent=1)


def _stamp(path: str) -> str:
    """입력 파일의 지문 — 내용 해시 대신 크기+mtime(대용량 영상이라 해싱은 낭비)."""
    try:
        st = os.stat(path)
        return f"{st.st_size}:{int(st.st_mtime)}"
    except OSError:
        return "missing"


def _bg_is_light(path: str) -> bool:
    """자막이 놓이는 화면 하단이 밝은 배경인가 — 이미지·영상 모두 첫 프레임으로 판정.

    아이보리 자막 박스는 네이비 슬라이드에서는 잘 읽히지만 브랜드 오프화이트 슬라이드
    위에서는 박스가 배경에 묻혀 글자만 떠 보인다. 밝으면 색을 뒤집는다. 하단 25%만
    보는 이유는 자막이 거기 놓이기 때문 — 상단이 짙어도 하단이 밝으면 안 읽힌다.
    """
    if not path or not os.path.exists(path):
        return False
    try:
        out = subprocess.run(
            ["ffmpeg", "-v", "error", "-i", path, "-frames:v", "1",
             "-vf", "crop=iw:ih*0.25:0:ih*0.75,scale=1:1",
             "-f", "rawvideo", "-pix_fmt", "gray", "-"],
            capture_output=True, timeout=60)
        return bool(out.stdout) and out.stdout[0] > 140
    except (OSError, subprocess.SubprocessError):
        return False


def _dialogue_checkpoints(narration: str, words: list, total_dur: float):
    """전문가가 말하는 구간에만 화면 요소 등장 시점을 놓는다. 대화가 아니면 None.

    진행자의 질문은 화제를 여는 도입이지 화면 요소를 꺼낼 대목이 아니다. 대사 전체에
    균등 배분하면 질문이 끝나기도 전에 요소가 다 나와, 정작 전문가가 설명할 때는 화면이
    정지해 있다(사용자 지적). 전문가 문장이 끝나는 시점들만 체크포인트로 준다 —
    첫 구간은 0부터 시작하므로 제목 등 앞쪽 요소는 여전히 질문 중에 떠 있다.
    """
    turns = [(w, t) for w, t in revoice_tts.parse_dialogue(narration) if w != "BEAT"]
    if not turns:
        return None
    units = []                       # (화자, 문장) — 발화 순서
    for who, text in turns:
        for s in re.split(r'(?<=[.?!])\s+', text):
            if s.strip():
                units.append((who, s.strip()))
    counts = [max(1, len(s.split())) for _, s in units]
    total_script = sum(counts)
    n_audio = len(words)
    # 전사 단어 수가 대본과 크게 어긋나면(오인식) 위치 추정을 믿을 수 없다 — 비례 배분.
    use_words = n_audio > 0 and abs(n_audio - total_script) / total_script <= 0.4

    cps, cum = [], 0
    for (who, _), c in zip(units, counts):
        cum += c
        if who != "전문가":
            continue
        if use_words:
            idx = max(0, min(int(round(cum / total_script * n_audio)), n_audio) - 1)
            cps.append(words[idx]["end"])
        else:
            cps.append(total_dur * cum / total_script)
    if not cps:
        return None
    cps[-1] = total_dur
    return cps


def _dialogue_cfg(m):
    """매니페스트에 다화자 설정이 있으면 그걸, 없으면 기본값(진행자=Kore/전문가=Charon)."""
    return (getattr(m, "dialogue_style", "") or revoice_tts.DIALOGUE_STYLE,
            getattr(m, "voices", None) or revoice_tts.DIALOGUE_VOICES)


def _spoken_text(narration: str) -> str:
    """실제로 발화되는 텍스트만 — 화자 라벨과 `>` 영상 인용 줄을 걷어낸다.

    체크포인트는 대본 단어 수와 whisper 전사 단어 수의 '비율'로 위치를 잡는데,
    전사본에는 "진행자:" 라벨도 영상 인용 줄도 없다. 대본 쪽만 그것들을 세면 비율이
    어긋나 화면 요소가 엉뚱한 시점에 등장한다. 1인 나레이션이면 원문을 그대로 돌려준다.
    """
    turns = revoice_tts.parse_dialogue(narration)
    if not turns:
        return narration
    return " ".join(t for who, t in turns if who != "BEAT")


def _voice_key(sc, m) -> str:
    """나레이션 음성이 달라지는 입력만 — 대본·스타일·보이스.

    다화자 씬일 때만 화자·보이스 배정을 키에 더한다. 1인 나레이션 씬의 키를 예전과
    똑같이 유지해야 기존 빌드(구 23슬라이드판)의 음성 캐시가 통째로 무효화되지 않는다.
    """
    parts = [sc.narration, m.tts_style, m.voice]
    if revoice_tts.parse_dialogue(sc.narration):
        style, voices = _dialogue_cfg(m)
        parts.append([style, sorted(voices.items())])
    return hashlib.sha1(
        json.dumps(parts, ensure_ascii=False).encode()
    ).hexdigest()[:16]


def _slide_no(path: str) -> int:
    """`deck/slide_07.png` → 7. 매칭 안 되면 None(=PPT 슬라이드가 아닌 파일)."""
    mo = _SLIDE_NO_RE.search((path or "").replace("\\", "/"))
    return int(mo.group(1)) if mo else None


def _deck_dir(slide_file: str) -> str:
    """슬라이드 PNG가 있는 덱 디렉터리. 프로젝트에 덱이 여러 개 있을 수 있다."""
    return os.path.dirname((slide_file or "").replace("\\", "/")) or "deck"


def _anim_path(slide_file: str) -> str:
    """`deck_260821/slide_07.png` → `deck_260821/anim/slide_07.mp4`.

    애니메이션 mp4를 슬라이드 PNG 옆에 두면 덱마다 자동으로 갈린다. 예전처럼
    deck/anim 한 곳에 몰아넣으면, 슬라이드 번호가 겹치는 다른 덱의 빌드가 서로의
    애니메이션을 조용히 덮어쓴다(구 23슬라이드판 vs 신 43슬라이드 팟캐스트판).
    """
    no = _slide_no(slide_file)
    if no is None:
        return ""
    return os.path.join(_deck_dir(slide_file), "anim", f"slide_{no:02d}.mp4")


def _scene_key(sc, m, wav: str) -> str:
    """세그먼트 영상이 달라지는 모든 입력."""
    inputs = [sc.file, m.source_file if sc.type in ("VO", "ORIGINAL") else "", wav]
    if sc.type == "SLIDE":
        inputs.append(_anim_path(sc.file))
    return hashlib.sha1(json.dumps(
        [_CODEGEN_VERSION, sc.__dict__, [_stamp(p) for p in inputs if p]],
        ensure_ascii=False, sort_keys=True
    ).encode()).hexdigest()[:16]


# ---------------------------------------------------------------- 게이트

def check_gate(m) -> list:
    """미확정 표식이 남은 씬 목록을 돌려준다(비어 있어야 렌더 가능)."""
    hits = []
    for sc in m.scenes:
        for marker in GATE_MARKERS:
            if marker in (sc.narration or "") or marker in (sc.note or ""):
                hits.append(f"{sc.id}: {marker}")
                break
    return hits


# ---------------------------------------------------------------- seed

def cmd_seed():
    from revoice.scene_seed import build_seed_manifest
    if os.path.exists(MANIFEST):
        print(f"[seed] {MANIFEST} 이미 있음 — 검수본 보호를 위해 덮어쓰지 않음. "
              f"재생성하려면 파일을 지우세요.")
        return
    src = download.download_source("eDV49nXi6CE")
    words_path = transcribe.transcribe(src)
    with open(words_path, encoding="utf-8") as f:
        words = json.load(f)
    m = build_seed_manifest()
    for sc in m.scenes:
        if sc.type == "VO":
            sc.narration = transcribe.text_between(
                words, parse_tc(sc.src_in), parse_tc(sc.src_out)) or "(대본 확인 필요)"
    save_manifest(m, MANIFEST)
    print(f"[seed] 매니페스트 생성 완료 → {MANIFEST}")
    print("       ▶ manifest.json 을 검토·수정(순서/대본/유형)한 뒤 'render'를 실행하세요.")


# ---------------------------------------------------------------- render

def _narrate(sc, m, cache: dict, force: bool, dry: bool = False) -> str:
    """VO/SLIDE 씬의 나레이션 WAV를 확보 — 대본이 안 바뀌었으면 재호출하지 않는다."""
    wav = os.path.join(AUDIO, f"{sc.id}.wav")
    key = _voice_key(sc, m)
    if not force and os.path.exists(wav) and cache.get(f"voice:{sc.id}") == key:
        return wav
    if dry:
        kind = "다화자" if revoice_tts.parse_dialogue(sc.narration) else "1인"
        print(f"[voice] {sc.id} 재합성 예정 ({kind} TTS 호출)")
        return wav
    pronounce = getattr(m, "pronounce", None) or None
    if revoice_tts.parse_dialogue(sc.narration):
        style, voices = _dialogue_cfg(m)
        revoice_tts.synth_dialogue(sc.narration, style, wav, voices, pronounce=pronounce)
    else:
        revoice_tts.synth_scene(sc.narration, m.tts_style, m.voice, wav, pronounce=pronounce)
    cache[f"voice:{sc.id}"] = key
    # 씬마다 즉시 저장 — 예전엔 나레이션을 전부 만든 뒤 애니메이션 단계에서 처음
    # 저장했다. 그래서 애니메이션에서 죽으면 이미 만들어 둔 WAV가 멀쩡히 있는데도
    # 기록이 없어 다음 실행이 44개를 통째로 다시 합성했다(실제로 그러다 429를 맞음).
    # TTS는 돈과 쿼터가 드는 호출이라, 성공한 건 그 자리에서 기록해야 한다.
    _save_cache(cache)
    print(f"[voice] {sc.id} → {wav}")
    return wav


def _ass_lines(path: str):
    """.ass 자막의 대사 텍스트 목록 — 없으면 None."""
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return [ln.split(",,0,0,0,,", 1)[1].strip()
                for ln in f if ln.startswith("Dialogue:")]


def cmd_trust_audio():
    """캐시 도입 이전에 만든 나레이션 WAV를 재합성 없이 이어 쓴다.

    빈 캐시로 시작하면 대본이 그대로인 씬까지 TTS를 다시 호출해 크레딧을 태운다.
    그렇다고 WAV 존재만 보고 믿으면 대본을 고친 씬에서 낡은 음성이 나가므로,
    옆에 남아 있는 자막(렌더 당시 대본으로 생성됨)을 지금 대본으로 다시 만든 것과
    대조해 실제로 일치하는 씬만 인정한다.
    """
    m = load_manifest(MANIFEST)
    cache = _load_cache()
    ok, stale = [], []
    for sc in m.scenes:
        if sc.type not in ("VO", "SLIDE"):
            continue
        wav = os.path.join(AUDIO, f"{sc.id}.wav")
        was = _ass_lines(os.path.join(BUILD, f"sub_{sc.id}.ass"))
        if not os.path.exists(wav) or was is None:
            stale.append(sc.id)
            continue
        now = _ass_lines(compose.create_scene_srt(
            sc.narration, _probe_dur(wav), os.path.join(BUILD, "_trustcmp.srt"),
            start_offset=compose._LEAD_IN))
        if was == now:
            cache[f"voice:{sc.id}"] = _voice_key(sc, m)
            ok.append(sc.id)
        else:
            stale.append(sc.id)
    tmp = os.path.join(BUILD, "_trustcmp.ass")
    if os.path.exists(tmp):
        os.remove(tmp)
    _save_cache(cache)
    print(f"[trust] 대본 일치 확인 {len(ok)}개 — 재합성 없이 사용: {', '.join(ok) or '없음'}")
    if stale:
        print(f"        재합성 필요 {len(stale)}개: {', '.join(stale)}")


def _composite_video_base(video: str, wav: str, tts_dur: float, srt: str, seg: str,
                          lead_in: float = None) -> None:
    """신규 화면(영상)에 나레이션·자막을 얹는다 — 구 자막이 없으니 마스킹은 안 한다.

    사전 렌더한 슬라이드 애니메이션이든, 다른 곳에서 만든 VO용 영상이든 같은 조립.
    정지 이미지 경로(build_slide_segment_cmd)와 똑같이 리드인만큼 나레이션을 늦춰야
    자막 표시 시점(create_scene_srt의 start_offset)과 실제 발화가 맞물린다 — 리드인 없이
    바로 붙이면 자막이 오디오보다 lead_in초 늦게 뜨는 것처럼 보인다.
    PPT 애니메이션은 요청한 노출 시간(AdvanceTime)대로 만들지만 인코더의 프레임레이트
    반올림으로 근소하게 짧아질 수 있다 — 영상이 먼저 끝나고 오디오만 남는 사고를 막기
    위해 마지막 프레임을 살짝 고정(freeze)해 여유를 둔다.
    """
    lead_in = compose._LEAD_IN if lead_in is None else lead_in
    vf = compose._VNORM + ",tpad=stop_mode=clone:stop_duration=0.6"
    if srt and os.path.exists(srt):
        vf += "," + compose.ass_filter(srt)
    af = compose._ANORM
    if lead_in > 0:
        af += f",adelay={int(lead_in * 1000)}:all=1"
    compose.run_ffmpeg([
        "ffmpeg", "-y", "-i", video, "-i", wav,
        "-vf", vf, "-af", af,
        "-c:v", "libx264", "-preset", "medium", "-pix_fmt", "yuv420p", "-r", compose.FPS,
        "-c:a", "aac", "-ar", "48000", "-ac", "2", "-t", f"{tts_dur + lead_in:.3f}", seg])


def _render_scene(sc, m, wav: str, seg: str) -> None:
    if sc.type in ("ORIGINAL", "EXTERNAL", "CLIP"):
        src = sc.file if sc.type in ("EXTERNAL", "CLIP") else m.source_file
        compose.run_ffmpeg(compose.build_original_segment_cmd(
            src, sc.src_in or "0:00", sc.src_out or "0:30", seg))
        return

    tts_dur = _probe_dur(wav)
    # 자막 색은 실제로 뒤에 깔릴 화면의 밝기로 정한다 — SLIDE는 애니메이션 mp4(없으면
    # 슬라이드 PNG), VO는 그 영상 자체.
    bg_ref = (_anim_path(sc.file) if sc.type == "SLIDE" else "") or sc.file
    # 같은 WAV의 전사 단어 타임스탬프를 자막 타이밍에 그대로 쓴다(캐시돼 있어 공짜).
    # 글자수 비례 배분은 발화 속도를 일정하다고 가정해 자막이 음성보다 앞서 나갔다.
    try:
        words = _narration_words(sc, wav)
    except Exception:
        words = None
    srt = compose.create_scene_srt(sc.narration, tts_dur,
                                   os.path.join(BUILD, f"sub_{sc.id}.srt"),
                                   start_offset=compose._LEAD_IN,
                                   light_bg=_bg_is_light(bg_ref),
                                   words=words)

    if sc.type == "SLIDE":
        # 지난 빌드에서 만들어 둔 애니메이션 파일이 남아 있어도 no-anim이면 쓰지 않는다.
        anim = "" if "no-anim" in (sc.flags or []) else _anim_path(sc.file)
        if anim and os.path.exists(anim):
            _composite_video_base(anim, wav, tts_dur, srt, seg)
        else:
            # PPT 요소 애니메이션을 못 만들었을 때의 대체 — 정지 이미지 + 느린 푸시인.
            # no-anim 씬은 푸시인도 빼고 완전 고정(zoompan의 1px 떨림 제거).
            zoom = 0.0 if "no-anim" in (sc.flags or []) else 0.06
            compose.run_ffmpeg(compose.build_slide_segment_cmd(
                sc.file, wav, tts_dur, seg, srt_file=srt, zoom=zoom))
        return

    # VO: file이 지정되면 그 영상을 화면으로 쓰고, 없으면 원본 설명회 영상 구간을 쓴다.
    if sc.file:
        _composite_video_base(sc.file, wav, tts_dur, srt, seg)
        return

    plan = reconcile(parse_tc(sc.src_out) - parse_tc(sc.src_in), tts_dur)
    compose.run_ffmpeg(compose.build_vo_segment_cmd(
        m.source_file, sc.src_in, sc.src_out, wav, plan, seg,
        srt_file=srt, is_new_slide=False))


def _sentence_checkpoints(narration: str, words: list, total_dur: float) -> list:
    """나레이션을 문장 단위로 나누고, 각 문장이 실제로 다 말해지는 시각을 돌려준다.

    문자열 매칭이 아니라 '단어 개수' 비율로 위치를 찾는다 — TTS가 원문을 그대로
    읽어도 whisper 전사 철자(숫자 표기·붙여쓰기 등)는 원문과 달라지기 쉬워 문자열
    매칭은 깨지기 쉽지만, 단어 개수 비율은 안정적이다. 이 체크포인트가 PPT 요소
    애니메이션의 등장 타이밍이 된다 — "화면 요소가 전부 처음 2초 안에 뜨고 나머지
    구간은 정적"이던 문제를, 나레이터가 그 대목을 말하는 시점에 맞춰 등장하게 바꾼다.
    """
    parts = [s.strip() for s in re.split(r'(?<=[.?!])\s+|(?<=—)\s+', narration.replace("\n", " "))
             if s.strip()]
    if not parts:
        return [total_dur]
    word_counts = [max(1, len(p.split())) for p in parts]
    total_words_script = sum(word_counts)
    total_words_audio = len(words)
    if total_words_audio == 0 or abs(total_words_audio - total_words_script) / total_words_script > 0.4:
        # whisper 전사가 원문과 단어 수부터 크게 어긋나면(오인식 등) 위치 추정을
        # 신뢰할 수 없다 — 균등 분배로 폴백. 그래도 "전부 처음 2초"보다는 낫다.
        n = len(parts)
        return [total_dur * (i + 1) / n for i in range(n)]
    cum = 0
    checkpoints = []
    for wc in word_counts:
        cum += wc
        idx = max(0, min(int(round(cum / total_words_script * total_words_audio)),
                        total_words_audio) - 1)
        checkpoints.append(words[idx]["end"])
    checkpoints[-1] = total_dur    # 마지막 체크포인트는 항상 씬 전체 길이로 고정
    return checkpoints


def _narration_words(sc, wav: str) -> list:
    """씬 나레이션 WAV를 전사해 단어 타임스탬프를 얻는다 — 대본이 바뀌면 다시 전사."""
    words_path = os.path.join(AUDIO, f"{sc.id}.words.json")
    key = hashlib.sha1(_stamp(wav).encode()).hexdigest()[:12]
    stamp_path = words_path + ".stamp"
    if os.path.exists(words_path) and os.path.exists(stamp_path):
        with open(stamp_path, encoding="utf-8") as f:
            if f.read().strip() == key:
                with open(words_path, encoding="utf-8") as wf:
                    return json.load(wf)
        os.remove(words_path)
    transcribe.transcribe(wav, out=words_path)
    with open(stamp_path, "w", encoding="utf-8") as f:
        f.write(key)
    with open(words_path, encoding="utf-8") as f:
        return json.load(f)


def _ensure_slide_animations(m, wavs: dict, cache: dict, only, rebuild_all: bool,
                             dry: bool, use_anim: bool) -> None:
    """대상 SLIDE 씬들의 요소 애니메이션 mp4를 한 PowerPoint 세션에서 일괄 생성.

    슬라이드마다 따로 호출하면 PowerPoint를 그만큼 열고 닫아 느리다 — 이번 빌드에서
    새로 필요한 슬라이드를 모아 한 번에 넘긴다(revoice/deck.py의 Hidden 토글 방식).
    각 슬라이드의 등장 타이밍은 그 씬 나레이션을 실제로 전사(whisper)해 문장별로
    체크포인트를 계산해 넘긴다 — 균등 배분이 아니라 나레이션 진행에 맞춘 등장.
    """
    if not use_anim:
        return
    # 덱 디렉터리는 매니페스트의 SLIDE 씬이 가리키는 곳에서 얻는다. 예전엔 "deck"으로
    # 하드코딩돼 있었는데, 프로젝트에 덱이 둘 이상이면(구 23슬라이드 deck/ vs 신
    # 43슬라이드 deck_260821/) 엉뚱한 PPT로 애니메이션을 만들어 넣는다.
    # 덱 슬라이드(slide_NN.png)인 것만 본다. SLIDE 타입이라고 다 덱은 아니다 —
    # 클립 뒤 홀드 프레임(clips_podcast/DP21_hold.png)도 SLIDE로 태우는데, 그건
    # 애니메이션 대상이 아니라 정지 이미지다. 예전엔 이것까지 세어 "덱이 둘"이라며
    # 렌더를 막았다.
    slide_files = [sc.file for sc in m.scenes
                   if sc.type == "SLIDE" and _slide_no(sc.file) is not None]
    if not slide_files:
        return
    decks = {_deck_dir(f) for f in slide_files}
    if len(decks) > 1:
        raise SystemExit(f"[anim] 한 매니페스트가 여러 덱을 가리킵니다: {sorted(decks)} "
                         f"— 슬라이드 애니메이션의 원본 PPT를 정할 수 없습니다.")
    deck_dir = decks.pop()
    outline_path = os.path.join(deck_dir, "outline.json")
    if not os.path.exists(outline_path):
        return
    with open(outline_path, encoding="utf-8") as f:
        pptx = json.load(f).get("source_pptx")
    # 다른 PC에서 클론했을 때 outline.json에 이전 PC의 절대경로가 남아 있을 수 있다.
    # 상대경로·현재 폴더·상위 폴더(m.source 파일명) 순으로 자동 복구한다.
    if pptx and not os.path.exists(pptx):
        rel_try = os.path.join(deck_dir, pptx)
        if os.path.exists(rel_try):
            pptx = rel_try
        else:
            name = os.path.basename(pptx.replace("\\", "/")) or os.path.basename(m.source or "")
            for base_dir in (".", "..", os.path.join("..", "..")):
                cand = os.path.join(base_dir, name)
                if name and os.path.exists(cand):
                    pptx = cand
                    break
    if not pptx or not os.path.exists(pptx):
        return
    if not deck.has_powerpoint():
        print("[anim]  PowerPoint COM 미감지(macOS/Linux 또는 미설치) — 정지/푸시인 슬라이드로 자동 진행합니다.")
        return

    todo = {}          # slide_no -> {duration, checkpoints}
    for sc in m.scenes:
        if sc.type != "SLIDE":
            continue
        # 표지·간지·숫자 요약처럼 한 화면으로 읽히는 장표와, 영상 뒤 부연설명용 슬라이드는
        # 요소가 하나씩 뜨면 오히려 산만하다 — 매니페스트에서 no-anim으로 지정한다.
        if "no-anim" in (sc.flags or []):
            continue
        no = _slide_no(sc.file)
        if no is None:
            continue
        targeted = only is None or sc.id in only
        if not targeted:
            continue
        wav = wavs.get(sc.id, "")
        if not wav or not os.path.exists(wav):
            continue
        dur = _probe_dur(wav) + compose._LEAD_IN
        key = hashlib.sha1(json.dumps(
            [_ANIM_CODEGEN_VERSION, _stamp(pptx), no, round(dur, 1), sc.narration],
            ensure_ascii=False).encode()
        ).hexdigest()[:16]
        anim = _anim_path(sc.file)
        # 캐시 키에 덱을 포함 — 슬라이드 번호만으로는 덱이 다른 빌드끼리 충돌한다.
        if rebuild_all or not os.path.exists(anim) \
                or cache.get(f"anim:{deck_dir}:{no}") != key:
            todo[no] = {"dur": dur, "key": key, "sc": sc, "wav": wav}

    if not todo:
        return
    if dry:
        for no in sorted(todo):
            print(f"[anim]  slide_{no:02d} 애니메이션 재렌더 예정 ({todo[no]['dur']:.1f}s)")
        return

    print(f"[anim]  {len(todo)}개 슬라이드 애니메이션 렌더 중 (전사 + PowerPoint COM)...")
    payload = {}
    for no, t in todo.items():
        words = _narration_words(t["sc"], t["wav"])
        # 리드인(정적) 만큼 체크포인트를 뒤로 밀어 — 애니메이션은 씬 전체(리드인+발화) 기준.
        raw_dur = t["dur"] - compose._LEAD_IN
        raw_cps = (_dialogue_checkpoints(t["sc"].narration, words, raw_dur)
                   or _sentence_checkpoints(_spoken_text(t["sc"].narration), words, raw_dur))
        checkpoints = [compose._LEAD_IN + c for c in raw_cps]
        payload[no] = {"duration": t["dur"], "checkpoints": checkpoints}

    rendered = deck.render_animated_slides(
        pptx, payload, out_dir=os.path.join(deck_dir, "anim"))
    for no, t in todo.items():
        if no in rendered:
            cache[f"anim:{deck_dir}:{no}"] = t["key"]
            print(f"[anim]  slide_{no:02d} → {rendered[no]}")
    _save_cache(cache)


def cmd_render(only=None, rebuild_all=False, force_gate=False, new_bgm=False, dry=False,
               use_anim=True, manifest_path=None):
    m = load_manifest(manifest_path or MANIFEST)
    os.makedirs(BUILD, exist_ok=True)
    # 프로젝트에 매니페스트가 여러 개 병존할 수 있다(예: 기존 리컷용 manifest.json과
    # 신규 PPT용 manifest_client.json) — 중간 산출물 이름을 output 파일명에서 파생시켜
    # 서로 다른 매니페스트의 빌드가 list.txt·베이스 영상·BGM을 덮어쓰지 않게 한다.
    prefix = os.path.splitext(os.path.basename(m.output))[0]

    gate = check_gate(m)
    if gate and not force_gate:
        print("[gate] 미확정 표식이 남아 있어 렌더를 중단합니다 "
              "(그대로 렌더하면 화면에 박힙니다):")
        for g in gate:
            print(f"        - {g}")
        print("       확정 후 다시 실행하거나, 의도한 것이면 --force 를 붙이세요.")
        raise SystemExit(2)

    cache = _load_cache()

    # TTS는 무료 티어 기준 모델당 하루 100회다(소진되면 23시간 잠김 — 백오프로 안 풀린다).
    # 몇 회를 쓸 참인지 먼저 알려 준다. 예전엔 모르고 돌렸다가 하루 작업이 막혔다.
    todo_tts = [sc.id for sc in m.scenes
                if sc.type in ("VO", "SLIDE") and (only is None or sc.id in only)
                and not (not (rebuild_all and (only is None or sc.id in only))
                         and os.path.exists(os.path.join(AUDIO, f"{sc.id}.wav"))
                         and cache.get(f"voice:{sc.id}") == _voice_key(sc, m))]
    if todo_tts:
        print(f"[tts]   이번 실행에서 TTS {len(todo_tts)}회 호출 예정 "
              f"(무료 티어 한도 100회/일) — {', '.join(todo_tts[:6])}"
              f"{' …' if len(todo_tts) > 6 else ''}")

    # 1단계: 나레이션 WAV 확보 (SLIDE 애니메이션 길이를 정하려면 TTS 길이를 먼저 알아야 함).
    wavs = {}
    for sc in m.scenes:
        if sc.type in ("VO", "SLIDE"):
            targeted = only is None or sc.id in only
            wavs[sc.id] = _narrate(sc, m, cache, force=rebuild_all and targeted, dry=dry)

    # 2단계: 그 길이에 맞춰 SLIDE 씬의 PPT 요소 애니메이션을 한 세션에서 일괄 렌더.
    _ensure_slide_animations(m, wavs, cache, only, rebuild_all, dry, use_anim)

    # 3단계: 씬별 최종 세그먼트 합성(신규/변경분만).
    seg_paths, rebuilt = [], 0

    for sc in m.scenes:
        seg = os.path.join(BUILD, f"seg_{sc.id}.mp4")
        seg_paths.append(seg)
        targeted = only is None or sc.id in only
        wav = wavs.get(sc.id, "")

        key = _scene_key(sc, m, wav)
        if not targeted or (not rebuild_all and os.path.exists(seg)
                            and cache.get(f"seg:{sc.id}") == key):
            print(f"[skip]  {sc.id}")
            continue

        if dry:
            print(f"[seg]   {sc.id} 재렌더 예정 ({sc.type})")
            rebuilt += 1
            continue

        _render_scene(sc, m, wav, seg)
        cache[f"seg:{sc.id}"] = key
        _save_cache(cache)      # 씬마다 저장 — 중간에 끊겨도 앞부분은 재활용된다
        rebuilt += 1
        print(f"[seg]   {sc.id} → {seg}")

    if dry:
        print(f"[dry]   재빌드 대상 {rebuilt}개 씬 — 실제로 만들려면 --dry-run 을 빼세요.")
        return

    missing = [p for p in seg_paths if not os.path.exists(p)]
    if missing:
        raise SystemExit(f"세그먼트 누락: {missing} — --all 로 전체 빌드하세요.")

    list_file = os.path.join(BUILD, f"{prefix}_list.txt")
    with open(list_file, "w", encoding="utf-8") as f:
        for p in seg_paths:
            f.write(f"file '{os.path.basename(p)}'\n")

    base = os.path.join(BUILD, f"{prefix}_base.mp4")
    compose.run_ffmpeg(compose.build_concat_cmd(list_file, base))

    # BGM은 매 빌드마다 새로 생성하면 씬 하나만 고쳐도 영상 전체의 배경음악이
    # 달라진다. 이미 만든 트랙이 있으면 그대로 쓴다.
    bgm = os.path.join("bgm", f"{prefix}.wav")
    if new_bgm or not os.path.exists(bgm):
        bgm = lyria_bgm.generate_bgm(m.bgm.get("prompt", "warm ambient"), out=bgm, seconds=90)
    else:
        print(f"[bgm]   기존 트랙 재사용 → {bgm} (바꾸려면 --new-bgm)")

    compose.run_ffmpeg(compose.build_bgm_duck_cmd(base, bgm, m.output))
    print(f"[done]  {rebuilt}개 씬 재빌드 · 최종 → {m.output}")


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    cmd = argv[0] if argv else ""
    if cmd == "seed":
        cmd_seed()
    elif cmd == "trust-audio":
        cmd_trust_audio()
    elif cmd == "render":
        only = None
        if "--only" in argv:
            only = set(argv[argv.index("--only") + 1].split(","))
        manifest_path = None
        if "--manifest" in argv:
            manifest_path = argv[argv.index("--manifest") + 1]
        cmd_render(only=only, rebuild_all="--all" in argv,
                   force_gate="--force" in argv, new_bgm="--new-bgm" in argv,
                   dry="--dry-run" in argv, use_anim="--static-slides" not in argv,
                   manifest_path=manifest_path)
    else:
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
