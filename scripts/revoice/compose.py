"""ffmpeg 명령 빌더(순수) + 기존 자막 마스킹 + 신규 자막 합성 헬퍼."""
import os
import re
import subprocess
from revoice.timecode import parse_tc
from revoice.timing import ReconcilePlan

W, H, FPS = "1920", "1080", "30"
_VNORM = f"scale={W}:{H}:force_original_aspect_ratio=decrease,pad={W}:{H}:(ow-iw)/2:(oh-ih)/2,setsar=1,fps={FPS}"
_ANORM = "aresample=48000,aformat=sample_rates=48000:channel_layouts=stereo"

# 원본 영상(eDV49.mp4)의 구 자막은 인터뷰 구간(노트 박스)과 VO 대체 구간(타이틀카드/브롤)의
# 스타일이 다르다. VO씬이 실제로 대체하는 구간의 구 자막은 심플한 화이트 박스(또는 무배경)+
# 검정 굵은 텍스트로, 실측 결과 y≈963~1038에 고정 배치됨(여러 씬 샘플링으로 확인).
# 화면 전체 폭을 덮을 필요는 없으므로 폭 1600px(중앙 정렬)·높이 110px로 축소해
# 구 자막을 가리되 화면 하단을 불필요하게 넓게 덮지 않는다. 새 자막 박스와 같은
# 아이보리 톤으로 채워 아래 ASS 자동 박스와 이어져 보이도록 한다.
_BAR_W = 1600
_BAR_X = (1920 - _BAR_W) // 2
_BAR_Y = 930
_BAR_H = 110
_MASK_BOX = f"drawbox=x={_BAR_X}:y={_BAR_Y}:w={_BAR_W}:h={_BAR_H}:color=0xF5F1E8@1.0:t=fill"

# 신규 나레이션 자막: SRT+force_style은 subtitles 필터가 내부적으로 임의의 PlayRes로
# 스케일링하면서 줄바꿈·위치가 어긋나는 문제가 있어(1920x1080 명시해도 재현됨),
# PlayResX/Y를 직접 못박은 .ass로 렌더링해 좌표계 모호성을 제거한다.
# WrapStyle=2(자동 줄바꿈 금지)로 두 줄로 밀리는 문제도 원천 차단 — 대신 청크를
# max_chars 이하로 미리 쪼개 한 줄에 들어가도록 보장한다.
# BorderStyle=3(불투명 박스)으로 텍스트 길이에 맞춰 배경 박스가 자동으로 좁아지거나
# 넓어지게 해, 화면 하단 전체를 덮는 고정폭 바 대신 실제 텍스트만큼만 덮도록 한다.
# 원본 영상(구 자막)의 톤(아이보리 박스+짙은 텍스트)에 맞춰 색을 지정했다.
# 텍스트는 순검정 대신 브랜드 그레이 팔레트의 가장 짙은 톤(#464646)으로 톤다운.
# 폰트는 브랜드 폰트 Wanted Sans ExtraBold(시스템 미설치라 _FONTS_DIR로 런타임 로드, 아래 참고).
#
# 배경 밝기에 따라 두 벌을 쓴다. 아이보리 박스는 네이비 슬라이드 위에서는 잘 읽히지만
# 크림색 슬라이드(브랜드 오프화이트 #F2F7F5) 위에서는 박스가 배경에 묻혀 글자만 떠
# 보인다. 밝은 배경에서는 색을 뒤집어 짙은 박스 + 오프화이트 텍스트를 쓴다.
# (.ass 색은 &HAABBGGRR — RGB 역순임에 주의.)
_ASS_HEAD = """[Script Info]
ScriptType: v4.00+
PlayResX: 1920
PlayResY: 1080
WrapStyle: 2
ScaledBorderAndShadow: yes

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Default,Wanted Sans ExtraBold,56,{fg},&H000000FF,{bg},{bg},0,0,0,0,100,100,0,0,3,20,0,2,80,80,70,1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
"""

# 어두운 배경용(기본) — 아이보리 박스 + 브랜드 그레이 텍스트.
_ASS_ON_DARK = _ASS_HEAD.format(fg="&H00464646", bg="&H00E8F1F5")
# 밝은 배경용 — 브랜드 네이비 박스 + 오프화이트 텍스트.
_ASS_ON_LIGHT = _ASS_HEAD.format(fg="&H00F5F7F2", bg="&H00432B1B")

_ASS_HEADER = _ASS_ON_DARK    # 하위 호환(기존 호출부)

import sys

# Wanted Sans는 시스템에 설치돼 있지 않아(브랜드 폰트, 프로젝트 assets에만 존재),
# ffmpeg의 ass 필터에 fontsdir로 직접 지정해 설치 없이 렌더링에 사용한다.
# 폰트 파일이 없는 범용 프로젝트(macOS/Windows/Linux)에서는 OS 기본 한글 폰트로 폴백한다.
_FONTS_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "assets", "fonts").replace("\\", "/")


def _has_bundled_font() -> bool:
    if not os.path.isdir(_FONTS_DIR):
        return False
    return any(f.lower().startswith("wantedsans") and f.lower().endswith(".ttf")
               for f in os.listdir(_FONTS_DIR))


def _subtitle_font_name() -> str:
    if _has_bundled_font():
        return "Wanted Sans ExtraBold"
    if sys.platform == "darwin":
        return "Apple SD Gothic Neo"
    if os.name == "nt":
        return "Malgun Gothic"
    return "Noto Sans CJK KR"


def _escape_ffmpeg_filter_path(p: str) -> str:
    p = p.replace("\\", "/")
    if len(p) >= 2 and p[1] == ":":
        return p[0] + r"\:" + p[2:].replace(":", r"\:")
    return p.replace(":", r"\:")


def ass_filter(srt_file: str) -> str:
    """ffmpeg 필터 문자열용 ass= 절.

    필터 인자 안에서 경로의 콜론은 옵션 구분자로 먹히므로 이스케이프해야 한다.
    fontsdir는 미설치 브랜드 폰트(Wanted Sans)를 설치 없이 로드하기 위한 것.
    """
    path = _escape_ffmpeg_filter_path(srt_file)
    if _has_bundled_font():
        fonts = _escape_ffmpeg_filter_path(_FONTS_DIR)
        return f"ass='{path}':fontsdir='{fonts}'"
    return f"ass='{path}'"


def _fmt_ass_tc(sec: float) -> str:
    h = int(sec // 3600)
    m = int((sec % 3600) // 60)
    s = int(sec % 60)
    cs = int(round((sec - int(sec)) * 100))
    if cs >= 100:
        s += 1
        cs -= 100
    return f"{h:d}:{m:02d}:{s:02d}.{cs:02d}"


# 뒤 단어와 한 덩어리로 읽히는 말들 — 자막 줄 끝에 홀로 남으면 안 된다.
# "…이 네 / 가지로 모입니다" 처럼 관형사·수사가 수량사와 갈라지면 눈에 걸린다.
_NO_TRAILING = {
    "이", "그", "저", "한", "두", "세", "네", "다섯", "여섯", "일곱", "여덟", "아홉", "열",
    "몇", "각", "매", "제", "총", "약", "더", "덜", "안", "못", "잘", "다", "좀", "왜",
    "이런", "그런", "저런", "어떤", "무슨", "새", "옛", "온", "전", "첫",
}


# "10월 / 21일"처럼 날짜가 줄바꿈에서 갈라지면 안 된다(사용자 지적 2026-10-24) — 년·월로 끝나는 낱말은 줄 끝에 못 둔다.
_DATE_HEAD = re.compile(r'\d+(년|월)')


def _pack_words(part: str, max_chars: int) -> list:
    """공백 경계에서 max_chars 이하로 욕심쟁이 재분할(절 경계가 없을 때의 최종 폴백).

    끝에 관형사·수사가 홀로 남으면 다음 줄로 넘긴다 — 그것만으로는 뜻이 서지 않아
    자막이 어정쩡하게 끊긴 것처럼 보인다(사용자 지적 2026-08-24).
    """
    words = part.split(' ')
    chunks, cur = [], ''
    for w in words:
        cand = f"{cur} {w}".strip()
        if cur and len(cand) > max_chars:
            # 줄 끝에 걸리는 낱말이 연달아 있을 수 있다("… 만큼 이 네") — 전부 넘긴다.
            head, moved = cur, []
            while ' ' in head and (head.rsplit(' ', 1)[1] in _NO_TRAILING
                               or _DATE_HEAD.fullmatch(head.rsplit(' ', 1)[1])):
                head, tail = head.rsplit(' ', 1)
                moved.insert(0, tail)
            if head:
                chunks.append(head)
                cur = ' '.join(moved + [w])
            else:                             # 통째로 걸리면 그냥 원래대로 끊는다
                chunks.append(cur)
                cur = w
        else:
            cur = cand
    if cur:
        chunks.append(cur)
    return chunks


def _split_long(part: str, max_chars: int) -> list:
    """max_chars를 넘는 구문을 쉼표(자연스러운 끊어읽기 지점) 경계로 우선 분할하고,
    그래도 한 절이 너무 길면 공백 경계로 재분할한다.

    쉼표 없이 공백 경계로만 자르면 "...운영하는 법, PTR과..."처럼 절 중간에서
    잘려 부자연스러운 자막이 나오는 문제가 있었다(사용자 피드백 2026-07-27).
    """
    if len(part) <= max_chars:
        return [part]

    # 절 하나가 max_chars를 살짝 넘는 정도(=합리적인 끊어읽기 단위)까지는
    # 억지로 쪼개지 않고 한 줄로 허용한다.
    hard_cap = int(max_chars * 1.4)

    clauses = [c.strip() for c in re.split(r'(?<=,)\s+', part) if c.strip()]
    if len(clauses) <= 1:
        return _pack_words(part, max_chars)

    chunks, cur = [], ''
    for c in clauses:
        cand = f"{cur} {c}".strip()
        if cur and len(cand) > max_chars:
            chunks.append(cur)
            cur = c
        else:
            cur = cand
    if cur:
        chunks.append(cur)

    refined = []
    for ch in chunks:
        if len(ch) <= hard_cap:
            refined.append(ch)
        else:
            refined.extend(_pack_words(ch, max_chars))
    return refined


_LEAD_IN = 0.5  # 씬 시작부터 나레이션이 시작되기까지의 정지 여백(초).


def _spans_from_words(parts: list, words: list, total_dur: float):
    """자막 조각들을 실제 발화 타임스탬프에 맞춘다. 신뢰할 수 없으면 None.

    조각 텍스트를 전사본과 문자열로 맞추려 들면 깨진다 — TTS가 원문대로 읽어도
    whisper 표기(숫자·붙여쓰기·오인식)는 원문과 다르기 때문. 그래서 체크포인트
    계산과 같은 방식으로 **단어 개수 비율**로 위치를 잡는다. 조각 경계마다 그 지점의
    단어 end 시각을 쓰면, 쉼과 완급이 그대로 반영된 타이밍이 나온다.
    """
    n_audio = len(words or [])
    counts = [max(1, len(p.split())) for p in parts]
    n_script = sum(counts)
    if n_audio == 0 or abs(n_audio - n_script) / n_script > 0.4:
        return None            # 전사가 대본과 크게 어긋남 — 믿지 말고 폴백

    # 단어 '개수 비율'로만 위치를 잡으면 대본과 전사의 띄어쓰기 차이가 누적돼 뒤로 갈수록
    # 자막이 밀린다(실측: 60초 지점에서 2.5초 지연). 그래서 조각의 첫 낱말을 전사에서
    # 글자로 찾아 시작 시각을 고정한다. 못 찾은 조각은 앞뒤 고정점 사이를 보간한다.
    def norm(s):
        return re.sub(r"[^0-9가-힣A-Za-z]", "", s)

    starts = [None] * len(parts)
    cur = 0
    for i, p in enumerate(parts):
        toks = p.split()
        if not toks:
            continue
        first = norm(toks[0])
        # 후보 키를 순서대로 시도한다. 첫 낱말이 "그"처럼 한 글자면 그것만으론 오매칭이
        # 잦아 다음 낱말을 붙인 형태를 먼저 보고, 안 맞으면 첫 낱말 단독으로 되돌아간다.
        keys = []
        if len(first) < 2 and len(toks) > 1:
            keys.append((first + norm(toks[1]))[:3])
        keys.append(first[:3])
        for key in keys:
            if not key:
                continue
            hit = None
            for j in range(cur, n_audio):
                wd = norm(words[j]["word"])
                if (key in wd) if len(key) >= 2 else (wd == key):
                    hit = j
                    break
            if hit is not None:
                starts[i] = float(words[hit]["start"])
                cur = hit + 1
                break

    if starts[0] is None:
        starts[0] = float(words[0]["start"])
    # 마지막 조각을 words[-1].start에 고정하면 안 된다 — 그건 '마지막 낱말'이 시작되는
    # 시각이라, 조각이 길 때 0.9초 만에 30자를 띄우는 사고가 난다(실측). 비워 두고
    # 아래 보간에 맡긴 뒤, 읽기 속도 보정에서 시작을 앞으로 당긴다.
    if starts[-1] is None and len(starts) > 1:
        anchored = [s for s in starts if s is not None]
        starts[-1] = max(anchored[-1], total_dur - 0.6) if anchored else total_dur - 0.6
    elif starts[-1] is None:
        starts[-1] = 0.0

    # 빈 곳은 양옆 고정점 사이를 글자수 비례로 채운다.
    i = 0
    while i < len(starts):
        if starts[i] is not None:
            i += 1
            continue
        a = i - 1
        b = i
        while b < len(starts) and starts[b] is None:
            b += 1
        span = starts[b] - starts[a]
        weights = [max(1, len(parts[k])) for k in range(a + 1, b + 1)]
        tot = sum(weights)
        acc = 0
        for k, wgt in zip(range(a + 1, b), weights):
            acc += wgt
            starts[k] = starts[a] + span * acc / tot
        i = b

    # 읽기 속도 보정 — 끝 시각은 다음 자막이 뜨는 시점이라 건드릴 수 없으니, 너무 빠른
    # 자막은 **시작을 앞으로 당겨** 읽을 시간을 준다(앞 자막을 침범하지 않는 선에서).
    MAX_CPS = 15.0                             # 한글 자막 가독 상한 대략치
    MIN_SHOW = 0.4
    for i in range(len(starts) - 1, -1, -1):
        en = starts[i + 1] if i + 1 < len(starts) else total_dur
        need = len(parts[i]) / MAX_CPS
        if en - starts[i] >= need:
            continue
        # **앞으로만 당긴다.** max()로 하면 여유가 없을 때 시작이 뒤로 밀려 끝보다 늦어지고
        # 길이가 음수인 자막이 나온다(실측 -0.3초). 앞 자막의 최소 노출은 지켜 준다.
        floor = starts[i - 1] + MIN_SHOW if i else 0.0
        starts[i] = min(starts[i], max(floor, en - need))

    # 단조 증가와 범위를 마지막으로 강제 — 어떤 보정도 이 불변식을 깨선 안 된다.
    for i in range(len(starts)):
        starts[i] = max(0.0, min(starts[i], total_dur))
        if i and starts[i] < starts[i - 1]:
            starts[i] = starts[i - 1]

    spans = []
    for i, st in enumerate(starts):
        en = starts[i + 1] if i + 1 < len(starts) else total_dur
        if en - st < 0.35:                     # 그래도 짧으면 최소 폭은 보장
            en = min(total_dur, st + 0.35)
        spans.append((max(0.0, st), min(en, total_dur)))
    spans[-1] = (spans[-1][0], total_dur)      # 마지막 자막은 씬 끝까지 유지
    return spans


def create_scene_srt(narration: str, total_dur: float, srt_path: str, max_chars: int = 28,
                      start_offset: float = 0.0, light_bg: bool = False,
                      words: list = None) -> str:
    """나레이션 문장을 구문 단위로 분할해 타임코드 .ass 자막 파일 생성(반환 경로는 .ass).

    한 줄에 안전하게 들어가도록 max_chars(기본 28자)를 넘는 구문은
    공백 경계에서 추가로 쪼갠다. WrapStyle=2라 자동 줄바꿈이 없으므로
    이 사전 분할이 곧 실제 표시 줄 단위가 된다. total_dur은 실제 발화 구간
    길이(리드인 제외)이고, start_offset만큼 모든 타임코드를 뒤로 민다 —
    리드인 동안에는 자막이 뜨지 않고 발화 시작과 정확히 맞물리게 하기 위함.
    """
    if not narration or not narration.strip():
        return None

    # [BEAT n]은 "여기서 n초 쉬어라"는 제작 지시 — "—"와 똑같은 이유로 화면에 뜨면
    # 안 된다. 팟캐스트판에서 클립 뒤 리액션 대사를 다음 씬 앞으로 옮기면서 이 마커가
    # 나레이션 첫 줄이 됐고, 8개 씬 자막에 "[BEAT 0.8s]"가 그대로 찍혀 나왔다.
    narration = re.sub(r"\[BEAT\s+[0-9.]+s?\]", " ", narration)

    # "진행자:" / "전문가:" 화자 라벨은 대본에서 누가 말하는지 나누려고 쓰는 표기이지
    # 시청자에게 읽힐 문장이 아니다. 목소리가 이미 둘로 구분되므로 화면에 또 쓰지 않는다.
    narration = re.sub(r"^\s*(진행자|전문가)\s*[:：]\s*", "", narration, flags=re.M)
    if not narration.strip():
        return None

    # 자막 조각 만들기. **줄(=화자 발언) 경계를 절대 넘지 않는다** — 넘으면 진행자
    # 질문 끝과 전문가 답 첫머리가 한 자막에 섞여 나온다(사용자 지적).
    # 줄 안에서는 문장(. ? !) → 너무 길면 쉼표/공백 순으로 쪼갠다.
    parts = []
    for line in (l.strip() for l in narration.splitlines()):
        if not line:
            continue
        raw = [p.strip() for p in re.split(r'(?<=[.?!])\s+|(?<=—)\s+', line) if p.strip()]
        for p in raw:
            p = re.sub(r'—\s*$', '', p).strip()
            if p:
                parts.extend(_split_long(p, max_chars))
    if not parts:
        parts = [narration.strip()]

    # 시간 배분 — 전사 단어 타임스탬프가 있으면 그것으로, 없으면 글자수 비례(폴백).
    #
    # 글자수 비례는 발화 속도가 일정하다고 가정하는데 실제로는 쉼·강조·완급이 있어
    # 자막이 음성보다 앞서 나간다(사용자 지적: "자막이 음성보다 빠른 구간이 다수").
    # 같은 WAV를 whisper로 전사한 단어 타임스탬프가 이미 있으므로 그걸 쓴다.
    spans = _spans_from_words(parts, words, total_dur) if words else None
    if spans is None:
        cc = [max(1, len(p)) for p in parts]
        tot = sum(cc)
        spans, t = [], 0.0
        for c in cc:
            e = min(total_dur, t + total_dur * c / tot)
            spans.append((t, e))
            t = e

    events = []
    for (t_start, t_end), text in zip(spans, parts):
        safe_text = text.replace("{", "(").replace("}", ")")
        events.append(
            f"Dialogue: 0,{_fmt_ass_tc(t_start + start_offset)},{_fmt_ass_tc(t_end + start_offset)},"
            f"Default,,0,0,0,,{safe_text}")

    ass_path = os.path.splitext(srt_path)[0] + ".ass"
    os.makedirs(os.path.dirname(ass_path), exist_ok=True)
    head = _ASS_ON_LIGHT if light_bg else _ASS_ON_DARK
    font_name = _subtitle_font_name()
    if font_name != "Wanted Sans ExtraBold":
        head = head.replace("Wanted Sans ExtraBold", font_name)
    with open(ass_path, "w", encoding="utf-8") as f:
        f.write(head + "\n".join(events) + "\n")
    return ass_path


def _window(src_in: str, src_out: str) -> tuple[str, str]:
    a, b = parse_tc(src_in), parse_tc(src_out)
    def fmt(x):
        return str(int(x)) if float(x).is_integer() else f"{x:.3f}"
    return fmt(a), fmt(b - a)


def build_vo_segment_cmd(src, src_in, src_out, wav, plan: ReconcilePlan, out, srt_file: str = None, is_new_slide: bool = False,
                          lead_in: float = _LEAD_IN) -> list:
    ss, dur = _window(src_in, src_out)
    vfilters = [_VNORM]

    # 원본 비디오 기반 씬일 경우 기존 자막을 100% 가리는 마스킹 박스 필터 적용
    if not is_new_slide:
        vfilters.append(_MASK_BOX)

    if plan.video_setpts is not None:
        vfilters.append(f"setpts={plan.video_setpts:.4f}*PTS")
    if plan.video_freeze > 0:
        vfilters.append(f"tpad=stop_mode=clone:stop_duration={plan.video_freeze:.3f}")
    if lead_in > 0:
        # 씬 시작~나레이션 시작 사이 정지 여백 — 첫 프레임을 lead_in초만큼 유지
        vfilters.append(f"tpad=start_mode=clone:start_duration={lead_in:.3f}")

    if srt_file and os.path.exists(srt_file):
        vfilters.append(ass_filter(srt_file))

    vchain = ",".join(vfilters)

    afilters = [_ANORM]
    if plan.audio_pad > 0:
        afilters.append(f"apad=pad_dur={plan.audio_pad:.3f}")
    if lead_in > 0:
        afilters.append(f"adelay={int(lead_in * 1000)}:all=1")
    achain = ",".join(afilters)

    return [
        "ffmpeg", "-y",
        "-ss", ss, "-t", dur, "-i", src,   # 0: source video window
        "-i", wav,                          # 1: new narration audio only
        "-filter_complex",
        f"[0:v]{vchain}[v];[1:a]{achain}[a]",
        "-map", "[v]", "-map", "[a]",
        "-t", f"{plan.scene_dur + lead_in:.3f}",
        "-c:v", "libx264", "-preset", "medium", "-pix_fmt", "yuv420p", "-r", FPS,
        "-c:a", "aac", "-ar", "48000", "-ac", "2",
        out,
    ]


_CUT_FADE = 0.25  # 원본 구간 앞뒤로 붙이는 짧은 페이드 — 하드컷의 어색함 제거용.


def build_original_segment_cmd(src, src_in, src_out, out, fade: float = _CUT_FADE) -> list:
    """원본 영상 구간을 그대로 잘라 쓴다(CLIP·ORIGINAL·EXTERNAL).

    유튜브 인터뷰처럼 임의 지점에서 잘라온 구간은 시작·끝이 하드컷으로 뚝 끊기면
    어색하다 — 앞뒤로 짧은 페이드를 걸어 자연스러운 인·아웃을 만든다.
    """
    ss, dur_s = _window(src_in, src_out)
    dur = parse_tc(src_out) - parse_tc(src_in)
    fade = min(fade, dur / 2 - 0.01) if dur > 0.1 else 0
    vf = _VNORM
    af = _ANORM
    if fade > 0:
        vf += f",fade=t=in:st=0:d={fade:.3f},fade=t=out:st={dur - fade:.3f}:d={fade:.3f}"
        af += f",afade=t=in:st=0:d={fade:.3f},afade=t=out:st={dur - fade:.3f}:d={fade:.3f}"
    return [
        "ffmpeg", "-y",
        "-ss", ss, "-t", dur_s, "-i", src,
        "-vf", vf,
        "-af", af,
        "-c:v", "libx264", "-preset", "medium", "-pix_fmt", "yuv420p", "-r", FPS,
        "-c:a", "aac", "-ar", "48000", "-ac", "2",
        out,
    ]


def build_concat_cmd(list_file, out) -> list:
    return [
        "ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i", list_file,
        "-c", "copy", out,
    ]


def build_bgm_duck_cmd(base, bgm, out, bgm_gain_db: float = -18.0) -> list:
    fc = (
        f"[1:a]volume={bgm_gain_db}dB,aresample=48000[bg];"
        f"[bg][0:a]sidechaincompress=threshold=0.05:ratio=8:attack=20:release=300[duck];"
        f"[0:a][duck]amix=inputs=2:duration=first:normalize=0[a]"
    )
    return [
        "ffmpeg", "-y",
        "-i", base,                          # 0: video + narration/interview audio
        "-stream_loop", "-1", "-i", bgm,     # 1: BGM
        "-filter_complex", fc,
        "-map", "0:v", "-map", "[a]",
        "-c:v", "copy", "-c:a", "aac", "-ar", "48000", "-ac", "2",
        "-shortest", out,
    ]


def run_ffmpeg(args: list) -> None:
    subprocess.run(args, check=True)


def build_slide_segment_cmd(png, wav, dur: float, out, srt_file: str = None,
                            zoom: float = 0.06, lead_in: float = _LEAD_IN,
                            fade: float = _CUT_FADE) -> list:
    """PPT 슬라이드 PNG + 새 나레이션 → 한 씬.

    정지 이미지를 그대로 40~90초 물리면 '슬라이드쇼'로 보이므로(1차 시도의 실패
    지점), 아주 느린 푸시인을 넣어 영상의 호흡을 만든다. zoom은 전체 구간에 걸친
    확대 비율(기본 6%) — 눈에 띄는 효과가 아니라 화면이 죽지 않을 정도만.
    슬라이드는 신규 화면이라 구 자막 마스킹(_MASK_BOX)은 쓰지 않는다.

    zoom=0이면 푸시인 없이 완전 고정 화면. zoompan은 크롭 원점을 프레임마다 정수로
    끊어 계산해 1px 단위로 떨리는데(2026-08-24 확인), 애니메이션을 쓰지 않는 씬에서는
    그 떨림만 남아 화면이 계속 흔들리는 것처럼 보인다.

    앞뒤로 build_original_segment_cmd와 같은 0.25초 페이드를 건다 — 이 씬은 페이드가
    없어서, CLIP(인터뷰 클립)처럼 이미 페이드인/아웃하는 씬과 이어질 때 밝은 화면에서
    검정으로 뚝 끊기는 게 도드라졌다(사용자 지적: 인터뷰 전환 컷이 튀어 보임).
    """
    total = dur + lead_in
    frames = max(1, int(round(total * float(FPS))))
    if zoom > 0:
        # zoompan은 확대 시 크롭 중심이 흔들리기 쉬워, 미리 2배로 키운 뒤 확대하고
        # 중심(iw/2-(iw/zoom/2))을 고정해 부드럽게 밀어 넣는다.
        vfilters = [
            f"scale={int(W)*2}:-2",
            f"zoompan=z='1+{zoom}*on/{frames}':d={frames}:s={W}x{H}"
            f":x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)':fps={FPS}",
            "setsar=1",
        ]
    else:
        vfilters = [f"scale={W}:{H}", f"fps={FPS}", "setsar=1"]
    seg_fade = min(fade, total / 2 - 0.01) if fade > 0 and total > 0.1 else 0
    if seg_fade > 0:
        vfilters.append(f"fade=t=in:st=0:d={seg_fade:.3f}")
        vfilters.append(f"fade=t=out:st={total - seg_fade:.3f}:d={seg_fade:.3f}")
    if srt_file and os.path.exists(srt_file):
        vfilters.append(ass_filter(srt_file))
    vchain = ",".join(vfilters)

    afilters = [_ANORM]
    if lead_in > 0:
        afilters.append(f"adelay={int(lead_in * 1000)}:all=1")
    achain = ",".join(afilters)

    return [
        "ffmpeg", "-y",
        "-loop", "1", "-i", png,
        "-i", wav,
        "-filter_complex", f"[0:v]{vchain}[v];[1:a]{achain}[a]",
        "-map", "[v]", "-map", "[a]",
        "-t", f"{total:.3f}",
        "-c:v", "libx264", "-preset", "medium", "-pix_fmt", "yuv420p", "-r", FPS,
        "-c:a", "aac", "-ar", "48000", "-ac", "2",
        out,
    ]
