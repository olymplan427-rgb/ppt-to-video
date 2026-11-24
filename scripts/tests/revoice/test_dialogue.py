"""다화자(팟캐스트) 대본 파싱 — API 호출 없이 검증 가능한 부분만.

synth_dialogue 자체는 Gemini 호출이라 여기서 안 돌린다. 대신 그 앞단인
parse_dialogue와, 체크포인트 계산에 쓰이는 _spoken_text를 고정한다 —
여기가 틀어지면 화면 요소가 엉뚱한 시점에 등장하거나 라벨이 낭독된다.
"""
from revoice.revoice_tts import parse_dialogue
from revoice.build import _spoken_text, _voice_key, _dialogue_cfg


P07 = """진행자: 그런데 이런 수업, 아이 입장에선 부담스럽지 않을까요?
전문가: 그 질문은 아이한테 직접 듣는 게 낫겠습니다.

> (인터뷰 원음 발췌 — "설명하다 보니 부족한 부분을 알 수 있게 되었어요.")

[BEAT 0.8s]
진행자: 틀려도 괜찮다는 마음이 먼저 생겼네요."""


def test_turns_in_order_with_beat():
    assert parse_dialogue(P07) == [
        ("진행자", "그런데 이런 수업, 아이 입장에선 부담스럽지 않을까요?"),
        ("전문가", "그 질문은 아이한테 직접 듣는 게 낫겠습니다."),
        ("BEAT", "0.8"),
        ("진행자", "틀려도 괜찮다는 마음이 먼저 생겼네요."),
    ]


def test_clip_quote_line_is_not_spoken():
    """`>` 인용 줄은 CLIP 원음이 재생될 자리 — TTS가 읽으면 원음과 겹쳐 들린다."""
    spoken = _spoken_text(P07)
    assert "인터뷰 원음 발췌" not in spoken
    assert "설명하다 보니" not in spoken


def test_speaker_labels_stripped_for_checkpoints():
    """대본에 '진행자:'가 남으면 whisper 전사와 단어 수가 어긋나 타이밍이 밀린다."""
    spoken = _spoken_text(P07)
    assert "진행자" not in spoken and "전문가" not in spoken
    assert spoken.startswith("그런데 이런 수업")
    assert spoken.endswith("먼저 생겼네요.")


def test_single_speaker_narration_passes_through_unchanged():
    """1인 나레이션 씬은 예전 동작 그대로 — 구 23슬라이드판 산출물 보호."""
    plain = "학부모가 학원을 고르는 기준이 바뀌었습니다. 오늘은 그 기준을 확인합니다."
    assert parse_dialogue(plain) == []
    assert _spoken_text(plain) == plain


class _Scene:
    def __init__(self, narration):
        self.narration = narration


class _M:
    tts_style = "차분하게: "
    voice = "Charon"


def test_voice_key_unchanged_for_single_speaker():
    """1인 씬 키가 바뀌면 기존 음성 캐시가 통째로 무효화돼 TTS를 전부 다시 호출한다."""
    import hashlib, json
    sc = _Scene("평범한 1인 나레이션입니다.")
    legacy = hashlib.sha1(json.dumps(
        [sc.narration, _M.tts_style, _M.voice], ensure_ascii=False).encode()
    ).hexdigest()[:16]
    assert _voice_key(sc, _M()) == legacy


def test_voice_key_tracks_voice_assignment_for_dialogue():
    """화자별 보이스를 바꾸면 다시 합성돼야 한다."""
    sc = _Scene(P07)
    m = _M()
    before = _voice_key(sc, m)
    m.voices = {"진행자": "Puck", "전문가": "Charon"}
    assert _voice_key(sc, m) != before


def test_default_dialogue_voices_are_two_distinct():
    """Gemini 다화자 모드는 화자 2명이 상한 — 서로 다른 보이스여야 구분된다."""
    _, voices = _dialogue_cfg(_M())
    assert set(voices) == {"진행자", "전문가"}
    assert len(set(voices.values())) == 2


# ---------------------------------------------------------------- TTS 재시도

class _Boom(Exception):
    def __init__(self, code):
        super().__init__(f"HTTP {code}")
        self.code = code


class _FlakyClient:
    """앞의 fail_times번은 code로 실패하고 그다음 성공하는 가짜 클라이언트."""
    def __init__(self, code, fail_times):
        self.code, self.left, self.calls = code, fail_times, 0
        self.models = self

    def generate_content(self, **kw):
        self.calls += 1
        if self.left > 0:
            self.left -= 1
            raise _Boom(self.code)
        return "ok"


def test_transient_error_is_retried(monkeypatch):
    """503 한 번에 렌더 전체가 죽던 문제 — 44씬 연속 합성의 실제 실패 원인이었다."""
    import time
    from revoice import revoice_tts as t
    monkeypatch.setattr(time, "sleep", lambda s: None)
    c = _FlakyClient(503, 2)
    assert t._generate(c, model="m") == "ok"
    assert c.calls == 3


def test_permanent_error_is_not_retried(monkeypatch):
    """400 같은 영구 오류를 6번 재시도하면 실패를 늦게 알게 될 뿐이다."""
    import time, pytest
    from revoice import revoice_tts as t
    monkeypatch.setattr(time, "sleep", lambda s: None)
    c = _FlakyClient(400, 99)
    with pytest.raises(_Boom):
        t._generate(c, model="m")
    assert c.calls == 1


def test_retry_gives_up_after_cap(monkeypatch):
    import time, pytest
    from revoice import revoice_tts as t
    monkeypatch.setattr(time, "sleep", lambda s: None)
    c = _FlakyClient(503, 99)
    with pytest.raises(_Boom):
        t._generate(c, model="m")
    assert c.calls == 6


def test_beat_marker_never_reaches_subtitles(tmp_path):
    """[BEAT n]은 TTS용 제작 지시 — 화면에 찍히면 안 된다(8개 씬에서 실제로 노출됐음)."""
    from revoice import compose
    out = compose.create_scene_srt(
        "[BEAT 0.8s]\n진행자: 틀려도 괜찮다는 마음이 먼저 생겼네요.",
        5.0, str(tmp_path / "s.srt"))
    body = open(out, encoding="utf-8").read()
    assert "BEAT" not in body
    assert "틀려도 괜찮다는" in body


def test_subtitle_of_only_a_beat_is_empty(tmp_path):
    from revoice import compose
    assert compose.create_scene_srt("[BEAT 0.8s]", 5.0, str(tmp_path / "s.srt")) is None


# ------------------------------------------------- 자막 · 애니메이션 동기화

def test_speaker_labels_stripped_from_subtitles(tmp_path):
    """목소리가 이미 둘로 구분되므로 화면에 "전문가:"를 또 쓰지 않는다."""
    from revoice import compose
    out = compose.create_scene_srt(
        "진행자: 규모가 크다는 게 무슨 의미입니까?\n전문가: 구조라는 뜻입니다.",
        6.0, str(tmp_path / "s.srt"))
    body = open(out, encoding="utf-8").read()
    assert "진행자" not in body and "전문가" not in body
    assert "규모가 크다는" in body and "구조라는 뜻입니다" in body


def test_subtitle_palette_flips_on_light_background(tmp_path):
    """크림색 슬라이드 위에서 아이보리 박스는 배경에 묻힌다 — 색을 뒤집어야 한다."""
    from revoice import compose
    dark = compose.create_scene_srt("전문가: 가나다.", 3.0, str(tmp_path / "d.srt"),
                                    light_bg=False)
    light = compose.create_scene_srt("전문가: 가나다.", 3.0, str(tmp_path / "l.srt"),
                                     light_bg=True)
    d, l = open(dark, encoding="utf-8").read(), open(light, encoding="utf-8").read()
    assert "&H00E8F1F5" in d          # 아이보리 박스
    assert "&H00432B1B" in l          # 네이비 박스
    style = lambda t: next(x for x in t.splitlines() if x.startswith("Style: Default"))
    assert style(d) != style(l)


def test_checkpoints_only_land_on_expert_speech():
    """진행자 질문 중에 요소가 다 나와버리던 문제 — 체크포인트는 전문가 구간에만."""
    from revoice.build import _dialogue_checkpoints
    narration = ("진행자: 첫 질문입니다 정말 그런가요?\n"
                 "전문가: 첫째 이유입니다. 둘째 이유입니다.")
    cps = _dialogue_checkpoints(narration, [], 10.0)
    assert len(cps) == 2                     # 전문가 문장 2개만
    assert cps[0] > 5.0                      # 질문(전체의 절반 가까이) 이후에야 시작
    assert cps[-1] == 10.0                   # 마지막은 씬 끝에 고정


def test_single_speaker_narration_falls_back_to_sentence_checkpoints():
    from revoice.build import _dialogue_checkpoints
    assert _dialogue_checkpoints("평범한 1인 나레이션입니다.", [], 10.0) is None


def test_hold_frame_scene_does_not_trip_multi_deck_guard():
    """홀드 프레임도 SLIDE 타입이지만 덱 슬라이드가 아니다 — 덱 판정에서 빠져야 한다."""
    from revoice.build import _slide_no, _deck_dir
    files = ["deck_260821/slide_07.png", "clips_podcast/DP21_hold.png"]
    decks = {_deck_dir(f) for f in files if _slide_no(f) is not None}
    assert decks == {"deck_260821"}


# ------------------------------------------------- 자막 싱크 · 줄 경계

def test_subtitle_never_mixes_two_speakers(tmp_path):
    """진행자 질문 끝과 전문가 답 첫머리가 한 자막에 섞이면 안 된다."""
    from revoice import compose
    out = compose.create_scene_srt(
        "진행자: 짧게 묻습니다\n전문가: 짧게 답합니다",
        6.0, str(tmp_path / "s.srt"))
    lines = [l.split(",", 9)[9] for l in open(out, encoding="utf-8").read().splitlines()
             if l.startswith("Dialogue:")]
    assert lines == ["짧게 묻습니다", "짧게 답합니다"]


def test_subtitle_timing_follows_word_timestamps(tmp_path):
    """글자수 비례가 아니라 실제 발화 시각을 따라야 자막이 앞서 나가지 않는다."""
    from revoice import compose
    # 앞 문장은 빨리, 뒤 문장은 한참 뒤에 말한 경우
    words = ([{"start": 0.0 + i * 0.2, "end": 0.2 + i * 0.2, "word": "가"} for i in range(4)] +
             [{"start": 6.0 + i * 0.2, "end": 6.2 + i * 0.2, "word": "나"} for i in range(4)])
    out = compose.create_scene_srt("가 가 가 가.\n나 나 나 나.", 8.0,
                                   str(tmp_path / "s.srt"), words=words)
    ev = [l for l in open(out, encoding="utf-8").read().splitlines() if l.startswith("Dialogue:")]
    second_start = ev[1].split(",")[1]
    assert second_start.startswith("0:00:06"), second_start   # 6초 쉼이 반영돼야 함


def test_subtitle_falls_back_when_transcript_mismatches(tmp_path):
    """전사가 대본과 크게 어긋나면 추정을 믿지 말고 글자수 비례로 돌아간다."""
    from revoice import compose
    out = compose.create_scene_srt("가 나 다 라 마 바 사 아.", 8.0,
                                   str(tmp_path / "s.srt"),
                                   words=[{"start": 0.0, "end": 0.1, "word": "?"}])
    assert out and "Dialogue:" in open(out, encoding="utf-8").read()


def test_determiner_never_stranded_at_line_end():
    """'…이 네 / 가지로 모입니다' 처럼 관형사가 수량사와 갈라지면 자막이 어색해진다."""
    from revoice.compose import _pack_words
    out = _pack_words("원장님들과 상담해 보면 고민이 신기할 만큼 이 네 가지로 모입니다.", 28)
    assert all(c.split()[-1] not in ("이", "네", "그", "한") for c in out), out
    assert " ".join(out) == "원장님들과 상담해 보면 고민이 신기할 만큼 이 네 가지로 모입니다."


def test_subtitle_spans_are_always_positive_and_ordered(tmp_path):
    """읽기 속도 보정이 시작을 뒤로 밀어 길이가 음수인 자막이 나온 적이 있다(실측 -0.3초)."""
    from revoice import compose
    # 짧은 시간에 말이 몰린 경우 — 보정 여유가 없는 최악 조건
    words = [{"start": i * 0.12, "end": i * 0.12 + 0.1, "word": "말%d" % i} for i in range(40)]
    out = compose.create_scene_srt(
        "\n".join("진행자: 말%d 말%d 말%d 아주 긴 문장입니다." % (i, i, i) for i in range(6)),
        4.8, str(tmp_path / "s.srt"), words=words)
    def tc(s):
        h, m, r = s.split(":"); return int(h) * 3600 + int(m) * 60 + float(r)
    ev = [l.split(",", 9) for l in open(out, encoding="utf-8").read().splitlines()
          if l.startswith("Dialogue:")]
    prev = -1.0
    for e in ev:
        st, en = tc(e[1]), tc(e[2])
        assert en > st, (st, en, e[9])          # 길이는 항상 양수
        assert st >= prev - 1e-6, (prev, st)    # 시작은 단조 증가
        prev = st


def test_no_anim_flag_forces_static_slide():
    """표지·간지·영상 정리 슬라이드는 요소 애니메이션을 쓰지 않는다."""
    from revoice.manifest import Scene
    sc = Scene(id="DP12", type="SLIDE", file="deck_260821/slide_12.png",
               narration="전문가: 가나다.", flags=["no-anim"])
    assert "no-anim" in sc.flags
    plain = Scene(id="DP13", type="SLIDE", file="deck_260821/slide_13.png",
                  narration="전문가: 가나다.")
    assert not plain.flags


def test_zoom_zero_removes_pushin():
    """no-anim 씬은 푸시인도 빼야 한다 — zoompan이 1px씩 떨려 화면이 흔들려 보인다."""
    from revoice import compose
    still = compose.build_slide_segment_cmd("a.png", "a.wav", 10.0, "o.mp4", zoom=0.0)
    assert "zoompan" not in " ".join(still)
    moving = compose.build_slide_segment_cmd("a.png", "a.wav", 10.0, "o.mp4")
    assert "zoompan" in " ".join(moving)   # 기본값은 그대로 푸시인
