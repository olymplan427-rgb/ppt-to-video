# ppt-to-video 파이프라인 상세 (v2)

## 0. 코드 위치 및 환경 설정

이 스킬의 `scripts/revoice/`가 실제 파이썬 패키지다. 작업할 프로젝트 폴더에 `revoice/`를
두고(`python -m revoice.<모듈>` 형태로 실행), 루트에 `.env` 파일을 두면 `GEMINI_API_KEY`,
`VIDEO_SRC_DIR`, `WHISPER_MODEL` 등이 Windows·macOS·Linux 어디서나 자동으로 로드된다.

- 테스트: `cd scripts && python -m pytest tests -q`
- 환경 점검: `python scripts/check_env.py`

## 1. 인제스트 — `revoice/deck.py`

```bash
python -m revoice.deck <파일.pptx> [출력폴더=deck] [--backend auto|powerpoint|libreoffice|reuse] [--reuse-pngs]
```

- `backend=auto` 기본 동작:
  1. **Windows + PowerPoint**: PowerPoint COM으로 1920×1080 PNG 추출(열려 있는 다른 PPT는 닫지 않음).
  2. **macOS / Linux / PowerPoint 미설치 Windows**: `soffice --headless --convert-to pdf` → PyMuPDF(`fitz`) 또는 `pdftoppm`으로 1920×1080 PNG 변환.
  3. **수동 PNG 재사용**: PowerPoint for Mac / Keynote에서 직접 PNG로 내보내 `deck/slide_01.png` ... 으로 넣어 두었으면 `--reuse-pngs`로 `outline.json`만 생성.
- `outline.json`의 `source_pptx`는 이식성을 위해 가능하면 상대경로로 저장되며, 다른 PC에서 열었을 때도 `build.py`가 파일명 기준으로 자동 복구한다.

## 2. 매니페스트 스키마 — `revoice/manifest.py`

```python
Scene(id, type, file="", src_in="", src_out="", narration="", note="", flags=[])
```

| type | file 필요 | src_in/out | narration | 화면 | 소리 |
|---|---|---|---|---|---|
| `SLIDE` | O (`deck/slide_NN.png`) | 불필요(길이=TTS) | 필수 | PPT 애니메이션(Windows) 또는 정지/푸시인 | 신규 TTS |
| `VO` | 선택 | O(기존 영상 구간) 또는 file 지정 시 그 영상 | 필수 | 기존영상+구자막마스킹, 또는 file 그대로 | 신규 TTS |
| `ORIGINAL` | - | O | - | 기존영상 구간 | 원음 |
| `CLIP` | O | O | - | 그 파일 구간 | 원음 |
| `EXTERNAL` | O | O | - | 그 파일 구간(범퍼) | 원음 |

Manifest 최상위 필드:
- `source`, `source_file`, `voice`, `tts_style`, `output`, `bgm{engine,prompt,duck}`, `scenes[]`
- **2인 팟캐스트 전용**: `voices` (예: `{"진행자": "Erinome", "전문가": "Alnilam"}`), `dialogue_style` (짧게 유지 — 600자 이상 길면 모델 무음 패딩 오류 발생)
- **TTS 발음 보정**: `pronounce` (예: `{"YES24": "예스24", "채점표": "채쩜표"}` — 자막은 원표기를 유지하고 TTS 입력만 치환)

### 2인 다화자 TTS 주의사항 (`revoice/revoice_tts.py`)

- 씬 나레이션에 `진행자:` / `전문가:` (또는 `voices` 키에 등록된 화자명)가 있으면 자동으로 `MultiSpeakerVoiceConfig`를 호출한다.
- `[BEAT 0.8s]` 마커는 TTS 구간 사이에 무음을 삽입하며 자막에서는 자동 제거된다.
- 503/429 일시 오류는 지수 백오프로 최대 6회 자동 재시도한다.
- Gemini 3.1 TTS가 간헐적으로 뒤에 붙이는 긴 0-패딩 무음은 `_trim_trailing_silence`가 자동 압축한다.

## 3. 자막 & PPT 요소 애니메이션

### Whisper 단어 싱크 자막 (`revoice/compose.py`)
- 합성된 WAV를 `faster-whisper`로 전사(`audio/<id>.words.json` 캐시)해 실제 발화 시각에 자막을 맞춘다.
- 화자 줄 경계는 절대 섞지 않고, 관형사/수사(`이`, `네` 등)와 날짜(`10월 21일`)가 줄 끝에서 갈라지지 않게 보호한다.
- 하단 25% 밝기를 자동 측정해 어두운 슬라이드엔 아이보리 박스, 밝은 슬라이드엔 네이비 박스를 적용한다.
- `assets/fonts/WantedSans-*.ttf`가 있으면 설치 없이 `fontsdir`로 로드하고, 없으면 OS 기본 한글 폰트(macOS `Apple SD Gothic Neo`, Windows `Malgun Gothic`, Linux `Noto Sans CJK KR`)로 자동 폴백한다.

### PPT 요소 애니메이션 (`revoice/deck.py` + `_animate_slides.ps1`)
- **Windows + PowerPoint**: 슬라이드 상단 30%(제목·로고)는 도입부에 일괄 등장하고, 본문 요소만 전문가 발화 체크포인트(`_dialogue_checkpoints`)에 맞춰 순차 페이드인된다. `flags: ["no-anim"]` 슬라이드는 완전 고정(줌 떨림 제거).
- **macOS / Linux / PowerPoint 미설치**: `deck.has_powerpoint()`가 자동으로 감지해 정지/푸시인 모드로 오류 없이 진행한다.

## 4. 렌더 & 증분 캐시 — `revoice/build.py`

```bash
python -m revoice.build render --manifest <파일> --dry-run   # API 호출 없이 미리보기
python -m revoice.build render --manifest <파일>              # 변경분만 재빌드
python -m revoice.build render --manifest <파일> --only P03   # 특정 씬만
python -m revoice.build render --manifest <파일> --static-slides  # PPT 애니메이션 끄기
python -m revoice.build render --manifest <파일> --new-bgm    # BGM 새로 생성
```

- **월별 수정판 캐시 보호 전략**: 기존 영상의 일부만 고칠 때 수정한 씬 ID에만 접미사(예: `DP01_12`)를 붙이면 기존 캐시(`audio/DP*.wav`, `build/seg_DP*.mp4`, `build/fin/fin_DP*.mp4`)를 덮어쓰지 않고 100% 재사용할 수 있다.
- PowerShell UTF-8 출력이 내장돼 있어 `PYTHONUTF8=0` 같은 수동 환경변수 우회가 필요 없다.

## 5. 마감(마스터링) — `revoice/finalize.py`

```bash
python -m revoice.finalize --manifest <파일> --bgm bgm/track.wav \
  --bgm-stop-before DP44A --copy-to ../../최종본.mp4
```

- 모든 세그먼트를 `-16.0 LUFS` 정적 게인 + `alimiter`로 균일화하고 `build/fin/_cache.json`에 캐시한다.
- 원음 소리가 매우 작은 클립에는 `special_audio`(고역통과 + `afftdn` + 상한 게인)를, 자료 강조가 필요한 씬에는 `overlays`(라운드 패널 카드 크로스페이드)를 프로그래밍 방식으로 지정할 수 있다.
- `--bgm-stop-before <씬ID>`를 주면 광고/별도 홍보영상 시작 직전에 설명회 BGM을 1.5초 페이드아웃 후 음소거해 원음과 BGM이 겹치지 않게 한다.

## 6. Windows vs macOS 기능 비교

| 기능 | Windows + PowerPoint | macOS / PowerPoint 미설치 PC |
|---|---|---|
| PPT → 1920×1080 PNG 추출 | PowerPoint COM (자동) | LibreOffice 헤드리스 (자동) 또는 `--reuse-pngs` |
| PPT 요소 순차 페이드인 | 지원 (COM `CreateVideo`) | 자동 폴백 (정지/푸시인) |
| 1인/2인 Gemini TTS + 발음 보정 | 동일 | 동일 |
| Whisper 단어 싱크 자막 | 동일 (`WHISPER_MODEL=medium` 조정 가능) | 동일 |
| `-16 LUFS` 마감 · 오버레이 · BGM 덕킹 | 동일 | 동일 |
