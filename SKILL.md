---
name: ppt-to-video
description: Use when the user wants to turn a PowerPoint/PPTX file into a narrated or 2-host podcast video with animated slide elements, Whisper-synced subtitles, and -16 LUFS mastering — or wants to re-cut/re-voice an existing presentation-style video on Windows or macOS.
---

# PPT to Video (v2)

새 PPTX를 받아 **1인 나레이션 또는 2인 팟캐스트 대화형 영상**으로 만들거나, 기존 설명회
영상을 씬 단위로 리컷/리보이스한다. Windows와 macOS를 모두 1급 지원한다.

**핵심 원칙 — 절대 어기지 말 것: 슬라이드를 HTML로 다시 그리지 않는다.** PPTX를
PowerPoint(Windows COM) 또는 LibreOffice(macOS/Linux/미설치 PC)로 원본 그대로 PNG/영상으로
뽑아 쓴다. 브랜드 자산(사진·폰트·레이아웃)을 100% 보존하기 위함이다.

## 최초 실행 시 — 환경 점검

**세션에서 처음 이 스킬을 쓸 때 먼저 실행:**

```bash
python scripts/check_env.py
```

`[MISSING]` 항목이 있으면 사용자에게 설치해도 되는지 먼저 물은 뒤, 출력에 나온 OS별 설치
명령을 실행한다. `GEMINI_API_KEY`는 프로젝트 루트의 `.env` 파일에 넣으면 Windows·macOS
어디서나 자동으로 읽힌다.

- **Windows + PowerPoint**: 슬라이드 PNG 추출 + 네이티브 요소 순차 페이드인 애니메이션 지원
- **macOS / Linux / PowerPoint 미설치 PC**: LibreOffice 헤드리스(`soffice`) + PyMuPDF로 슬라이드 PNG를 자동 추출하고, 애니메이션은 정지/푸시인 폴백으로 매끄럽게 동작 (`--reuse-pngs`로 PowerPoint/Keynote 수동 내보내기 PNG도 바로 사용 가능)

## 워크플로우

전체 단계·매니페스트 스키마·마감(`finalize`)·트러블슈팅은
[references/pipeline.md](references/pipeline.md) 참고. 요약:

1. **인제스트**: `python -m revoice.deck <파일.pptx> [출력폴더]`
   - Windows PowerPoint COM / macOS LibreOffice 자동 선택 → `slide_NN.png` + `outline.json`
   - 이미 수동으로 PNG를 뽑아 두었다면 `--reuse-pngs` 옵션 사용
2. **대본 기획 (1인 또는 2인 팟캐스트)**:
   - `outline.json`을 참고 자료로만 쓰고, 발표자 노트를 그대로 베끼지 않는다.
   - 2인 대화형은 `진행자: ...` / `전문가: ...` (또는 매니페스트 `voices`에 지정한 커스텀 화자명) 형식을 쓰면 `revoice.build`가 자동으로 Gemini 다화자 TTS(`MultiSpeakerVoiceConfig`)로 합성한다.
   - 자막은 원표기를 유지하되 발음만 고쳐야 하는 단어(예: `YES24`→`예스24`)는 매니페스트의 `pronounce` 사전에 등록한다.
3. **매니페스트 구성**: `SLIDE` / `VO` / `ORIGINAL` / `CLIP` / `EXTERNAL` 씬을 배합해 `manifest.json` 생성.
   - 기존 영상의 **수정판(예: 12월판)**을 만들 때는 **수정한 씬 ID에만 월별 접미사(예: `DP01_12`)**를 붙여 원본 캐시(`audio/`, `build/seg_*`, `build/fin/*`)를 100% 보호한다.
4. **미리보기 (API 비용 0원)**: `python -m revoice.build render --manifest <파일> --dry-run`
   - TTS 호출 횟수와 재빌드 대상 씬을 먼저 확인한다.
5. **세그먼트 렌더**: `python -m revoice.build render --manifest <파일>`
   - 변경분만 증분 빌드 (TTS → Whisper 단어 타임스탬프 전사 → PPT 애니메이션 → 자막/세그먼트 합성).
   - 특정 씬만 수정할 때는 `--only <씬ID>` 사용.
6. **마감(마스터링)**: `python -m revoice.finalize --manifest <파일> --bgm <bgm.wav> [--bgm-stop-before <광고씬ID>] [--copy-to <최종본.mp4>]`
   - 전 씬 `-16 LUFS` 음량 정규화 + 리미터 + 선택 오버레이 카드 + 광고 원음 직전 BGM 페이드아웃 + 최종 합본.
7. **음성·자막 검수**: 수정된 씬은 Whisper 전사 대조 및 오디오/프레임 확인으로 오발음·자막 끊김을 잡는다.

## 하지 말 것

- 슬라이드를 HTML/React로 재현 — 브랜드 룩이 어긋난다.
- 확정 안 된 수치·문구를 `⚠확인`/`[확인]` 표식 없이 그대로 나레이션에 넣기 — 렌더 게이트가 자동 차단한다.
- 씬 하나 고칠 때 `--all`로 전체 재빌드 — TTS 쿼터(무료 100회/일)와 시간을 낭비한다. `--only` 또는 수정 씬 접미사를 활용한다.
- 개인 로컬 절대경로(`C:\Users\...`)를 스크립트나 매니페스트에 하드코딩 — `.env`의 `VIDEO_SRC_DIR`와 상대경로를 쓴다.

## 알려진 제약

- **PPT 요소 애니메이션은 Windows + 데스크톱 PowerPoint 전용.** macOS/Linux/미설치 PC에서는 자동으로 정지/푸시인 슬라이드로 폴백되며, 슬라이드 PNG 추출은 LibreOffice 또는 `--reuse-pngs`로 동일하게 동작한다.
- **유튜브 640p 초과 화질은 PO 토큰 필요**: 로컬 원본 파일이 있으면 우선 사용하고, 없으면 640p 또는 `--cookies-from-browser chrome`을 안내한다.
- **PowerPoint `CreateVideo` 간헐 실패**: 4회 재시도 + 파일 잠금 대기가 내장돼 있다. 배치 렌더 로그에서 `렌더 실패` 여부를 꼭 확인한다.
