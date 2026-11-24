# ppt-to-video (v2)

*다른 언어로 보기: [English](README.md) | **한국어***

PowerPoint 슬라이드를 **1인 나레이션 또는 2인 팟캐스트 대화형 설명회 영상**으로 만드는
AI 에이전트 스킬(Claude Code / Antigravity 공용) 및 Python 파이프라인이다.
**Windows와 macOS를 모두 지원**하며, 슬라이드는 절대 HTML로 다시 그리지 않는다.

## 주요 기능

- **원본 슬라이드 100% 보존**: Windows에서는 PowerPoint COM(네이티브 요소 페이드인 애니메이션 포함), macOS/Linux/미설치 PC에서는 LibreOffice 헤드리스(+PyMuPDF) 또는 수동 PNG(`--reuse-pngs`)로 자동 동작한다.
- **1인 나레이션 & 2인 팟캐스트 TTS**: Gemini TTS(`MultiSpeakerVoiceConfig`)로 두 화자의 자연스러운 대화를 합성하고, 503/429 재시도 · 무음 패딩 트리밍 · 발음 보정 사전(`pronounce`)을 내장했다.
- **Whisper 단어 싱크 자막**: 합성 음성을 `faster-whisper`로 전사해 실제 발화 시각에 자막을 맞추고, 배경 밝기에 따라 자막 박스 색상을 자동 반전한다.
- **증분 캐시 빌드**: 대본이나 화면이 바뀐 씬만 다시 만들어 TTS 쿼터와 시간을 절약한다.
- **마감 마스터링(`revoice.finalize`)**: 전 씬 `-16 LUFS` 음량 정규화, 자료 카드 오버레이, 광고 구간 직전 BGM 페이드아웃, 음성 아래 BGM 자동 덕킹을 처리한다.

## 설치 (Claude Code / Antigravity)

```bash
# Claude Code용
git clone https://github.com/olymplan427-rgb/ppt-to-video.git ~/.claude/skills/ppt-to-video

# Antigravity용 (원하는 경우 동일 저장소 클론)
git clone https://github.com/olymplan427-rgb/ppt-to-video.git ~/.gemini/antigravity/skills/ppt-to-video
```

가상환경 생성 후 패키지 설치 및 환경 점검:

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate   |   macOS/Linux: source .venv/bin/activate
pip install -r ~/.claude/skills/ppt-to-video/requirements.txt
python ~/.claude/skills/ppt-to-video/scripts/check_env.py
```

프로젝트 루트에 `.env` 파일을 만들고 API 키를 넣는다:

```env
GEMINI_API_KEY=your_api_key_here
```

## 기본 실행 루프

```bash
python -m revoice.deck slides.pptx deck                               # 1. 슬라이드 PNG + outline.json
python -m revoice.build render --manifest manifest.json --dry-run     # 2. 미리보기 (API 호출 0)
python -m revoice.build render --manifest manifest.json               # 3. 변경 씬만 증분 빌드
python -m revoice.finalize --manifest manifest.json --bgm bgm/bgm.wav # 4. -16 LUFS 마감 합본
```

## 테스트

```bash
cd scripts && python -m pytest tests -q
```

## 라이선스

MIT — [LICENSE](LICENSE) 참고.
