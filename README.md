# ppt-to-video (v2)

*Read in other languages: **English** | [한국어](README.ko.md)*

An AI agent skill (Claude Code & Antigravity) and Python pipeline that turns PowerPoint
decks into **single-narrator or 2-host podcast briefing videos** with Whisper-synced
subtitles, incremental caching, and `-16 LUFS` mastering. Supports both **Windows and macOS**
without ever redrawing slides in HTML.

## Highlights

- **Direct PPTX authoring & editing (`revoice.ppt_edit`)**: Format-preserving run-level text replacement, brand-consistent slide duplication, bulk theme hex recoloring, high-DPI PDF page insertion, and fast per-slide PNG visual QA previews.
- **Pixel-faithful slide rendering**: Uses native PowerPoint COM on Windows (including element-by-element fade-in animations synced to speech checkpoints) and automatically falls back to LibreOffice headless + PyMuPDF (or `--reuse-pngs`) on macOS/Linux/no-PowerPoint machines.
- **Single-speaker & 2-host podcast TTS**: Gemini TTS with `MultiSpeakerVoiceConfig`, automatic 429/503 retries, silence padding trimmer, and a TTS-only pronunciation dictionary (`pronounce`) that keeps subtitles clean.
- **Whisper word-synced subtitles**: Aligns subtitles to real spoken timestamps via `faster-whisper`, prevents awkward line breaks, and flips subtitle contrast automatically for light vs. dark slide bottoms.
- **Incremental build cache**: Rebuilds only changed scenes (`--dry-run`, `--only`).
- **Mastering (`revoice.finalize`)**: `-16 LUFS` loudness normalization, timed overlay cards, pre-ad BGM fadeout, and sidechain BGM ducking.

## Quick Setup

```bash
# Claude Code
git clone https://github.com/olymplan427-rgb/ppt-to-video.git ~/.claude/skills/ppt-to-video

# Antigravity (optional)
git clone https://github.com/olymplan427-rgb/ppt-to-video.git ~/.gemini/antigravity/skills/ppt-to-video
```

Install Python dependencies and run the cross-platform environment checker:

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate   |   macOS/Linux: source .venv/bin/activate
pip install -r ~/.claude/skills/ppt-to-video/requirements.txt
python ~/.claude/skills/ppt-to-video/scripts/check_env.py
```

Create a `.env` file in your project root:

```env
GEMINI_API_KEY=your_api_key_here
```

## Basic Loop

```bash
python -m revoice.deck slides.pptx deck
python -m revoice.build render --manifest manifest.json --dry-run
python -m revoice.build render --manifest manifest.json
python -m revoice.finalize --manifest manifest.json --bgm bgm/bgm.wav
```

## Tests

```bash
cd scripts && python -m pytest tests -q
```

## License

MIT — see [LICENSE](LICENSE).
