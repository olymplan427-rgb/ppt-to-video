# -*- coding: utf-8 -*-
"""ppt-to-video 환경 점검 (Windows / macOS / Linux 공용).

무엇이 빠졌는지만 보고한다 — 동의 없이 자동으로 설치하지 않는다.
실행: python check_env.py
"""
import os
import platform
import shutil
import subprocess
import sys

# revoice.env가 같은 폴더/상위 폴더에 있으면 .env 로더 재사용
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)


def _ok(label, detail=""):
    print(f"[OK]      {label}" + (f" — {detail}" if detail else ""))


def _warn(label, detail=""):
    print(f"[INFO]    {label}" + (f" — {detail}" if detail else ""))


def _missing(label, fix):
    print(f"[MISSING] {label}")
    print(f"          설치: {fix}")


def _is_mac():
    return sys.platform == "darwin"


def _is_win():
    return os.name == "nt"


def check_slide_renderer():
    """슬라이드 PNG 내보내기 / 애니메이션 렌더러 점검."""
    has_ppt = False
    if _is_win():
        try:
            r = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command",
                 "try { $a = New-Object -ComObject PowerPoint.Application; "
                 "if ($a.Presentations.Count -eq 0) { $a.Quit() }; "
                 "Write-Output 'OK' } catch { Write-Output 'FAIL' }"],
                capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=20)
            has_ppt = "OK" in r.stdout
        except Exception:
            has_ppt = False

    lo = shutil.which("soffice") or shutil.which("libreoffice")
    if not lo:
        for cand in (
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ):
            if os.path.exists(cand):
                lo = cand
                break

    if has_ppt:
        _ok("PowerPoint COM (Windows)", "슬라이드 PNG + 요소 애니메이션 모두 지원")
        return True
    if lo:
        _ok("LibreOffice (크로스플랫폼)", f"{lo} — 슬라이드 PNG 자동 변환 지원 (요소 애니메이션은 정지/푸시인 폴백)")
        return True

    if _is_mac():
        _missing(
            "슬라이드 렌더러 (LibreOffice 권장)",
            "brew install --cask libreoffice  (또는 PowerPoint/Keynote에서 PNG 수동 내보내기 후 --reuse-pngs 사용)")
    elif _is_win():
        _missing(
            "PowerPoint COM 또는 LibreOffice",
            "Microsoft 365 데스크톱 앱(요소 애니메이션 포함) 또는 `winget install TheDocumentFoundation.LibreOffice`")
    else:
        _missing("LibreOffice", "sudo apt install libreoffice  (또는 배포판 패키지 관리자)")
    return False


def check_ffmpeg():
    ff = shutil.which("ffmpeg")
    fp = shutil.which("ffprobe")
    if ff and fp:
        _ok("ffmpeg / ffprobe", ff)
        return True
    if _is_mac():
        fix = "brew install ffmpeg"
    elif _is_win():
        fix = "winget install Gyan.FFmpeg  (설치 후 터미널 재시작)"
    else:
        fix = "sudo apt install ffmpeg"
    _missing("ffmpeg / ffprobe", fix)
    return False


def check_node():
    path = shutil.which("node")
    if path:
        _ok("Node.js (선택)", path)
        return True
    fix = "brew install node" if _is_mac() else (
        "winget install OpenJS.NodeJS.LTS" if _is_win() else "sudo apt install nodejs"
    )
    _warn("Node.js 미설치 (선택)", f"유튜브 클립 다운로드 시에만 필요 — {fix}")
    return True  # 선택 항목이므로 실패 처리하지 않음


def check_python_packages():
    pkgs = [
        ("pptx", "python-pptx", True),
        ("google.genai", "google-genai", True),
        ("faster_whisper", "faster-whisper", True),
        ("PIL", "Pillow", True),
        ("numpy", "numpy", True),
        ("pymupdf", "pymupdf", not _is_win()),
        ("yt_dlp", "yt-dlp", False),
    ]
    all_ok = True
    for mod, pip_name, required in pkgs:
        try:
            __import__(mod)
            _ok(f"python: {pip_name}")
        except ImportError:
            if required:
                _missing(f"python: {pip_name}", f"pip install {pip_name}")
                all_ok = False
            else:
                _warn(f"python: {pip_name} 미설치 (선택)", f"필요 시 `pip install {pip_name}`")
    return all_ok


def check_gemini_key():
    key = ""
    dot_path = ""
    try:
        from revoice import env as _env
        dot_path = _env.load_dotenv()
        key = _env.get_api_key()
    except Exception:
        key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY") or ""

    if key:
        src = f".env ({dot_path})" if dot_path and os.environ.get("GEMINI_API_KEY") == key else "환경변수"
        _ok("GEMINI_API_KEY", f"{len(key)}자 ({src})")
        return True
    _missing(
        "GEMINI_API_KEY",
        "https://aistudio.google.com/apikey 에서 발급 후 프로젝트 루트 `.env` 파일에 "
        "`GEMINI_API_KEY=발급받은키` 입력")
    return False


def check_video_src_dir():
    try:
        from revoice import env as _env
        _env.load_dotenv()
    except Exception:
        pass
    vdir = os.environ.get("VIDEO_SRC_DIR", "").strip()
    if not vdir:
        # 기본 상대경로 확인 (영상 소스 폴더가 프로젝트 루트에 있는 구조)
        for cand in ("영상 소스", os.path.join("..", "..", "영상 소스")):
            if os.path.isdir(cand):
                vdir = os.path.abspath(cand)
                break
    if vdir and os.path.isdir(vdir):
        cnt = len([f for f in os.listdir(vdir) if not f.startswith(".")])
        _ok("영상 소스 폴더", f"{vdir} ({cnt}개 파일)")
    elif vdir:
        _warn("VIDEO_SRC_DIR 경로 없음", f"{vdir} — 인터뷰/광고 클립을 쓰는 프로젝트라면 경로를 확인하세요")


def main():
    print(f"=== ppt-to-video 환경 점검 ({platform.system()} {platform.release()}, Python {platform.python_version()}) ===\n")
    results = {
        "슬라이드 렌더러": check_slide_renderer(),
        "ffmpeg (합성·인코딩, 필수)": check_ffmpeg(),
        "Python 패키지": check_python_packages(),
        "GEMINI_API_KEY (TTS·BGM 생성, 필수)": check_gemini_key(),
    }
    check_node()
    check_video_src_dir()
    print()
    missing = [k for k, v in results.items() if not v]
    if not missing:
        print("모든 필수 항목 준비 완료 — 바로 작업 가능합니다.")
    else:
        print(f"빠진 필수 항목 {len(missing)}개. 위 안내대로 설치한 뒤 다시 점검하세요.")
    return 0 if not missing else 1


if __name__ == "__main__":
    sys.exit(main())
