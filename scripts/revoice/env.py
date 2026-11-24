# -*- coding: utf-8 -*-
"""환경변수·.env 로더 (외부 의존성 없음, Windows/macOS/Linux 공용).

프로젝트 루트나 작업 폴더의 `.env` 파일을 자동으로 읽어 `os.environ`에 채운다.
이미 셸 환경변수로 지정된 값은 덮어쓰지 않는다(셸 > .env > OS 레지스트리).
"""
import os
from pathlib import Path

_LOADED_PATHS = set()


def _parse_dotenv(path: Path) -> dict:
    out = {}
    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        return out
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, val = line.split("=", 1)
        key = key.strip()
        if key.startswith("export "):
            key = key[7:].strip()
        val = val.strip()
        if len(val) >= 2 and val[0] == val[-1] and val[0] in ("'", '"'):
            val = val[1:-1]
        elif " #" in val:
            val = val.split(" #", 1)[0].strip()
        if key:
            out[key] = val
    return out


def load_dotenv(start_dir: str = None, max_up: int = 4) -> str:
    """가장 가까운 `.env` 파일을 찾아 `os.environ`에 반영하고 경로를 반환."""
    roots = []
    if start_dir:
        roots.append(Path(start_dir).resolve())
    roots.append(Path.cwd().resolve())

    checked = set()
    for base in roots:
        cur = base
        for _ in range(max_up + 1):
            candidate = cur / ".env"
            key = str(candidate)
            if key not in checked:
                checked.add(key)
                if candidate.is_file():
                    if key not in _LOADED_PATHS:
                        for k, v in _parse_dotenv(candidate).items():
                            os.environ.setdefault(k, v)
                        _LOADED_PATHS.add(key)
                    return key
            if cur.parent == cur:
                break
            cur = cur.parent
    return ""


def get_api_key(start_dir: str = None) -> str:
    """GEMINI_API_KEY 탐색: 셸 환경변수 → .env → Windows 사용자 레지스트리."""
    load_dotenv(start_dir)
    key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not key and os.name == "nt":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment") as k:
                for name in ("GEMINI_API_KEY", "GOOGLE_API_KEY"):
                    try:
                        key, _ = winreg.QueryValueEx(k, name)
                        if key:
                            break
                    except FileNotFoundError:
                        continue
        except Exception:
            pass
    return key or ""


def resolve_video_src_dir(default_rel: str = "", start_dir: str = None) -> str:
    """영상 소스 디렉터리 결정: VIDEO_SRC_DIR(.env/환경변수) → 기본 상대경로."""
    load_dotenv(start_dir)
    env_dir = os.environ.get("VIDEO_SRC_DIR", "").strip()
    if env_dir:
        return os.path.abspath(os.path.expanduser(env_dir))
    return os.path.abspath(default_rel) if default_rel else ""
