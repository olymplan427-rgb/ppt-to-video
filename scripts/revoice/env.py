# -*- coding: utf-8 -*-
"""환경변수·.env 로더 (외부 의존성 없음, Windows/macOS/Linux 공용).

프로젝트 루트나 작업 폴더의 `.env` 파일을 자동으로 읽어 현재 프로세스의 `os.environ`에 반영한다.
- `.env`에 값이 비어 있지 않게 명시된 항목은 전역 Windows/OS 환경변수보다 우선 적용된다
  (Windows 전역 환경변수는 건드리지 않으므로 다른 프로젝트와 충돌하지 않는다).
- `.env`에서 비워 둔 항목(`KEY=`)은 기존 OS/셸 환경변수를 그대로 유지한다.
"""
import os
from pathlib import Path

_LOADED_PATHS = set()
_DOTENV_KEYS = set()


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


def load_dotenv(start_dir: str = None, max_up: int = 4, override: bool = True) -> str:
    """가장 가까운 `.env` 파일을 찾아 `os.environ`에 반영하고 경로를 반환.

    override=True(기본): `.env`에 값이 채워진 항목은 전역 OS 환경변수보다 우선 적용한다.
    비어 있는 항목(`KEY=`)은 기존 환경변수를 덮어쓰지 않는다.
    """
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
                            if not v:
                                continue
                            if override:
                                os.environ[k] = v
                                _DOTENV_KEYS.add(k)
                            elif k not in os.environ or not os.environ[k]:
                                os.environ[k] = v
                                _DOTENV_KEYS.add(k)
                        _LOADED_PATHS.add(key)
                    return key
            if cur.parent == cur:
                break
            cur = cur.parent
    return ""


def is_from_dotenv(key: str) -> bool:
    """해당 키가 현재 프로세스에서 프로젝트 `.env`로부터 로드되었는지 여부."""
    return key in _DOTENV_KEYS


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
    """영상 소스 디렉터리 결정: VIDEO_SRC_DIR → SHARED_DRIVE_DIR/영상 소스 → 기본 상대경로."""
    load_dotenv(start_dir)
    env_dir = os.environ.get("VIDEO_SRC_DIR", "").strip()
    if env_dir:
        return os.path.abspath(os.path.expanduser(env_dir))
    shared_dir = os.environ.get("SHARED_DRIVE_DIR", "").strip()
    if shared_dir:
        cand = os.path.join(os.path.abspath(os.path.expanduser(shared_dir)), "영상 소스")
        if os.path.isdir(cand):
            return cand
    return os.path.abspath(default_rel) if default_rel else ""
