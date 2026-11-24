# -*- coding: utf-8 -*-
"""PPTX 인제스트 — 회사가 준 발표자료를 '그대로' 렌더해서 씬 소재로 쓴다.

Windows + PowerPoint 환경에서는 PowerPoint COM으로 원본 폰트·사진·도형과 요소
애니메이션을 그대로 살리고, macOS/Linux 또는 PowerPoint가 없는 PC에서는
LibreOffice 헤드리스(+PyMuPDF/pdftoppm) 또는 사전 내보낸 PNG를 자동 사용한다.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

W, H = 1920, 1080
_HAS_PPT_CACHE = None


def _ps(script: str, timeout: int = 120) -> str:
    """PowerShell 실행 — UTF-8 출력을 강제해 cp949 디코딩 오류를 원천 차단한다."""
    wrapped = (
        "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8\n"
        "$OutputEncoding = [System.Text.Encoding]::UTF8\n"
        + script
    )
    r = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", wrapped],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=timeout)
    if r.returncode != 0:
        raise RuntimeError(f"powershell failed: {r.stderr.strip() or r.stdout.strip()}")
    return r.stdout


def has_powerpoint() -> bool:
    """Windows PowerPoint COM 사용 가능 여부(프로세스당 1회 캐시)."""
    global _HAS_PPT_CACHE
    if _HAS_PPT_CACHE is not None:
        return _HAS_PPT_CACHE
    if os.name != "nt":
        _HAS_PPT_CACHE = False
        return False
    try:
        out = _ps(
            "try { "
            "$a = New-Object -ComObject PowerPoint.Application; "
            "if ($a.Presentations.Count -eq 0) { $a.Quit() }; "
            "Write-Output 'OK' "
            "} catch { Write-Output 'FAIL' }",
            timeout=20)
        _HAS_PPT_CACHE = "OK" in out
    except Exception:
        _HAS_PPT_CACHE = False
    return _HAS_PPT_CACHE


def find_libreoffice() -> str:
    """크로스플랫폼 LibreOffice(soffice) 실행 파일 탐색."""
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return ""


def _list_slide_pngs(out_dir: str) -> list:
    if not os.path.isdir(out_dir):
        return []
    return sorted(
        os.path.join(out_dir, f) for f in os.listdir(out_dir)
        if f.startswith("slide_") and f.endswith(".png")
    )


def _export_slides_powerpoint(pptx: str, out_dir: str) -> list:
    """Windows PowerPoint COM으로 슬라이드 PNG 내보내기(열린 PPT는 닫지 않음)."""
    pptx_esc = os.path.abspath(pptx).replace("'", "''")
    out_esc = os.path.abspath(out_dir).replace("'", "''")
    os.makedirs(out_dir, exist_ok=True)
    _ps(f"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open('{pptx_esc}', $true, $false, $false)
try {{
  foreach ($i in 1..$pres.Slides.Count) {{
    $pres.Slides.Item($i).Export(('{out_esc}\\slide_{{0:d2}}.png' -f $i), 'PNG', {W}, {H})
  }}
}} finally {{
  $pres.Close()
  if ($app.Presentations.Count -eq 0) {{ $app.Quit() }}
}}
""", timeout=300)
    return _list_slide_pngs(out_dir)


def _pdf_to_pngs(pdf_path: str, out_dir: str) -> list:
    """PDF 각 페이지를 1920x1080 slide_NN.png로 렌더(PyMuPDF 우선, pdftoppm 폴백)."""
    os.makedirs(out_dir, exist_ok=True)
    try:
        try:
            import pymupdf as fitz  # PyMuPDF >= 1.24
        except ImportError:
            import fitz  # legacy PyMuPDF
        doc = fitz.open(pdf_path)
        try:
            for idx, page in enumerate(doc, 1):
                rect = page.rect
                sx = W / rect.width if rect.width else 2.0
                sy = H / rect.height if rect.height else 2.0
                pix = page.get_pixmap(matrix=fitz.Matrix(sx, sy), alpha=False)
                out_png = os.path.join(out_dir, f"slide_{idx:02d}.png")
                if pix.width == W and pix.height == H:
                    pix.save(out_png)
                else:
                    from PIL import Image
                    im = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    im = im.resize((W, H), Image.LANCZOS)
                    im.save(out_png)
        finally:
            doc.close()
        return _list_slide_pngs(out_dir)
    except ImportError:
        pass

    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm:
        with tempfile.TemporaryDirectory() as tmp:
            prefix = os.path.join(tmp, "p")
            subprocess.run(
                [pdftoppm, "-png", "-r", "150", pdf_path, prefix],
                check=True, capture_output=True)
            pages = sorted(f for f in os.listdir(tmp) if f.endswith(".png"))
            from PIL import Image
            for idx, name in enumerate(pages, 1):
                im = Image.open(os.path.join(tmp, name)).convert("RGB")
                if im.size != (W, H):
                    im = im.resize((W, H), Image.LANCZOS)
                im.save(os.path.join(out_dir, f"slide_{idx:02d}.png"))
        return _list_slide_pngs(out_dir)

    raise RuntimeError(
        "LibreOffice PDF를 PNG로 변환하려면 PyMuPDF(`pip install pymupdf`) "
        "또는 `pdftoppm`(poppler)이 필요합니다."
    )


def _export_slides_libreoffice(pptx: str, out_dir: str) -> list:
    """macOS/Linux/PowerPoint 미설치 PC용: LibreOffice headless → PDF → 1920x1080 PNG."""
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError("LibreOffice(soffice)를 찾지 못했습니다.")
    pptx = os.path.abspath(pptx)
    with tempfile.TemporaryDirectory() as tmp:
        r = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, pptx],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=300)
        if r.returncode != 0:
            raise RuntimeError(f"LibreOffice 변환 실패: {r.stderr.strip() or r.stdout.strip()}")
        pdfs = [os.path.join(tmp, f) for f in os.listdir(tmp) if f.lower().endswith(".pdf")]
        if not pdfs:
            raise RuntimeError("LibreOffice가 PDF를 생성하지 못했습니다.")
        return _pdf_to_pngs(pdfs[0], out_dir)


def export_slides(pptx: str, out_dir: str, backend: str = "auto") -> list:
    """PPTX의 각 슬라이드를 out_dir/slide_NN.png 로 내보내고 경로 목록 반환.

    backend:
      - 'auto'        : Windows PowerPoint COM → LibreOffice → 기존 PNG 재사용 순으로 자동 선택
      - 'powerpoint'  : Windows PowerPoint COM 강제
      - 'libreoffice' : LibreOffice 헤드리스 강제 (macOS/Linux/Windows 공용)
      - 'reuse'       : 이미 내보내 둔 out_dir/slide_NN.png 그대로 사용
    """
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)

    if backend == "reuse":
        existing = _list_slide_pngs(out_dir)
        if not existing:
            raise RuntimeError(f"{out_dir}에 기존 slide_NN.png가 없습니다.")
        return existing

    if backend == "powerpoint":
        return _export_slides_powerpoint(pptx, out_dir)

    if backend == "libreoffice":
        return _export_slides_libreoffice(pptx, out_dir)

    # auto
    if os.name == "nt" and has_powerpoint():
        return _export_slides_powerpoint(pptx, out_dir)

    if find_libreoffice():
        print("[deck] PowerPoint COM 미사용 — LibreOffice 헤드리스로 슬라이드 PNG를 추출합니다.")
        return _export_slides_libreoffice(pptx, out_dir)

    existing = _list_slide_pngs(out_dir)
    if existing:
        print(f"[deck] 기존 슬라이드 PNG {len(existing)}장을 재사용합니다 ({out_dir}).")
        return existing

    if sys.platform == "darwin":
        hint = (
            "macOS에서는 아래 둘 중 하나를 선택하세요:\n"
            "  1) LibreOffice 설치: `brew install --cask libreoffice` 후 재실행\n"
            "  2) PowerPoint/Keynote에서 '파일 > 내보내기 > PNG'로 "
            f"{out_dir}/slide_01.png ... 형식으로 저장한 뒤 `--reuse-pngs`로 실행"
        )
    else:
        hint = (
            "Microsoft PowerPoint 또는 LibreOffice(`soffice`)를 설치하거나, "
            f"{out_dir}/slide_01.png ... 형식으로 PNG를 직접 내보낸 뒤 `--reuse-pngs`로 실행하세요."
        )
    raise RuntimeError("슬라이드 PNG를 내보낼 렌더러(PowerPoint/LibreOffice)가 없습니다.\n" + hint)


def extract_outline(pptx: str) -> list:
    """슬라이드별 {n, title, body[], notes} — 나레이션 기획의 원자료."""
    from pptx import Presentation
    out = []
    for n, slide in enumerate(Presentation(pptx).slides, 1):
        lines = [ln.strip() for sh in slide.shapes if sh.has_text_frame
                 for ln in sh.text_frame.text.splitlines() if ln.strip()]
        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        out.append({"n": n, "title": lines[0] if lines else "",
                    "body": lines[1:], "notes": notes})
    return out


def _portable_pptx_path(pptx: str) -> str:
    """다른 PC에서도 동작하도록 가능하면 CWD 기준 상대경로(POSIX 슬래시)로 저장."""
    abs_p = os.path.abspath(pptx)
    try:
        rel = os.path.relpath(abs_p, os.getcwd())
        return rel.replace("\\", "/")
    except ValueError:
        return abs_p.replace("\\", "/")


def ingest(pptx: str, out_dir: str = "deck", backend: str = "auto") -> dict:
    """슬라이드 PNG + outline.json 을 만들고 요약을 반환."""
    pngs = export_slides(pptx, out_dir, backend=backend)
    outline = extract_outline(pptx)
    for item in outline:                       # 슬라이드 번호로 PNG 경로 연결
        p = os.path.join(out_dir, f"slide_{item['n']:02d}.png")
        item["png"] = p.replace("\\", "/") if os.path.exists(p) else ""
    path = os.path.join(out_dir, "outline.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"source_pptx": _portable_pptx_path(pptx), "slides": outline},
                  f, ensure_ascii=False, indent=2)
    return {"pngs": len(pngs), "outline": path, "slides": len(outline)}


def render_animated_slides(pptx: str, durations: dict, out_dir: str = "deck/anim",
                           effect_duration: float = 0.5, stagger: float = 0.25,
                           width: int = W, fps: int = 30) -> dict:
    """슬라이드별 요소 순차 페이드인 애니메이션 mp4 생성 — Windows+PowerPoint 전용.

    macOS/Linux 또는 PowerPoint 미설치 환경에서는 빈 dict를 반환해 호출부가
    정지/푸시인 폴백으로 자연스럽게 넘어가게 한다.
    """
    if not has_powerpoint():
        print("[deck] PowerPoint COM 미감지 — 요소 애니메이션 대신 정지/푸시인 슬라이드를 사용합니다.")
        return {}

    pptx = os.path.abspath(pptx)
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    durations_path = os.path.join(out_dir, "_durations.json")
    with open(durations_path, "w", encoding="utf-8") as f:
        json.dump({str(k): v for k, v in durations.items()}, f)

    script = os.path.join(os.path.dirname(__file__), "_animate_slides.ps1")
    timeout = 60 + 45 * len(durations)   # 슬라이드당 인코딩+오버헤드 여유
    r = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-File", script,
         "-PptxPath", pptx, "-DurationsJson", durations_path, "-OutDir", out_dir,
         "-EffectDuration", str(effect_duration), "-Stagger", str(stagger),
         "-Width", str(width), "-Fps", str(fps)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=timeout)
    if r.returncode != 0:
        raise RuntimeError(f"슬라이드 애니메이션 렌더 실패: {r.stderr.strip() or r.stdout.strip()}")

    out = {}
    for line in r.stdout.splitlines():
        parts = line.split(maxsplit=2)
        if len(parts) == 3 and parts[0] == "RENDERED":
            out[int(parts[1])] = parts[2].strip()
        elif parts and parts[0] == "FAILED":
            print(f"[deck] 슬라이드 {parts[1]} 애니메이션 렌더 실패(재시도 포함) — "
                  f"정지 이미지로 대체됩니다.")
        elif parts and parts[0] == "RETRY":
            print(f"[deck] {' '.join(parts[1:])} — 재시도 중")
    return out


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    backend = "auto"
    pos = []
    i = 0
    while i < len(argv):
        if argv[i] == "--reuse-pngs":
            backend = "reuse"
            i += 1
        elif argv[i] == "--backend" and i + 1 < len(argv):
            backend = argv[i + 1]
            i += 2
        else:
            pos.append(argv[i])
            i += 1
    if not pos:
        print("usage: python -m revoice.deck <파일.pptx> [출력디렉터리] [--backend auto|powerpoint|libreoffice|reuse] [--reuse-pngs]")
        raise SystemExit(1)
    r = ingest(pos[0], pos[1] if len(pos) > 1 else "deck", backend=backend)
    print(f"[deck] 슬라이드 {r['slides']}장 · PNG {r['pngs']}장 → {r['outline']}")


if __name__ == "__main__":
    main()
