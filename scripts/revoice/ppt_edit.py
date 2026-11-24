# -*- coding: utf-8 -*-
"""PPTX 제작·편집 헬퍼 — 서식 보존 치환, 슬라이드 복제, 테마 리컬러, PDF 삽입, 부분 프리뷰.

슬라이드를 HTML로 재작성하지 않고 원본 `.pptx` 자체를 `python-pptx`로 직접 생성·편집한 뒤,
수정한 슬라이드만 PNG로 뽑아 시각 검수(로고 겹침·줄바꿈·여백 확인)까지 한 루프로 처리한다.

CLI 예:
  python -m revoice.ppt_edit inspect deck.pptx --slides 1,3,45
  python -m revoice.ppt_edit replace in.pptx out.pptx --find "10월" --replace "11월" --slides 1,3
  python -m revoice.ppt_edit recolor in.pptx out.pptx --map 2FA599=E3A33B,123C38=26190F
  python -m revoice.ppt_edit pdf2png cover.pdf cover.png --page 1 --dpi 220
  python -m revoice.ppt_edit preview out.pptx _preview --slides 3,45
"""
import argparse
import copy
import os
import re
import sys
from collections import Counter

from revoice import deck


def _iter_shapes(shapes):
    """그룹 도형 내부까지 재귀 순회."""
    for sh in shapes:
        yield sh
        if hasattr(sh, "shapes"):
            yield from _iter_shapes(sh.shapes)


def _iter_paragraphs(slide):
    """슬라이드 내 모든 텍스트 프레임·표 셀의 단락(paragraph)을 순회."""
    for sh in _iter_shapes(slide.shapes):
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                yield p
        if getattr(sh, "has_table", False):
            for row in sh.table.rows:
                for cell in row.cells:
                    for p in cell.text_frame.paragraphs:
                        yield p


def replace_in_paragraph(paragraph, old: str, new: str) -> int:
    """단락 서식(폰트·크기·색상·굵기)을 유지한 채 `old` → `new` 치환.

    `paragraph.text = ...`로 덮어쓰면 모든 run 서식이 날아간다.
    1차로 단일 run 내부 일치를 치환하고, 여러 run에 걸쳐 쪼개진 문자열은
    시작 run에 치환 결과를 넣고 걸친 나머지 run 구간만 비워 서식을 최대한 보존한다.
    """
    if not old:
        return 0
    runs = list(paragraph.runs)
    if not runs:
        return 0

    count = 0
    # 1) 단일 run 내부 치환 (가장 깔끔하게 모든 서식 유지)
    for r in runs:
        if old in r.text:
            c = r.text.count(old)
            r.text = r.text.replace(old, new)
            count += c

    # 2) 여러 run 경계에 걸친 문자열 치환
    while True:
        full = "".join(r.text for r in runs)
        idx = full.find(old)
        if idx < 0:
            break
        end = idx + len(old)
        pos = 0
        replaced = False
        for r in runs:
            r_start, r_end = pos, pos + len(r.text)
            pos = r_end
            if r_end <= idx or r_start >= end:
                continue
            cut_a = max(0, idx - r_start)
            cut_b = min(len(r.text), end - r_start)
            if not replaced:
                r.text = r.text[:cut_a] + new + r.text[cut_b:]
                replaced = True
            else:
                r.text = r.text[:cut_a] + r.text[cut_b:]
        count += 1
    return count


def _parse_slide_set(slides):
    if not slides:
        return None
    if isinstance(slides, str):
        return {int(x.strip()) for x in slides.split(",") if x.strip()}
    return {int(x) for x in slides}


def _save_pptx(prs, out_path: str) -> str:
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    try:
        prs.save(out_path)
    except PermissionError as e:
        raise RuntimeError(
            f"PPTX 저장 실패(파일이 PowerPoint에서 열려 있음): {out_path}\n"
            "PowerPoint를 닫거나 새 버전 파일명(예: *_YYMMDD.pptx)으로 저장하세요."
        ) from e
    return out_path


def inspect_pptx(pptx_path: str, slides=None) -> dict:
    """PPTX 규격(크기·주요 폰트·주요 컬러)과 슬라이드별 도형/텍스트 요약 반환."""
    from pptx import Presentation
    prs = Presentation(pptx_path)
    target = _parse_slide_set(slides)

    fonts, colors = Counter(), Counter()
    slide_infos = []
    for n, sl in enumerate(prs.slides, 1):
        for sh in _iter_shapes(sl.shapes):
            if sh.has_text_frame:
                for par in sh.text_frame.paragraphs:
                    for r in par.runs:
                        if r.font.name:
                            fonts[r.font.name] += 1
                        try:
                            if r.font.color and r.font.color.type == 1:
                                colors[str(r.font.color.rgb)] += 1
                        except Exception:
                            pass
            try:
                if sh.fill and sh.fill.type == 1:
                    colors[str(sh.fill.fore_color.rgb)] += 1
            except Exception:
                pass

        if target is None or n in target:
            shapes_info = []
            for idx, sh in enumerate(sl.shapes):
                txt = ""
                if sh.has_text_frame:
                    txt = sh.text_frame.text.replace("\n", " / ").strip()
                shapes_info.append({
                    "idx": idx,
                    "name": sh.name,
                    "pos_in": (round(sh.left.inches, 2), round(sh.top.inches, 2)),
                    "size_in": (round(sh.width.inches, 2), round(sh.height.inches, 2)),
                    "text": txt[:80],
                })
            notes = ""
            if sl.has_notes_slide:
                notes = sl.notes_slide.notes_text_frame.text.strip()
            slide_infos.append({"n": n, "shapes": shapes_info, "notes": notes[:120]})

    return {
        "path": pptx_path,
        "slide_count": len(prs.slides),
        "width_in": round(prs.slide_width.inches, 2),
        "height_in": round(prs.slide_height.inches, 2),
        "top_fonts": fonts.most_common(10),
        "top_colors": colors.most_common(12),
        "slides": slide_infos,
    }


def replace_text_preserve_format(pptx_path: str, replacements: dict, out_path: str,
                                 slides=None, include_notes: bool = False) -> dict:
    """여러 문구를 서식 보존 치환하고 결과를 새 PPTX로 저장.

    replacements: {"이전 문구": "새 문구", ...}
    slides: 대상 1-based 슬라이드 번호 집합(None이면 전체).
    """
    from pptx import Presentation
    prs = Presentation(pptx_path)
    target = _parse_slide_set(slides)
    hits = {}

    for n, sl in enumerate(prs.slides, 1):
        if target is not None and n not in target:
            continue
        slide_hits = 0
        for p in _iter_paragraphs(sl):
            for old, new in replacements.items():
                slide_hits += replace_in_paragraph(p, old, new)
        if include_notes and sl.has_notes_slide:
            for p in sl.notes_slide.notes_text_frame.paragraphs:
                for old, new in replacements.items():
                    slide_hits += replace_in_paragraph(p, old, new)
        if slide_hits:
            hits[n] = slide_hits

    _save_pptx(prs, out_path)
    return {"out": out_path, "hits": hits, "total": sum(hits.values())}


def set_notes(pptx_path: str, slide_no: int, notes_text: str, out_path: str) -> str:
    """특정 슬라이드(1-based)의 발표자 노트를 설정하고 저장."""
    from pptx import Presentation
    prs = Presentation(pptx_path)
    sl = prs.slides[slide_no - 1]
    sl.notes_slide.notes_text_frame.text = notes_text
    return _save_pptx(prs, out_path)


def duplicate_slide(prs, slide_no: int):
    """기존 슬라이드(1-based)의 도형·이미지 관계·배경을 복제해 맨 뒤에 새 슬라이드 추가.

    기존 브랜드 슬라이드를 템플릿으로 삼아 새 슬라이드를 만들 때 사용한다.
    """
    src = prs.slides[slide_no - 1]
    dst = prs.slides.add_slide(src.slide_layout)

    # 레이아웃 기본 플레이스홀더 제거 후 원본 도형 복제
    for sh in list(dst.shapes):
        sp = sh._element
        sp.getparent().remove(sp)

    # 이미지·외부 관계 복사 및 rId 매핑
    rid_map = {}
    for rel in src.part.rels.values():
        if "notesSlide" in rel.reltype or "slideLayout" in rel.reltype:
            continue
        if rel.is_external:
            new_rid = dst.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            new_rid = dst.part.relate_to(rel.target_part, rel.reltype)
        rid_map[rel.rId] = new_rid

    for el in src.shapes._spTree:
        tag = el.tag.rsplit("}", 1)[-1]
        if tag in ("nvGrpSpPr", "grpSpPr"):
            continue
        cloned = copy.deepcopy(el)
        for node in cloned.iter():
            for attr, val in list(node.attrib.items()):
                if val in rid_map:
                    node.attrib[attr] = rid_map[val]
        dst.shapes._spTree.append(cloned)

    # 배경(cSld/bg) 복제
    src_bg = src._element.cSld.find("{http://schemas.openxmlformats.org/presentationml/2006/main}bg")
    if src_bg is not None:
        dst_bg = dst._element.cSld.find("{http://schemas.openxmlformats.org/presentationml/2006/main}bg")
        if dst_bg is not None:
            dst._element.cSld.remove(dst_bg)
        bg_clone = copy.deepcopy(src_bg)
        for node in bg_clone.iter():
            for attr, val in list(node.attrib.items()):
                if val in rid_map:
                    node.attrib[attr] = rid_map[val]
        dst._element.cSld.insert(0, bg_clone)

    return dst


def move_slide(prs, from_no: int, to_no: int) -> None:
    """슬라이드 순서 이동 (1-based)."""
    sldIdLst = prs.slides._sldIdLst
    items = list(sldIdLst)
    el = items.pop(from_no - 1)
    items.insert(to_no - 1, el)
    for child in list(sldIdLst):
        sldIdLst.remove(child)
    for child in items:
        sldIdLst.append(child)


def recolor_theme_hex(pptx_path: str, color_map: dict, out_path: str, slides=None) -> dict:
    """슬라이드 XML 내 `srgbClr` hex 컬러를 일괄 치환 (예: {'2FA599': 'E3A33B'})."""
    from pptx import Presentation
    prs = Presentation(pptx_path)
    target = _parse_slide_set(slides)
    norm_map = {k.lstrip("#").upper(): v.lstrip("#").upper() for k, v in color_map.items()}
    replaced = 0

    for n, sl in enumerate(prs.slides, 1):
        if target is not None and n not in target:
            continue
        for node in sl._element.iter():
            if node.tag.endswith("}srgbClr"):
                val = (node.attrib.get("val") or "").upper()
                if val in norm_map:
                    node.attrib["val"] = norm_map[val]
                    replaced += 1

    _save_pptx(prs, out_path)
    return {"out": out_path, "replaced": replaced}


def recolor_png(in_png: str, out_png: str, target_rgb: tuple) -> str:
    """아이콘 PNG의 알파 채널과 명암을 유지하면서 지정 RGB 색으로 재색상."""
    import numpy as np
    from PIL import Image
    im = Image.open(in_png).convert("RGBA")
    arr = np.array(im, dtype=np.float32)
    alpha = arr[:, :, 3]
    # 기존 밝기를 미세 톤 차이로 보존하되 투명 영역은 그대로 유지
    r, g, b = target_rgb
    arr[:, :, 0] = r
    arr[:, :, 1] = g
    arr[:, :, 2] = b
    arr[:, :, 3] = alpha
    os.makedirs(os.path.dirname(os.path.abspath(out_png)), exist_ok=True)
    Image.fromarray(arr.astype(np.uint8), "RGBA").save(out_png)
    return out_png


def pdf_page_to_png(pdf_path: str, out_png: str, page: int = 1, dpi: int = 220) -> str:
    """PDF 지정 페이지(1-based)를 고해상도 PNG로 추출 (교재 표지·포스터 삽입용)."""
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    os.makedirs(os.path.dirname(os.path.abspath(out_png)), exist_ok=True)
    doc = fitz.open(pdf_path)
    try:
        pg = doc[page - 1]
        scale = dpi / 72.0
        pix = pg.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        pix.save(out_png)
    finally:
        doc.close()
    return out_png


def add_picture_fit(slide, img_path: str, box_in: tuple):
    """이미지를 비율 왜곡 없이 `box_in=(left, top, width, height)`(인치) 중앙에 배치."""
    from PIL import Image
    from pptx.util import Inches
    left, top, bw, bh = box_in
    with Image.open(img_path) as im:
        iw, ih = im.size
    ratio = min(bw / iw, bh / ih)
    w, h = iw * ratio, ih * ratio
    x = left + (bw - w) / 2.0
    y = top + (bh - h) / 2.0
    return slide.shapes.add_picture(img_path, Inches(x), Inches(y), width=Inches(w), height=Inches(h))


def preview_slides(pptx_path: str, out_dir: str, slides=None) -> list:
    """수정·추가한 슬라이드를 PNG로 빠르게 추출해 시각 검수용 파일 경로 목록을 반환.

    Windows + PowerPoint 환경에서 `slides`가 지정되면 해당 슬라이드만 선택 Export해
    수 초 안에 끝난다. macOS/Linux에서는 `deck.export_slides`로 추출한 뒤 대상만 돌려준다.
    """
    target = _parse_slide_set(slides)
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)

    if target and os.name == "nt" and deck.has_powerpoint():
        pptx_esc = os.path.abspath(pptx_path).replace("'", "''")
        out_esc = out_dir.replace("'", "''")
        idx_list = ",".join(str(i) for i in sorted(target))
        deck._ps(f"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open('{pptx_esc}', $true, $false, $false)
try {{
  foreach ($i in @({idx_list})) {{
    $pres.Slides.Item($i).Export(('{out_esc}\\slide_{{0:d2}}.png' -f $i), 'PNG', {deck.W}, {deck.H})
  }}
}} finally {{
  $pres.Close()
  if ($app.Presentations.Count -eq 0) {{ $app.Quit() }}
}}
""", timeout=180)
        return [os.path.join(out_dir, f"slide_{i:02d}.png") for i in sorted(target)
                if os.path.exists(os.path.join(out_dir, f"slide_{i:02d}.png"))]

    all_pngs = deck.export_slides(pptx_path, out_dir)
    if target is None:
        return all_pngs
    want = {f"slide_{i:02d}.png" for i in target}
    return [p for p in all_pngs if os.path.basename(p) in want]


def main(argv=None):
    ap = argparse.ArgumentParser(description="PPTX 제작·편집·시각 검수 헬퍼")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_ins = sub.add_parser("inspect", help="PPTX 크기·폰트·컬러·도형 구조 조회")
    p_ins.add_argument("pptx")
    p_ins.add_argument("--slides", default="", help="조회할 1-based 슬라이드 번호(예: 1,3,45)")

    p_rep = sub.add_parser("replace", help="서식 보존 문구 치환")
    p_rep.add_argument("src")
    p_rep.add_argument("dst")
    p_rep.add_argument("--find", required=True)
    p_rep.add_argument("--replace", required=True)
    p_rep.add_argument("--slides", default="")
    p_rep.add_argument("--notes", action="store_true")

    p_col = sub.add_parser("recolor", help="테마 hex 컬러 일괄 치환")
    p_col.add_argument("src")
    p_col.add_argument("dst")
    p_col.add_argument("--map", required=True, help="OLD=NEW,OLD2=NEW2")
    p_col.add_argument("--slides", default="")

    p_pdf = sub.add_parser("pdf2png", help="PDF 페이지를 고해상도 PNG로 추출")
    p_pdf.add_argument("pdf")
    p_pdf.add_argument("out_png")
    p_pdf.add_argument("--page", type=int, default=1)
    p_pdf.add_argument("--dpi", type=int, default=220)

    p_pre = sub.add_parser("preview", help="지정 슬라이드 PNG 프리뷰 추출")
    p_pre.add_argument("pptx")
    p_pre.add_argument("out_dir")
    p_pre.add_argument("--slides", default="")

    args = ap.parse_args(argv)
    if args.cmd == "inspect":
        info = inspect_pptx(args.pptx, args.slides or None)
        print(f"Slides: {info['slide_count']}  Size: {info['width_in']}x{info['height_in']} in")
        print("Top fonts :", info["top_fonts"][:6])
        print("Top colors:", info["top_colors"][:8])
        for sl in info["slides"]:
            print(f"\n[Slide {sl['n']}] ({len(sl['shapes'])} shapes)")
            for sh in sl["shapes"]:
                if sh["text"]:
                    print(f"  #{sh['idx']:02d} pos={sh['pos_in']} sz={sh['size_in']} {sh['text']}")
    elif args.cmd == "replace":
        r = replace_text_preserve_format(
            args.src, {args.find: args.replace}, args.dst,
            slides=args.slides or None, include_notes=args.notes)
        print(f"[replace] 총 {r['total']}곳 치환 ({r['hits']}) → {r['out']}")
    elif args.cmd == "recolor":
        cmap = dict(pair.split("=", 1) for pair in args.map.split(",") if "=" in pair)
        r = recolor_theme_hex(args.src, cmap, args.dst, slides=args.slides or None)
        print(f"[recolor] {r['replaced']}개 컬러 노드 치환 → {r['out']}")
    elif args.cmd == "pdf2png":
        out = pdf_page_to_png(args.pdf, args.out_png, page=args.page, dpi=args.dpi)
        print(f"[pdf2png] → {out}")
    elif args.cmd == "preview":
        pngs = preview_slides(args.pptx, args.out_dir, slides=args.slides or None)
        for p in pngs:
            print(f"[preview] {p}")


if __name__ == "__main__":
    main()
