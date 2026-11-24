# -*- coding: utf-8 -*-
"""revoice.ppt_edit 단위 테스트 — 서식 보존 치환, 슬라이드 복제/이동, 컬러 치환, 이미지 배치."""
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt


def _sample_pptx(path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    tb = sl.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(2))
    p = tb.text_frame.paragraphs[0]
    r1 = p.add_run()
    r1.text = "10월 "
    r1.font.name = "Pretendard ExtraBold"
    r1.font.size = Pt(40)
    r1.font.color.rgb = RGBColor(0xE3, 0xA3, 0x3B)
    r2 = p.add_run()
    r2.text = "창업 전략 설명회"
    r2.font.name = "Pretendard"
    r2.font.size = Pt(32)
    r2.font.color.rgb = RGBColor(0x26, 0x19, 0x0F)
    sl.notes_slide.notes_text_frame.text = "10월 설명회 발표자 노트"
    prs.save(path)
    return path


def test_replace_text_preserves_run_formatting_and_cross_run(tmp_path):
    from revoice import ppt_edit
    src = _sample_pptx(str(tmp_path / "in.pptx"))
    dst = str(tmp_path / "out.pptx")

    # run 경계에 걸친 "10월 창업" 치환 + 노트 포함
    res = ppt_edit.replace_text_preserve_format(
        src, {"10월 창업": "11월 창업", "10월 설명회": "11월 설명회"},
        dst, include_notes=True)
    assert res["total"] == 2

    prs = Presentation(dst)
    sl = prs.slides[0]
    p = sl.shapes[0].text_frame.paragraphs[0]
    assert p.text == "11월 창업 전략 설명회"
    assert p.runs[0].font.name == "Pretendard ExtraBold"
    assert p.runs[0].font.size == Pt(40)
    assert str(p.runs[0].font.color.rgb) == "E3A33B"
    assert "11월 설명회" in sl.notes_slide.notes_text_frame.text


def test_duplicate_and_move_slide(tmp_path):
    from revoice import ppt_edit
    src = _sample_pptx(str(tmp_path / "in.pptx"))
    prs = Presentation(src)
    dup = ppt_edit.duplicate_slide(prs, 1)
    assert len(prs.slides) == 2
    assert dup.shapes[0].text_frame.text == "10월 창업 전략 설명회"

    dup.shapes[0].text_frame.paragraphs[0].runs[0].text = "12월 "
    ppt_edit.move_slide(prs, 2, 1)
    assert prs.slides[0].shapes[0].text_frame.text.startswith("12월")
    assert prs.slides[1].shapes[0].text_frame.text.startswith("10월")


def test_recolor_theme_hex(tmp_path):
    from revoice import ppt_edit
    src = _sample_pptx(str(tmp_path / "in.pptx"))
    dst = str(tmp_path / "recolored.pptx")
    res = ppt_edit.recolor_theme_hex(src, {"#E3A33B": "2FA599"}, dst)
    assert res["replaced"] == 1
    prs = Presentation(dst)
    assert str(prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.color.rgb) == "2FA599"


def test_add_picture_fit_preserves_aspect_ratio(tmp_path):
    from PIL import Image
    from revoice import ppt_edit
    img_path = str(tmp_path / "wide.png")
    Image.new("RGB", (400, 200), (200, 100, 50)).save(img_path)

    prs = Presentation()
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    pic = ppt_edit.add_picture_fit(sl, img_path, (1.0, 1.0, 4.0, 4.0))
    assert abs(pic.width.inches - 4.0) < 0.02
    assert abs(pic.height.inches - 2.0) < 0.02
    assert abs(pic.top.inches - 2.0) < 0.02


def test_inspect_pptx_returns_summary(tmp_path):
    from revoice import ppt_edit
    src = _sample_pptx(str(tmp_path / "in.pptx"))
    info = ppt_edit.inspect_pptx(src, slides="1")
    assert info["slide_count"] == 1
    assert info["width_in"] == 13.33
    assert any(f == "Pretendard ExtraBold" for f, _ in info["top_fonts"])
