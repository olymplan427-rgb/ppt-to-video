# PPTX 직접 제작·수정 가이드 (`revoice.ppt_edit`)

영상을 만들기 전에 PPT 발표자료 자체의 문구를 수정하거나, 신규 슬라이드를 추가하거나,
브랜드 컬러 테마를 바꿔야 할 때 **원본 `.pptx` 자체를 직접 편집**하는 절차다.
(슬라이드를 HTML로 재작성하지 않는다.)

## 1. 먼저 덱 구조와 브랜드 규격 파악 (`inspect`)

편집 전에 대상 PPTX의 크기·폰트·컬러 팔레트·도형 좌표를 확인한다:

```bash
python -m revoice.ppt_edit inspect slides.pptx --slides 1,3,10
```

## 2. 핵심 편집 도구 (`revoice/ppt_edit.py`)

### ① 서식 보존 문구 치환 (`replace_text_preserve_format`)
`text_frame.text = ...`로 덮어쓰면 기존 폰트·크기·색상이 전부 사라진다.
반드시 run 단위로 치환하는 헬퍼를 사용한다(여러 run에 걸쳐 쪼개진 문자열도 처리):

```bash
python -m revoice.ppt_edit replace in.pptx out.pptx --find "10월" --replace "11월" --slides 1,3 --notes
```

### ② 기존 슬라이드 복제 후 신규 슬라이드 제작 (`duplicate_slide`, `move_slide`)
빈 슬라이드에서 처음부터 그리는 것보다 **기존 브랜드 슬라이드를 복제**해 헤더·로고·배경·폰트를
그대로 물려받은 뒤 텍스트와 이미지만 교체하는 편이 브랜드 일관성이 훨씬 높다.

```python
from pptx import Presentation
from revoice import ppt_edit

prs = Presentation("deck_v1.pptx")
new_sl = ppt_edit.duplicate_slide(prs, 44)  # 44번 슬라이드 스타일 복제
ppt_edit.pdf_page_to_png("brochure.pdf", "cover.png", page=1, dpi=220)
ppt_edit.add_picture_fit(new_sl, "cover.png", (0.8, 1.8, 4.5, 5.0))
prs.save("deck_v2.pptx")
```

### ③ 컬러 테마 일괄 치환 (`recolor_theme_hex`, `recolor_png`)
- 슬라이드 XML의 `srgbClr` hex 값을 일괄 교체한다:
  ```bash
  python -m revoice.ppt_edit recolor in.pptx out.pptx --map 2FA599=E3A33B,123C38=26190F
  ```
- 슬라이드에 삽입된 단색 아이콘 PNG는 `recolor_png(in_png, out_png, (R, G, B))`로 알파 채널을 유지한 채 재색상한다.

### ④ 파일 잠금 방지
사용자가 PowerPoint로 원본 파일을 열어 두고 있을 수 있다(`PermissionError`).
항상 새 버전 파일명(예: `presentation_261215.pptx`)으로 저장한다.

## 3. 수정 후 시각 검수 루프 (필수)

`python-pptx` 편집만 하고 바로 영상 렌더로 넘어가지 않는다 — 긴 제목이 우측 로고와 겹치거나
텍스트 박스를 넘치는 문제가 잦다.

```bash
python -m revoice.ppt_edit preview deck_v2.pptx _preview --slides 3,45
```

생성된 `_preview/slide_NN.png`를 직접 열어 아래 4가지를 확인한다:
1. 제목이 우측 상단 로고·배지와 겹치지 않는가? (겹치면 폰트 크기 축소 예: 40pt → 35pt)
2. 문장이 의도치 않은 곳에서 줄바꿈되거나 카드 영역을 넘치지 않는가?
3. 하단 자막 영역(화면 아래 ~15%)에 핵심 숫자·캡션이 가려지지 않는가?
4. 삽입 이미지가 비율 왜곡이나 저해상도 깨짐 없이 선명한가?
