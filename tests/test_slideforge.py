import sys, os, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from slideforge import parse_md, build

MD = """# 封面标题
副标题一行

## 内容页
- 要点一
- 要点二
  - 二级要点

> 一句金句

# 第二部分

## 步骤页
1. 第一步
2. 第二步
"""

def test_parse():
    slides = parse_md(MD)
    kinds = [s["kind"] for s in slides]
    assert kinds == ["cover", "content", "quote", "section", "content"], kinds
    assert slides[1]["items"][2] == (1, "二级要点")

def test_build_roundtrip():
    from pptx import Presentation
    slides = parse_md(MD)
    with tempfile.TemporaryDirectory() as td:
        out = build(slides, os.path.join(td, "deck.pptx"), theme="light", brand="@t")
        prs = Presentation(out)
        assert len(prs.slides.__iter__.__self__._sldIdLst) == 5
        assert abs(prs.slide_width.inches - 13.333) < 0.01

if __name__ == "__main__":
    test_parse(); test_build_roundtrip(); print("ok")
