#!/usr/bin/env python3
"""生成 lhg-ppt 冒烟测试用的 pptx fixtures（仅标准库）。

bad.pptx：埋点齐全的"问题 PPT"——小字号、版权字体、低对比度、
对不齐、单页字多、超 30 页、无备注。
good.pptx：干净的"好 PPT"——规范字号、免费字体、对齐、有备注、页数适中。

注意：生成的是最小可用 OOXML（仅保证本插件审计脚本可解析），
不是 PowerPoint 完全合规文件。重新生成：python3 make_fixtures.py
"""

import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
IN = 914400  # 1 英寸 = 914400 EMU

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _run_xml(text, sz_pt=None, font=None, color=None):
    attrs = []
    inner = ""
    if sz_pt is not None:
        attrs.append('sz="%d"' % int(sz_pt * 100))
    if color:
        inner += '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % color
    if font:
        inner += '<a:latin typeface="%s"/><a:ea typeface="%s"/>' % (font, font)
    rpr = "<a:rPr %s>%s</a:rPr>" % (" ".join(attrs), inner)
    # XML 转义
    t = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return "<a:r>%s<a:t>%s</a:t></a:r>" % (rpr, t)


def _shape_xml(sid, x, y, w, h, runs):
    runs_xml = "".join(_run_xml(**r) for r in runs)
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="shape%d"/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/>'
        '<a:ext cx="%d" cy="%d"/></a:xfrm></p:spPr>'
        '<p:txBody><a:bodyPr/><a:p>%s</a:p></p:txBody></p:sp>'
    ) % (sid, sid, x, y, w, h, runs_xml)


def _slide_xml(shapes, bg="FFFFFF"):
    sp_xml = "".join(
        _shape_xml(i + 2, **sh) for i, sh in enumerate(shapes))
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<p:sld xmlns:p="%s" xmlns:a="%s" xmlns:r="%s">'
        '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="%s"/>'
        "</a:solidFill></p:bgPr></p:bg>"
        "<p:spTree>%s</p:spTree></p:cSld></p:sld>"
    ) % (P, A, R, bg, sp_xml)


def _slide_rels_xml(with_notes, notes_name):
    rels = (
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
        'Target="../slideLayouts/slideLayout1.xml"/>'
    )
    if with_notes:
        rels += (
            '<Relationship Id="rId2" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" '
            'Target="../notesSlides/%s"/>' % notes_name
        )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        "%s</Relationships>" % rels
    )


def _notes_xml(text):
    t = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<p:notes xmlns:p="%s" xmlns:a="%s"><p:cSld><p:spTree>'
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="notes"/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/>'
        "</a:xfrm></p:spPr>"
        '<p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1200"/>'
        "<a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>"
        "</p:spTree></p:cSld></p:notes>" % (P, A, t)
    )


_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
{slides}{notes}
</Types>"""

_RELS = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"""

_MASTER = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<p:sldMaster xmlns:p="%s" xmlns:a="%s"><p:cSld><p:spTree>'
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>'
    "<p:grpSpPr/></p:spTree></p:cSld></p:sldMaster>" % (P, A)
)
_LAYOUT = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<p:sldLayout xmlns:p="%s" xmlns:a="%s"><p:cSld><p:spTree>'
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>'
    "<p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>" % (P, A)
)


def build(path, slides):
    """slides: list of dict(bg, shapes, notes)."""
    n = len(slides)
    slides_ct = "".join(
        '<Override PartName="/ppt/slides/slide%d.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        % (i + 1) for i in range(n))
    notes_ct = "".join(
        '<Override PartName="/ppt/notesSlides/notesSlide%d.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'
        % (i + 1) for i, s in enumerate(slides) if s.get("notes"))
    ct = _CONTENT_TYPES.format(slides=slides_ct, notes=notes_ct)

    sld_ids = "".join(
        '<p:sldId id="%d" r:id="rId%d"/>' % (256 + i, i + 2) for i in range(n))
    pres_rels = (
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" '
        'Target="slideMasters/slideMaster1.xml"/>'
        + "".join(
            '<Relationship Id="rId%d" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
            'Target="slides/slide%d.xml"/>' % (i + 2, i + 1) for i in range(n)))
    pres = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<p:presentation xmlns:p="%s" xmlns:r="%s">'
        '<p:sldMasterIdLst><p:sldMasterId r:id="rId1"/></p:sldMasterIdLst>'
        "<p:sldIdLst>%s</p:sldIdLst>"
        '<p:sldSz cx="12192000" cy="6858000"/>'
        "</p:presentation>" % (P, R, sld_ids))
    pres_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        "%s</Relationships>" % pres_rels)

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", _RELS)
        z.writestr("ppt/presentation.xml", pres)
        z.writestr("ppt/_rels/presentation.xml.rels", pres_rels_xml)
        z.writestr("ppt/slideMasters/slideMaster1.xml", _MASTER)
        z.writestr("ppt/slideLayouts/slideLayout1.xml", _LAYOUT)
        for i, s in enumerate(slides):
            z.writestr("ppt/slides/slide%d.xml" % (i + 1),
                       _slide_xml(s["shapes"], s.get("bg", "FFFFFF")))
            notes_name = "notesSlide%d.xml" % (i + 1)
            z.writestr("ppt/slides/_rels/slide%d.rels" % (i + 1),
                       _slide_rels_xml(bool(s.get("notes")), notes_name))
            if s.get("notes"):
                z.writestr("ppt/notesSlides/" + notes_name,
                           _notes_xml(s["notes"]))


def _shape(x_in, runs, y_in=1.0, w_in=11.0, h_in=1.0):
    return {"x": int(x_in * IN), "y": int(y_in * IN),
            "w": int(w_in * IN), "h": int(h_in * IN), "runs": runs}


def _run(text, sz_pt=20, font="思源黑体", color="333333"):
    return {"text": text, "sz_pt": sz_pt, "font": font, "color": color}


def make_bad():
    long_text = "这是一段故意写得很长的文字。" * 20  # 240 字，触发 PPT-TXT-001
    slides = [
        # 第1页：小字号 + 版权字体 + 低对比度
        {"shapes": [
            _shape(0.7, [_run("小到看不清的标题", sz_pt=8, font="微软雅黑")]),
            _shape(0.7, [_run("浅灰文字", sz_pt=10, font="微软雅黑", color="CCCCCC")],
                   y_in=2.5),
        ]},
        # 第2页：单页文字过多 + 对不齐（3 个文本框左边缘差 >1 英寸）
        {"shapes": [
            _shape(0.7, [_run(long_text, sz_pt=16)], y_in=1.0, h_in=3.0),
            _shape(2.0, [_run("第二栏", sz_pt=20)], y_in=4.5),
            _shape(3.5, [_run("第三栏", sz_pt=20)], y_in=5.5),
        ]},
        # 第3页：偏小字号（14pt，触发 FONT-002）
        {"shapes": [
            _shape(0.7, [_run("偏小的正文", sz_pt=14)], y_in=2.0),
        ]},
    ]
    # 再加 29 页普通页，凑满 32 页触发 PPT-PG-001
    for i in range(29):
        slides.append({"shapes": [
            _shape(0.7, [_run("第 %d 页内容" % (i + 4), sz_pt=20)], y_in=2.0),
        ]})
    build(os.path.join(HERE, "bad.pptx"), slides)


def make_good():
    slides = []
    titles = ["封面：Q3 用户增长复盘", "核心结论：获客成本下降 40%",
              "渠道分析：短视频起量", "留存提升：8 个百分点",
              "下季度规划", "谢谢"]
    for t in titles:
        slides.append({
            "shapes": [
                _shape(0.7, [_run(t, sz_pt=30)], y_in=0.8),
                _shape(0.7, [_run("要点一：数据支撑结论", sz_pt=20),
                              _run("要点二：方法可复用", sz_pt=20)], y_in=2.5, h_in=2.0),
            ],
            "notes": "讲这页时的 3 句话：先说结论，再说数据，最后说行动。",
        })
    build(os.path.join(HERE, "good.pptx"), slides)


if __name__ == "__main__":
    make_bad()
    make_good()
    print("fixtures 已生成：bad.pptx / good.pptx")
