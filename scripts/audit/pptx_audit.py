#!/usr/bin/env python3
"""lhg-ppt 确定性 PPT 设计审计：9 条规则，零第三方依赖。

用 zipfile + XML 直接解析 .pptx（不用装 python-pptx 也能跑），
输出 JSON 与人类可读报告；有 blocker 时以 exit code 2 退出（供 hooks 与 CI 使用）。

用法：
    python3 pptx_audit.py --target <pptx文件或目录> [--format json|text] [--fail-on blocker|warning|nit]

规则文案均为本项目原创（方法论思路受 ppt-master 等开源项目启发，
未复制任何第三方原文案与代码）。
"""

import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

VERSION = "1.0.0"
TOOL = "lhg-ppt/pptx-audit"

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

EMU_PER_INCH = 914400
# 疑似需商用授权的字体（子串匹配，不区分大小写）
RISKY_FONTS = ["微软雅黑", "microsoft yahei", "方正", "founder",
               "汉仪", "hanyi", "造字工房", "文鼎", "wending"]
CJK_RE = re.compile(r"[\u4e00-\u9fff]")
WORD_RE = re.compile(r"[A-Za-z0-9]+")


class Finding:
    __slots__ = ("rule_id", "severity", "category", "title", "file", "line", "snippet", "detail")

    def __init__(self, rule_id, severity, category, title, file, line, snippet, detail):
        self.rule_id = rule_id
        self.severity = severity
        self.category = category
        self.title = title
        self.file = file
        self.line = line  # 这里是"第几页"
        self.snippet = snippet
        self.detail = detail

    def to_dict(self):
        return {
            "rule_id": self.rule_id, "severity": self.severity,
            "category": self.category, "title": self.title,
            "file": self.file, "line": self.line,
            "snippet": self.snippet, "detail": self.detail,
        }


RULES = []
_RULE_BY_ID = {}


def rule(rule_id, severity, category, title, detail):
    def deco(fn):
        r = {"id": rule_id, "severity": severity,
             "category": category, "title": title, "detail": detail}
        _RULE_BY_ID[rule_id] = r
        r["check"] = fn
        RULES.append(r)
        return fn
    return deco


def _finding(rule_id, path, page, snippet):
    r = _RULE_BY_ID[rule_id]
    return Finding(r["id"], r["severity"], r["category"], r["title"],
                   path, page, (snippet or "")[:160], r["detail"])


# ---------------- pptx 解析 ----------------

class Slide:
    def __init__(self, index, part_name):
        self.index = index          # 第几页（1-based）
        self.part_name = part_name
        self.bg = "FFFFFF"
        self.runs = []              # [{text, sz(100ths pt)|None, font|None, color|None}]
        self.shape_xs = []          # 文本框左边缘 x（EMU）
        self.has_notes = False

    @property
    def char_count(self):
        n = 0
        for r in self.runs:
            n += len(CJK_RE.findall(r["text"]))
            n += len(WORD_RE.findall(r["text"]))
        return n


def _lum(hex6):
    h = (hex6 or "FFFFFF")[-6:]
    try:
        rgb = [int(h[i:i + 2], 16) / 255.0 for i in (0, 2, 4)]
    except ValueError:
        return 1.0
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * f(rgb[0]) + 0.7152 * f(rgb[1]) + 0.0722 * f(rgb[2])


def _contrast(c1, c2):
    l1, l2 = _lum(c1), _lum(c2)
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


def parse_pptx(path):
    """返回 Slide 列表；解析失败抛异常（调用方捕获后记为文件级错误）。"""
    slides = []
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        # 从 [Content_Types] 找 slides（按编号排序）
        try:
            ct_root = ET.fromstring(z.read("[Content_Types].xml"))
        except KeyError:
            raise ValueError("不是有效的 pptx：缺 [Content_Types].xml")
        slide_parts = []
        for ov in ct_root.findall("ct:Override", NS):
            pn, ctype = ov.get("PartName", ""), ov.get("ContentType", "")
            if re.fullmatch(r"/ppt/slides/slide\d+\.xml", pn) and "slide" in ctype:
                slide_parts.append(pn.lstrip("/"))
        slide_parts.sort(key=lambda p: int(re.search(r"slide(\d+)", p).group(1)))

        for idx, part in enumerate(slide_parts, 1):
            s = Slide(idx, part)
            try:
                root = ET.fromstring(z.read(part))
            except KeyError:
                continue
            # 背景色
            bg = root.find("p:cSld/p:bg/p:bgPr/a:solidFill/a:srgbClr", NS)
            if bg is not None and bg.get("val"):
                s.bg = bg.get("val")
            # 形状与文本
            for sp in root.findall(".//p:sp", NS):
                tx = sp.find("p:txBody", NS)
                if tx is None:
                    continue
                off = sp.find("p:spPr/a:xfrm/a:off", NS)
                if off is not None and off.get("x") is not None:
                    try:
                        s.shape_xs.append(int(off.get("x")))
                    except ValueError:
                        pass
                for r in tx.findall(".//a:r", NS):
                    rPr = r.find("a:rPr", NS)
                    sz, font, color = None, None, None
                    if rPr is not None:
                        if rPr.get("sz"):
                            try:
                                sz = int(rPr.get("sz"))
                            except ValueError:
                                pass
                        ea = rPr.find("a:ea", NS)
                        latin = rPr.find("a:latin", NS)
                        if ea is not None and ea.get("typeface"):
                            font = ea.get("typeface")
                        elif latin is not None and latin.get("typeface"):
                            font = latin.get("typeface")
                        clr = rPr.find("a:solidFill/a:srgbClr", NS)
                        if clr is not None and clr.get("val"):
                            color = clr.get("val")
                    text = "".join((t.text or "") for t in r.findall("a:t", NS))
                    if text:
                        s.runs.append({"text": text, "sz": sz,
                                       "font": font, "color": color})
            # 演讲者备注：slide 的 rels 里有 notesSlide 即算有
            m = re.search(r"slide(\d+)\.xml$", part)
            rels_name = "ppt/slides/_rels/slide%s.rels" % m.group(1) if m else ""
            if rels_name in names:
                try:
                    rels = ET.fromstring(z.read(rels_name))
                    for rel in rels.findall("rel:Relationship", NS):
                        if "notesSlide" in (rel.get("Type") or ""):
                            s.has_notes = True
                            break
                except ET.ParseError:
                    pass
            slides.append(s)
    return slides


# ---------------- 规则 ----------------

@rule("PPT-FONT-001", "blocker", "字号", "正文字号过小（<12pt）",
      "某页存在小于 12pt 的文字，投影/投屏时最后一排看不清。提到 18pt 以上。")
def r_font_001(slides, path, ctx):
    out = []
    for s in slides:
        tiny = [r for r in s.runs if r["sz"] is not None and r["sz"] < 1200]
        if tiny:
            out.append(_finding("PPT-FONT-001", path, s.index,
                               "最小 %.1fpt：%s" % (min(r["sz"] for r in tiny) / 100,
                                                  tiny[0]["text"][:40])))
    return out


@rule("PPT-FONT-002", "warning", "字号", "正文字号偏小（<18pt）",
      "某页存在 12–18pt 的文字，偏小。中文正文建议 18–24pt，注释不低于 12pt。")
def r_font_002(slides, path, ctx):
    out = []
    for s in slides:
        small = [r for r in s.runs if r["sz"] is not None and 1200 <= r["sz"] < 1800]
        if small:
            out.append(_finding("PPT-FONT-002", path, s.index,
                               "%d 处偏小（如 %.1fpt：%s）" % (
                                   len(small), small[0]["sz"] / 100, small[0]["text"][:40])))
    return out


@rule("PPT-FNT-003", "warning", "版权", "疑似商用版权字体",
      "检测到可能需单独购买商用授权的字体（如微软雅黑/方正/汉仪）。换免费可商用字体"
      "（思源黑体/阿里巴巴普惠体），或确认已有授权。商用前以字体官网最新声明为准。")
def r_fnt_003(slides, path, ctx):
    out = []
    for s in slides:
        hit = set()
        for r in s.runs:
            f = (r["font"] or "").lower()
            for risky in RISKY_FONTS:
                if risky.lower() in f:
                    hit.add(r["font"])
        if hit:
            out.append(_finding("PPT-FNT-003", path, s.index,
                               "字体：%s" % "、".join(sorted(hit))))
    return out


@rule("PPT-TXT-001", "warning", "文字量", "单页文字过多（>180字）",
      "某页文字超过 180 字（一页一观点被打破）。拆成两页，或把细节移到演讲者备注。")
def r_txt_001(slides, path, ctx):
    out = []
    for s in slides:
        if s.char_count > 180:
            out.append(_finding("PPT-TXT-001", path, s.index,
                               "本页约 %d 字" % s.char_count))
    return out


@rule("PPT-PG-001", "warning", "篇幅", "总页数过多（>30页）",
      "全篇超过 30 页。汇报 PPT 宜 15–25 页；超了先砍页，再考虑拆成两场。")
def r_pg_001(slides, path, ctx):
    if len(slides) > 30:
        return [_finding("PPT-PG-001", path, 0, "共 %d 页" % len(slides))]
    return []


@rule("PPT-PG-002", "nit", "篇幅", "总页数过少（<5页）",
      "全篇不足 5 页。确认封面/目录/内容/结尾是否齐全，别把 30 分钟汇报压缩成 5 页。")
def r_pg_002(slides, path, ctx):
    if 0 < len(slides) < 5:
        return [_finding("PPT-PG-002", path, 0, "共 %d 页" % len(slides))]
    return []


@rule("PPT-ALN-001", "warning", "对齐", "同页文本框左对齐不一致",
      "某页 ≥3 个文本框，左边缘偏差超过 1 英寸（2.54cm）。对齐到网格，肉眼'差不多'不算对齐。")
def r_aln_001(slides, path, ctx):
    out = []
    for s in slides:
        if len(s.shape_xs) >= 3 and max(s.shape_xs) - min(s.shape_xs) > EMU_PER_INCH:
            out.append(_finding("PPT-ALN-001", path, s.index,
                               "%d 个文本框左边缘最大偏差 %.1fcm" % (
                                   len(s.shape_xs),
                                   (max(s.shape_xs) - min(s.shape_xs)) / 360000)))
    return out


@rule("PPT-CON-001", "warning", "对比度", "文字与背景对比度不足",
      "某页文字与背景对比度 <3:1（启发式估算），投影上一打就糊。正文与背景建议 ≥4.5:1。")
def r_con_001(slides, path, ctx):
    out = []
    for s in slides:
        for r in s.runs:
            if r["color"] and _contrast(r["color"], s.bg) < 3.0:
                out.append(_finding("PPT-CON-001", path, s.index,
                                   "对比度 %.1f:1：%s" % (
                                       _contrast(r["color"], s.bg), r["text"][:40])))
                break
    return out


@rule("PPT-NOT-001", "nit", "备注", "页面缺演讲者备注",
      "某页没有演讲者备注。每页写 3 句备注（讲这页要说的要点），忘词时救命。")
def r_not_001(slides, path, ctx):
    out = []
    for s in slides:
        if not s.has_notes:
            out.append(_finding("PPT-NOT-001", path, s.index, "无备注"))
    return out


# ---------------------------------------------------------------- 运行器

def collect_files(target):
    if os.path.isfile(target):
        return [target]
    out = []
    for root, dirs, files in os.walk(target):
        dirs[:] = [d for d in dirs if d not in {".git", "__pycache__"}]
        for f in files:
            if f.lower().endswith(".pptx"):
                out.append(os.path.join(root, f))
    return sorted(out)


def run(target):
    files = collect_files(target)
    ctx = {"target": target}
    findings = []
    for path in files:
        rel = os.path.relpath(path, target) if os.path.isdir(target) else os.path.basename(path)
        try:
            slides = parse_pptx(path)
        except Exception as e:
            findings.append(Finding("PPT-PARSE-ERR", "warning", "解析",
                                    "pptx 解析失败", rel, 0, str(e)[:120],
                                    "文件可能损坏或不是标准 pptx。换 PowerPoint/WPS 另存一次再试。"))
            continue
        ctx["slides"] = slides
        for r in RULES:
            try:
                for f in r["check"](slides, rel, ctx) or []:
                    findings.append(f)
            except Exception:
                continue
    return files, findings, ctx


def main(argv=None):
    ap = argparse.ArgumentParser(description="lhg-ppt 确定性 PPT 设计审计")
    ap.add_argument("--target", default=".", help="pptx 文件或目录")
    ap.add_argument("--format", choices=["json", "text"], default="text")
    ap.add_argument("--fail-on", choices=["blocker", "warning", "nit"], default="blocker")
    args = ap.parse_args(argv)

    target = os.path.abspath(args.target)
    if not (os.path.isfile(target) or os.path.isdir(target)):
        print("target 不存在：%s" % args.target, file=sys.stderr)
        return 1

    files, findings, ctx = run(target)

    sev_rank = {"blocker": 3, "warning": 2, "nit": 1}
    counts = {"blocker": 0, "warning": 0, "nit": 0}
    for f in findings:
        counts[f.severity] = counts.get(f.severity, 0) + 1

    result = {
        "tool": TOOL,
        "version": VERSION,
        "target": args.target,
        "summary": {
            "files_scanned": len(files),
            "rules_run": len(RULES),
            "blocker": counts["blocker"],
            "warning": counts["warning"],
            "nit": counts["nit"],
        },
        "findings": [f.to_dict() for f in sorted(
            findings, key=lambda x: (-sev_rank.get(x.severity, 0), x.file, x.line))],
    }

    if args.format == "json":
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        s = result["summary"]
        print("PPT 审计：扫描 %d 文件，运行 %d 条规则" % (s["files_scanned"], s["rules_run"]))
        print("blocker=%d warning=%d nit=%d" % (s["blocker"], s["warning"], s["nit"]))
        for f in result["findings"]:
            where = "第%d页" % f["line"] if f["line"] else "全文"
            print("[%s][%s] %s %s %s" % (
                f["severity"].upper(), f["rule_id"], f["file"], where, f["title"]))
            print("    %s" % f["detail"])
            if f["snippet"]:
                print("    → %s" % f["snippet"])

    level = sev_rank[args.fail_on]
    bad = any(sev_rank.get(f.severity, 0) >= level for f in findings)
    return 2 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
