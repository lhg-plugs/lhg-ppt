---
name: lhg-ppt-build
description: python-pptx逐页生成实操：原生可编辑形状/文本框/图表/表格的生成规范、母版常量、版式函数。拒绝"贴图式PPT"。
license: MIT
compatibility: 适用于任何支持 Markdown 指令的 AI 助手（Claude Code、豆包、Workbuddy 等）
metadata:
  version: 1.0.0
  author: 刘洪光
---

# lhg-ppt-build：python-pptx 逐页生成实操

> 目标：用 python-pptx 生成**原生可编辑**的 .pptx——每个字都是文本框，
> 每张图都是原生图表，拿到手就能改。
> 核心纪律：**整页贴图 = 不合格；文字不可选 = 不合格；数据编造 = 红线**。

## 1. 何时使用

- `/lhg-ppt:ppt` 的 Phase 3 → 全程按本 skill 执行
- 大纲已确认、风格三件套已定 → 开始逐页生成
- 用户说"这页字太小/对不齐" → 按第 5 节调生成脚本重跑，不要手工改 pptx

## 2. 输入与输出

- **输入**：确认过的大纲（lhg-ppt-outline 输出）、风格三件套（lhg-ppt-aesthetics 输出）
- **输出**：`output/<主题>-<日期>.pptx` + 生成脚本 `output/build_<主题>.py`（留着，下次改稿重跑）
- **环境**：本机装 python-pptx（`pip install python-pptx`，只需装一次）

## 3. 生成脚本骨架

一个可维护的生成脚本分四段。风格常量**全部**来自 lhg-ppt-aesthetics，不许在页面代码里硬编码颜色字号：

```python
# --- 0. 风格常量（来自 lhg-ppt-aesthetics，本页唯一可改颜色字号的地方）---
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

TITLE_FONT, BODY_FONT = "思源黑体", "思源黑体"   # 免费可商用，见 aesthetics 3.1
MAIN, ACCENT, TEXT = RGBColor(0x1F,0x4E,0x79), RGBColor(0xED,0x7D,0x31), RGBColor(0x33,0x33,0x33)
TITLE_SIZE, BODY_SIZE, NOTE_SIZE = Pt(30), Pt(20), Pt(12)
MARGIN = Inches(0.7)                              # 安全边距

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)  # 16:9
```

```python
# --- 1. 版式函数（一页只调函数，不写重复代码）---
def add_title_slide(prs, title, subtitle):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白版式
    # 主标题文本框：add_textbox(left, top, width, height)
    ...

def add_bullets_slide(prs, title, points, note=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # 页标题 + 要点文本框（≤3条）+ 右下角数据口径 + 演讲者备注
    ...

def add_chart_slide(prs, title, conclusion, categories, values, note=""):
    # 原生图表：chart_data = ChartData(); 不是截图！
    ...
```

```python
# --- 2. 按大纲逐页调用（大纲一页 = 一次函数调用）---
add_title_slide(prs, "Q3 用户增长复盘", "增长部 · 张三 · 2026-10")
add_bullets_slide(prs, "Q3 获客成本下降 40%，主因是短视频渠道起量",
                  ["短视频渠道获客占比从 10% 升至 45%",
                   "单用户获客成本从 18 元降至 11 元",
                   "次月留存率提升 8 个百分点"],
                  note="数据来源：内部 BI，2026Q3")
...
```

```python
# --- 3. 保存（二进制不进 git，见 .gitignore）---
prs.save("output/主题-20261011.pptx")
```

## 4. 页面类型清单（按需选用）

| 类型 | 函数 | 什么时候用 |
|---|---|---|
| 封面 | add_title_slide | 每份 1 页 |
| 要点页 | add_bullets_slide | 90% 的内容页 |
| 数据页 | add_chart_slide | 有数字结论时（折线/柱状/饼） |
| 对比页 | add_compare_slide | 两方案/前后对比（左右分栏） |
| 时间线 | add_timeline_slide | 里程碑、规划 |
| 封底 | add_end_slide | 谢谢+联系方式 |

**没有"图片页"类型**：需要配图时，用纯色块+大字标题代替；
真要图片，用 `add_picture` 且图片必须有版权（自制/免费图库），并在备注里写来源。

## 5. 生成铁律（违反一条就返工）

1. **原生可编辑**：所有文字都是文本框，所有图表都是 python-pptx 原生图表。
   整页 `add_picture` 贴图 → 不合格。
2. **字号**：正文 ≥18pt，标题 28–32pt（常量区统一改，不许单页特调）。
3. **对齐**：文本框 left 对齐到网格（MARGIN 的整数倍），不用"看着差不多"。
4. **一页一观点**：要点 >3 条 → 拆成两页调函数两次，不要挤。
5. **备注**：每页写演讲者备注（讲这页时要说的 3 句话），审计 PPT-NOT-001 会查。
6. **数据诚实**：图表数据来自用户给的真实数据；没有数据的页写"数据待补充"，
   **不许编数字凑图表**。
7. **改稿走脚本**：用户说改 → 改 `build_<主题>.py` 重跑，不要手工改 pptx
   （手工改的下次重跑就丢了）。

### 演讲者备注怎么写（每页 3 句，审计 PPT-NOT-001 会查）
1. 第一句：这页的结论（一句话）。
2. 第二句：支撑结论的关键数据/案例（防忘数字）。
3. 第三句：过渡到下一页的衔接语（"看完现状，我们看问题出在哪"）。
备注是讲给自己看的，口语化，不要写书面语。

## 6. 图表数据规范

- 图表数据单独放脚本顶部的 `DATA = {...}` 区，和页面代码分离，方便用户核对。
- 每个图表配一句话结论当标题（"获客成本下降 40%"），不要写"近 6 个月数据"。
- 数据口径写进该页备注 + 右下角小字（12pt 灰色）。

## 7. 常见误区（反模式）

- ❌ 用 `slide_layouts` 默认版式不调样式——默认版式是英文审美，按本 skill 常量重写
- ❌ 颜色字号写死在每页代码里——改一版风格要改 30 处；全部收敛到常量区
- ❌ 生成脚本用完就删——下次改稿还得重写，留着 `build_<主题>.py`
- ❌ 把 .pptx 提交进 git——二进制进仓库，diff 没法看；输出目录加 .gitignore
- ❌ 图表用截图——"原生可编辑"是本插件的底线，截图一律重做
- ❌ 编数据让图表好看——**红线**，发现一次整份 PPT 信用破产

## 8. 自检与反馈

生成完成后：

1. 用 PowerPoint/WPS 打开，全选一页 → 都是可编辑文本框吗？→ 有整页图片就返工
2. 跑审计脚本：`pptx_audit.py --target output/<文件>.pptx`，blocker 清零
3. 备注检查：每页都有演讲者备注吗？
4. 改稿测试：改一个标题重跑脚本，新 pptx 生出来了吗？（验证脚本可重复）
5. 文件大小：单文件 >50MB？→ 图片没压缩，压到 200dpi 以内

反馈渠道：以 `[QC] <一句话问题>` 为标题提交到 https://github.com/lhg-plugs/lhg-ppt/issues，
正文写清"场景 → 问题 → 期望"。好用的版式函数也欢迎贡献（附效果描述）。

## 出品：刘洪光

本 skill 由真人出镜 IP「刘洪光」（安徽合肥）出品，归属 [lhg-skills](https://github.com/lhg-skills)。

- GitHub 主页：https://github.com/lhg-skills —— 全部 skill 开源在此，欢迎 star
- 抖音：搜「刘老师讲AI实战」
- 微信：lhgsmsw
