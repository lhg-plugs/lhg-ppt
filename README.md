<!-- About 简介（推送后填写到仓库 About 栏）：
PPT制作专家：述职融资答辩一键成稿 ｜ 一键安装：/plugin marketplace add lhg-plugs/lhg-ppt
-->
<!-- topics（推送后打）：claude-code, claude-plugin, agent-skills, ppt, presentation, python-pptx -->

# lhg-ppt：PPT 制作专家

> **一句话介绍**：给 AI 助手装上"汇报组合拳"——一句话/一篇文档→金字塔大纲→中文审美母版→python-pptx 逐页生成**原生可编辑** PPT→设计自检，述职/融资/答辩一键成稿。
>
> **一键安装**：在 Claude Code 里运行
> ```
> /plugin marketplace add lhg-plugs/lhg-ppt
> /plugin install lhg-ppt@lhg-plugs
> ```

## 它是什么

`lhg-ppt` 是一个 Claude Code 插件（plugin 为主、skill 为辅）：

- **3 个命令**（小白入口）：`ppt` 一句话成稿一条龙，`ppt-outline` 只出大纲，`ppt-audit` 设计体检
- **4 个平台中立 skill**（方法论，可移植到豆包/Workbuddy 等）：金字塔大纲、中文审美母版、python-pptx 逐页生成、设计自检
- **3 个专家 agent**：大纲专家、生成专家、设计评审
- **2 个硬门 hooks**：commit 前 pptx 审计拦截、写 deck 脚本提醒（均可一键豁免）
- **9 条确定性审计规则**：纯 Python、零第三方依赖，直接解析 pptx 查字号/对比度/对齐/文字量/页数

差异化（相对 ppt-master 等英文生态）：**中文审美母版**（免费可商用字体清单+版权雷区）、
**说人话的小白交互**（大纲确认卡、风格确认卡）、**汇报场景模板**（述职/融资/答辩）。

## 3 个命令用法

| 命令 | 作用 | 示例 |
|---|---|---|
| `/lhg-ppt:ppt` | 一句话/一篇文档→大纲→母版→逐页生成→自检，一条龙 | `/lhg-ppt:ppt "Q3 用户增长复盘" --style 商务蓝` |
| `/lhg-ppt:ppt-outline` | 只出大纲：金字塔结构，每页一句话标题+要点+数据 | `/lhg-ppt:ppt-outline "融资路演"` |
| `/lhg-ppt:ppt-audit` | 设计体检：9 条规则审计 + 投影/内容/合规人工走查 | `/lhg-ppt:ppt-audit output/复盘-20261011.pptx` |

## 自查说明

自查 = 确定性脚本（机器）+ 人工走查（人）：

1. 先跑 `scripts/audit/pptx_audit.py --target <pptx文件>`，零成本（不用装 python-pptx 也能跑）；
2. 发现分三级：**blocker**（字号<12pt，必须改生成脚本重跑）、**warning**（对比度/对齐/文字量/疑似版权字体，修或写豁免理由）、**nit**（缺演讲者备注）；
3. 再按 `lhg-ppt-review` 第 5 节做人工走查：投影实测、错别字、数据口径、字体版权。

blocker 清零才算"可汇报"；为清零改脚本规则或阈值 → 作弊，不许。

## 目录结构

```
lhg-ppt/
├── .claude-plugin/          # plugin.json（插件清单）+ marketplace.json（本仓即市场）
├── commands/                # 3 个命令：ppt / ppt-outline / ppt-audit
├── skills/                  # 4 个平台中立 skill（lhg-ppt-outline/aesthetics/build/review）
├── agents/                  # 3 个专家 agent（outline-writer / deck-builder / design-reviewer）
├── hooks/                   # 2 个硬门（commit 前 pptx 审计拦截 / 写 deck 脚本提醒）+ 豁免说明
├── scripts/audit/           # 9 条确定性审计规则（零第三方依赖，zipfile+XML 解析 pptx）
├── fixtures/                # 冒烟测试用例（bad.pptx / good.pptx + make_fixtures.py 生成器）
├── smoke_test.py            # 一键冒烟测试
├── ATTRIBUTION.md           # 灵感来源与许可声明
└── LICENSE                  # MIT
```

## 什么时候用 / 什么时候不用（FAQ）

**什么时候用：**

- 你下周要述职/答辩/路演，大纲还没影 → 用 `/lhg-ppt:ppt`，一句话开始，大纲确认后再生成
- 你有一篇长文档，想转成能讲的 PPT → 先走提炼+大纲，不许分页粘贴
- 你的 PPT 总被吐槽"字小""看不清" → `/lhg-ppt:ppt-audit` 先体检，9 条规则 10 秒出结果
- 你想要"中国人看着顺眼"的版式 → 中文审美母版（字体/配色/版式）开箱即用

**什么时候不用：**

- 你要的是英文风/欧美极简风 PPT → 本插件的审美规范按中文汇报场景调的，不对味别硬套
- 你想一键生成 100 页的长报告 → 本插件按 15–25 页汇报场景设计，超 30 页审计会报警
- 你希望 AI 直接出"图片式"PPT（整页贴图）→ **本插件不干这个**，必须原生可编辑，这是底线

**和 ppt-master 的区别（诚实版）：**

- ppt-master（hugohe3，约 5.9 万 stars，MIT）是英文生态最强的"文档→原生 PPTX"工作流，
  本插件借鉴了它的工作流思想，但**没有复制它的任何代码与文案**。
- 本插件的差异化不在"生成"本身，而在三件事：① 中文审美母版（字体/配色/版式按中文汇报场景调优）；
  ② 说人话的小白交互（大纲确认卡、风格确认卡，一步一确认）；③ 汇报场景模板（述职/融资/答辩）+
  确定性设计自检（9 条规则，零依赖）。
- 如果你英文好、习惯自己调版式，ppt-master 可能更适合你；如果你要"中文汇报、开箱即用、有人把关"，
  试试本插件。

**版权字体风险（必读）：**

- 本插件默认使用免费可商用字体（思源黑体/阿里巴巴普惠体等），审计脚本的 PPT-FNT-003
  会提示疑似商用版权字体（如微软雅黑/方正/汉仪）——**提示不是摆设**。
- 字体厂商可能变更授权条款，商用前请以各字体官网最新声明为准，本插件不构成法律意见。
- 详见 skill `lhg-ppt-aesthetics` 第 6 节。

## lhg-plugs 插件矩阵

| 插件 | 一句话 | 安装 |
|---|---|---|
| [lhg-ppt](https://github.com/lhg-plugs/lhg-ppt) | PPT 制作专家：述职融资答辩一键成稿 | `/plugin install lhg-ppt@lhg-plugs` |
| *（后续插件在此追加，保持表格格式一致）* | | |

## 自检与反馈

本插件自带质检：每次使用后可跑 `python3 smoke_test.py` 一键验证（JSON 合法、skill/command/agent 完整性、9 条规则、fixtures 命中、hooks 行为）。

发现 bug 或有改进建议，以 `[QC] <一句话问题>` 为标题提交到 https://github.com/lhg-plugs/lhg-ppt/issues，正文写清"复现步骤 → 实际表现 → 期望表现"。

## 出品：刘洪光

本插件由真人出镜 IP「刘洪光」（安徽合肥）出品，归属 [lhg-skills](https://github.com/lhg-skills)。

- GitHub 主页：https://github.com/lhg-skills —— 全部 skill 开源在此，欢迎 star
- 视频号：搜「刘洪光实名上网」
- 微信：lhgsmsw
