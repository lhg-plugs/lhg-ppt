---
description: PPT设计体检：解析pptx检查字号/对比度/对齐/文字量/页数，输出分级报告。
argument-hint: [pptx文件] [--fix]
---

# /lhg-ppt:ppt-audit —— 设计体检

用户输入在 `$ARGUMENTS` 中：`[pptx文件路径] [--fix]`。缺省扫描 `output/` 下最新的 .pptx。

## 流程（加载 lhg-ppt-review 全程指导）

### 1. 跑脚本
```bash
python3 <插件目录>/scripts/audit/pptx_audit.py --target <pptx文件> --format text
```
脚本用 zipfile+XML 直接解析 pptx（零第三方依赖），检查：
字号下限、单页文字量、总页数、对齐、对比度启发式、版权字体提示、演讲者备注。

### 2. 分级与修复
- blocker（正文字号 <12pt 等）：逐条修；`--fix` 缺省即自动修（改生成脚本重跑，最小改动）。
- warning：修掉或写豁免理由；nit：顺手修。

### 3. 人工走查（脚本覆盖不到的）
按 lhg-ppt-review 第 5 节清单执行：
- 投影/投屏试看：最后一排能看清吗
- 错别字：标题零容忍
- 数据口径：各页同一指标一致吗
- 字体版权：用了商用字体吗（见 lhg-ppt-aesthetics 第 6 节）

### 4. 输出报告
按 lhg-ppt-review 第 6 节模板输出，存为 `output/<主题>-audit-YYYYMMDD.md`，
结论写清"可汇报 / 需修改"。

## 硬规则
- 有 blocker 就说"可汇报" → 不许。
- 为清零改脚本规则或阈值 → 作弊，不许。
