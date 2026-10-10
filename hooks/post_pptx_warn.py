#!/usr/bin/env python3
"""lhg-ppt 硬门②（警告级，不拦截）：写 deck 生成脚本 / pptx 二进制时提醒。

触发：PostToolUse / Write|Edit。只做提醒，永远 exit 0、不拦截写入。

- 写 .py 且内容含 python-pptx/deck/build_ 关键词 → 提醒跑 /lhg-ppt:ppt-audit
- 写 .pptx 二进制 → 提醒二进制不进 git、改稿走脚本重跑

逃生方式：LHG_PPT_HOOKS_OFF=1（环境变量）可关闭本提醒；
或 /plugin disable lhg-ppt 停用整个插件。
零第三方依赖。
"""

import json
import os
import sys

DECK_PY_HINTS = ("python-pptx", "from pptx", "import pptx", "build_", "deck",
                 "add_slide", "prs.save")


def main() -> int:
    if os.environ.get("LHG_PPT_HOOKS_OFF") == "1":
        return 0

    try:
        raw = sys.stdin.read()
        payload = json.loads(raw) if raw.strip() else {}
    except Exception:
        return 0

    tool_input = payload.get("tool_input", {}) or {}
    file_path = tool_input.get("file_path", "") or ""
    if not file_path:
        return 0
    low = file_path.lower()
    msg = ""

    if low.endswith(".pptx"):
        msg = (
            "PPT 提醒（lhg-ppt）：%s 是二进制文件。\n"
            "1) 二进制不进 git：把 output/ 加进 .gitignore；\n"
            "2) 改稿走脚本：改 build_*.py 重跑，不要手工改 pptx（下次重跑会丢）；\n"
            "3) 交付前跑 /lhg-ppt:ppt-audit 做一次设计体检。\n"
            "关闭本提醒：export LHG_PPT_HOOKS_OFF=1（当前会话有效）。"
        ) % file_path
    elif low.endswith(".py"):
        content = tool_input.get("content", "") or ""
        # PostToolUse 的 Edit 可能没有 content；读文件兜底（失败则放行）
        if not content:
            try:
                with open(file_path, "r", encoding="utf-8", errors="replace") as fh:
                    content = fh.read(4096)
            except OSError:
                return 0
        if any(h in content for h in DECK_PY_HINTS):
            msg = (
                "PPT 提醒（lhg-ppt）：%s 疑似 deck 生成脚本。\n"
                "生成后记得跑 python3 <插件目录>/scripts/audit/pptx_audit.py 做设计自检，\n"
                "或直接跑 /lhg-ppt:ppt-audit（含人工走查清单）。\n"
                "关闭本提醒：export LHG_PPT_HOOKS_OFF=1（当前会话有效）。"
            ) % file_path

    if not msg:
        return 0
    out = {
        "hookSpecificOutput": {
            "hookEventName": "PostToolUse",
            "additionalContext": msg,
        }
    }
    sys.stdout.write(json.dumps(out, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
