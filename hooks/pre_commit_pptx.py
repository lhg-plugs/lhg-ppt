#!/usr/bin/env python3
"""lhg-ppt 硬门①：git commit 前检查暂存区里的 .pptx，二进制有 blocker 则拦截。

触发：PreToolUse / Bash。先判断本次 Bash 是否为 git commit，
不是则 0 开销直接放行；是则看暂存区有无 .pptx，有则跑插件自带的
pptx 审计脚本，有 blocker 才拦截。

逃生方式（按推荐度排序）：
  1. 修掉 blocker 后重新 commit（推荐：改 build_*.py 重跑生成）。
  2. 紧急情况：LHG_PPT_HOOKS_OFF=1 git commit -m "..."（单次豁免本次 hook）。
  3. /plugin disable lhg-ppt（整个插件停用，hooks 不再触发）。

设计原则：脚本内部出错（无 git、审计脚本缺失、超时等）fail-open 直接放行；
真正的门禁在 /lhg-ppt:ppt-audit。零第三方依赖。
"""

import json
import os
import re
import subprocess
import sys

AUDIT_TIMEOUT = 60
COMMIT_RE = re.compile(r"(^|[;&|])\s*git\s+.*\bcommit\b")


def _git(args, cwd):
    try:
        p = subprocess.run(["git"] + args, capture_output=True, text=True,
                           timeout=15, cwd=cwd)
        return p.stdout if p.returncode == 0 else ""
    except Exception:
        return ""


def main() -> int:
    if os.environ.get("LHG_PPT_HOOKS_OFF") == "1":
        return 0

    try:
        raw = sys.stdin.read()
        payload = json.loads(raw) if raw.strip() else {}
    except Exception:
        return 0  # fail-open

    tool_input = payload.get("tool_input", {}) or {}
    command = tool_input.get("command", "") or ""
    if not COMMIT_RE.search(command):
        return 0  # 非 git commit：快速放行

    plugin_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    audit_script = os.path.join(plugin_root, "scripts", "audit", "pptx_audit.py")
    if not os.path.isfile(audit_script):
        return 0  # fail-open

    project_dir = os.environ.get("CLAUDE_PROJECT_DIR") or os.getcwd()
    staged = _git(["diff", "--cached", "--name-only", "--diff-filter=ACM"], project_dir)
    pptx_files = [l.strip() for l in staged.split("\n")
                  if l.strip().lower().endswith(".pptx")]
    if not pptx_files:
        return 0  # 暂存区无 pptx：放行

    blockers = []
    for rel in pptx_files:
        target = rel if os.path.isabs(rel) else os.path.join(project_dir, rel)
        if not os.path.isfile(target):
            continue
        try:
            proc = subprocess.run(
                [sys.executable, audit_script, "--target", target, "--format", "json"],
                capture_output=True, text=True, timeout=AUDIT_TIMEOUT)
            data = json.loads(proc.stdout) if proc.stdout.strip() else {}
        except Exception:
            continue  # 单个审计失败：fail-open
        for f in data.get("findings", []):
            if f.get("severity") == "blocker":
                blockers.append((rel, f))

    if not blockers:
        return 0

    lines = [
        "提交被 lhg-ppt 硬门拦截：暂存区的 .pptx 有 %d 个 blocker（汇报会翻车的问题），"
        "修完再提交。" % len(blockers),
        "建议：二进制 pptx 本就不该进仓库——改 build_*.py 重跑生成，用 .gitignore 忽略 output/。",
        "",
    ]
    for rel, f in blockers[:10]:
        lines.append("- [%s][%s] 第%s页 %s" % (rel, f.get("rule_id"), f.get("line"), f.get("title")))
    if len(blockers) > 10:
        lines.append("…另有 %d 个，完整清单请跑 /lhg-ppt:ppt-audit" % (len(blockers) - 10))
    lines += ["", "紧急豁免（单次）：LHG_PPT_HOOKS_OFF=1 git commit"]
    out = {
        "hookSpecificOutput": {
            "hookEventName": "PreToolUse",
            "permissionDecision": "deny",
            "permissionDecisionReason": "\n".join(lines),
        }
    }
    sys.stdout.write(json.dumps(out, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
