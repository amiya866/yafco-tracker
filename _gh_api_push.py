# -*- coding: utf-8 -*-
# github.com 被墙、代理未开时的绕行部署：用 gh(api.github.com 直连) 走 Contents API 逐文件推送。
# 用法：python _gh_api_push.py   （推送当前本地 HEAD~1..HEAD 提交涉及的文件到 main）
import base64
import json
import subprocess
import sys
from pathlib import Path

REPO = "wangziquan-del/yafco-tracker"
BRANCH = "main"
BASE = Path(__file__).parent


def gh(args, input_data=None):
    cmd = ["gh", "api"] + args
    r = subprocess.run(cmd, input=input_data, capture_output=True, text=True, encoding="utf-8")
    return r


def remote_sha(path):
    r = gh([f"repos/{REPO}/contents/{path}?ref={BRANCH}"])
    if r.returncode != 0:
        return None
    return json.loads(r.stdout)["sha"]


def main():
    files = subprocess.run(["git", "diff", "--name-only", "HEAD~1", "HEAD"],
                           capture_output=True, text=True, cwd=BASE).stdout.split()
    files = [f for f in files if f.strip()]
    print(f"待推送 {len(files)} 个文件")
    last_sha = None
    for f in files:
        f = f.strip()
        local = BASE / f
        sha = remote_sha(f)
        if not local.exists():
            # 本地已删除 → 远端同步删除
            if sha is None:
                print(f"  跳过（双端均无）{f}")
                continue
            body = json.dumps({"message": f"删除 {f}（数据更新）", "sha": sha, "branch": BRANCH})
            r = gh(["-X", "DELETE", f"repos/{REPO}/contents/{f}", "--input", "-"], input_data=body)
        else:
            content = base64.b64encode(local.read_bytes()).decode()
            msg = {"message": f"更新 {f}（数据更新 2026-08-04）", "content": content, "branch": BRANCH}
            if sha:
                msg["sha"] = sha
            r = gh(["-X", "PUT", f"repos/{REPO}/contents/{f}", "--input", "-"],
                   input_data=json.dumps(msg))
        if r.returncode != 0:
            print(f"  FAIL {f}: {r.stderr[:300]}")
            sys.exit(1)
        last_sha = json.loads(r.stdout)["commit"]["sha"]
        print(f"  OK {f} -> {last_sha[:8]}")
    print(f"完成，最终提交 {last_sha}")


if __name__ == "__main__":
    main()
