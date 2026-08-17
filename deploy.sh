#!/bin/bash
# 一键：重建数据 → 提交 → 推送 GitHub Pages（外网 https://wangziquan-del.github.io/yafco-tracker/）
set -e
cd "$(dirname "$0")"
python build_site.py
git add -A
git commit -m "数据更新 $(date '+%Y-%m-%d %H:%M')" || echo "无变更，跳过提交"
git push
echo "已推送，GitHub Pages 约 1 分钟后生效：https://wangziquan-del.github.io/yafco-tracker/"
