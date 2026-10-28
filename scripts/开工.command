#!/bin/bash
# 开工：下载另一台设备的最新修改
cd /Users/Shared/Projects/GrandJOY || { echo "⚠️ 找不到项目目录"; read -p "按回车键关闭"; exit 1; }
echo "正在下载最新修改（git pull）..."
if ! git pull --rebase origin main; then
  echo "⚠️ 下载失败，可能是网络问题或两台设备有冲突。"
  echo "   请把本窗口内容截图发给 TRAE 里的 AI 处理。"
  read -p "按回车键关闭"
  exit 1
fi
echo ""
echo "✅ 已是最新。可以打开 TRAE 开始工作了。"
read -p "按回车键关闭窗口"
