#!/bin/bash
# 收工：把今天的修改上传同步
cd /Users/Shared/Projects/GrandJOY || { echo "⚠️ 找不到项目目录"; read -p "按回车键关闭"; exit 1; }
echo "第1步：登记修改（git add）..."
git add -A
if git commit -m "日常同步 $(date '+%Y-%m-%d %H:%M')"; then
  echo "✅ 已提交修改。"
else
  echo "（本次没有新修改，跳过提交）"
fi
echo "第2步：先下载（防止覆盖另一台设备的修改）..."
if ! git pull --rebase origin main; then
  echo "⚠️ 下载失败，可能两台设备改到了同一个文件。"
  echo "   请把本窗口内容截图发给 TRAE 里的 AI 处理，不要强行覆盖。"
  read -p "按回车键关闭"
  exit 1
fi
echo "第3步：上传（git push）..."
if git config remote.ssh.url >/dev/null 2>&1; then
  git push ssh main
else
  git push origin main
fi
if [ $? -ne 0 ]; then
  echo "⚠️ 上传失败，请把本窗口内容截图发给 TRAE 里的 AI 处理。"
  read -p "按回车键关闭"
  exit 1
fi
echo ""
echo "✅ 已上传成功，另一台设备 pull 后就能看到今天的成果。"
read -p "按回车键关闭窗口"
