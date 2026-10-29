#!/usr/bin/env bash
# =============================================================================
# 全国运动/男装战略品牌合同台账 —— 一键更新
#
# 用法：
#   1) 把新一版 BI 导出的 xlsx 放进本目录的「源数据/」子目录
#   2) 双击本文件，或在终端执行：
#        ./脚本/更新台账.sh                 # 自动取源数据里最新的 xlsx
#        ./脚本/更新台账.sh <文件路径>       # 指定文件
#
# 动作：
#   1) 生成本地 Excel 台账（产出/全国运动男装战略品牌合同台账_<版本>.xlsx）
#   2) 覆盖刷新飞书同名表（我的空间 / 大悦商业 / 战略台账）
#   3) 追加一条版本记录到「版本记录」子表
# =============================================================================
set -uo pipefail

TASK_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SCRIPT_DIR="$TASK_DIR/脚本"
SRC_DIR="$TASK_DIR/源数据"
OUT_DIR="$TASK_DIR/产出"

PY="/Users/zhijian/.grandjoybuddy/binaries/python/envs/default/bin/python"
if [ ! -x "$PY" ]; then
  PY="/Users/zhijian/.grandjoybuddy/binaries/python/versions/3.13.12/bin/python3"
fi

# 飞书目标表（我的空间 / 大悦商业 / 战略台账）
SHEET_TOKEN="JDKGs49LKhZRL9ts7mscg14inbg"
SHEET_URL="https://my.feishu.cn/sheets/${SHEET_TOKEN}"

echo "==========================================================="
echo " 运动/男装战略品牌合同台账 · 更新"
echo "==========================================================="

# 1. 选定源文件
SRC="${1:-}"
if [ -z "$SRC" ]; then
  SRC=$(ls -1t "$SRC_DIR"/*.xlsx 2>/dev/null | head -1)
fi
if [ -z "$SRC" ] || [ ! -f "$SRC" ]; then
  echo "✗ 未找到源文件。请把 BI 导出的 xlsx 放进："
  echo "  $SRC_DIR"
  exit 1
fi
echo "源文件：$SRC"

# 2. 生成本地 Excel
echo ""
echo "[1/3] 生成本地 Excel 台账..."
"$PY" "$SCRIPT_DIR/生成台账xlsx.py" --源文件 "$SRC" --说明 "定期更新" || exit 1

# 3. 刷新飞书表（覆盖式）
echo ""
echo "[2/3] 刷新飞书表..."
"$PY" "$SCRIPT_DIR/推送飞书.py" --源文件 "$SRC" --表格token "$SHEET_TOKEN" || exit 1

# 4. 重建看板
echo ""
echo "[3/4] 重建台账看板..."
"$PY" "$SCRIPT_DIR/加看板.py" --表格token "$SHEET_TOKEN" --源文件 "$SRC" || exit 1

# 5. 条件格式 + 子表顺序
echo ""
echo "[4/4] 配置条件格式与子表顺序..."
"$PY" "$SCRIPT_DIR/加条件格式.py" --表格token "$SHEET_TOKEN" || echo "  (条件格式配置失败，可手工重跑)"

# 6. 追加版本记录
echo ""
echo "[追加] 版本记录..."
"$PY" "$SCRIPT_DIR/追加版本记录.py" --表格token "$SHEET_TOKEN" --源文件 "$SRC" || echo "  (版本记录追加失败，可手工补)"

echo ""
echo "✓ 更新完成"
echo "  本地：$OUT_DIR"
echo "  飞书：$SHEET_URL"
