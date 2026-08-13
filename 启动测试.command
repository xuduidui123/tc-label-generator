#!/bin/bash
# ============================================================
# Tesco 标贴系统 · 本地测试一键启动（macOS 双击运行）
# 首次运行会自动建虚拟环境并安装依赖，之后启动很快。
# ============================================================
cd "$(dirname "$0")" || exit 1

echo "======================================"
echo "  Tesco 标贴系统 · 本地测试"
echo "======================================"

# 1) 检查 python3
if ! command -v python3 >/dev/null 2>&1; then
  echo "❌ 未找到 python3。请先安装 Python 3（https://www.python.org/downloads/ 或 brew install python）。"
  echo "按回车键关闭…"; read; exit 1
fi

# 2) 检查 LibreOffice（生成 PDF 必需）
if [ ! -e "/Applications/LibreOffice.app/Contents/MacOS/soffice" ] \
   && ! command -v soffice >/dev/null 2>&1 \
   && ! command -v libreoffice >/dev/null 2>&1; then
  echo "⚠️  未检测到 LibreOffice——可以打开网页界面，但『生成 PDF』会失败。"
  echo "    安装方法：brew install --cask libreoffice   或到 https://www.libreoffice.org 下载。"
  echo ""
fi

# 3) 首次创建虚拟环境
if [ ! -d ".venv" ]; then
  echo "🔧 首次运行：创建虚拟环境并安装依赖（需几分钟，请耐心等待）…"
  python3 -m venv .venv || { echo "❌ 创建虚拟环境失败"; read; exit 1; }
  source .venv/bin/activate
  python -m pip install --upgrade pip >/dev/null
  pip install -r requirements.txt || { echo "❌ 依赖安装失败"; read; exit 1; }
else
  source .venv/bin/activate
  # 确保 streamlit 在（以防依赖有更新）
  python -c "import streamlit" 2>/dev/null || pip install -r requirements.txt
fi

# 4) 启动，浏览器会自动打开 http://localhost:8501
echo ""
echo "🚀 正在启动……浏览器会自动打开。测试完成后，回到本窗口按 Ctrl+C 关闭。"
echo ""
streamlit run app.py

echo ""
echo "已停止。按回车键关闭本窗口…"; read
