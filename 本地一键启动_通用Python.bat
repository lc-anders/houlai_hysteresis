@echo off
chcp 65001 >nul
title houlai_hysteresis 滞回曲线处理与抗震评价程序

echo ========================================
echo houlai_hysteresis v1.0
echo 滞回曲线处理与抗震评价程序
echo ========================================
echo.

cd /d "%~dp0"

echo 正在检查并安装依赖...
python -m pip install -r requirements.txt

echo.
echo 正在启动程序...
echo 浏览器若未自动打开，请访问 http://localhost:8501
python -m streamlit run app.py

pause
