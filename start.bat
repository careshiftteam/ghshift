@echo off
rem CareShift 起動スクリプト(Windows用)
rem 外部から接続できるよう CARESHIFT_HOST を 0.0.0.0 にセットしてから起動します。
rem このウィンドウを閉じるとCareShiftも停止します。

cd /d "%~dp0"
set CARESHIFT_HOST=0.0.0.0
python app.py
pause
