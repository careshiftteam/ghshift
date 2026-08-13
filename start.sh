#!/bin/bash
# CareShift 本番起動スクリプト
# 外部から接続できるよう CARESHIFT_HOST=0.0.0.0 をセットしてから起動する。
# 直接 `python3 app.py` を実行すると127.0.0.1のみになり、外部から繋がらないので注意。
#
# 起動前に、古いプロセスが残っていないか必ず確認すること：
#   ps aux | grep app.py
#   （残っていれば）sudo kill <PID>

cd "$(dirname "$0")"
export CARESHIFT_HOST=0.0.0.0
python3 app.py
