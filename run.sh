#!/bin/sh
# macOS 启动脚本，对应 Windows 版的 run.cmd
cd "$(dirname "$0")" || exit 1
exec python3 app.py
