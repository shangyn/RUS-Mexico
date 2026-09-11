#!/usr/bin/env bash
# 发货额度统计（quota_app）更新脚本
# 用法: bash update.sh
set -e

APP_DIR="$(cd "$(dirname "$0")" && pwd)"
PORT=5050
cd "$APP_DIR"

echo "==> 拉取最新代码"
git pull --ff-only

echo "==> 停止 ${PORT} 端口上的旧进程"
fuser -k ${PORT}/tcp 2>/dev/null || true
sleep 1

echo "==> 启动 waitress"
nohup waitress-serve --host=127.0.0.1 --port=${PORT} \
  --channel-timeout=600 --threads=8 quota_app:app > app.log 2>&1 &
NEW_PID=$!
sleep 2

echo "==> 启动完成 PID=${NEW_PID}"
echo "==> 日志: ${APP_DIR}/app.log"