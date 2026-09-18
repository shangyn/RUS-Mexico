#!/bin/bash
# 发货额度统计 · 服务器一键更新
#   拉取最新代码 → 安装/更新依赖 → 重启 waitress
#
# 用法：
#   cd /www/wwwroot/quota.xxx.com
#   bash update.sh
#
# 说明：requirements.txt 有变化时必须走这个脚本（或手动执行 pip install），
#       否则新功能会因缺少依赖而不可用（例如凭证 PDF 需要 fpdf2）。
set -e

cd "$(dirname "$0")"

PORT=5050
APP="quota_app:app"
PY="python3.11"

echo "==> 1/3 拉取最新代码"
git pull

echo "==> 2/3 安装 / 更新依赖"
"$PY" -m pip install -r requirements.txt

echo "==> 3/3 重启 waitress"
pkill -f waitress-serve || true
sleep 1
nohup waitress-serve --host=127.0.0.1 --port="$PORT" --channel-timeout=600 --threads=8 "$APP" > app.log 2>&1 &
sleep 2

if pgrep -f waitress-serve > /dev/null; then
  echo "更新完成：waitress 已启动（端口 $PORT）"
  echo "如页面异常，请查看 app.log"
else
  echo "启动失败：waitress 未运行，请查看 app.log" >&2
  exit 1
fi
