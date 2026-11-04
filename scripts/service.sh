#!/usr/bin/env bash
# WordAI 开发服务 · launchd LaunchAgent 管理
# 用法: bash scripts/service.sh install|uninstall|status|restart|logs
#
# 为什么用 LaunchAgent 而不是 LaunchDaemon：
# WordAI 是 GUI 用户级进程（npm/vite/Word 加载项都在用户会话里），
# 用 Daemon 需要 root 且拿不到用户的 HOME 和登录态，必失败。

set -euo pipefail

LABEL="com.wordai.dev"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
LOG_DIR="$HOME/Library/Logs/WordAI"
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
NPM="$(command -v npm || true)"
UID_NUM="$(id -u)"
DOMAIN="gui/$UID_NUM"

if [ -z "$NPM" ]; then
  echo "✗ 找不到 npm，先装 Node（brew install node）" >&2
  exit 1
fi

# launchd 默认 PATH 只有 /usr/bin:/bin:/usr/sbin:/sbin，
# 不显式注入 Homebrew 路径，npm 及其 shebang 里的 node 都会找不到。
NODE_BIN_DIR="$(dirname "$NPM")"
FULL_PATH="$NODE_BIN_DIR:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"

port_busy() {
  lsof -nP -iTCP:"$1" -sTCP:LISTEN >/dev/null 2>&1
}

cmd_install() {
  mkdir -p "$HOME/Library/LaunchAgents" "$LOG_DIR"

  # 幂等：已加载就先卸掉，避免 bootstrap 报 "already loaded"
  launchctl bootout "$DOMAIN/$LABEL" 2>/dev/null || true

  # 端口被别的东西（比如手动 npm run dev）占着，装上也起不来，先拦住
  for p in 5173 8787; do
    if port_busy "$p"; then
      echo "✗ 端口 $p 已被占用，先停掉再装：" >&2
      lsof -nP -iTCP:"$p" -sTCP:LISTEN >&2
      exit 1
    fi
  done

  cat > "$PLIST" <<PLIST_EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>$LABEL</string>

  <key>ProgramArguments</key>
  <array>
    <string>$NPM</string>
    <string>run</string>
    <string>dev</string>
  </array>

  <key>WorkingDirectory</key>
  <string>$ROOT</string>

  <key>EnvironmentVariables</key>
  <dict>
    <key>PATH</key>
    <string>$FULL_PATH</string>
    <key>HOME</key>
    <string>$HOME</string>
  </dict>

  <key>RunAtLoad</key>
  <true/>

  <!-- 只有异常退出才拉起；正常停止后不会被强行重启 -->
  <key>KeepAlive</key>
  <dict>
    <key>SuccessfulExit</key>
    <false/>
  </dict>

  <!-- 崩溃重试最小间隔，避免端口冲突时疯狂刷日志 -->
  <key>ThrottleInterval</key>
  <integer>10</integer>

  <key>StandardOutPath</key>
  <string>$LOG_DIR/dev.log</string>
  <key>StandardErrorPath</key>
  <string>$LOG_DIR/dev.err.log</string>
</dict>
</plist>
PLIST_EOF

  plutil -lint "$PLIST" >/dev/null
  launchctl bootstrap "$DOMAIN" "$PLIST"
  echo "✓ 已安装并加载：$PLIST"
  echo "  日志：$LOG_DIR/dev.log"
  echo "  等待服务就绪…"
  for i in $(seq 1 30); do
    if port_busy 8787 && port_busy 5173; then
      echo "✓ 5173 + 8787 已就绪"
      return 0
    fi
    sleep 1
  done
  echo "✗ 30 秒内端口未就绪，看日志：tail -30 $LOG_DIR/dev.log" >&2
  exit 1
}

cmd_uninstall() {
  launchctl bootout "$DOMAIN/$LABEL" 2>/dev/null || true
  rm -f "$PLIST"
  echo "✓ 已卸载（日志保留在 $LOG_DIR）"
}

cmd_status() {
  if [ ! -f "$PLIST" ]; then
    echo "✗ 未安装（$PLIST 不存在）"
    return 0
  fi
  echo "— launchd 状态 —"
  launchctl print "$DOMAIN/$LABEL" 2>/dev/null \
    | grep -E '^\s+(state|pid|last exit code|program) =' || echo "（未加载）"
  echo "— 端口 —"
  for p in 5173 8787; do
    if port_busy "$p"; then
      echo "  $p ✓ LISTEN"
    else
      echo "  $p ✗ 未监听"
    fi
  done
  echo "— 探测 —"
  curl -sk -o /dev/null -w "  https 5173: %{http_code}\n" https://localhost:5173/index.html || true
  curl -s  -o /dev/null -w "  http  8787: %{http_code}\n" http://localhost:8787/ || true
}

cmd_restart() {
  launchctl kickstart -k "$DOMAIN/$LABEL"
  echo "✓ 已重启"
}

case "${1:-}" in
  install)   cmd_install ;;
  uninstall) cmd_uninstall ;;
  status)    cmd_status ;;
  restart)   cmd_restart ;;
  logs)      tail -f "$LOG_DIR/dev.log" "$LOG_DIR/dev.err.log" ;;
  *)
    echo "用法: bash scripts/service.sh {install|uninstall|status|restart|logs}" >&2
    exit 1
    ;;
esac
