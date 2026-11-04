#!/usr/bin/env bash
# WordAI 一键安装自检:环境 → 证书 → 依赖 → Pi 凭证 → 侧载 manifest
# 幂等,可以反复跑。用法: npm run setup
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

ok()   { printf '  ✓ %s\n' "$1"; }
warn() { printf '  ⚠️  %s\n' "$1"; }
die()  { printf '\n✗ %s\n\n' "$1" >&2; exit 1; }
step() { printf '\n==> %s\n' "$1"; }

# ── 1. 系统 ────────────────────────────────────────────────────────────────
step "检查系统"
[ "$(uname -s)" = "Darwin" ] || die "WordAI 目前只支持 macOS(Mac 版 Word + wef 侧载路径都是 macOS 专属)。"
ok "macOS $(sw_vers -productVersion)"

# ── 2. Node ────────────────────────────────────────────────────────────────
command -v node >/dev/null 2>&1 || die "没有 Node。先装: brew install node"
NODE_MAJOR="$(node -p 'process.versions.node.split(".")[0]')"
if [ "$NODE_MAJOR" -lt 20 ]; then
  die "Node $(node -v) 太旧,需要 >= 20。升级: brew upgrade node"
fi
ok "Node $(node -v) · npm $(npm -v)"

# ── 3. mkcert ──────────────────────────────────────────────────────────────
command -v mkcert >/dev/null 2>&1 || die "没有 mkcert。先装: brew install mkcert nss"
ok "mkcert $(mkcert -version 2>/dev/null | head -1)"

# ── 4. 本地 https 证书 ─────────────────────────────────────────────────────
step "准备本地 https 证书"
if [ -f "$ROOT/certs/cert.pem" ] && [ -f "$ROOT/certs/key.pem" ]; then
  ok "certs/ 已存在,跳过生成"
else
  mkdir -p "$ROOT/certs"
  mkcert -cert-file "$ROOT/certs/cert.pem" -key-file "$ROOT/certs/key.pem" \
    localhost 127.0.0.1 ::1 >/dev/null
  ok "已生成 certs/cert.pem + key.pem"
fi

# Word 的 WKWebView 会严格校验,根证书不受信任就直接拒绝加载
CA_ROOT="$(mkcert -CAROOT 2>/dev/null || true)"
if [ -n "$CA_ROOT" ] && [ -f "$CA_ROOT/rootCA.pem" ] \
   && security verify-cert -c "$CA_ROOT/rootCA.pem" >/dev/null 2>&1; then
  ok "mkcert 根证书已受系统信任"
else
  cat <<'MSG'

  ⚠️  mkcert 根证书尚未受信任,Word 会拒绝加载 https://localhost:5173。

  请单独执行下面这条命令(要输 Mac 登录密码),然后重新跑 npm run setup:

      sudo "$(command -v mkcert)" -install

MSG
  exit 1
fi

# ── 5. 依赖 ────────────────────────────────────────────────────────────────
step "安装依赖"
if [ -d node_modules ] && [ -f package-lock.json ]; then
  ok "node_modules 已存在(需要更新时跑 npm install)"
else
  npm install
  ok "依赖安装完成"
fi

# ── 6. Pi 内核与凭证 ───────────────────────────────────────────────────────
step "检查 Pi 内核凭证"
PI_AGENT_DIR="${PI_AGENT_DIR:-$HOME/.pi/agent}"
if [ -f "$PI_AGENT_DIR/auth.json" ] || [ -f "$PI_AGENT_DIR/settings.json" ]; then
  ok "找到 Pi 配置:$PI_AGENT_DIR"
else
  cat <<'MSG'

  ⚠️  没找到 Pi 的凭证配置,WordAI 现在还不能调用模型。

  WordAI 不存 API Key,模型和凭证全部由 Pi 自己管理。请先:

      curl -fsSL https://pi.dev/install.sh | sh
      pi            # 首次运行按提示登录 / 配置模型

  配好后再回来跑 npm run setup。

MSG
  exit 1
fi

# ── 7. 侧载 manifest ───────────────────────────────────────────────────────
step "侧载 Word 加载项"
bash "$ROOT/scripts/install-addin.sh"

cat <<'MSG'

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 安装完成。日常启动:

     npm run dev          # 前端 5173 + 代理 8787,保持窗口开着
     npm run service:install   # 可选:交给 launchd 常驻,开机自启

 打开 Word:插入 → 我的加载项 → 「开发人员加载项」里的 WordAI
 选中一段文字,在侧栏发指令。
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

MSG
