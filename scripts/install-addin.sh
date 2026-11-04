#!/usr/bin/env bash
# WordAI 侧载脚本:把 manifest 放进 Mac Word 的 wef 目录
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
WEF="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
MANIFEST="$ROOT/manifest.xml"

echo "==> WordAI 加载项侧载"

# 0. 依赖工具
if ! command -v mkcert >/dev/null 2>&1; then
  echo "✗ 找不到 mkcert,先装:brew install mkcert nss" >&2
  exit 1
fi

# 1. 证书存在性
if [ ! -f "$ROOT/certs/cert.pem" ]; then
  echo "未找到证书,正在生成…"
  mkdir -p "$ROOT/certs"
  mkcert -cert-file "$ROOT/certs/cert.pem" -key-file "$ROOT/certs/key.pem" localhost 127.0.0.1 ::1
fi

# 2. 根证书是否已被系统信任(Word 的 WKWebView 会严格校验)
CA_ROOT="$(mkcert -CAROOT 2>/dev/null || true)"
TRUSTED=0
if [ -n "$CA_ROOT" ] && [ -f "$CA_ROOT/rootCA.pem" ]; then
  if security verify-cert -c "$CA_ROOT/rootCA.pem" >/dev/null 2>&1; then TRUSTED=1; fi
fi

if [ "$TRUSTED" -ne 1 ]; then
  cat <<'MSG'

⚠️  mkcert 根证书尚未被系统信任。Word 会直接拒绝加载 https://localhost:5173。

请执行下面这条命令(会要求输入 Mac 登录密码),然后重新运行 npm run addin:

    sudo "$(command -v mkcert)" -install

MSG
  exit 1
fi
echo "✓ 本地证书已受信任"

# 3. 写入 manifest
mkdir -p "$WEF"
if [ -f "$WEF/manifest.xml" ]; then
  BAK="$WEF/manifest.xml.bak.$(date +%Y%m%d%H%M%S)"
  cp "$WEF/manifest.xml" "$BAK"
  echo "✓ 已备份旧 manifest → $BAK"
fi
cp "$MANIFEST" "$WEF/manifest.xml"
echo "✓ manifest 已写入 $WEF/manifest.xml"

cat <<'MSG'

接下来:
  1) npm run dev        # 保持这个终端窗口开着
  2) 完全退出 Word(⌘Q,确认 Dock 图标下没有小圆点)再重新打开
  3) 插入 → 我的加载项 → 在「开发人员加载项」里双击 WordAI
  4) 选中一段文字,在侧栏发指令

MSG
