# WordAI

> Mac Word 右侧栏的 AI 写作助手。**Pi 当内核**(模型/规划/多轮),**Office.js 当手**(读写文档),本地 Node 服务托管。

**AI 只能提交草稿,必须你勾选后点「替换原文」才会写入文档。** 没有静默改写。

## 能力

| 能力 | 说明 |
|---|---|
| 读全文结构 | 标题层级 + 段落数 + 字数 |
| 读正文 / 检索 | 按段落号读,单次 ≤6000 字;全文检索带上下文 |
| 选区改写 | 逐处勾选后写入,替换 / 批注 / 插入下方三选一 |
| 批注审阅 | Word 原生批注,不改正文 |
| 结构化插入 | 标题 / 段落 / 列表 / 表格 |
| 全文审查 | 清单 + 跳转 + 一键转批注 |
| 参考文献 | 拖文件进侧栏,抽文本建索引,模型按需读原文(含 PDF) |
| 一键撤销 | 改写 / 批注 / 插入三类,撤销前回读校验 |
| 多会话 | 1 个当前 + 历史池,切换时 Pi 上下文一起切回 |

## 环境要求

| 依赖 | 要求 | 说明 |
|---|---|---|
| 系统 | **macOS** | 只支持 Mac 版 Word。Windows 的 wef 路径和启动方式都不同,当前不支持 |
| Word | Office 365 / 2019+ | 需支持加载项(插入 → 我的加载项) |
| Node.js | >= 20 | |
| [mkcert](https://github.com/FiloSottile/mkcert) | 任意版本 | `brew install mkcert nss`,本地 https 证书 |
| [Pi](https://pi.dev) | 任意版本 | 模型和 API Key 全交给 Pi,**WordAI 自己不存任何密钥** |

## 快速开始

```bash
# 1. 克隆到任意目录
git clone https://github.com/xylinwc/WordAI.git
cd WordAI

# 2. 装 Pi 并配好模型凭证(只需一次)
curl -fsSL https://pi.dev/install.sh | sh
pi                      # 首次运行按提示登录 / 选模型

# 3. 一键安装自检:证书 → 依赖 → 凭证 → 侧载 manifest
npm run setup

# 4. 启动服务(保持窗口开着)
npm run dev             # 前端 https://localhost:5173 + 代理 http://localhost:8787
```

如果 setup 提示根证书未受信任,单独执行一次(要输 Mac 登录密码)后重跑:

```bash
sudo "$(command -v mkcert)" -install
npm run setup
```

## 在 Word 里用

1. 完全退出 Word(⌘Q,确认 Dock 图标下没有小圆点)再打开
2. 插入 → 我的加载项 →「开发人员加载项」里双击 **WordAI**
3. 在文档里选中要改的文字,侧栏发指令
4. 检查草稿卡片 → 勾选 → 点「替换原文」

模型可在输入坞下拉切换,只影响 WordAI,不改 Pi CLI 的默认模型。

## 常驻运行(可选)

不想每次手动开终端:

```bash
npm run service:install   # 交给 launchd,登录后自启
npm run service:status    # 看状态 + 端口 + 探测
npm run service:logs      # 跟踪日志
npm run service:uninstall # 卸载
```

日志在 `~/Library/Logs/WordAI/dev.log`。**装了 launchd 就别再 `npm run dev`**,端口会冲突。

## 原理

```
Word 侧栏 (Office.js, https://localhost:5173)
   ├─ 读选区 / 写回选区            ← 全部在本机,不经网络
   └─ /api/agent/run + /stream ──► Node 代理 (8787) ──► Pi 内核
                      ◄──────────────  SSE 流式事件
```

「先预览后确认」由工具设计保证:模型手里没有任何写入能力,只能调 `propose_*` 产出草稿;
真正的 `insertText` 只在你勾选并点「应用选中 N 处」时执行。

定位策略:每处改动带 `paragraphIndex + find` 双信息,写入前校验 `find` 确实存在于该段落,
不一致就拒写 —— 索引会因编辑漂移,文本不会。

## 命令

| 命令 | 作用 |
|---|---|
| `npm run setup` | 一键安装自检(证书 / 依赖 / 凭证 / 侧载) |
| `npm run dev` | 前端 + 代理一起跑 |
| `npm run dev:web` / `dev:api` | 只跑前端 / 只跑代理 |
| `npm run addin` | 重新侧载 manifest |
| `npm run selftest` | 接口自测:服务、会话、附件、批量删(不耗额度) |
| `npm run build` / `npm start` | 生产构建 / 单端口托管 dist |

## 排错

| 现象 | 处理 |
|---|---|
| 侧栏白屏 / 加载不出来 | 确认服务在跑,浏览器打开 https://localhost:5173 无证书警告 |
| 提示「请先选中文字」 | 选区必须落在正文里,不能在页眉 / 文本框 |
| 提示「选区已变化」 | 生成期间你点过别处,重新选中原文再点替换 |
| 模型调用失败 | 跑 `pi` 确认已登录;WordAI 只读 `~/.pi/agent/` 里的配置 |
| 端口 5173 / 8787 被占 | `lsof -nP -iTCP:5173 -sTCP:LISTEN`,停掉占用进程再启动 |
| 改了代码侧栏没变化 | 面板还在跑旧 bundle,重启 `npm run dev`;Word 里缓存的按 ⌘Q 重开 |
| 改完 manifest 不生效 | ⌘Q 完全退出 Word;必要时删 `wef/` 里的旧 manifest 再 `npm run addin` |

## 已知限制

- **只支持 macOS + Mac 版 Word**,Windows 未适配
- 端口固定 5173 / 8787(manifest 里写死了 5173),改端口需要同步改 `manifest.xml`
- 整段替换会丢段内上标 / 加粗(Office.js 的 `Range.search()` 搜索串上限 255 字)
- 未保存的文档不落盘会话(没有稳定身份)
- 会话历史上限 30 个,满了不自动删,需要手动清理

## 目录结构

```
manifest.xml          Word 加载项清单(端口 / UUID / 权限)
server/               Node 代理:Pi 内核桥接、路由、会话、参考文献
src/                  侧栏前端(React + Office.js)
scripts/              安装、侧载、launchd、自测
public/assets/        侧栏与加载项图标
certs/                本地 https 证书(gitignored,setup 自动生成)
```

> 侧载 UUID 固定为 `731b28f8-81ea-4b51-bc31-c6fb1975ca19`。与别人已装的加载项冲突时,
> 改 `manifest.xml` 里的 `<Id>` 即可。

## 隐私

所有文档内容、附件、会话记录都留在本机(`~/.wordai/`)。只有发给模型的文本会离开你的电脑,
具体去哪个 provider 由你自己的 Pi 配置决定。仓库里不含任何密钥,`.env` 和 `certs/` 已 gitignore。

## License

[MIT](LICENSE)
