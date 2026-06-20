# Mac PowerPoint Live Builder

**默认中文** | [English version](#english-version)

让 AI Agent 在 macOS 上直接控制 Microsoft PowerPoint：你提出需求，Agent 在 PowerPoint 窗口里逐步生成幻灯片，最后交付可编辑的 `.pptx`、单页预览图和总览图。

这个仓库是一个完整安装包，包含：

- Codex/Agent skill：告诉 Agent 如何规划、生成和校验 PPT。
- 本地 PowerPoint MCP server：真正通过 AppleScript 控制 Mac PowerPoint。
- 一键安装脚本：复制 skill、安装 MCP、写入 Codex 配置，并打印通用 MCP/WorkBuddy 配置示例。
- WorkBuddy 桥接模式：当 Agent 沙盒拦截 PowerPoint 自动化时，通过本机 localhost bridge 继续实现 live 控制。

## 适合谁

适合想要这种体验的人：

```text
在 Agent 里输入：
“做一个三页新能源汽车研究 PPT，信息密度特别高”

然后看到：
PowerPoint 窗口里一页一页、一个元素一个元素生成，
最终得到可继续人工编辑的 .pptx 文件。
```

## 系统要求

- macOS
- Microsoft PowerPoint for Mac
- Python 3.10 或更高版本
- 一个支持 skill / MCP 工作流的 Agent 产品
- 推荐安装 Homebrew `poppler`，用于导出缩略图：

```bash
brew install poppler
```

如果你在 Codex 环境中使用，安装器会优先尝试使用 Codex 自带的 Python 和 `pdftoppm`。

## 快速安装

### 方式一：下载 ZIP

1. 在 GitHub 页面点击 **Code -> Download ZIP**。
2. 解压 ZIP。
3. 双击 `install.command`。
4. 按终端提示操作。
5. 如果你用 Codex，重启 Codex；如果你用 WorkBuddy 或其它 Agent，把终端打印的 MCP JSON 配置复制到该产品的 MCP 设置后重启。
6. 第一次控制 PowerPoint 时，macOS 会询问是否允许自动化控制 PowerPoint，请点击允许。

WorkBuddy 如果已经打开自动化权限但仍然报 `-10004`，请使用下面的桥接模式。

### 方式二：命令行安装

```bash
git clone https://github.com/Ghhe97/mac-powerpoint-live-builder.git
cd mac-powerpoint-live-builder
./install.command
```

安装完成后，在 Codex 中可以这样使用：

```text
Use $mac-powerpoint-live-builder 做一个三页新能源汽车研究 PPT，信息密度特别高。
```

也可以直接说：

```text
做一个三页新能源汽车研究 PPT，要求在 PowerPoint 窗口里边生成边可见。
```

## 安装器会做什么

`install.command` 会自动执行：

1. 把 `skill/mac-powerpoint-live-builder` 复制到：

   ```text
   ~/.codex/skills/mac-powerpoint-live-builder
   ```

2. 运行 skill 内置安装器：

   ```text
   ~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py
   ```

3. 创建 MCP 虚拟环境：

   ```text
   ~/.local/share/powerpoint-live-mcp/.venv
   ```

4. 安装 MCP 运行依赖。
5. 创建本地 `powerpoint-live-mcp` 启动脚本。
6. 写入或更新 Codex MCP 配置：

   ```text
   ~/.codex/config.toml
   ```

7. 验证 MCP server 暴露了 `pptx_*` 工具。
8. 打印通用 stdio MCP JSON 和 WorkBuddy `mcp.json` server block 示例。

注意：工具列表验证只说明 MCP server 能启动、能暴露工具；不等于当前 Agent 会话已经挂载了这些工具。很多 Agent 产品需要重启后才会加载新的 MCP 配置。

## 其它 Agent 产品怎么用

这个仓库的核心是本地 stdio MCP server，所以不只适用于 Codex。

如果你的 Agent 产品支持 stdio MCP，可以运行：

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --print-config
```

然后把终端输出的 MCP 配置填入对应产品的 MCP 设置并重启 Agent。

通用配置形态是：

```text
transport: stdio
command: ~/.local/share/powerpoint-live-mcp/.venv/bin/powerpoint-live-mcp
```

如果产品支持环境变量，建议配置：

```text
PATH=/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin
```

WorkBuddy 风格的 `mcp.json` 通常需要类似下面的 server block：

```json
"powerpoint-live-mcp": {
  "command": "/Users/YOU/.local/share/powerpoint-live-mcp/.venv/bin/powerpoint-live-mcp",
  "env": {
    "PATH": "/Users/YOU/.local/share/powerpoint-live-mcp/.venv/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
  }
}
```

### WorkBuddy 桥接模式

我们在真实测试里发现：有些 WorkBuddy 会把 MCP 放进更严格的执行沙盒。此时 macOS 自动化列表里虽然显示 WorkBuddy 已允许控制 PowerPoint，但 MCP 子进程仍可能被拦截，表现为 `-10004` / 权限违例。

解决办法是启动一个本机 bridge。它只监听 `127.0.0.1`，用 token 文件保护；WorkBuddy MCP 只连这个本地 bridge，真正执行 AppleScript 的进程在沙盒外。

先开一个终端窗口并保持它运行：

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/start_bridge.command
```

启动时会先跑一个很短的 PowerPoint Automation self-test。如果失败，窗口里会直接提示应该给哪个启动器授权。

注意：因为这个 bridge 是从终端启动的，所以 macOS 里还需要允许“终端”控制 Microsoft PowerPoint：

```text
系统设置 -> 隐私与安全性 -> 自动化 -> 终端 -> Microsoft PowerPoint
```

再写入 WorkBuddy bridge 配置：

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --write-workbuddy-config --bridge-mode
```

然后重启 WorkBuddy。验证：

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --doctor --smoke-powerpoint --bridge-mode
```

如果这个验证通过，才说明不是只“列出了工具”，而是真的可以 live 控制 PowerPoint。

如果 bridge 模式报 `osascript timed out after 60s`，通常就是上面的“终端 -> Microsoft PowerPoint”权限还没打开。打开后重启 bridge，再跑验证。

如果要确认不只是“能列出工具”，还是真的能控制 PowerPoint，可以运行：

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --doctor --smoke-powerpoint
```

这个命令会通过 MCP 创建并关闭一个很小的 PowerPoint 演示文稿。

## 仓库结构

```text
mac-powerpoint-live-builder/
  README.md
  install.command
  start_bridge.command
  skill/
    mac-powerpoint-live-builder/
      SKILL.md
      agents/
      references/
      scripts/
        install_mcp.py
        check_pptx_mcp.py
        powerpoint_bridge.py
        start_bridge.command
      vendor/
        powerpoint-live-mcp/
```

其中：

- `SKILL.md`：Agent 使用说明和工作流。
- `references/workflow.md`：PPT live 生成、排版和 QA 规范。
- `references/install.md`：MCP 安装、自检和其它 Agent 配置说明。
- `scripts/install_mcp.py`：安装 MCP server。
- `vendor/powerpoint-live-mcp`：随 skill 附带的本地 MCP server 代码。

## 常见问题

### 需要 API Token 吗？

不需要。这个 MCP 是本地 stdio 服务，不需要 API token。

### MCP 从哪里下载？

MCP 代码已经内置在仓库里：

```text
skill/mac-powerpoint-live-builder/vendor/powerpoint-live-mcp
```

安装时不需要从其它 GitHub 仓库下载 MCP。只有 Python 运行依赖可能需要联网安装。

### 为什么要重启 Agent？

很多 Agent 产品只在启动时读取 MCP 配置。安装器写完配置后，需要重启 Agent 才能看到新的 `pptx_*` 工具。

### 为什么第一次会弹权限？

macOS 会保护应用之间的自动化控制。第一次让 Agent 控制 PowerPoint 时，需要允许 Automation 权限。

### `install_mcp.py --check` 通过，但 Agent 里看不到 `pptx_*` 怎么办？

这说明 MCP server 可以启动，但当前 Agent 会话还没有挂载它。把 MCP 配置复制到 Agent 设置后重启 Agent。仅仅 `--check` 通过，不代表当前对话已经能调用 MCP 工具。

### 遇到 `-1708` 或 `"activate"不能继续` 怎么办？

这是 PowerPoint 拒绝前台激活，不一定是权限问题。最新版 MCP server 已把 `activate` 包在 `try/end try` 中；请更新后重新安装。

### 遇到 `-10004` 或权限违例怎么办？

打开 macOS：系统设置 -> 隐私与安全性 -> 自动化，确认启动 MCP 的 Agent 应用被允许控制 Microsoft PowerPoint。

如果你已经打开权限但 WorkBuddy 仍然失败，通常不是 PowerPoint 或 MCP 坏了，而是 WorkBuddy 的 MCP 子进程在沙盒里。请使用上面的 WorkBuddy 桥接模式。

如果桥接模式从终端启动后变成超时，请确认“终端”也被允许控制 Microsoft PowerPoint。

### 如果 live 失败，能不能先离线生成 PPTX？

可以，但必须明确标记为非 live fallback。不要把 `python-pptx` 或其它离线生成结果说成是在 PowerPoint 窗口里逐页生成的。

### PowerPoint 没打开可以用吗？

建议先打开 Microsoft PowerPoint。安装器可以安装 MCP，但真正生成 PPT 时需要 PowerPoint 可被 AppleScript 控制。

## 卸载

删除 skill：

```bash
rm -rf ~/.codex/skills/mac-powerpoint-live-builder
```

删除 MCP 环境：

```bash
rm -rf ~/.local/share/powerpoint-live-mcp
```

如需清理 Codex MCP 配置，打开：

```text
~/.codex/config.toml
```

删除由 `powerpoint-live-mcp` 管理的配置块。

---

## English Version

Mac PowerPoint Live Builder lets AI agents control Microsoft PowerPoint for Mac directly. You describe the deck, the Agent builds it inside the visible PowerPoint window, and the final output remains an editable `.pptx` with slide thumbnails for visual QA.

This repository is a complete distribution package. It includes:

- An Agent skill that defines the planning, generation, and QA workflow.
- A local PowerPoint MCP server that controls PowerPoint through AppleScript.
- A one-command installer that copies the skill, installs the MCP server, writes Codex config, and prints generic MCP/WorkBuddy config examples.
- WorkBuddy bridge mode for Agent sandboxes that block direct PowerPoint Automation.

## Who This Is For

Use this when you want an experience like:

```text
Ask your Agent:
"Create a dense three-slide research deck about new energy vehicles."

Then watch:
PowerPoint builds the deck slide by slide and element by element,
ending with an editable .pptx file.
```

## Requirements

- macOS
- Microsoft PowerPoint for Mac
- Python 3.10+
- An Agent product that supports skills and/or stdio MCP
- Recommended: Homebrew `poppler` for thumbnail export

```bash
brew install poppler
```

When running inside Codex, the installer will also try to use Codex's bundled Python and `pdftoppm` runtime.

## Quick Start

### Option 1: Download ZIP

1. Click **Code -> Download ZIP** on GitHub.
2. Unzip the package.
3. Double-click `install.command`.
4. Follow the terminal prompts.
5. Restart Codex. For WorkBuddy or another Agent, copy the printed MCP JSON into that product's MCP settings, then restart the Agent.
6. On first use, allow macOS Automation permission to control Microsoft PowerPoint.

If WorkBuddy still reports `-10004` after Automation permission is enabled, use bridge mode below.

### Option 2: Command Line

```bash
git clone https://github.com/Ghhe97/mac-powerpoint-live-builder.git
cd mac-powerpoint-live-builder
./install.command
```

Then in Codex:

```text
Use $mac-powerpoint-live-builder to create a dense three-slide research deck about new energy vehicles.
```

## What The Installer Does

`install.command` will:

1. Copy `skill/mac-powerpoint-live-builder` to:

   ```text
   ~/.codex/skills/mac-powerpoint-live-builder
   ```

2. Run the bundled MCP installer:

   ```text
   ~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py
   ```

3. Create an MCP virtual environment at:

   ```text
   ~/.local/share/powerpoint-live-mcp/.venv
   ```

4. Install MCP runtime dependencies.
5. Create the local `powerpoint-live-mcp` executable.
6. Write/update Codex MCP config:

   ```text
   ~/.codex/config.toml
   ```

7. Verify that the MCP server exposes the expected `pptx_*` tools.
8. Print generic stdio MCP JSON and a WorkBuddy-style `mcp.json` server block.

Tool inventory verification means the MCP server can start and list tools. It does not mean the current Agent session has mounted those tools. Many Agent apps require a restart after MCP config changes.

## Other Agent Products

The core runtime is a local stdio MCP server, so it can work outside Codex if your Agent product supports stdio MCP.

Run:

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --print-config
```

Then copy the printed MCP command or JSON into your Agent's MCP settings and restart the Agent.

Generic MCP settings:

```text
transport: stdio
command: ~/.local/share/powerpoint-live-mcp/.venv/bin/powerpoint-live-mcp
```

Recommended `PATH`:

```text
PATH=/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin
```

A WorkBuddy-style `mcp.json` server block usually looks like:

```json
"powerpoint-live-mcp": {
  "command": "/Users/YOU/.local/share/powerpoint-live-mcp/.venv/bin/powerpoint-live-mcp",
  "env": {
    "PATH": "/Users/YOU/.local/share/powerpoint-live-mcp/.venv/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
  }
}
```

### WorkBuddy Bridge Mode

Some WorkBuddy builds run MCP inside a stricter execution sandbox. In that case,
macOS Automation may show WorkBuddy as allowed, but the MCP subprocess can still
fail with `-10004`.

Bridge mode starts a local server outside the Agent sandbox. It binds only to
`127.0.0.1` and uses a token file. WorkBuddy's MCP process talks to the local
bridge; the bridge performs the AppleScript call.

Keep this running in a Terminal window:

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/start_bridge.command
```

Startup runs a short PowerPoint Automation self-test. If it fails, the bridge
window prints the launcher permission that needs to be enabled.

Because the bridge is launched by Terminal, macOS must also allow Terminal to
control Microsoft PowerPoint:

```text
System Settings -> Privacy & Security -> Automation -> Terminal -> Microsoft PowerPoint
```

Write WorkBuddy bridge config:

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --write-workbuddy-config --bridge-mode
```

Then restart WorkBuddy. Verify:

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --doctor --smoke-powerpoint --bridge-mode
```

Only treat live mode as working after this smoke test creates and closes a tiny
PowerPoint presentation.

If bridge mode reports `osascript timed out after 60s`, Terminal is usually not
authorized to control PowerPoint. Enable that checkbox, restart the bridge, and
run the smoke test again.

To verify real PowerPoint control, not just tool listing, run:

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --doctor --smoke-powerpoint
```

This creates and closes a tiny PowerPoint presentation through MCP.

## Repository Layout

```text
mac-powerpoint-live-builder/
  README.md
  install.command
  start_bridge.command
  skill/
    mac-powerpoint-live-builder/
      SKILL.md
      agents/
      references/
      scripts/
        install_mcp.py
        check_pptx_mcp.py
        powerpoint_bridge.py
        start_bridge.command
      vendor/
        powerpoint-live-mcp/
```

## FAQ

### Does this require an API token?

No. The MCP server is a local stdio service.

### Where does the MCP server come from?

It is bundled in this repository:

```text
skill/mac-powerpoint-live-builder/vendor/powerpoint-live-mcp
```

The installer does not download the MCP code from another GitHub repository. It may download Python runtime dependencies.

### Why restart the Agent app?

Many Agent apps only load MCP configuration at startup. Restart after installation so the new `pptx_*` tools are visible.

### Why does macOS ask for permission?

macOS protects app-to-app automation. The first time your Agent controls PowerPoint, allow Automation permission.

### `install_mcp.py --check` passes, but the Agent cannot see `pptx_*`

The MCP server can start, but the active Agent session has not mounted it. Add the MCP config to the Agent settings and restart the Agent.

### What does `-1708` or `"activate" can't continue` mean?

PowerPoint rejected foreground activation. Update to the latest MCP server; it wraps `activate` in `try/end try`.

### What does `-10004` or permission violation mean?

Open macOS System Settings -> Privacy & Security -> Automation, then allow the Agent app that launched MCP to control Microsoft PowerPoint.

If WorkBuddy still fails after that checkbox is enabled, use WorkBuddy bridge mode. The likely blocker is the MCP subprocess sandbox, not the PowerPoint MCP code itself.

If bridge mode is launched from Terminal and then times out, also allow Terminal to control Microsoft PowerPoint.

### Can I fall back to offline PPTX generation?

Yes, but label it as a non-live fallback. Do not claim the deck was visibly built in PowerPoint unless MCP/AppleScript actually controlled the running app.

## Uninstall

Remove the skill:

```bash
rm -rf ~/.codex/skills/mac-powerpoint-live-builder
```

Remove the MCP environment:

```bash
rm -rf ~/.local/share/powerpoint-live-mcp
```

Then remove the `powerpoint-live-mcp` block from:

```text
~/.codex/config.toml
```
