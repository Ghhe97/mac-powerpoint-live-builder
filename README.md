# Mac PowerPoint Live Builder

**默认中文** | [English version](#english-version)

让 AI Agent 在 macOS 上直接控制 Microsoft PowerPoint：你提出需求，Agent 在 PowerPoint 窗口里逐步生成幻灯片，最后交付可编辑的 `.pptx`、单页预览图和总览图。

这个仓库是一个完整安装包，包含：

- Codex/Agent skill：告诉 Agent 如何规划、生成和校验 PPT。
- 本地 PowerPoint MCP server：真正通过 AppleScript 控制 Mac PowerPoint。
- 一键安装脚本：复制 skill、安装 MCP、写入 Codex 配置、验证工具可用。

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
5. 重启 Codex 或你的 Agent 产品。
6. 第一次控制 PowerPoint 时，macOS 会询问是否允许自动化控制 PowerPoint，请点击允许。

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

## 其它 Agent 产品怎么用

这个仓库的核心是本地 stdio MCP server，所以不只适用于 Codex。

如果你的 Agent 产品支持 stdio MCP，可以运行：

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py
```

然后把终端输出的 MCP 配置填入对应产品的 MCP 设置。

通用配置形态是：

```text
transport: stdio
command: ~/.local/share/powerpoint-live-mcp/.venv/bin/powerpoint-live-mcp
```

如果产品支持环境变量，建议配置：

```text
PATH=/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin
```

## 仓库结构

```text
mac-powerpoint-live-builder/
  README.md
  install.command
  skill/
    mac-powerpoint-live-builder/
      SKILL.md
      agents/
      references/
      scripts/
        install_mcp.py
        check_pptx_mcp.py
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
- A one-command installer that copies the skill, installs the MCP server, writes Codex config, and verifies the tool inventory.

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
5. Restart Codex or your Agent app.
6. On first use, allow macOS Automation permission to control Microsoft PowerPoint.

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

## Other Agent Products

The core runtime is a local stdio MCP server, so it can work outside Codex if your Agent product supports stdio MCP.

Run:

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py
```

Then copy the printed MCP command into your Agent's MCP settings.

Generic MCP settings:

```text
transport: stdio
command: ~/.local/share/powerpoint-live-mcp/.venv/bin/powerpoint-live-mcp
```

Recommended `PATH`:

```text
PATH=/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin
```

## Repository Layout

```text
mac-powerpoint-live-builder/
  README.md
  install.command
  skill/
    mac-powerpoint-live-builder/
      SKILL.md
      agents/
      references/
      scripts/
        install_mcp.py
        check_pptx_mcp.py
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
