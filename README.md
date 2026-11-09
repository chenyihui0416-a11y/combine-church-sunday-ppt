# Sunday PPT Skill

自动组合中文教会主日 PowerPoint，并保留敬拜、讲道及回应诗歌的原生页面。项目采用通用 `SKILL.md` 结构，可供 Codex 和 Claude Code 使用。

## 支持范围

| 环境 | 状态 | PowerPoint 自动化 |
| --- | --- | --- |
| macOS + Microsoft PowerPoint | 已测试 | AppleScript |
| Windows + Microsoft PowerPoint | 已实现 | PowerShell COM |
| Codex | 支持 | `~/.codex/skills/` |
| Claude Code | 支持 | `~/.claude/skills/` 或项目 `.claude/skills/` |

Linux 和仅安装 LibreOffice 的电脑不支持原生 PowerPoint 复制，因为无法保证动画、主题、图片效果和背景完全保留。

## 安装

需要 Python 3.10 或以上版本及桌面版 Microsoft PowerPoint。

```bash
python install.py --target both
```

更新已有安装：

```bash
python install.py --target both --force
```

Claude Code 也可以安装到单个项目：

```bash
python install.py --target claude --project /path/to/project
```

处理 PDF 报告图片时需要 Poppler 的 `pdftoppm`；安装 Pillow 后会自动进行与既有模板一致的裁切。没有 PDF 时不需要这两个额外工具。

## 每周使用

把 template、经文 DOCX、四首歌的敬拜 PPTX、bulletin DOCX、可选讲道 PPTX及报告图片放在同一文件夹，然后在 Codex 或 Claude Code 中要求：

> 使用 combine-church-sunday-ppt 制作本周主日 PPT。

也可直接运行：

```bash
python skill/scripts/combine_sunday_ppt.py /weekly/folder \
  --output /weekly/folder/SundayPPT_YYYYMMDD_AM.pptx
```

为了保护教会资料，本仓库和安装后的 skill **不包含任何 PPTX**。模板必须保存在每周工作文件夹中，文件名包含 `template`；也可以用 `--template /private/path/template.pptx` 明确指定。请勿把真实模板、周报、敬拜、讲道或成品上传到本仓库。

脚本自动选择 macOS 或 Windows 后端，并在输出中报告每个阶段的耗时。正常流程只需运行一次；只有文件识别不正确时才先运行 `input_manifest.py` 检查。

## 基准测试

2026-11-08 在 macOS PowerPoint 上，以 54 页成品测试：

| 版本 | 总耗时 |
| --- | ---: |
| 优化前 | 26.48 秒 |
| 优化后首次运行 | 14.21 秒 |
| 经文缓存后的运行 | 12.67 秒 |

主要优化是避免逐页重复打开 PowerPoint、避免重复扫描同一 PPTX，并缓存不变的经文内容。具体时间会随 PowerPoint 版本、文件大小和电脑性能变化。

## 共同维护

排版规则和输入识别的修改方法见 `skill/references/customization.md`。提交修改前运行：

```bash
python -m unittest discover -s tests
python skill/scripts/input_manifest.py /path/to/sample-folder
```

涉及原生页面复制的修改必须在真实 PowerPoint 中完成端到端测试，并保留生成后的内部一致性验证。
