# waytouniverse-skills

开源 AI 技能合集，支持 **Claude Code** / **OpenClaw (QClaw)** / **OpenAI Codex** 三种 AI 编码工具。

## 技能列表

| 技能 | 说明 | 目录 |
|------|------|------|
| [PPT 生成器](./ppt-generator) | 根据 PPT 大纲，用 AI 逐页生成视觉风格统一的 PPT 图片 | `ppt-generator/` |
| [公文格式转换](./chinese-gov-docx) | 按照 GB/T 9704 标准生成中国党政机关公文格式 Word 文档 | `chinese-gov-docx/` |

## 安装方法

根据你使用的 AI 工具，将技能文件夹复制到对应目录：

```bash
# Claude Code
cp -r <技能目录> ~/.claude/skills/

# OpenClaw (QClaw)
cp -r <技能目录> ~/.agents/skills/

# OpenAI Codex
cp -r <技能目录> ~/.codex/skills/
```

例如安装 PPT 生成器：

```bash
# Claude Code
cp -r ppt-generator ~/.claude/skills/
```

也支持符号链接：

```bash
ln -s /你的路径/waytouniverse-skills/ppt-generator ~/.claude/skills/ppt-generator
```

## 安装路径速查

| 工具 | 技能目录 |
|------|----------|
| Claude Code | `~/.claude/skills/` |
| OpenClaw (QClaw) | `~/.agents/skills/` |
| OpenAI Codex | `~/.codex/skills/` |

## 许可证

MIT
