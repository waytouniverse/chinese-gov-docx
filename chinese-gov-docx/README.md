---
name: official-docx
description: "Generate official government-style Word documents (.docx) with strict formatting requirements. Trigger ONLY when user explicitly says '输出公文版本的word' or '生成公文格式word' or similar phrases indicating need for official document format. This skill creates properly formatted official documents following Chinese government document standards with specific fonts, spacing, and layout requirements."
trigger: "输出公文版本的word|生成公文格式word|生成公文word|转成公文word|按公文格式生成word|转成公文格式word"
license: Proprietary
---

# Official Government Document Word Generator

Generate Chinese government-style Word documents with strict formatting compliance.

## When to Use

**TRIGGER ONLY** when user explicitly requests:
- "输出公文版本的word"
- "生成公文格式word"
- "生成公文word"
- "转成公文word"
- "按公文格式生成word"
- "转成公文格式word"

## Document Format Specifications

### Font Requirements
- **主标题**: 方正小标宋简体, 二号 (22pt)，居中，无首行缩进
- **一级标题**: 黑体, 三号 (16pt) - 格式：一、二、三、，首行缩进2字符
- **二级标题**: 楷体_GB2312, 三号 (16pt) - 格式：（一）（二）（三），首行缩进2字符
- **三级标题**: 仿宋_GB2312 + Times New Roman (数字), 三号 (16pt) - 格式：1.2.3.，首行缩进2字符
- **四级标题**: 仿宋_GB2312 + Times New Roman (数字), 三号 (16pt) - 格式：（1）（2）（3），首行缩进2字符
- **正文中文**: 仿宋_GB2312, 三号，首行缩进2字符
- **正文数字/字母**: Times New Roman, 三号，首行缩进2字符
- **表格内容**: 仿宋_GB2312 (中文) + Times New Roman (数字/字母), 三号，无首行缩进（表格整体居中）
- **日期**: 仿宋_GB2312, 三号，居中显示，无首行缩进

### Page Layout
- **上页边距**: 3.7cm
- **下页边距**: 3.5cm
- **左页边距**: 2.8cm
- **右页边距**: 2.6cm
- **行距**: 固定值28.7磅
- **正文对齐**: 两端对齐
- **首行缩进**: 2字符
- **标题**: 首行无缩进，左对齐
- **日期**: 居中对齐

### Text Processing Rules
1. 全文不加粗（标题除外）
2. 清除数字、字母前后空格
3. 清除各类奇怪符号#-|
4. 双引号转为全角汉字双引号""（U+201C左引号，U+201D右引号）
5. 标题序号后使用顿号或句点，序号后不留空格
6. 反斜杠转义字符还原（如1\. → 1.）

## Usage

### Prerequisites
Ensure `docx` npm package is available:
```bash
npm list docx || npm install docx
```

### Input Requirements
User must provide:
1. Document title (主标题)
2. Document content with clear heading structure
3. Output file path

### Heading Hierarchy Rules
**IMPORTANT**: The skill automatically normalizes heading levels to ensure strict hierarchy:

The heading order must be: **一、 → （一） → 1. → （1）**

If the input text violates this hierarchy, the skill automatically adjusts:
- If only **（一）** exists without **一、**, converts **（一）** to **一、**
- If only **1.** exists without **（一）**, converts **1.** to **（一）** (or to **一、** if no level-1 headings exist)
- If only **（1）** exists without **1.**, promotes it to the appropriate level

This rule applies to **ALL** numbered items in the document, including those in body text.

### Content Format
Provide content with clear heading markers:
```
一、一级标题内容
（一）二级标题内容
1.三级标题内容
正文内容...

（二）二级标题内容
正文...

二、另一个一级标题
...
```

### Markdown Table Support
Tables in Markdown format are automatically converted to Word tables with proper formatting:
```
| 列1 | 列2 | 列3 |
|-----|-----|-----|
| 数据1 | 数据2 | 数据3 |
```

### Special Content Handling

#### 目录处理
文档中的"目 录"或"目录"部分会被识别，目录内的所有内容（包括标题序号）都作为正文处理，使用仿宋_GB2312字体，不按标题格式渲染。

#### 日期居中
格式为"2026年02月"或"2026年1月"的日期行会自动居中显示。

#### 主标题去重
如果文档第一行包含"规划|报告|方案|意见|通知|决定|决议|纲要|计划"等关键词，会自动跳过第一行，避免主标题重复。

#### 表格字体
表格内文字自动切换字体：中文使用仿宋_GB2312，数字和字母使用Times New Roman。

### Example Invocation
```
用户: 输出公文版本的word，标题是"关于XXX的报告"，内容如下：
一、基本情况
（一）工作进展
1.已完成事项
具体内容...
保存到 /path/to/output.docx
```

### Output
Generates a properly formatted .docx file at the specified path with:
- Correct fonts for each element type
- Proper heading hierarchy
- Official document spacing and alignment
- Clean text processing (quotes, spaces)
- Tables with proper font switching
- Centered dates
- Catalog content as body text

## Implementation

This skill uses the `docx` Node.js library to generate OOXML documents programmatically.

Key implementation details:
- Font switching for Chinese/English text using regex matching
- Precise paragraph spacing (DXA units)
- Proper indentation (twips)
- Line spacing calculations (28.7磅 = 574 twips)
- Chinese quote conversion using U+201C (left) and U+201D (right)

### Core Functions

```javascript
// Creates main title with 方正小标宋简体, 二号 font
function createMainTitle(text)

// Creates centered paragraph for dates
function createCenteredParagraph(text)

// Heading level functions with specific fonts
function createHeading1(text)  // 黑体
function createHeading2(text)  // 楷体_GB2312
function createHeading3(text)  // 仿宋_GB2312 + Times New Roman
function createHeading4(text)  // 仿宋_GB2312 + Times New Roman

// Body paragraph with automatic font switching
function createBodyParagraph(text)

// Table cell with automatic font switching
function createTableCellChildren(text)

// Text processing pipeline
function processText(text)

// Chinese quote conversion - uses U+201C and U+201D
function toChineseQuotes(str)

// Reads file with special handling for catalog and dates
function readContentFromFile(filePath)
```

### Key Processing Logic

1. **Catalog Detection**: Detects "目 录" or "目录" and marks content within catalog section with `__CATALOG__` prefix
2. **Date Detection**: Detects date format `YYYY年MM月` and marks with `__CENTER__` prefix
3. **Title Deduplication**: Skips first line if it contains title keywords to avoid duplicate main title
4. **Heading Normalization**: Ensures strict hierarchy from level 1 to level 4
5. **Font Switching**: Uses regex `/([一-龥　-〿＀-￯]+)|([a-zA-Z0-9\.\/:]+)|(.)/g` to switch between Chinese and English fonts
6. **Space Cleaning**: Removes spaces after serial numbers (一、, （一）, 1., etc.)
7. **Quote Conversion**: Converts `"` to `"` (U+201C) and `"` (U+201D) alternately

## Changelog

### Latest Updates
- **Optimized table row height** - Added vertical centering and paragraph spacing (before/after: 100 twips) for better readability
- **Added first line indent to all heading levels** (一、, （一）, 1., （1）) - all headings now have 2-character first line indent
- Fixed duplicate main title issue by skipping first line if it contains title keywords
- Added date centering support for `YYYY年MM月` format
- Fixed Chinese quote characters to use correct Unicode (U+201C/U+201D)
- Added table font switching (Chinese: 仿宋_GB2312, Numbers: Times New Roman)
- Fixed catalog content to render as body text instead of headings
- Added automatic space removal after serial numbers
- Improved Markdown table to Word table conversion
