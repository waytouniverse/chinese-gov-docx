#!/usr/bin/env node

/**
 * 公文Word文档生成器
 * 按照中国党政机关公文格式标准生成Word文档
 */

const { Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableCell, TableRow, WidthType, VerticalAlign } = require('docx');
const fs = require('fs');

// 字号定义：二号=22pt=44半点，三号=16pt=32半点
const SIZE_ERHAO = 44;
const SIZE_SANHAO = 32;

// 行距：28.7磅 = 574 twips
const LINE_SPACING = 574;

// 首行缩进2字符 ≈ 640 twips
const FIRST_LINE_INDENT = 640;

// 公文页边距（厘米转DXA，1cm ≈ 567）
const marginTop = Math.round(3.7 * 567);
const marginBottom = Math.round(3.5 * 567);
const marginLeft = Math.round(2.8 * 567);
const marginRight = Math.round(2.6 * 567);

/**
 * 将英文双引号转为中文双引号
 * 成对的英文双引号转换为中文左右引号：”文本” -> “文本”
 */
function toChineseQuotes(str) {
  let result = str;
  let count = 0;
  // 先处理英文双引号 “ (U+0022)
  result = result.replace(/"/g, () => {
    count++;
    return count % 2 === 1 ? '“' : '”';  // 左引号 U+201C，右引号 U+201D
  });
  return result;
}

/**
 * 中文数字转阿拉伯数字
 */
const chineseToArabic = {
  '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
  '六': 6, '七': 7, '八': 8, '九': 9, '十': 10
};

const arabicToChinese = ['', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'];

function chineseNumToArabic(str) {
  if (/^[一二三四五六七八九十]$/.test(str)) {
    return chineseToArabic[str];
  }
  return parseInt(str) || 0;
}

function arabicToChineseNum(num) {
  if (num >= 1 && num <= 10) {
    return arabicToChinese[num];
  }
  return num.toString();
}

/**
 * 检测行的序号类型和值
 * 返回：{ type: 1|2|3|4|null, value: number, original: string }
 * type: 1=一、 2=（一） 3=1. 4=（1）
 */
function detectHeadingLevel(line) {
  const trimmed = line.trim();

  // 一级：一、二、三、
  const match1 = trimmed.match(/^([一二三四五六七八九十])、/);
  if (match1) {
    return { type: 1, value: chineseNumToArabic(match1[1]), original: match1[0] };
  }

  // 二级：（一）（二）（三）
  const match2 = trimmed.match(/^（([一二三四五六七八九十]+)）/);
  if (match2) {
    return { type: 2, value: chineseNumToArabic(match2[1]), original: match2[0] };
  }

  // 三级：1. 2. 3.
  const match3 = trimmed.match(/^(\d+)\./);
  if (match3) {
    return { type: 3, value: parseInt(match3[1]), original: match3[0] };
  }

  // 四级：（1）（2）（3）
  const match4 = trimmed.match(/^（(\d+)）/);
  if (match4) {
    return { type: 4, value: parseInt(match4[1]), original: match4[0] };
  }

  return null;
}

/**
 * 将序号提升一级
 * 2->1, 3->2, 4->3
 */
function promoteHeading(line, levelInfo) {
  const trimmed = line.trim();

  if (levelInfo.type === 2) {
    // （一）-> 一、
    return trimmed.replace(/^（[一二三四五六七八九十]+）/, arabicToChineseNum(levelInfo.value) + '、');
  } else if (levelInfo.type === 3) {
    // 1. -> （一）
    return trimmed.replace(/^\d+\./, '（' + arabicToChineseNum(levelInfo.value) + '）');
  } else if (levelInfo.type === 4) {
    // （1）-> 1.
    return trimmed.replace(/^（\d+）/, levelInfo.value + '.');
  }

  return line;
}

/**
 * 规范化序号层级
 * 确保序号严格按照 一、（一）1.（1）的顺序出现
 *
 * 核心逻辑：找出文档中实际使用的最高层级，然后将所有序号
 * 从该层级开始向下排列，保证层级连续性
 *
 * 注意：对于已经有正确公文格式的文档，不做任何修改
 */
function normalizeHeadings(contentLines) {
  // 检测所有行的序号信息
  const headingInfos = contentLines.map((line, index) => ({
    index,
    line,
    info: detectHeadingLevel(line)
  }));

  // 找出实际存在的层级（原始层级）
  const existingLevels = new Set();
  headingInfos.forEach(h => {
    if (h.info) {
      existingLevels.add(h.info.type);
    }
  });

  if (existingLevels.size === 0) {
    return contentLines;
  }

  // 找出最小层级（最高级）
  const minLevel = Math.min(...existingLevels);

  // 如果已经有正确的一级标题（一、），则不调整
  if (minLevel === 1) {
    return contentLines;
  }

  // 计算层级偏移量
  // 目标：从层级1开始排列
  const levelOffset = minLevel - 1;

  // 处理每一行
  return headingInfos.map(({ index, line, info }) => {
    if (!info) return line;

    let currentLine = line;
    let currentInfo = info;

    // 根据层级偏移量多次提升
    for (let i = 0; i < levelOffset; i++) {
      const newInfo = detectHeadingLevel(currentLine);
      if (!newInfo) break;

      currentLine = promoteHeading(currentLine, newInfo);
      currentInfo = detectHeadingLevel(currentLine);
    }

    return currentLine;
  });
}

/**
 * 清除数字、字母前后空格
 */
function cleanSpaces(str) {
  return str.replace(/\s*([a-zA-Z0-9])\s*/g, '$1');
}

/**
 * 清除 Markdown 格式标记
 * 处理：**粗体**、*斜体*、~~删除线~~、`代码`、[链接](url)、>引用、-列表、*列表等
 * 注意：保留看起来像公文标题的 "1. 标题内容" 格式（数字+点+空格+非空格字符）
 */
function cleanMarkdown(str) {
  // 处理分隔线 ***、---、___（独占一行或前后有空格的）
  str = str.replace(/^[*\-_]{3,}\s*$/gm, '');

  // 处理 **粗体** 和 __粗体__
  str = str.replace(/\*\*(.+?)\*\*/g, '$1');
  str = str.replace(/__(.+?)__/g, '$1');

  // 处理 *斜体* 和 _斜体_
  str = str.replace(/\*(.+?)\*/g, '$1');
  str = str.replace(/_(.+?)_/g, '$1');

  // 处理 ~~删除线~~
  str = str.replace(/~~(.+?)~~/g, '$1');

  // 处理 `行内代码`
  str = str.replace(/`(.+?)`/g, '$1');

  // 先处理 ![图片描述](url) -> 保留图片描述（必须在 [链接] 之前处理）
  str = str.replace(/!\[([^\]]*)\]\([^)]+\)/g, '$1');

  // 再处理 [链接文本](url) -> 保留链接文本
  str = str.replace(/\[([^\]]+)\]\([^)]+\)/g, '$1');

  // 处理 > 引用标记（行首）
  str = str.replace(/^>\s*/gm, '');

  // 处理列表标记 - 列表项、* 列表项、+ 列表项（行首）
  str = str.replace(/^[\-\*\+]\s+/gm, '');

  // 处理数字列表标记 1. 2. 等（行首）
  // 但保留看起来像公文标题的格式（如 "1. 工作内容"，数字+点+空格+中文字符）
  // 只删除纯数字列表（如 "1. " 后面没有内容或只有空格）
  str = str.replace(/^\d+\.\s*$/gm, '');

  // 处理 HTML 标签如 <br>, <br/>, <hr> 等
  str = str.replace(/<\/?[^>]+>/g, '');

  // 最后清理残留的 Markdown 符号（如单独的 *、**、_、__ 等）
  str = str.replace(/\*{1,2}/g, '');
  str = str.replace(/_{1,2}/g, '');

  return str;
}

/**
 * 处理文本，应用各类清理规则
 */
function processText(text) {
  // 先清理 Markdown 格式
  text = cleanMarkdown(text);
  // 再处理反斜杠转义（如 1\. -> 1.）- 顺序很重要！
  text = text.replace(/\\([\.\*\+\-\!\[\]\(\)\`\#\>])/g, '$1');
  text = toChineseQuotes(text);
  text = cleanSpaces(text);
  // 清除奇怪符号（不再清理 -，因为已在 cleanMarkdown 中处理）
  text = text.replace(/[#|]/g, '');
  return text;
}

/**
 * 创建正文段落（自动切换字体）
 */
function createBodyParagraph(text) {
  text = processText(text);

  const children = [];
  // 匹配中文、英文数字、其他字符
  const regex = /([一-龥　-〿＀-￯]+)|([a-zA-Z0-9\.\/:]+)|(.)/g;
  let match;

  while ((match = regex.exec(text)) !== null) {
    if (match[1]) {
      // 中文和中文标点 - 仿宋_GB2312
      children.push(new TextRun({
        text: match[1],
        font: "仿宋_GB2312",
        size: SIZE_SANHAO
      }));
    } else if (match[2]) {
      // 数字、字母、英文标点 - Times New Roman
      children.push(new TextRun({
        text: match[2],
        font: "Times New Roman",
        size: SIZE_SANHAO
      }));
    } else if (match[3]) {
      // 其他字符
      children.push(new TextRun({
        text: match[3],
        font: "仿宋_GB2312",
        size: SIZE_SANHAO
      }));
    }
  }

  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: LINE_SPACING, lineRule: "exact" },
    indent: { firstLine: FIRST_LINE_INDENT },
    children: children.length > 0 ? children : [new TextRun({ text: text, font: "仿宋_GB2312", size: SIZE_SANHAO })]
  });
}

/**
 * 创建一级标题（黑体）- 格式：一、二、三、
 */
function createHeading1(text) {
  text = processText(text);
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 200, after: 100, line: LINE_SPACING, lineRule: "exact" },
    indent: { firstLine: FIRST_LINE_INDENT },
    children: [new TextRun({ text: text, font: "黑体", size: SIZE_SANHAO })]
  });
}

/**
 * 创建二级标题（楷体_GB2312）- 格式：（一）（二）（三）
 */
function createHeading2(text) {
  text = processText(text);
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 150, after: 100, line: LINE_SPACING, lineRule: "exact" },
    indent: { firstLine: FIRST_LINE_INDENT },
    children: [new TextRun({ text: text, font: "楷体_GB2312", size: SIZE_SANHAO })]
  });
}

/**
 * 创建三级标题 - 格式：1.2.3.
 * 数字使用 Times New Roman，中文使用仿宋_GB2312
 */
function createHeading3(text) {
  text = processText(text);

  const children = [];
  // 匹配数字和点（如 "1." "2." "3."）、中文、其他字符
  const regex = /(\d+\.\s*)|([一-龥　-〿＀-￯]+)|([a-zA-Z]+)|(.)/g;
  let match;

  while ((match = regex.exec(text)) !== null) {
    if (match[1]) {
      // 数字和点（如 "1. "）- Times New Roman
      children.push(new TextRun({
        text: match[1],
        font: "Times New Roman",
        size: SIZE_SANHAO
      }));
    } else if (match[2]) {
      // 中文 - 仿宋_GB2312
      children.push(new TextRun({
        text: match[2],
        font: "仿宋_GB2312",
        size: SIZE_SANHAO
      }));
    } else if (match[3]) {
      // 英文字母 - Times New Roman
      children.push(new TextRun({
        text: match[3],
        font: "Times New Roman",
        size: SIZE_SANHAO
      }));
    } else if (match[4]) {
      // 其他字符
      children.push(new TextRun({
        text: match[4],
        font: "仿宋_GB2312",
        size: SIZE_SANHAO
      }));
    }
  }

  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 100, after: 50, line: LINE_SPACING, lineRule: "exact" },
    indent: { firstLine: FIRST_LINE_INDENT },
    children: children.length > 0 ? children : [new TextRun({ text: text, font: "仿宋_GB2312", size: SIZE_SANHAO })]
  });
}

/**
 * 创建四级标题 - 格式：（1）（2）（3）
 * 数字和括号使用 Times New Roman，中文使用仿宋_GB2312
 */
function createHeading4(text) {
  text = processText(text);

  const children = [];
  // 匹配括号内的数字（如 "（1）"）、中文、英文字母、其他字符
  const regex = /（\d+）|([一-龥　-〿＀-￯]+)|([a-zA-Z]+)|(.)/g;
  let match;

  while ((match = regex.exec(text)) !== null) {
    if (match[1]) {
      // 括号内的数字如 "（1）" - 整体处理，需要分别处理括号和数字
      const fullMatch = match[0]; // 完整匹配如 "（1）"
      // 括号用仿宋，数字用 Times New Roman
      children.push(new TextRun({
        text: "（",
        font: "仿宋_GB2312",
        size: SIZE_SANHAO
      }));
      children.push(new TextRun({
        text: match[1], // 数字部分
        font: "Times New Roman",
        size: SIZE_SANHAO
      }));
      children.push(new TextRun({
        text: "）",
        font: "仿宋_GB2312",
        size: SIZE_SANHAO
      }));
    } else if (match[2]) {
      // 中文 - 仿宋_GB2312
      children.push(new TextRun({
        text: match[2],
        font: "仿宋_GB2312",
        size: SIZE_SANHAO
      }));
    } else if (match[3]) {
      // 英文字母 - Times New Roman
      children.push(new TextRun({
        text: match[3],
        font: "Times New Roman",
        size: SIZE_SANHAO
      }));
    } else if (match[4]) {
      // 其他字符
      children.push(new TextRun({
        text: match[4],
        font: "仿宋_GB2312",
        size: SIZE_SANHAO
      }));
    }
  }

  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 50, after: 50, line: LINE_SPACING, lineRule: "exact" },
    indent: { firstLine: FIRST_LINE_INDENT },
    children: children.length > 0 ? children : [new TextRun({ text: text, font: "仿宋_GB2312", size: SIZE_SANHAO })]
  });
}

/**
 * 创建主标题（方正小标宋简体，二号）
 * 按照公文标准格式排版：
 * - 单行标题：左右留白，约空4-6个字符
 * - 多行标题：采用沙漏式/倒金字塔排列，中间对齐
 */
function createMainTitle(text) {
  text = processText(text);
  const titleLength = text.length;

  // 公文标题排版规则：
  // - 单行（≤20字）：左右各缩进约4-6个字符（约1280-1920 twips）
  // - 两行（21-28字）：上下行字数大致相等，上短下长或等长
  // - 三行及以上：采用倒金字塔或沙漏式

  const paragraphs = [];

  if (titleLength <= 28) {
    // 单行或短标题：左右缩进4个字符（1280 twips），让Word自动换行
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 300, line: LINE_SPACING, lineRule: "exact" },
      indent: { left: 1280, right: 1280 },
      children: [new TextRun({
        text: text,
        font: "方正小标宋简体",
        size: SIZE_ERHAO
      })]
    }));
  } else {
    // 三行及以上：采用沙漏式排列
    // 每行约14-16字，中间最长，上下稍短
    const charsPerLine = 16;
    const lines = [];

    // 简单分词断行
    let currentLine = '';
    for (let i = 0; i < text.length; i++) {
      currentLine += text[i];
      if (currentLine.length >= charsPerLine && i < text.length - 1) {
        // 检查下一个字是否适合作为行首
        const nextChar = text[i + 1];
        const noStartChars = ['的', '和', '与', '及', '或', '了', '在', '中', '对', '为'];
        if (!noStartChars.includes(nextChar)) {
          lines.push(currentLine);
          currentLine = '';
        }
      }
    }
    if (currentLine) {
      lines.push(currentLine);
    }

    // 如果最后一行太短，合并到前一行
    if (lines.length > 2 && lines[lines.length - 1].length < 8) {
      lines[lines.length - 2] += lines[lines.length - 1];
      lines.pop();
    }

    // 生成各行，采用沙漏式（中间行最长）
    lines.forEach((line, index) => {
      // 根据行号计算缩进，形成沙漏效果
      let leftIndent = 640; // 基础缩进2个字符
      if (lines.length > 2) {
        if (index === 0 || index === lines.length - 1) {
          leftIndent = 1280; // 首末行缩进4个字符
        } else if (index === Math.floor(lines.length / 2)) {
          leftIndent = 320; // 中间行缩进1个字符（最长）
        } else {
          leftIndent = 640; // 其他行缩进2个字符
        }
      }

      const isLast = index === lines.length - 1;
      paragraphs.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: {
          before: 0,
          after: isLast ? 300 : 0,
          line: LINE_SPACING,
          lineRule: "exact"
        },
        indent: { left: leftIndent, right: leftIndent },
        children: [new TextRun({
          text: line,
          font: "方正小标宋简体",
          size: SIZE_ERHAO
        })]
      }));
    });
  }

  return paragraphs;
}

/**
 * 创建右对齐段落（用于联系人信息）
 */
function createRightAlignedParagraph(text) {
  text = processText(text);
  return new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { line: LINE_SPACING, lineRule: "exact" },
    children: [new TextRun({ text: text, font: "仿宋_GB2312", size: SIZE_SANHAO })]
  });
}

/**
 * 创建居中段落（用于日期等）
 */
function createCenteredParagraph(text) {
  text = processText(text);
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: LINE_SPACING, lineRule: "exact" },
    children: [new TextRun({ text: text, font: "仿宋_GB2312", size: SIZE_SANHAO })]
  });
}

/**
 * 检测是否为 Markdown 表格行
 */
function isTableRow(line) {
  const trimmed = line.trim();
  return trimmed.startsWith('|') && trimmed.endsWith('|');
}

/**
 * 检测是否为表格分隔行（如 |---|---|）
 */
function isTableSeparator(line) {
  const trimmed = line.trim();
  if (!isTableRow(trimmed)) return false;
  const content = trimmed.slice(1, -1);
  return /^[\s\-:|]+$/.test(content) && content.includes('-');
}

/**
 * 解析 Markdown 表格行
 */
function parseTableRow(line) {
  const trimmed = line.trim();
  const content = trimmed.slice(1, -1);
  return content.split('|').map(cell => cell.trim());
}

/**
 * 创建表格单元格文本（自动切换字体）
 * 中文 - 仿宋_GB2312，数字/字母 - Times New Roman
 */
function createTableCellChildren(text) {
  const children = [];
  // 匹配中文、数字/字母/英文标点、其他字符
  const regex = /([一-龥　-〿＀-￯]+)|([a-zA-Z0-9\.\/:]+)|(.)/g;
  let match;

  while ((match = regex.exec(text)) !== null) {
    if (match[1]) {
      // 中文 - 仿宋_GB2312
      children.push(new TextRun({
        text: match[1],
        font: "仿宋_GB2312",
        size: SIZE_SANHAO
      }));
    } else if (match[2]) {
      // 数字、字母、英文标点 - Times New Roman
      children.push(new TextRun({
        text: match[2],
        font: "Times New Roman",
        size: SIZE_SANHAO
      }));
    } else if (match[3]) {
      // 其他字符
      children.push(new TextRun({
        text: match[3],
        font: "仿宋_GB2312",
        size: SIZE_SANHAO
      }));
    }
  }

  return children.length > 0 ? children : [new TextRun({
    text: text,
    font: "仿宋_GB2312",
    size: SIZE_SANHAO
  })];
}

/**
 * 创建表格
 */
function createTable(headers, rows) {
  const columnCount = headers.length;

  // 表格行高优化：单元格垂直居中 + 段落间距
  const tableCellSpacing = { before: 100, after: 100, line: 240, lineRule: "auto" };

  // 创建表头行 - 表格整体是正文，表头也使用仿宋_GB2312，支持自动切换字体
  const headerCells = headers.map(header => new TableCell({
    children: [new Paragraph({
      children: createTableCellChildren(header),
      alignment: AlignmentType.CENTER,
      spacing: tableCellSpacing
    })],
    shading: {
      fill: "F2F2F2"
    },
    verticalAlign: VerticalAlign.CENTER
  }));

  const headerRow = new TableRow({
    children: headerCells
  });

  // 创建数据行 - 自动切换字体
  const dataRows = rows.map(row => {
    const cells = row.map(cell => new TableCell({
      children: [new Paragraph({
        children: createTableCellChildren(cell),
        alignment: AlignmentType.CENTER,
        spacing: tableCellSpacing
      })],
      verticalAlign: VerticalAlign.CENTER
    }));
    return new TableRow({ children: cells });
  });

  // 计算列宽（均分）
  const columnWidth = 100 / columnCount;

  return new Table({
    rows: [headerRow, ...dataRows],
    width: {
      size: 100,
      type: WidthType.PERCENTAGE
    },
    columnWidths: Array(columnCount).fill(columnWidth)
  });
}

/**
 * 解析内容并生成段落
 */
function parseContent(contentLines) {
  const paragraphs = [];
  let i = 0;

  // 先规范化序号层级
  const normalizedLines = normalizeHeadings(contentLines);

  while (i < normalizedLines.length) {
    const line = normalizedLines[i];
    const trimmedLine = line.trim();

    // 检测是否为表格开始
    if (isTableRow(trimmedLine) && !isTableSeparator(trimmedLine)) {
      // 收集整个表格
      const tableLines = [];
      let j = i;

      // 收集表格所有行
      while (j < normalizedLines.length) {
        const currentLine = normalizedLines[j].trim();
        if (isTableRow(currentLine)) {
          if (!isTableSeparator(currentLine)) {
            tableLines.push(currentLine);
          }
          j++;
        } else {
          break;
        }
      }

      // 如果有表头行（至少2行：表头+分隔符+数据，或1行表头）
      if (tableLines.length >= 1) {
        const headers = parseTableRow(tableLines[0]);
        const rows = tableLines.slice(1).map(parseTableRow);
        paragraphs.push(createTable(headers, rows));
        i = j;
        continue;
      }
    }

    // 先清理 Markdown，然后检查是否为空
    const cleanedLine = cleanMarkdown(trimmedLine);
    if (!cleanedLine || cleanedLine.length === 0) {
      // 清理后为空，添加空段落保持行距
      paragraphs.push(new Paragraph({ children: [new TextRun("")] }));
      i++;
      continue;
    }

    // 检测标题级别（使用原始行，因为标题检测需要原始格式）
    // 但如果行以 __CATALOG__ 开头，表示是目录内容，作为正文处理
    if (trimmedLine.startsWith('__CATALOG__')) {
      // 目录内容作为正文，移除标记
      const catalogContent = trimmedLine.replace('__CATALOG__', '').trim();
      paragraphs.push(createBodyParagraph(catalogContent));
    } else if (trimmedLine.startsWith('__CENTER__')) {
      // 居中内容（如日期），移除标记并居中显示
      const centerContent = trimmedLine.replace('__CENTER__', '').trim();
      paragraphs.push(createCenteredParagraph(centerContent));
    } else if (/^[一二三四五六七八九十]+、/.test(trimmedLine)) {
      // 一级标题：一、二、三、
      paragraphs.push(createHeading1(trimmedLine));
    } else if (/^（[一二三四五六七八九十]+）/.test(trimmedLine)) {
      // 二级标题：（一）（二）（三）
      paragraphs.push(createHeading2(trimmedLine));
    } else if (/^\d+\./.test(trimmedLine)) {
      // 三级标题：1.2.3.
      paragraphs.push(createHeading3(trimmedLine));
    } else if (/^（\d+）/.test(trimmedLine)) {
      // 四级标题：（1）（2）（3）
      paragraphs.push(createHeading4(trimmedLine));
    } else {
      // 普通正文
      paragraphs.push(createBodyParagraph(trimmedLine));
    }

    i++;
  }

  return paragraphs;
}

/**
 * 生成公文Word文档
 */
async function generateDocument(title, contentLines, outputPath) {
  const children = [
    ...createMainTitle(title),
    ...parseContent(contentLines)
  ];

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: { top: marginTop, right: marginRight, bottom: marginBottom, left: marginLeft }
        }
      },
      children: children
    }]
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  console.log(`公文Word文档已生成：${outputPath}`);
}

/**
 * 从文件读取内容
 * 过滤掉 Markdown 一级标题行（# 标题名），并将 ##、### 等降级处理
 * 特殊处理：目录部分的内容作为正文，不识别为标题
 */
function readContentFromFile(filePath) {
  try {
    const content = fs.readFileSync(filePath, 'utf-8');
    const lines = content.split('\n');

    // 找到目录标题的位置
    let catalogTitleIndex = -1;
    for (let i = 0; i < lines.length; i++) {
      const cleanedLine = cleanMarkdown(lines[i]).trim();
      if (/^(目\s*录|目录)$/.test(cleanedLine)) {
        catalogTitleIndex = i;
        break;
      }
    }

    // 找到所有一级标题（一、二、三等）的位置
    const level1Titles = []; // { num: '一', index: 5 }
    for (let i = 0; i < lines.length; i++) {
      const cleanedLine = cleanMarkdown(lines[i]).trim();
      const match = cleanedLine.match(/^([一二三四五六七八九十]+)、/);
      if (match) {
        level1Titles.push({ num: match[1], index: i });
      }
    }

    // 确定目录区域结束位置
    let catalogEndIndex = -1;
    if (catalogTitleIndex >= 0 && level1Titles.length > 0) {
      // 找到第一个在实际内容中重复的一级标题
      // 即：在目录标题之后出现的第一个一级标题，且该标题在后面还会出现
      for (let i = 0; i < level1Titles.length; i++) {
        const current = level1Titles[i];
        // 只考虑目录标题之后的一级标题
        if (current.index <= catalogTitleIndex) continue;

        // 检查这个一级标题是否在后面还会出现
        let laterIndex = -1;
        for (let j = i + 1; j < level1Titles.length; j++) {
          if (level1Titles[j].num === current.num) {
            laterIndex = level1Titles[j].index;
            break;
          }
        }

        // 如果这个一级标题在后面还会出现，说明当前位置是目录内容
        // 而后面出现的位置是实际正文
        // 目录区域应该是从目录标题后到实际正文一级标题之前
        if (laterIndex > 0) {
          catalogEndIndex = laterIndex;
          break;
        }
      }

      // 如果没有找到重复的一级标题，说明目录到文档结束
      if (catalogEndIndex === -1) {
        catalogEndIndex = lines.length;
      }
    }

    // 处理每一行
    const processedLines = lines
      .map((line, index) => {
        // 先清理 Markdown 格式（但不删除数字列表，因为可能是标题）
        line = cleanMarkdown(line);

        // 处理反斜杠转义（如 1\. -> 1.）- 在cleanMarkdown之后处理
        line = line.replace(/\\([\.\*\+\-\!\[\]\(\)\`\#\>])/g, '$1');

        // 转换英文双引号为中文双引号
        line = toChineseQuotes(line);

        // 如果在目录区域内（目录标题之后，第一个重复一级标题之前）
        if (catalogTitleIndex >= 0 &&
            index > catalogTitleIndex &&
            index < catalogEndIndex) {
          // 目录中的内容标记为目录区域，后续作为正文处理
          // 但空行不添加标记，让其被过滤掉
          if (line.trim().length > 0) {
            line = '__CATALOG__' + line;
          }
        }

        // 删除中文序号后的空格（一、 二、 等）
        line = line.replace(/^([一二三四五六七八九十]+)、\s+/g, '$1、');
        // 删除二级标题序号后的空格（（一） （二） 等）
        line = line.replace(/^(（[一二三四五六七八九十]+）)\s+/g, '$1');
        // 删除三级标题序号后的空格（1. 2. 等）
        line = line.replace(/^(\d+\.)\s+/g, '$1');

        return line.trim();
      })
      .filter(line => {
        if (line.length === 0) return false;
        // 过滤一级标题：以 # 开头但不含 ##
        if (/^# [^#]/.test(line)) return false;
        return true;
      })
      .map(line => {
        // 移除 Markdown 标题前缀 ##、###、#### 等
        // 将 ## 标题降级为普通文本（去掉前缀）
        return line.replace(/^#{2,6}\s*/, '');
      });

    // 跳过与主标题重复的第一行（如果第一行看起来像文档标题）
    // 并检测日期行标记为居中
    const resultLines = [];
    let skipFirstLine = false;

    // 检查第一行是否看起来像标题（包含规划、报告、方案等关键词）
    if (processedLines.length > 0) {
      const firstLine = processedLines[0];
      // 如果第一行包含典型的标题关键词，则跳过
      if (/规划|报告|方案|意见|通知|决定|决议|纲要|计划/.test(firstLine)) {
        skipFirstLine = true;
      }
    }

    for (let i = 0; i < processedLines.length; i++) {
      const line = processedLines[i];

      // 跳过第一行（如果检测到是标题）
      if (i === 0 && skipFirstLine) {
        continue;
      }

      // 检测日期格式（如"2026年02月"、"2026年1月"等）并标记为居中
      if (/^\d{4}年\d{1,2}月$/.test(line)) {
        resultLines.push('__CENTER__' + line);
      } else {
        resultLines.push(line);
      }
    }

    return resultLines;
  } catch (error) {
    console.error(`读取文件失败：${error.message}`);
    process.exit(1);
  }
}

// 命令行参数处理
function main() {
  const args = process.argv.slice(2);

  if (args.length < 3) {
    console.log('用法：node generate_official_doc.js <标题> <内容文件> <输出路径>');
    console.log('或：node generate_official_doc.js --title "标题" --content "内容文本" --output "输出路径"');
    process.exit(1);
  }

  let title, contentLines, outputPath;

  if (args[0] === '--title') {
    // 使用 --key value 格式
    const argMap = {};
    for (let i = 0; i < args.length; i += 2) {
      if (args[i].startsWith('--') && args[i + 1]) {
        argMap[args[i]] = args[i + 1];
      }
    }
    title = argMap['--title'];
    outputPath = argMap['--output'];
    if (argMap['--content-file']) {
      contentLines = readContentFromFile(argMap['--content-file']);
    } else if (argMap['--content']) {
      // 处理换行符：先尝试按字面量\n分割，如果失败则按实际换行符分割
      const rawContent = argMap['--content'];
      if (rawContent.includes('\\n')) {
        contentLines = rawContent.split('\\n').map(line => line.trim()).filter(line => line.length > 0);
      } else {
        contentLines = rawContent.split('\n').map(line => line.trim()).filter(line => line.length > 0);
      }
    }
  } else {
    // 使用位置参数
    title = args[0];
    const contentFile = args[1];
    outputPath = args[2];
    contentLines = readContentFromFile(contentFile);
  }

  if (!title || !contentLines || !outputPath) {
    console.error('参数不完整，请提供标题、内容和输出路径');
    process.exit(1);
  }

  generateDocument(title, contentLines, outputPath).catch(err => {
    console.error('生成文档失败：', err);
    process.exit(1);
  });
}

main();
