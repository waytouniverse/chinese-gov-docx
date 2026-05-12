---
name: ppt-generator
description: >
  根据用户提供的PPT大纲（文件路径或文本内容），逐页生成统一风格的PPT图片。
  核心流程：先生成封面页 → 用AI读取封面提取统一风格 → 用封面图作为参考图+风格描述逐页生成内页，确保视觉高度统一。
  当用户说"生成PPT"、"把大纲做成PPT"、"根据大纲生成演示文稿"、"帮我做PPT"、
  "PPT大纲生成图片"、"一页一页生成PPT"、"将Markdown转为PPT"等任何与PPT生成相关的请求时触发。
  也当用户提供一个.md文件说"做成PPT"或"生成演示文稿"时触发。
---

# PPT 生成器

根据用户提供的PPT大纲，逐页生成统一风格的PPT图片。
核心机制：**封面图作为视觉锚点 + 风格描述作为文字引导**，双通道确保整套PPT视觉高度统一。

## 工作流程

### 1. 接收大纲

从以下方式获取PPT大纲：
- 用户直接提供的文本内容
- 用户提供的文件路径（.md 文件）
- 当前对话中提到的文件

**重要：不要通过分析大纲关键词来推断风格。** 封面风格必须由生图AI根据封面页内容自行决定，而不是由你来预设。

**如果用户明确指定了风格**（如"暗黑悬疑风"、"政务风"、"科技风"），在封面提示词中加入用户的风格描述。

### 2. 生成封面页

**封面是整套PPT风格的"定调"页面。** 封面风格不由你来推断或预设，而是让生图AI根据封面页内容自行决定。

**生成策略**——在提示词中加入**精确的布局规范**，让AI在自由发挥风格的同时，保持一套PPT的可继承结构：

```
提示词结构：
[页面类型 + 精确布局规范] + [标题/副标题内容] + [让AI自由发挥风格]
```

具体示例：
```
A professional PPT cover slide.

Layout specification (strict):
- Top 25%: Title area — large bold text, left-aligned or center
- 25-35%: Subtitle area — smaller text directly below the title
- 35-85%: Main visual area — visual elements, illustrations, or imagery
- Bottom 15%: Footer area — presenter name, date, decorative elements

Title: [主标题内容].
Subtitle: [副标题内容].
Footer: [汇报人/日期].

16:9 widescreen format.
```

**关键原则**：
- **不预设颜色、字体、背景风格**——让AI根据标题主题自由决定
- **必须指定精确布局规范**——上述4个区域的百分比是内页继承的布局骨架，内页必须遵守相同比例
- 让生图AI根据标题和主题内容自行选择配色、字体、背景、氛围

**如果用户明确指定了风格**，在提示词中加入：
```
Overall visual direction: [用户指定的风格描述，如 dark cinematic documentary style, government blue tech style, etc.]
```

生成后保存为 `00_封面.png`（或 `00_Cover.png`），**记录封面图的完整路径**，后续每页内页都需要用它作为参考图。

### 3. 提取统一风格（核心步骤）

**这是确保整套PPT风格统一的关键。**

读取生成的封面图片，使用多模态AI（gpt-5.5 Responses API）提取**双通道风格信息**：

```python
# 伪代码示意
import base64
from openai import OpenAI

with open('封面.png', 'rb') as f:
    image_base64 = base64.b64encode(f.read()).decode('utf-8')

client = OpenAI(
    api_key='<YOUR_API_KEY>',
    base_url='<YOUR_BASE_URL>'
)

response = client.responses.create(
    model='gpt-5.5',
    input=[
        {
            'role': 'user',
            'content': [
                {
                    'type': 'input_text',
                    'text': '''Analyze this PPT cover slide and extract its visual style. Output TWO sections:

1. STYLE DESCRIPTION (one paragraph in English):
Describe the overall visual style using atmospheric, tactile, and cultural references — like a creative brief. Include mood, texture, cultural references, and what kind of presentation this feels like.

2. DESIGN SPEC (structured list):
- Primary color: (hex or description)
- Secondary/accent color: (hex or description)
- Background: (gradient direction, colors, texture)
- Title font: (weight, size feel, position alignment)
- Body font: (weight, size relative to title)
- Decorative elements: (style of icons, lines, shapes)
- Layout ratio: (title area / content area / footer area percentages)
- Key visual motifs: (recurring visual symbols or patterns)'''
                },
                {
                    'type': 'input_image',
                    'image_url': f'data:image/png;base64,{image_base64}'
                }
            ]
        }
    ],
    max_output_tokens=2000
)

style_prompt = response.output_text
```

**提取结果包含两部分**：
- **STYLE DESCRIPTION**：英文画面化风格描述（200-400词），用于引导生图AI的氛围和质感
- **DESIGN SPEC**：具体设计要素清单（颜色、字体、排版比例），用于精确复现

**重要：风格描述必须输出英文**，因为 gpt-image-2 主要用英文训练，中文风格描述效果不稳定。

### 4. 解析大纲，确定每页内容

解析大纲为页面列表，为每页提取：
- 页面标题
- 核心内容要点（精简为3-5个要点）
- 页面类型：cover / interior / ending

**严格遵循大纲原则**：
- **大纲有几页就生成几页，不增不减**——不要自动补充大纲里没有的尾页/Q&A页，也不要删减大纲中已有的页面
- **每页内容必须来自大纲原文**——不要用模板化内容替换大纲中的具体内容，大纲写了什么就生成什么
- **如果大纲最后一页是Q&A/感谢页**，则使用大纲中该页的原文内容生成，不要用通用模板覆盖

内容规则：
- 单页提示词不超过500词，过长会导致生成失败
- 表格用 "comparison table" 描述
- 卡片布局用 "card-based layout"
- 架构图用 "layered architecture diagram"
- 时间线用 "timeline"

### 5. 逐页生成内页（核心优化）

对每一页内页（包括尾页），使用**封面图作为参考图 + 风格文字描述**双通道确保一致性：

#### 5.1 提示词结构（分层清晰）

```
[风格锚定]
Same visual identity as the reference cover image. Match its exact color palette, font personality, background treatment, and decorative elements.

[设计要素]
[步骤3提取的 DESIGN SPEC 内容，完整保留]

[画面氛围]
[步骤3提取的 STYLE DESCRIPTION 内容，完整保留]

[页面类型+布局]
This is an interior content page.
Layout (same zones as cover): top 25% title bar (same bold style as cover title), 25-35% subtitle/tagline area, 35-85% content area with [具体布局类型], bottom 15% footer/decoration. Maintain identical visual language.

[内容]
Content for this page: [该页标题和核心内容，3-5个要点]
```

#### 5.2 调用生图时传入封面参考图（最关键）

**每页内页生成时，必须把封面图作为 reference_images 传入**。这是确保风格统一的核心手段：

```bash
# 内页生成命令（注意第6个参数是封面图路径）
python3 {{SKILL_PATH}}/scripts/generate.py "$PROMPT" 16:9 2k "$OUTPUT_DIR" "$COVER_IMAGE_PATH"
```

封面图路径就是步骤2中保存的 `00_封面.png` 的完整路径。

**关键原则**：
- **封面参考图是最强的风格锚定**——图生图比纯文字描述稳定得多，务必每页都传入
- **不要精简风格描述**——完整保留步骤3返回的两部分内容
- **每页提示词开头声明"Same visual identity as the reference cover image"**
- **布局约束使用与封面相同的区域百分比**——top 25%, 25-35%, 35-85%, bottom 15%

**生成策略**：
- 一次只生成一页（避免并发导致失败）
- 每页生成后等待完成再生成下一页
- 文件命名为 `01_[标题简写].png`、`02_[标题简写].png` 等
- 如果遇到超时失败，简化内容描述后重试（但保留完整的风格描述和参考图）

### 6. 尾页处理

**只有当大纲中明确包含尾页/Q&A页时才生成**，不要自动添加大纲中没有的尾页。

如果大纲中有尾页，使用与内页完全相同的提示词结构和参考图，**内容必须使用大纲原文**，不要用通用模板替换：

```
Same visual identity as the reference cover image. Match its exact color palette, font personality, background treatment, and decorative elements.

[DESIGN SPEC]

[STYLE DESCRIPTION]

This is the final ending/Q&A page. Layout: same zones as cover, with title "Q&A" in the title area and centered content below.

Content: [大纲中尾页的原文内容，逐字使用，不要替换或省略]
```

**禁止行为**：
- 不要自动添加大纲中没有的"感谢聆听"或"Q&A"页
- 不要用模板化内容（如 "Large centered text 感谢聆听"）替换大纲中的具体内容
- 大纲写了"联系方式 / 二维码（如有）:"，就必须在提示词中包含这段内容

### 7. 输出整理

所有页面生成完成后，向用户汇报：
- 总页数
- 保存目录路径
- 每页的文件名列表
- 提示用户查看效果

## 配置信息

### 风格提取API（gpt-5.5 多模态）

- **base_url**: `<YOUR_BASE_URL>`（需替换为实际的 API 地址）
- **api_key**: `<YOUR_API_KEY>`（需替换为实际的 API 密钥）
- **model**: `gpt-5.5`
- **endpoint**: Responses API (`/v1/responses`)

### 图片生成API（gpt-image-2）

通过调用 `gpt-image` skill 使用。支持 `reference_images` 参数（最多16张参考图），**内页生成时务必传入封面图**。

## 风格一致性保障机制（总结）

| 手段 | 作用 | 优先级 |
|------|------|--------|
| 封面图作为参考图传入每页 | 最强锚定，图生图直接匹配颜色/字体/材质 | 最高 |
| DESIGN SPEC 具体设计要素 | 精确复现颜色值、字体特征、排版比例 | 高 |
| STYLE DESCRIPTION 氛围描述 | 引导整体质感和文化参照 | 中 |
| 统一布局百分比规范 | 确保排版结构一致 | 高 |
| "Same visual identity" 声明 | 文字层面强化风格一致性要求 | 辅助 |

## 注意事项

1. **封面参考图是第一优先级**：每页内页生成都必须传入封面图作为参考，这是风格统一的最核心手段。绝不能省略。

2. **风格提取必须双通道**：gpt-5.5 返回的 STYLE DESCRIPTION + DESIGN SPEC 两部分都要完整保留用于每页生图。如果提示词过长，可以精简"内容描述"部分，但风格描述和设计要素必须完整。

3. **风格描述输出英文**：gpt-image-2 主要用英文训练，英文风格描述效果更稳定。DESIGN SPEC 中的颜色描述等可用中英混合。

4. **中文文本处理**：提示词中的中文文本用英文描述位置和内容（如 "title text '产品概览'"），避免AI混淆。

5. **失败重试**：如果某页生成失败，简化内容描述后重试。但**不要省略封面参考图和风格描述**。常见失败原因：提示词过长、并发过高、内容敏感。

6. **尺寸与格式**：统一使用 `16:9` 比例，`2k` 分辨率。

7. **输出目录**：根据PPT标题自动命名输出文件夹。从封面页主标题中提取核心关键词（去除引号、特殊符号），作为文件夹名。例如：
   - 标题 `"死了么"——现象级产品爆火拆解` → 文件夹 `死了么PPT/`
   - 标题 `政务龙虾——广东省AI智能体政务应用实践` → 文件夹 `政务龙虾PPT/`
   - 如果用户指定了目录，使用用户指定的目录

8. **不要读取生成的图片文件**：生成后只报告文件路径，不要使用 Read 工具查看图片内容（会导致对话报错），除非用户明确要求查看。

## 完整示例

**用户输入**：
```
帮我生成这个PPT：

# "死了么"产品调研 PPT大纲

## 封面页
标题："死了么"——现象级产品爆火拆解
副标题：精准痛点 × 争议命名 × 零营销破圈
调研时间：26年5月 / 汇报人：叶远旬

## 第1页：产品概览
产品定位：独居人群"每日签到+超时通知"安全守护工具
核心功能：用户每日签到 → 超时未签到 → 自动邮件通知紧急联系人
商业模式：8元买断制
上线时间：2025年6月10日
关键数据：上线不到一周走红，下载量暴涨100倍，登顶苹果付费软件排行榜

## 第2页：爆火四要素总览
痛点精准：1.23亿独居人群的安全焦虑
命名破圈：黑色幽默打破生死避讳，争议自带传播
成本叙事：3名95后兼职、千元成本、草根逆袭
情绪共鸣：折射社会原子化下的集体焦虑
```

**执行流程**：
1. 解析大纲，提取封面页内容（标题、副标题、汇报人/日期）
2. **不预设风格**，直接将封面内容 + 精确布局规范给生图技能 → 生成 `00_封面.png`，**记录路径**
3. 读取封面图片 → 调用 gpt-5.5 Responses API 提取双通道风格 → 得到 STYLE DESCRIPTION + DESIGN SPEC
4. 解析大纲剩余页面
5. 用风格描述 + 封面参考图 + 第1页内容 → 生成 `01_产品概览.png`
6. 用风格描述 + 封面参考图 + 第2页内容 → 生成 `02_爆火四要素.png`
7. 汇报结果
