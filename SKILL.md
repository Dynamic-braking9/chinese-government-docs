---
name: chinese-government-docs
description: "Use when generating Chinese government or enterprise official documents (政企公文). Produces GB/T 9704-2012 compliant .docx files with proper margins, fonts, numbering, red headers, and standard document types."
version: 1.0.0
author: Dynamic-braking9
license: MIT
metadata:
  hermes:
    tags: [word, docx, government, official-docs, chinese, GB/T-9704, 公文]
    related_skills: [document-generation]
---

# 政企公文生成 (Chinese Government & Enterprise Official Documents)

基于 GB/T 9704-2012《党政机关公文格式》标准，程序化生成规范排版的 Word 公文文档。

## Overview

中国党政机关公文有严格的格式规范（GB/T 9704-2012），涵盖纸张、页边距、字体字号、标题层级、版头格式、附件说明、落款等。本 skill 提供完整的 Python 实现方案，一键生成符合标准的 .docx 文件。

支持的公文类型：通知、通报、决定、命令（令）、公告、通告、意见、请示、批复、函、会议纪要、报告。

## When to Use

- 生成政府机关正式公文（通知、通报、请示、批复等）
- 生成企业内部红头文件、管理制度、会议纪要
- 批量生成格式统一的公文模板
- 教学或培训场景中的公文格式示范

Don't use for: 学术论文、商务报告、营销文档——这些用 `document-generation` skill。

## GB/T 9704-2012 核心规范速查

### 纸张与页面

| 参数 | 标准值 |
|------|--------|
| 纸张 | A4 (210mm × 297mm) |
| 上边距 | 37mm (3.7cm) |
| 下边距 | 35mm (3.5cm) |
| 左边距 | 28mm (2.8cm) |
| 右边距 | 26mm (2.6cm) |
| 行数/页 | 22 行 |
| 字数/行 | 28 字 |
| 行间距 | 28.95pt（固定值，约 29pt） |

### 字体字号对照

| 公文要素 | 字体 | 字号 | pt值 | 备注 |
|----------|------|------|------|------|
| 份号 | 仿宋 | 3号 | 16pt | 左上角 |
| 密级/紧急程度 | 黑体 | 3号 | 16pt | 份号下方 |
| 发文机关标志 | 小标宋 | 1号 | 26pt | 红色，居中（版头） |
| 发文字号 | 仿宋 | 3号 | 16pt | 版头红色线下方 |
| 标题 | 小标宋 | 2号 | 22pt | 红色分隔线下，居中 |
| 主送机关 | 仿宋 | 3号 | 16pt | 标题下空一行 |
| 正文 | 仿宋 | 3号 | 16pt | 首行缩进2字符 |
| 一级标题 | 黑体 | 3号 | 16pt | 如"一、""二、" |
| 二级标题 | 楷体 | 3号 | 16pt | 如"（一）""（二）" |
| 三级标题 | 仿宋加粗 | 3号 | 16pt | 如"1.""2." |
| 四级标题 | 仿宋 | 3号 | 16pt | 如"（1）""（2）" |
| 附件说明 | 仿宋 | 3号 | 16pt | 左空2字 |
| 发文机关/落款 | 仿宋 | 3号 | 16pt | 右空4字 |
| 成文日期 | 仿宋 | 3号 | 16pt | 右空4字 |
| 页码 | 仿宋 | 4号 | 14pt | 居中，"—×—"格式 |
| 附注 | 仿宋 | 3号 | 16pt | 左空2字，加括号 |
| 抄送机关 | 仿宋 | 4号 | 14pt | 左右各空1字 |
| 印发机关/日期 | 仿宋 | 4号 | 14pt | 左右各空1字 |

### 公文结构层次编号

```
一、  ……（一级标题，黑体）
（一）……（二级标题，楷体）
1.   ……（三级标题，仿宋加粗）
（1）……（四级标题，仿宋）
```

### 版头（红头）格式

```
┌──────────────────────────────────┐
│  × 份  × 密  × 急               │ ← 份号/密级/紧急程度（左上角）
│                                  │
│    ██████████████████████████    │ ← 发文机关标志（红色大字）
│                                  │
│  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━   │ ← 红色分隔线（版头线）
│  ×政发〔2026〕1号                │ ← 发文字号
│                                  │
│  ────────────────────────────   │ ← 红色分隔线（正文开始线）
└──────────────────────────────────┘
```

## Python 实现

### 依赖

```bash
pip install python-docx
```

### 核心模板

见 `templates/generate_gongwen.py` — 完整的公文生成脚本，支持：
- 版头红头格式
- 自动字体设置（仿宋/黑体/楷体/小标宋）
- 标准页边距与行间距
- 标题层级自动格式化
- 落款与成文日期
- 页码（—×—格式）
- 附件说明
- 抄送与印发信息

### 快速使用

```python
from templates.generate_gongwen import GongwenGenerator

gen = GongwenGenerator("output.docx")

gen.set_header(
    org_name="国务院办公厅",
    doc_number="国办发〔2026〕15号",
    secrecy="秘密",
    urgency="特急"
)

gen.set_title("关于加快推进数字化转型的通知")
gen.set_main_recipient("各省、自治区、直辖市人民政府：")
gen.add_body("一、总体要求", level=1)
gen.add_body("坚持以习近平新时代中国特色社会主义思想为指导……")
gen.add_body("（一）基本原则", level=2)
gen.add_body("坚持统筹规划、分步实施……")

gen.set_signature(
    org_name="国务院办公厅",
    date="2026年5月26日"
)
gen.add_attachment_note("1. 数字化转型实施方案")
gen.add_cc("各省、自治区、直辖市人民政府，国务院各部委、直属机构。")
gen.add_print_info("国务院办公厅", "2026年5月26日印发")

gen.save()
```

## 常用字体说明

| 字体名称 | 西文名 | 用途 |
|----------|--------|------|
| 仿宋 | FangSong | 正文、四级标题 |
| 仿宋_GB2312 | FangSong_GB2312 | 旧版标准正文（兼容） |
| 黑体 | SimHei | 一级标题、密级 |
| 楷体 | KaiTi | 二级标题 |
| 华文中宋 / 方正小标宋 | STZhongsong / FZXiaoBiaoSong | 发文机关标志、标题 |

> ⚠️ 方正小标宋体在非政府采购电脑上可能未安装，可用"华文中宋"（STZhongsong）替代。

## Common Pitfalls

1. **中文字体显示为方框**：必须同时设置 `w:ascii`、`w:hAnsi` 和 `w:eastAsia` 三个属性。只设 `font.name` 不够。

2. **行间距不对**：GB/T 9704 要求固定值 28.95pt，不是 1.5 倍行距。用 `paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY`。

3. **首行缩进单位搞混**：Word 中 1 个中文字符 ≈ 字号 pt 值。3号字(16pt)缩进2字符 = 32pt = 0.56cm，但 python-docx 用 `Cm(0.74)` 或 `Pt(32)` 都可。

4. **页边距单位**：python-docx 的 `Cm()` 是厘米，别跟毫米搞混。37mm = `Cm(3.7)`。

5. **红色版头**：发文机关标志颜色为 `RGBColor(0xFF, 0x00, 0x00)`，分隔线用表格底部边框或段落下边框实现。

6. **页码格式**：公文页码格式为"—×—"（居中，如 —1—、—2—），需用 `section.footer` 操作 XML。

7. **落款对齐**：发文机关和成文日期居右对齐，右侧留4字空格（约 64pt 或 `right_indent=Cm(4)`）。

## Verification Checklist

- [ ] 纸张 A4，页边距符合标准（上37/下35/左28/右26 mm）
- [ ] 正文仿宋3号(16pt)，行距固定28.95pt
- [ ] 标题小标宋2号(22pt)，一级标题黑体，二级标题楷体
- [ ] 版头发文机关标志红色，字号1号(26pt)
- [ ] 发文字号格式正确（如"×政发〔2026〕1号"，用六角括号）
- [ ] 正文首行缩进2字符
- [ ] 落款右对齐，右空4字
- [ ] 页码"—×—"格式，居中，4号仿宋
- [ ] 中文字符全部正常显示（无方框）
- [ ] 文件在 Word / WPS 中打开排版正确
