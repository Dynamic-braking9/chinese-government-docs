# Chinese Government Document Generator (chinese-government-docs)

> **[中文版 / Chinese Version → README.zh.md](README.zh.md)**


A Hermes Agent Skill for generating standardized Chinese government and enterprise official documents (.docx) compliant with **GB/T 9704-2012** (*Format for Official Documents of Party and Government Organs*).

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🏛️ Red Header (红头) | Official header with institution name in red, document number, secrecy/urgency labels |
| 📐 GB/T 9704 Margins | A4, top 37mm / bottom 35mm / left 28mm / right 26mm |
| 🔤 Standard Fonts | Auto-switches between FangSong (仿宋), SimHei (黑体), KaiTi (楷体), STZhongsong (华文中宋) per heading level |
| 📏 Fixed Line Spacing | 28.95pt exact line height per spec, 22 lines × 28 chars per page |
| 🔢 4-Level Headings | 一、（一）1. （1） hierarchy with correct fonts per level |
| 📄 Page Numbers | —×— format, centered, 4号 FangSong |
| ✍️ Signature Block | Right-aligned institution name + date, 4-char right indent |
| 📎 Attachments & CC | Attachment notes, CC list, print info with separating lines |

## 🚀 Quick Start

### Prerequisites

```bash
pip install python-docx
```

### Usage as Python Module

```python
from templates.generate_gongwen import GongwenGenerator

gen = GongwenGenerator("output.docx")

# Red header (版头)
gen.set_header(
    org_name="国务院办公厅",           # Issuing institution
    doc_number="国办发〔2026〕15号",    # Document number (use 〔〕brackets)
    secrecy="",                        # Secrecy level: 秘密/机密/绝密
    urgency=""                         # Urgency: 特急/加急
)

# Title (2号 STZhongsong, red, centered)
gen.set_title("关于加快推进数字化转型的通知")

# Main recipient (仿宋3号, left-aligned)
gen.set_main_recipient("各省、自治区、直辖市人民政府：")

# Body text — level determines font:
#   0 = paragraph (FangSong 3号, 2-char indent)
#   1 = primary heading (SimHei bold, e.g. "一、总体要求")
#   2 = secondary heading (KaiTi, e.g. "（一）基本原则")
#   3 = tertiary heading (FangSong bold, e.g. "1. 目标")
#   4 = quaternary heading (FangSong, e.g. "（1）细节")
gen.add_body("一、总体要求", level=1)
gen.add_body("坚持以……")
gen.add_body("（一）基本原则", level=2)
gen.add_body("坚持统筹规划、分步实施……")

# Signature (right-aligned, 4-char right indent)
gen.set_signature("国务院办公厅", "2026年5月26日")

# Attachment note
gen.add_attachment_note("1. 数字化转型重点任务分工方案")

# CC (抄送) and print info
gen.add_cc("中央办公厅，全国人大常委会办公厅。")
gen.add_print_info("国务院办公厅秘书局", "2026年5月27日印发")

gen.save()
```

### Run the Demo

```bash
python templates/generate_gongwen.py
```

Generates a sample notification document `示例公文_通知.docx`.

## 📁 Repository Structure

```
chinese-government-docs/
├── SKILL.md                        # Hermes Agent skill definition
├── README.md                       # This file
├── push_to_github.sh               # One-click GitHub upload script
├── templates/
│   └── generate_gongwen.py         # Core generator (500+ lines)
└── 示例公文_通知.docx               # Demo output
```

## 🔤 Font Reference

The GB/T 9704-2012 standard mandates specific fonts for each document element. All CJK fonts require setting **both** `w:ascii`/`w:hAnsi` (Western) and `w:eastAsia` (East Asian) XML attributes — the generator handles this automatically.

| Element | Western Font | East Asia Font | Size |
|---------|-------------|----------------|------|
| Body text | FangSong | 仿宋 | 3号 (16pt) |
| Institution header | STZhongsong | 华文中宋 | 1号 (26pt), red |
| Document title | STZhongsong | 华文中宋 | 2号 (22pt), red |
| Level-1 heading | SimHei | 黑体 | 3号 (16pt), bold |
| Level-2 heading | KaiTi | 楷体 | 3号 (16pt) |
| Level-3 heading | FangSong | 仿宋 | 3号 (16pt), bold |
| Page number | FangSong | 仿宋 | 4号 (14pt) |
| CC / Print info | FangSong | 仿宋 | 4号 (14pt) |

> **Note:** The standard calls for 方正小标宋体 (FZXiaoBiaoSong) for headers/titles. Since this font is typically only on government-procured machines, STZhongsong (华文中宋) is used as a visually similar substitute.

## 📖 Install as Hermes Agent Skill

```bash
git clone https://github.com/Dynamic-braking9/chinese-government-docs.git \
    ~/.hermes/skills/productivity/chinese-government-docs
```

## 📄 Supported Document Types

通知 · 通报 · 决定 · 命令（令）· 公告 · 通告 · 意见 · 请示 · 批复 · 函 · 会议纪要 · 报告

## API Reference

### `GongwenGenerator(filepath: str)`

Create a new document generator targeting the given output path.

| Method | Parameters | Description |
|--------|-----------|-------------|
| `set_header()` | `org_name`, `doc_number`, `secrecy`, `urgency` | Red header block with institution name, number, labels |
| `set_title()` | `title` | Document title (2号 red STZhongsong, centered) |
| `set_main_recipient()` | `recipient` | Addressee line (3号 FangSong, left-aligned) |
| `add_body()` | `text`, `level=0` | Body text or heading (0-4 level hierarchy) |
| `add_attachment_note()` | `note` | Attachment reference |
| `set_signature()` | `org_name`, `date` | Issuing org + date (right-aligned) |
| `add_annotation()` | `text` | Footnote annotation in parentheses |
| `add_cc()` | `cc_text` | CC recipients (with separator line) |
| `add_print_info()` | `print_org`, `print_date` | Printer + date |
| `save()` | — | Write .docx file (auto-adds page numbers) |

## License

MIT
