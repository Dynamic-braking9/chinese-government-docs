# 政企公文生成 Skill (chinese-government-docs)

Hermes Agent Skill：基于 **GB/T 9704-2012《党政机关公文格式》** 标准，程序化生成规范排版的 Word (.docx) 公文文档。

## ✨ 功能特性

| 功能 | 说明 |
|------|------|
| 🏛️ 红头版头 | 发文机关标志红色大字、发文字号、密级/紧急程度标签 |
| 📐 标准页边距 | A4 纸，上37mm / 下35mm / 左28mm / 右26mm |
| 🔤 标准字体 | 根据标题层级自动切换仿宋、黑体、楷体、华文中宋 |
| 📏 固定行距 | 28.95pt 固定值，每页 22 行 × 28 字 |
| 🔢 四级标题 | 一、（一）1.（1）层次编号，各级对应不同字体 |
| 📄 页码格式 | —×— 格式，居中，4号仿宋 |
| ✍️ 落款签名 | 发文机关 + 成文日期，右对齐，右空4字 |
| 📎 附件/抄送 | 附件说明、抄送机关、印发信息，带分隔线 |

## 🚀 快速开始

### 安装依赖

```bash
pip install python-docx
```

### 作为 Python 模块使用

```python
from templates.generate_gongwen import GongwenGenerator

gen = GongwenGenerator("output.docx")

# 版头（红头区域）
gen.set_header(
    org_name="国务院办公厅",           # 发文机关名称
    doc_number="国办发〔2026〕15号",    # 发文字号（用六角括号〔〕）
    secrecy="",                        # 密级：秘密/机密/绝密
    urgency=""                         # 紧急程度：特急/加急
)

# 标题（2号华文中宋，红色居中）
gen.set_title("关于加快推进数字化转型的通知")

# 主送机关（仿宋3号，左顶格）
gen.set_main_recipient("各省、自治区、直辖市人民政府：")

# 正文 — level 决定字体：
#   0 = 普通段落（仿宋3号，首行缩进2字符）
#   1 = 一级标题（黑体加粗，如"一、总体要求"）
#   2 = 二级标题（楷体，如"（一）基本原则"）
#   3 = 三级标题（仿宋加粗，如"1. 目标"）
#   4 = 四级标题（仿宋，如"（1）细节"）
gen.add_body("一、总体要求", level=1)
gen.add_body("坚持以……")
gen.add_body("（一）基本原则", level=2)
gen.add_body("坚持统筹规划、分步实施……")

# 落款（右对齐，右空4字）
gen.set_signature("国务院办公厅", "2026年5月26日")

# 附件说明
gen.add_attachment_note("1. 数字化转型重点任务分工方案")

# 抄送与印发信息
gen.add_cc("中央办公厅，全国人大常委会办公厅。")
gen.add_print_info("国务院办公厅秘书局", "2026年5月27日印发")

gen.save()
```

### 运行示例

```bash
python templates/generate_gongwen.py
```

将生成一份标准通知公文 `示例公文_通知.docx`。

## 📁 仓库结构

```
chinese-government-docs/
├── SKILL.md                        # Hermes Agent Skill 定义文件
├── README.md                       # 英文说明
├── README.zh.md                    # 中文说明（本文件）
├── push_to_github.sh               # 一键上传脚本
├── templates/
│   └── generate_gongwen.py         # 核心生成脚本（500+ 行）
└── 示例公文_通知.docx               # 示例输出
```

## 🔤 字体说明

GB/T 9704-2012 标准对每个公文要素有严格的字体规定。所有中文字体需要同时设置 `w:ascii`/`w:hAnsi`（西文）和 `w:eastAsia`（东亚）三个 XML 属性，生成器已自动处理。

| 公文要素 | 西文字体 | 中文字体 | 字号 |
|----------|----------|----------|------|
| 正文 | FangSong | 仿宋 | 3号（16pt） |
| 发文机关标志 | STZhongsong | 华文中宋 | 1号（26pt），红色 |
| 公文标题 | STZhongsong | 华文中宋 | 2号（22pt），红色 |
| 一级标题 | SimHei | 黑体 | 3号（16pt），加粗 |
| 二级标题 | KaiTi | 楷体 | 3号（16pt） |
| 三级标题 | FangSong | 仿宋 | 3号（16pt），加粗 |
| 页码 | FangSong | 仿宋 | 4号（14pt） |
| 抄送/印发 | FangSong | 仿宋 | 4号（14pt） |

> **说明：** 标准规定标题使用方正小标宋体（FZXiaoBiaoSong），该字体通常仅在政府采购电脑上预装。本实现使用华文中宋（STZhongsong）作为视觉效果接近的替代方案。

## 📖 安装为 Hermes Agent Skill

```bash
git clone https://github.com/Dynamic-braking9/chinese-government-docs.git     ~/.hermes/skills/productivity/chinese-government-docs
```

## 📄 支持的公文类型

通知 · 通报 · 决定 · 命令（令）· 公告 · 通告 · 意见 · 请示 · 批复 · 函 · 会议纪要 · 报告

## API 参考

### `GongwenGenerator(filepath: str)`

创建公文生成器实例，指定输出文件路径。

| 方法 | 参数 | 说明 |
|------|------|------|
| `set_header()` | `org_name`, `doc_number`, `secrecy`, `urgency` | 红头版头：发文机关名称、发文字号、密级标签 |
| `set_title()` | `title` | 公文标题（2号红色华文中宋，居中） |
| `set_main_recipient()` | `recipient` | 主送机关（3号仿宋，左对齐） |
| `add_body()` | `text`, `level=0` | 正文内容或标题（0-4级层次） |
| `add_attachment_note()` | `note` | 附件说明 |
| `set_signature()` | `org_name`, `date` | 落款：发文机关 + 成文日期（右对齐） |
| `add_annotation()` | `text` | 附注（加圆括号） |
| `add_cc()` | `cc_text` | 抄送机关（带分隔线） |
| `add_print_info()` | `print_org`, `print_date` | 印发机关与日期 |
| `save()` | — | 保存 .docx 文件（自动添加页码） |

## 许可证

MIT
