# 政企公文生成 Skill (chinese-government-docs)

Hermes Agent skill for generating standardized Chinese government and enterprise official documents (政企公文).

## 📋 功能特性

基于 **GB/T 9704-2012《党政机关公文格式》** 标准，程序化生成规范排版的 Word (.docx) 公文文档。

### 支持的公文类型

通知 · 通报 · 决定 · 命令（令）· 公告 · 通告 · 意见 · 请示 · 批复 · 函 · 会议纪要 · 报告

### 格式规范

| 要素 | 标准 |
|------|------|
| 纸张 | A4 (210mm × 297mm) |
| 页边距 | 上37mm / 下35mm / 左28mm / 右26mm |
| 正文 | 仿宋 3号(16pt)，行距固定28.95pt |
| 标题 | 小标宋 2号(22pt)，红色居中 |
| 一级标题 | 黑体 3号 |
| 二级标题 | 楷体 3号 |
| 页码 | —×— 格式，居中，4号仿宋 |

## 🚀 快速开始

### 安装依赖

```bash
pip install python-docx
```

### 使用方式

#### 方式一：作为 Python 模块使用

```python
from templates.generate_gongwen import GongwenGenerator

gen = GongwenGenerator("output.docx")

# 版头
gen.set_header(
    org_name="国务院办公厅",
    doc_number="国办发〔2026〕15号",
    secrecy="",
    urgency=""
)

# 标题
gen.set_title("关于加快推进数字化转型的通知")

# 主送机关
gen.set_main_recipient("各省、自治区、直辖市人民政府：")

# 正文（支持4级标题层级）
gen.add_body("一、总体要求", level=1)
gen.add_body("坚持以……")
gen.add_body("（一）基本原则", level=2)
gen.add_body("坚持统筹规划……")

# 落款
gen.set_signature("国务院办公厅", "2026年5月26日")

# 抄送与印发
gen.add_cc("中央办公厅，全国人大常委会办公厅。")
gen.add_print_info("国务院办公厅秘书局", "2026年5月27日印发")

gen.save()
```

#### 方式二：直接运行示例

```bash
python templates/generate_gongwen.py
```

将生成一份标准通知公文示例 `示例公文_通知.docx`。

## 📁 文件结构

```
chinese-government-docs/
├── SKILL.md                        # Skill 定义文件
├── README.md                       # 本文件
├── templates/
│   └── generate_gongwen.py         # 核心生成脚本
└── 示例公文_通知.docx               # 示例输出
```

## 🔤 字体说明

| 用途 | 字体 |
|------|------|
| 正文 | 仿宋 (FangSong) |
| 一级标题 | 黑体 (SimHei) |
| 二级标题 | 楷体 (KaiTi) |
| 发文机关标志/标题 | 华文中宋 (STZhongsong) |

> ⚠️ 方正小标宋体在非政府采购电脑上可能未安装，本实现使用"华文中宋"替代，视觉效果接近。

## 📖 安装为 Hermes Agent Skill

```bash
# 克隆到 Hermes skill 目录
git clone https://github.com/Dynamic-braking9/chinese-government-docs.git \
    ~/.hermes/skills/productivity/chinese-government-docs
```

## 📄 License

MIT License
