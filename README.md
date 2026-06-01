# paper-formatter

中文学位论文 Word 排版工具。核心能力是把一份**格式模板**的排版样式，选择性地套用到你自己的**论文正文**上：保留模板的封面/前置页，用你的正文替换模板正文，逐段“模板里有对应格式就套用，没有就保持原样”。

基于 Python + [python-docx](https://python-docx.readthedocs.io/) + PyYAML 实现。

## 功能模式

| 入口 | 作用 |
|------|------|
| `template_main.py` | **模板套格式模式（v0.7，主力）**：保留模板封面/前置页 + 填充封面字段 + 用你的论文正文替换模板正文，逐段选择性套用模板格式 |
| `main.py` | **规则排版模式**：按 `config` 里的正则规则识别段落类型，强制套用 yaml 中定义的字体、字号、行距等 |
| `run_auto.py` | 一键运行（固定跑 `examples/123.docx` + `templates/111.docx`），结果写入 `output/` |

## 目录结构

```
paper-formatter/
├── main.py                        # 规则排版入口
├── template_main.py               # 模板套格式入口（主力）
├── run_auto.py                    # 一键运行
├── config/default_cn_thesis.yaml  # 排版规则 + 样式定义 + 前置页设置
├── src/
│   ├── classifier.py              # 正则规则编译与段落分类
│   ├── paragraph_detect.py        # 综合“样式名 + 文本内容”判定段落类型
│   ├── format_library.py          # 从模板抽取各类型的“格式范例段落”
│   ├── format_clone.py            # 复制段落/run/字体格式（含中文 eastAsia 字体）
│   ├── cover_fill.py              # 从论文封面提取字段并填入模板封面
│   ├── template_mode.py           # v0.7 核心流程编排
│   ├── template_profile.py        # 从模板推断标题/正文对应的样式名
│   ├── formatter.py               # 规则排版器（配合 main.py）
│   ├── styles.py / style_constants.py  # 样式应用与候选样式名表
├── scripts/                       # diagnose_docx.py、create_sample.py 辅助脚本
├── examples/ templates/           # 示例论文与模板 docx（不纳入版本库）
└── output/                        # 输出文件与运行报告（不纳入版本库）
```

## 安装

```bash
python -m venv .venv
.venv\Scripts\activate        # Windows
pip install -r requirements.txt
```

## 使用

模板套格式模式（推荐）：

```bash
python template_main.py 你的论文.docx 格式模板.docx -o output/from_template.docx
```

只查看模板里有哪些段落样式：

```bash
python template_main.py --list-styles 格式模板.docx
```

规则排版模式：

```bash
python main.py 你的论文.docx -o output/result.docx -c config/default_cn_thesis.yaml
```

## v0.7 核心流程（`apply_template_format`）

1. 复制**模板**作为输出文件，保留「摘要/第1章」之前的封面、声明等前置页；
2. 从你的论文封面提取题目、英文题目、学院、专业、姓名、学号、指导教师、日期等字段，填入模板封面占位；
3. 识别模板正文起点，只从模板**正文区**提取格式范例，避免封面大标题污染正文/章节格式；
4. 删除模板正文区，按原顺序接入你的正文段落和表格；
5. 如果模板摘要页有“题名/单位/英文题名/英文单位”题头，会按模板版式换成你的论文信息后保留；
6. 对每段正文：优先克隆模板范例格式；若无范例则套用模板里的 Word 样式（如 `Heading 2`、`Heading 3`、`正文`）；
7. 输出 docx，并生成 `output/template_apply_report.json` 报告。

## v0.7 已处理的格式细节

- `heading1` 范例优先选择真正的章节标题，如 `第 1 章 导论`，支持 `第1章` / `第 1 章` 这类带空格写法；
- 避免把 `第1章是绪论：...` 这类正文说明句误判成一级标题；
- `body` 范例会避开居中段落和封面段落，优先选择模板正文中的真实正文；
- `摘 要：正文`、`Abstract:正文`、`关键词：...`、`Key words:...` 这类段落会按 run 复制格式，只让标签加粗，不会把整段错误加粗；
- 复制 run 格式时会保留中文 `eastAsia` 字体，减少 Word 中中文字体丢失的问题；
- 如果目标输出文件正被 Word 打开占用，会自动另存为带时间戳的新文件。

验证示例：

```bash
python template_main.py examples\123.docx templates\111.docx -o output\from_template.docx
```

报告中重点看：

- `template_format_library`：模板各类段落实际选中了哪个格式范例；
- `format_applied`：每类段落最终套用了哪种模板格式；
- `abstract_page_preface`：摘要页题头是否已按模板结构插入。

> 说明：你提供的 [ysc/word](https://github.com/ysc/word) 是 Java 中文**分词**库，与 Word 排版无关。本项目借鉴的是 word_chat 等「模板 docx + 内容 docx → 规范输出」类项目的设计。

## 待办 / 已知问题

- 表格已支持从正文迁移；图片、公式尚未支持；
- 页眉页脚依赖模板保留，正文新增段落不会自动更新目录页码。
