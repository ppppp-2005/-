# paper-formatter

中文学位论文 Word 排版工具。核心能力是把一份**格式模板**的排版样式，选择性地套用到你自己的**论文正文**上：保留模板的封面/前置页，用你的正文替换模板正文，逐段“模板里有对应格式就套用，没有就保持原样”。

基于 Python + [python-docx](https://python-docx.readthedocs.io/) + PyYAML 实现。

## 功能模式

| 入口 | 作用 |
|------|------|
| `template_main.py` | **模板套格式模式（v0.6，主力）**：保留模板封面/前置页 + 用你的论文正文替换，逐段选择性套用模板格式 |
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
│   ├── format_clone.py            # 复制段落/字体格式（含中文 eastAsia 字体）
│   ├── template_mode.py           # v0.6 核心流程编排
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

## v0.6 核心流程（`apply_template_format`）

1. 复制**模板**作为输出文件，保留“摘要/第1章”之前的封面、声明等前置页；
2. 删除模板正文部分；
3. 读取论文每一段，识别类型（`heading1` / `body` / `abstract_title` …）；
4. 模板里若有该类型的格式范例 → `clone_paragraph_format` 套用；没有 → 保持原格式；
5. 输出 docx，并生成 `output/template_apply_report.json` 报告。

## 待办 / 已知问题

- `run_result.txt` 与报告 json 里的版本号仍显示 v0.5，与当前 v0.6 代码不一致，需同步。
- `references_body` 有时被误识别为 `Heading 2`（见 `format_library.py` 的参考文献条目匹配）。
