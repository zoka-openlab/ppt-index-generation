# PPT Index Generation Skill

[中文](#中文说明) | [English](#english)

Turn selected PowerPoint files into an AI-readable material library, then search and collect relevant original slides by topic.

The project is designed for people who already have valuable content scattered across old decks. It helps Codex locate relevant slides, summarize where the material lives, and extract selected pages unchanged into focused material packs.

## English

### What it does

1. Indexes only the `.pptx` files explicitly provided by the user.
2. Creates a master Markdown index and a page-level index for each deck.
3. Searches those indexes by topic and reports the relevant source pages.
4. Extracts selected original slides unchanged into focused material packs.
5. Optionally lets the user remove duplicates or reorder existing pages before producing a confirmed pack.

Source presentations are treated as read-only. The Skill does not rewrite content, redesign pages, or generate new slides. The bundled extraction script rejects any attempt to use a source file as its own output.

### Packaging rule

The default output is one material pack per source deck plus one cross-deck Markdown summary. All matching pages from the same source deck are grouped together; the Skill does not create one file per slide. A single combined PPTX is created only when the user explicitly requests it and the source decks have compatible slide sizes and formats.

### Install

Requirements:

- Codex with local Skill support
- Python 3.10+
- `python-pptx`

```bash
python3 -m pip install -r requirements.txt
mkdir -p ~/.codex/skills
cp -R skill/ppt-index-generation ~/.codex/skills/
```

Restart Codex after installation, then invoke the Skill explicitly:

```text
Use $ppt-index-generation to index these PPTX files and find slides about customer outcomes.
```

See [examples/README.md](examples/README.md) for command-line examples.

### Privacy

Generated indexes contain absolute source paths and extracted slide text. They are useful locally but may be sensitive. Do not commit `PPT_Index/`, `outputs/`, or source presentations. The included `.gitignore` excludes them by default.

This repository contains no real company decks, customer data, project names, or user file paths. Tests create fictional presentations in temporary folders and delete them automatically.

### Limitations

- Text embedded inside images is not indexed unless another OCR tool is used.
- SmartArt, unusual embedded objects, and some third-party PowerPoint features may not be fully described.
- Candidate extraction preserves complete source slides; it does not make every element editable.
- Source decks with incompatible sizes or styles are kept as separate extracted packs.
- The Skill intentionally does not create covers, directories, transitions, summaries, or newly designed slides.

## 中文说明

### 它解决什么问题

很多有价值的公司介绍、案例、图表和成果页，长期散落在旧 PPT 里。真正困难的不是再写一遍，而是快速找到旧材料中已经做好的页面，并在不破坏源文件的前提下把它们定向汇总出来。

这个 Skill 专注于三件事：

1. 为用户指定的 PPT 建立 AI 可读索引。
2. 根据主题检索相关内容，列出来源文件和原始页码。
3. 将选中的原始页面保持不变地抽取出来；如有重复，只做保留、删除和排序。

它不负责重写页面、重新设计版式、补封面目录或生成新的 PPT 内容。

### 材料包规则

默认按照源 PPT 分别生成材料包，并额外提供一份跨文件总说明。同一份源 PPT 中命中的所有页面汇总到一个材料包，不会每页生成一个文件。只有用户明确要求，并且不同来源 PPT 的页面尺寸和格式兼容时，才合并成一份总材料包。

### 安装

需要 Python 3.10+、`python-pptx`，以及支持本地 Skill 的 Codex。

```bash
python3 -m pip install -r requirements.txt
mkdir -p ~/.codex/skills
cp -R skill/ppt-index-generation ~/.codex/skills/
```

安装后重启 Codex，可以这样使用：

```text
使用 $ppt-index-generation 索引这些 PPT 文件，并帮我搜集与公司能力相关的候选页。
```

### 隐私边界

索引文件会记录源 PPT 的绝对路径和提取文本，因此只适合保存在本地受控目录中，不应直接上传 GitHub。仓库已默认忽略 PPT、索引目录和输出目录。

公开仓库中不包含任何真实 PPT、公司资料、客户信息、项目名称或个人文件路径。自动测试只使用运行时临时生成的虚构材料。

## Project structure

```text
ppt-index-generation/
├── skill/ppt-index-generation/   # Installable Codex Skill
├── tests/                        # Synthetic, privacy-safe tests
├── examples/                     # Generic usage examples
├── .github/workflows/            # Continuous integration
├── README.md
├── CONTRIBUTING.md
├── SECURITY.md
└── LICENSE
```

## Contributing

Issues and pull requests are welcome. Please read [CONTRIBUTING.md](CONTRIBUTING.md) before submitting changes.
