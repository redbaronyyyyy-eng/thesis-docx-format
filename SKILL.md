---
name: thesis-docx-format
description: Use when converting a thesis/dissertation draft .docx into the final school-specified format — accepting all tracked changes and stripping comments, detecting heading levels in drafts that have no heading styles at all, rebuilding cover/title page/declaration from a school profile, wiring multi-section page numbering (Roman front matter, Arabic body, unnumbered TOC), applying fonts/sizes/line-spacing per spec, auditing compliance in three passes, and verifying the exported PDF. Profile-driven (one JSON per school), stdlib-only, portable to other agents such as Codex.
---

# 论文草稿 docx → 规范格式 docx

## 0. 唯一最重要的原则

**确定性的活交给脚本，判断性的活才交给模型，拿不准的一律拦下来问人。**

接受修订、分节页码、字体字号、卫生修复、审查报告——全是确定性计算，`scripts/tfmt.py` 做。
模型只做一件事：**核对 `02_outline.tsv` 里的标题分级**。没确认就 `apply`，脚本直接拒绝执行。

三条红线：

1. **不改内容。** 标点、错别字、译名一律只报不改（`checks.text_patterns`）。
2. **每步自证。** accept 比对"原件最终视图 vs 接受后全文"，apply 比对"草稿正文是否原样保留"，对不上就退出，不产出文件。
3. **分页只认一个引擎。** 目录页码是域的缓存值，谁排的版就得谁导 PDF，见 `references/04-pdf.md`。

## 1. 环境与可移植性

脚本只用 Python 3 标准库（`xml.etree` + `zipfile`），不需要 pip 装任何东西。整个目录可以原样拷给
Codex 等其他 agent。唯一的可选依赖是 `pdfcheck` 用的 PyMuPDF，没有它其余流程照常跑。

```bash
S=~/.claude/skills/thesis-docx-format   # 或仓库所在目录
python3 $S/scripts/tfmt.py --help
bash $S/tests/run_regression.sh          # 合成草稿跑通全流程
```

## 2. 五步流程

```bash
python3 $S/scripts/tfmt.py init ./P --draft 草稿.docx --profile tfsu-mti
python3 $S/scripts/tfmt.py accept  --project ./P     # 接受修订 + 删批注 + 自证
python3 $S/scripts/tfmt.py outline --project ./P     # 识别标题 → 02_outline.tsv
#    ← 核对级别列，把 "# CONFIRMED: no" 改成 yes（这一步不能跳）
python3 $S/scripts/tfmt.py apply   --project ./P     # 套格式 + 建前置件 + 分节
python3 $S/scripts/tfmt.py audit   --project ./P --docx ./P/03_formatted.docx
```

| 步骤 | 做什么 | 产物 |
| --- | --- | --- |
| init | 复制草稿，建工程与 `metadata.json` | `00_input.docx` |
| accept | 接受全部修订、删批注、合并被删段落标记 | `01_accepted.docx` |
| outline | 两阶段识别标题，产出**需人工确认**的提案 | `02_outline.tsv` |
| apply | 标题/正文格式、前置件、分节页码、卫生修复 | `03_formatted.docx` |
| audit | 三轮审查（结构 / 字体 / 目录内容） | `report_audit.md` |
| pdfcheck | 核验用户导出的 PDF 与目录页码是否一致 | 终端报告 |

随时看进度：`python3 $S/scripts/tfmt.py status --project ./P`

## 3. 标题识别：两阶段 + 强制人工闸门

草稿最糟的情况是**一个标题样式都没有**，全文都是 Normal、字号一致。所以不逐段猜，先找体系：

**阶段一 发现编号体系。** 扫全文抽取段首编号（`第X章` / `1.1` / `1.1.1` / `一、` / `（一）` / `Chapter N`），
按**序列一致性**筛选：一级必须 1,2,3… 递增，子级在父级下重新计数，允许整篇重启一次。
不成体系的编号直接不采信——这条能自动排掉"案例1：…案例27："这种平坦序列的假标题。

**阶段二 逐段打分。** 特殊名称（致谢/摘要/Abstract/目录/参考文献/附录N）> 编号+序列合法 >
已有 `outlineLvl` > 字号大/加粗/居中/短行。负信号：`【…】`标签、`案例N：`、以句末标点结尾、超长。

产出 `02_outline.tsv`，级别列的取值：

| 值 | 含义 |
| --- | --- |
| `1` `2` `3` | 一/二/三级标题 |
| `0` | 正文。形状像标题的短行会标成"候选"列出来，**漏判的标题在这里捞** |
| `F` | 草稿自带的封面/扉页/声明，apply 会删掉改由规范包重建 |
| `T` | 目录域内容，apply 不动它 |
| `X` | 删除该段 |

**没把 `# CONFIRMED: no` 改成 `yes`，`apply` 会直接退出。** 这是有意的：标题分错会让整篇格式和目录全错。

## 4. 规范包（换学校只改这个）

`profiles/<name>.json` 是数据，`scripts/` 是通用内核。`tfsu-mti.json` 是天外 MTI 的完整实现，
新建学校时复制 `profiles/_template.json` 改。它声明：

- `page` 纸张与页边距　`body` 正文字体字号行距首行缩进　`headings` 三级标题
- `outline` 识别用的特殊名称/负信号/编号级别映射
- `front_matter` **封面、扉页、声明的逐段版式**（含制表位坐标），配 `metadata.json` 填值
- `sections` 分节计划与页码格式　`header`/`footer` 页眉页脚
- `hygiene` `checks` `limits` 卫生修复、只报不改的内容检查、硬性指标

细节见 `references/`：`01-outline.md` 标题识别　`02-sections.md` 分节页码
`03-hygiene.md` 卫生修复　`04-pdf.md` 导出 PDF。

## 5. 明确不做的事

- 不改内容、不代写、不查重、不翻译
- 不自动更新目录页码（域的缓存值必须在 Word/WPS 里按 F9）
- 不保证跨排版引擎分页一致，也不替用户导 PDF
- 表格/图片/脚注只检测并报告，不自动改版
