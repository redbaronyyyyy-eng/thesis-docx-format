# thesis-docx-format

把论文草稿 `.docx` 变成符合学校规范的定稿 `.docx`。确定性的活全部由脚本完成，
唯一需要判断的一步（标题分级）强制由人确认。

[English →](README.md)

面向中文高校的论文排版（规范会规定字体、字号、行距、分节页码和严格的前置件顺序），
但内核与学校无关：**一所学校一个 JSON 规范包，底下是通用引擎。**

## 它做什么

| 步骤 | 做什么 |
| --- | --- |
| `accept` | 接受全部修订、删批注、合并被删段落标记，然后自证结果与 Word「最终视图」逐字一致 |
| `outline` | **在完全没有标题样式的草稿里**识别标题层级，产出需你确认的提案 |
| `apply` | 套标题/正文格式，按规范包重建封面、扉页、原创性声明，建立分节页码，做卫生修复 |
| `audit` | 三轮审查：结构/分节/页码、字体/字号/行距、目录与内容一致性 |
| `pdfcheck` | 核验导出的 PDF：页码序列、嵌入字体，以及目录每条页码与标题实际落页是否一致 |

**不改一个字的内容。** 内容层面的发现（中文语境里的半角括号、疑似笔误、全角空格填空）
一律只报不改。

## 安装

零依赖、免编译，只用 Python 3.9+ 标准库。

```bash
git clone https://github.com/<你>/thesis-docx-format.git
cd thesis-docx-format
bash tests/run_regression.sh          # 合成草稿跑通全流程
```

作为 Claude Code skill：

```bash
git clone https://github.com/<你>/thesis-docx-format.git ~/.claude/skills/thesis-docx-format
```

给 Codex 或别的 agent 用：把目录拷进去即可。每一步都是一条普通命令行，不依赖任何 agent 的专有工具。

## 用法

```bash
S=.   # 或 ~/.claude/skills/thesis-docx-format
python3 $S/scripts/tfmt.py init ./P --draft 草稿.docx --profile tfsu-mti
python3 $S/scripts/tfmt.py accept  --project ./P
python3 $S/scripts/tfmt.py outline --project ./P     # → P/02_outline.tsv
#   ← 核对级别列，把 "# CONFIRMED: no" 改成 yes
python3 $S/scripts/tfmt.py apply   --project ./P
python3 $S/scripts/tfmt.py audit   --project ./P --docx ./P/03_formatted.docx
```

没确认就 `apply`，脚本直接退出（退出码 1）。这是有意的：标题层级一错，目录、页眉、分页、
页码会跟着全错，而且错得不显眼。

`02_outline.tsv` 的级别列取值：`1/2/3` 标题、`0` 正文、`F` 草稿自带前置件（删掉改由规范包重建）、
`T` 目录域（不动）、`X` 删除该段。

## 没有标题样式时怎么认标题

最糟的真实输入是：全文都是 `Normal`，字号也一样。所以不逐段猜，分两阶段：

1. **先找编号体系。** 扫全文抽取段首编号（`第X章` / `1.1` / `1.1.1` / `一、` / `（一）` / `Chapter N`），
   只采信**序列一致**的体系：一级 1,2,3… 递增，子级在父级下重新计数，允许整篇重启一次。
   仅这一条就能排掉"案例1：…案例27："这种形状酷似三级标题的平坦序列。
2. **再逐段打分。** 特殊名称（致谢/摘要/目录/参考文献/附录）> 编号+序列合法 > 已有 `outlineLvl`
   > 字号/加粗/居中/短行。负信号：`【…】`标签、`案例N：`、以句末标点结尾、超长。

文档里已有成体系编号时，形状类兜底会**关掉**——否则封面上的"某某大学"会被判成二级标题。

形状像标题但没判成标题的段，仍会以**候选**列出来，漏判的标题在那里捞。

## 规范包

`profiles/<name>.json` 是数据，`scripts/` 是引擎。`tfsu-mti.json` 是一份完整可用的实现。
换学校就复制 `profiles/_template.json`（里面有 7 步填写指引）来改。规范包声明：页面设置、
正文排版、三级标题、前置件逐段版式（**含封面填空线的制表位坐标**）、分节与页码方案、
页眉页脚、卫生修复规则、只报不改的内容检查、硬性指标（摘要字数、正文字数、题目长度）。

## 在信任任何同类工具之前值得一读

`references/` 记录了这套实现所围绕的那些坑：

- **`01-outline.md`** — 为什么序列一致性优于逐段打分；目录域如何搅乱序列；
  为什么重启判定写成 `b <= a` 是错的。
- **`02-sections.md`** — `pgNumType` 的格式**不跨节继承**；不指定 `headerReference` 的节会继承
  上一节的页眉；三种换页机制叠在同一边界上就会多出空白页；分节符前的空段会把它挤到下一页。
- **`03-hygiene.md`** — `"Times New Roman Regular"` 是 PostScript 样式名不是字体族名，
  Word 会静默回退且不报错；`w14:textFill` 在另一个命名空间，按 `w:` 查找删不掉却会覆盖 `w:color`；
  加下划线的**空格**在行尾会被裁掉，填空线必须用加下划线的**制表符**画。
- **`04-pdf.md`** — 目录页码是排版引擎缓存进去的。169 页论文实测：LibreOffice 48 条只对上 10 条，
  到附录偏差 6 页。**PDF 必须由排版出这份目录的那个程序导出。**

## 保证与不做的事

每一步都自证，证不过就不产出文件：`accept` 与 Word 最终视图逐字符比对；
`apply` 逐段比对草稿正文，不一致就报出第一处变化的段号并中止。

不做：改内容、更新目录页码（那是域刷新，要在 Word/WPS 里按 F9）、替你导 PDF、
重排表格/图片/脚注（只检测并报告）。

## 许可

MIT，见 [LICENSE](LICENSE)。仓库不附带任何学校的规范原件，
见 [`profiles/tfsu-mti/README.md`](profiles/tfsu-mti/README.md)。
