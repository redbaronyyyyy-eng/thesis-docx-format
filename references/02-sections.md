# 分节与页码

## 天外 MTI 是五节结构

| 节 | 范围 | 页眉 | 页码 |
| --- | --- | --- | --- |
| 1 | 封面 + 英文扉页 | 无 | 无 |
| 2 | 原创性声明 | 无 | 罗马 Ⅰ（`upperRoman` start=1）|
| 3 | 致谢 → Abstract | 无 | 罗马 Ⅱ–Ⅴ（`upperRoman` **不带 start**）|
| 4 | 目录 | 无 | 无（目录不计入页码）|
| 5 | 正文 → 附录 | 中文论文题目 | `- 1 -` 阿拉伯，宋体 |

## 四个必须知道的坑

1. **`pgNumType` 的 `fmt` 不跨节继承。** 第 3 节只写页脚不写 `<w:pgNumType w:fmt="upperRoman"/>`，
   Ⅰ 之后会变回 2、3、4。要续排就写 `fmt` 但不写 `start`。
2. **不指定 header/footerReference 的节会继承上一节。** 要"没有页眉"必须显式指向一个**空白**
   header 部件，不能省略。
3. **三种换页机制会打架。** 分节符（nextPage）、`pageBreakBefore`、`<w:br w:type="page"/>`
   叠在同一个边界上，每多一个就多一页空白。apply 里统一成：分节边界只用分节符，
   非分节边界只用 `pageBreakBefore`，并在最后清理重复。
4. **分节符前的空段会挤出空白页。** 目录域末尾那个只含 `fldChar end` 的空段尤其典型：
   它有正常行高，把分节符挤到下一页。要把**连续的**空段全部压成
   `spacing line=20 lineRule=exact` + `sz=2`，只压一个不够。

## 草稿自带的分节符必须先清掉

草稿里往往已经有旧的 `pPr/sectPr`，不清掉就会和新分节叠加，页码起点全乱。
apply 第 0 步就把所有段内 `sectPr` 删光，末尾的 body-level `sectPr` 整体替换。
