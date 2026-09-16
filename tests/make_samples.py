#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""合成测试草稿（只用标准库）。覆盖最糟糕的情况：完全没有标题样式。"""
import os, sys, zipfile
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'samples')
os.makedirs(OUT, exist_ok=True)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DRELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ''' + NS + '''><w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="21"/>
</w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>'''
def esc(s): return s.replace('&','&amp;').replace('<','&lt;')
def P(text, sz=None, bold=False, jc=None, ins=None, dele=None):
    rpr = '<w:rPr>%s%s</w:rPr>' % ('<w:b/>' if bold else '', '<w:sz w:val="%d"/>' % sz if sz else '')
    run = '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, esc(text))
    if ins: run = '<w:ins w:id="1" w:author="T" w:date="2026-01-01T00:00:00Z">%s</w:ins>' % run
    if dele:
        run = ('<w:del w:id="2" w:author="T" w:date="2026-01-01T00:00:00Z"><w:r>%s'
               '<w:delText xml:space="preserve">%s</w:delText></w:r></w:del>' % (rpr, esc(dele))) + run
    ppr = '<w:pPr>%s</w:pPr>' % ('<w:jc w:val="%s"/>' % jc if jc else '')
    return '<w:p>%s%s</w:p>' % (ppr, run)
def doc(paras):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>'
            '</w:body></w:document>' % (NS, ''.join(paras)))
def write(name, paras):
    path = os.path.join(OUT, name)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/styles.xml', STYLES); z.writestr('word/document.xml', doc(paras))
    print('  ', path)

BODY = '这是一段正文，用来占位，长度足够让它不会被误判成标题。' * 3
# ① 最糟：完全没有任何标题样式，全是 Normal，字号也一样
s1 = [P('某某大学'), P('硕士学位论文'), P('学 生 姓 名：张三'), P('指导教师姓名：李四'),
      P('致谢'), P(BODY), P('Acknowledgements'), P(BODY), P('摘要'), P(BODY),
      P('关键词：甲；乙；丙'), P('Abstract'), P(BODY), P('Key words: a; b; c')]
for ch, name in ((1,'引言'), (2,'方法'), (3,'结果'), (4,'结论')):
    s1.append(P('第%s章 %s' % ('一二三四'[ch-1], name)))
    s1.append(P(BODY))
    for se in (1, 2):
        s1.append(P('%d.%d 小节标题%d' % (ch, se, se))); s1.append(P(BODY))
        for ss in (1, 2):
            s1.append(P('%d.%d.%d 更小的标题' % (ch, se, ss))); s1.append(P(BODY))
s1 += [P('参考文献'), P('张三 (2020). 某书. 北京: 某出版社.'), P('附录1 原文'), P(BODY), P('附录2 译文'), P(BODY)]
write('s1_no_styles.docx', s1)

# ② 中文序号体系 一、（一）
s2 = [P('致谢'), P(BODY), P('摘要'), P(BODY), P('Abstract'), P(BODY)]
for i, name in enumerate(['引言', '文献综述', '结论'], 1):
    s2.append(P('第%s章 %s' % ('一二三'[i-1], name))); s2.append(P(BODY))
    for j in range(1, 3):
        s2.append(P('%s、二级标题%d' % ('一二三'[j-1], j))); s2.append(P(BODY))
        s2.append(P('（%s）三级标题' % '一二'[j-1])); s2.append(P(BODY))
s2 += [P('参考文献'), P('李四 (2021). 某文. 某刊, 1, 1-10.')]
write('s2_cn_numbering.docx', s2)

# ③ 带修订与段落标记删除
s3 = [P('致谢'), P(BODY), P('摘要'), P(BODY), P('Abstract'), P(BODY),
      P('第一章 引言'), P('保留的文字。', ins=True, dele='被删掉的旧文字。'),
      P(BODY), P('1.1 背景'), P(BODY), P('参考文献'), P('王五 (2022). 测试.')]
write('s3_tracked.docx', s3)
print('样本生成完毕 →', OUT)
