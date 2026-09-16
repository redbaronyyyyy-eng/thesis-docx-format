#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""thesis-docx-format 主 CLI —— 论文草稿 docx → 规范格式 docx。

只用 Python 3 标准库。每个阶段都自证（文本比对），对不上就拒绝产出。
"""
import argparse, json, os, re, shutil, sys, copy
import xml.etree.ElementTree as ET
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docxkit as D
from docxkit import q, W, XML  # W/XML 供本模块内联 XML 片段使用

SKILL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CJK = re.compile(r'[㐀-鿿]')
FULLW = re.compile(r'[　-〿＀-￯]')

# ══════════════════════════ 工程与状态 ══════════════════════════
def pj(proj, *a): return os.path.join(proj, *a)
def load_json(p, default=None):
    if not os.path.exists(p): return default
    with open(p, encoding='utf-8') as f: return json.load(f)
def dump_json(p, obj):
    with open(p, 'w', encoding='utf-8') as f: json.dump(obj, f, ensure_ascii=False, indent=2)

def load_profile(name):
    p = name if os.path.exists(name) else os.path.join(SKILL, 'profiles', name + '.json')
    if not os.path.exists(p): sys.exit(f'找不到规范包: {name}（在 {SKILL}/profiles/ 下）')
    return load_json(p)

def state(proj): return load_json(pj(proj, 'project.json'), {})
def set_state(proj, **kw):
    s = state(proj); s.update(kw); dump_json(pj(proj, 'project.json'), s)

# ══════════════════════════ init ══════════════════════════
def cmd_init(a):
    os.makedirs(a.project, exist_ok=True)
    shutil.copy(a.draft, pj(a.project, '00_input.docx'))
    prof = load_profile(a.profile)
    meta_p = pj(a.project, 'metadata.json')
    if not os.path.exists(meta_p):
        dump_json(meta_p, prof.get('metadata_template', {}))
    set_state(a.project, profile=a.profile, draft=os.path.abspath(a.draft), stage='init')
    print(f'工程已建立: {a.project}')
    print(f'  规范包: {a.profile}')
    print(f'  下一步: tfmt.py accept --project {a.project}')
    print(f'  如需生成封面/扉页/声明，请先填写 {meta_p}')

# ══════════════════════════ accept：接受修订 + 删批注 ══════════════════════════
REV_PARTS = re.compile(r'^word/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$')
COMMENT_PARTS = ['comments', 'commentsExtended', 'commentsIds', 'commentsExtensible', 'people']

def accept_tree(root, log):
    # 1 删除 w:del / w:moveFrom（段落标记上的除外，稍后处理）
    pm = D.parent_map(root)
    for tag in ('del', 'moveFrom'):
        for e in [x for x in root.iter(q(tag))]:
            par = pm.get(e)
            if par is None or D.ln(par) == 'rPr': continue
            par.remove(e); log['del'] += 1
    # 2 就地展开 w:ins / w:moveTo
    for tag in ('ins', 'moveTo'):
        for _ in range(20):                     # 嵌套 ins 需要反复展开，直到一个不剩
            pm = D.parent_map(root)
            items = [x for x in root.iter(q(tag))]
            if not items: break
            for e in items:
                par = pm.get(e)
                if par is None: continue
                if D.ln(par) == 'rPr':
                    par.remove(e); continue
                i = list(par).index(e)
                for k, ch in enumerate(list(e)): par.insert(i + k, ch)
                par.remove(e); log['ins'] += 1
    # 3 格式修订记录 / 批注锚点
    pm = D.parent_map(root)
    for t in ('rPrChange','pPrChange','tblPrChange','trPrChange','tcPrChange','sectPrChange',
              'numberingChange','tblGridChange','cellIns','cellDel','cellMerge',
              'moveFromRangeStart','moveFromRangeEnd','moveToRangeStart','moveToRangeEnd',
              'commentRangeStart','commentRangeEnd','commentReference'):
        for e in [x for x in root.iter(q(t))]:
            par = pm.get(e)
            if par is not None:
                par.remove(e); log['fmt' if 'Change' in t else 'comment'] += 1
    for n in root.iter(q('delText')): n.tag = q('t')
    for n in root.iter(q('delInstrText')): n.tag = q('instrText')

def merge_deleted_marks(root, log):
    """段落标记被删 → 与下一段合并，并继承下一段的 pPr（Word 的行为）。"""
    while True:
        target = None
        for p in root.iter(q('p')):
            pr = p.find(q('pPr'))
            if pr is None: continue
            rp = pr.find(q('rPr'))
            if rp is not None and rp.find(q('del')) is not None:
                target = (p, rp); break
        if target is None: break
        p, rp = target
        rp.remove(rp.find(q('del')))
        pm = D.parent_map(root); par = pm.get(p)
        sibs = list(par); i = sibs.index(p)
        nxt = sibs[i+1] if i+1 < len(sibs) else None
        if nxt is None or nxt.tag != q('p'): continue
        for ch in list(nxt):
            if ch.tag == q('pPr'): continue
            p.append(ch)
        npr = nxt.find(q('pPr')); opr = p.find(q('pPr'))
        if opr is not None: p.remove(opr)
        if npr is not None: p.insert(0, npr)
        par.remove(nxt); log['merge'] += 1

def cmd_accept(a):
    proj = a.project
    src = pj(proj, '00_input.docx')
    pkg = D.Package(src)
    before = D.final_view_text(pkg.tree('word/document.xml'))
    log = dict(ins=0, del_=0, fmt=0, comment=0, merge=0); log['del'] = 0
    for n in [x for x in pkg.names if REV_PARTS.match(x)]:
        root = pkg.tree(n); accept_tree(root, log)
        if n == 'word/document.xml': merge_deleted_marks(root, log)
    for n in ('word/settings.xml',):
        if pkg.has(n):
            root = pkg.tree(n)
            pm = D.parent_map(root)
            for t in ('trackChanges', 'revisionView'):
                for e in [x for x in root.iter(q(t))]:
                    if pm.get(e) is not None: pm[e].remove(e)
    for base in COMMENT_PARTS:
        if pkg.has('word/%s.xml' % base):
            pkg.drop('word/%s.xml' % base); pkg.drop_part_refs(base)
    out = pj(proj, '01_accepted.docx'); pkg.save(out)
    after = D.doc_text(D.Package(out).tree('word/document.xml'))
    ok = (before == after)
    print(f'接受修订: 插入展开 {log["ins"]}，删除清除 {log["del"]}，格式修订 {log["fmt"]}，'
          f'批注锚点 {log["comment"]}，合并段落 {log["merge"]}')
    print(f'自证（原件最终视图 vs 接受后全文）: {"逐字一致 ✓" if ok else "不一致 ✗"}  {len(before)} / {len(after)}')
    if not ok:
        sys.exit('！文本对不上，已停止。请勿使用 01_accepted.docx')
    set_state(proj, stage='accepted')
    print(f'→ {out}\n  下一步: tfmt.py outline --project {proj}')

# ══════════════════════════ outline：标题识别 ══════════════════════════
CN_NUM = '零一二三四五六七八九十百'
def cn2int(s):
    if s.isdigit(): return int(s)
    if not s: return None
    if s == '十': return 10
    if '十' in s:
        a, _, b = s.partition('十')
        return (CN_NUM.index(a) if a else 1) * 10 + (CN_NUM.index(b) if b else 0)
    return CN_NUM.index(s) if s in CN_NUM else None

SCHEMES = [
    ('chapter_cn', re.compile(r'^第\s*([' + CN_NUM + r'\d]+)\s*[章篇部]'), 1),
    ('chapter_en', re.compile(r'^(?:Chapter|CHAPTER)\s+(\d+|[IVXL]+)\b'), 1),
    ('dotted3',    re.compile(r'^(\d+)\.(\d+)\.(\d+)\s*[^\d]'), 3),
    ('dotted2',    re.compile(r'^(\d+)\.(\d+)\s*[^\d.]'), 2),
    ('dotted1',    re.compile(r'^(\d+)[.、]\s*\S'), 1),
    ('cn_dun',     re.compile(r'^([' + CN_NUM + r']+)\s*、'), 2),
    ('paren_cn',   re.compile(r'^[（(]\s*([' + CN_NUM + r'\d]+)\s*[)）]'), 3),
]

def para_facts(root, styles, profile):
    out = []
    for i, p in enumerate(D.paragraphs(root)):
        t = D.para_text(p).strip()
        runs = [r for r in p.iter(q('r')) if D.run_text(r).strip()]
        szs = {styles.rval(p, r, 'sz') for r in runs} or {None}
        out.append(dict(
            idx=i, text=t, style=D.style_id(p),
            outline=styles.pval(p, 'outlineLvl', 'val'),
            size=max((int(s) for s in szs if s and s.isdigit()), default=0),
            bold=all(styles.bold(p, r) for r in runs) if runs else False,
            jc=styles.pval(p, 'jc', 'val') or 'left',
            length=len(t), empty=(t == ''),
            ends_punct=bool(t) and t[-1] in '。！？；：.!?;,，、',
            has_break=any(b.get(q('type')) == 'page' for b in p.iter(q('br'))),
        ))
    return out

def discover_schemes(facts):
    """哪些编号体系在本文里是"成体系"的：序号必须单调递增（子级在父级下重新计数）。"""
    hits = {}
    for f in facts:
        if f['empty'] or f['length'] > 60: continue
        for name, rx, lvl in SCHEMES:
            m = rx.match(f['text'])
            if m:
                nums = tuple(cn2int(x) for x in m.groups())
                if any(n is None for n in nums): continue
                hits.setdefault(name, []).append((f['idx'], nums, lvl))
                break
    good = {}
    for name, lst in hits.items():
        if len(lst) < 2: continue
        firsts = [n[0] for _, n, _ in lst]
        # 允许重启一次（目录/正文各出现一遍，或分卷重新编号）
        restarts = sum(1 for a, b in zip(firsts, firsts[1:]) if b < a and b == 1)
        monotone = all(b - a in (0, 1) or (b == 1 and a >= b) for a, b in zip(firsts, firsts[1:]))
        seq_ok = monotone and restarts <= 1 and firsts[0] in (1, 2)
        good[name] = dict(count=len(lst), seq_ok=seq_ok, level=lst[0][2])
    return good, hits

TOC_STYLE = re.compile(r'^toc\s', re.I)
def mark_regions(root, styles, prof, facts):
    """标出目录域区间与草稿自带前置件区间。"""
    paras = D.paragraphs(root)
    region = ['body'] * len(paras)
    # 目录域：靠 TOC 样式 + fldChar 区间双保险
    in_toc, depth = False, 0
    for i, p in enumerate(paras):
        name = (styles.name.get(D.style_id(p) or '', '') or '')
        instr = ''.join((x.text or '') for x in p.iter(q('instrText')))
        starts = 'TOC' in instr and '\\o' in instr
        if starts: in_toc = True
        if TOC_STYLE.match(name) or in_toc: region[i] = 'toc'
        # 目录域内套着 HYPERLINK / PAGEREF 子域，必须按 begin/end 配对计数，不能见 end 就收
        for fc in p.iter(q('fldChar')):
            t = fc.get(q('fldCharType'))
            if t == 'begin': depth += 1
            elif t == 'end':
                depth -= 1
                if in_toc and depth <= 0: in_toc, depth = False, 0
    # 前置件：首个 front 角色标题之前，且命中封面/声明特征词
    pats = prof['outline']['special_titles']
    first_front = None
    for i, f in enumerate(facts):
        for role in prof['outline']['front_matter_titles']:
            if role in pats and re.fullmatch(pats[role]['pattern'], f['text']):
                first_front = i; break
        if first_front is not None: break
    marks = prof['outline'].get('front_matter_markers', [])
    if first_front:
        for i in range(first_front):
            if region[i] == 'toc': continue
            if any(re.search(mk, facts[i]['text']) for mk in marks): region[i] = 'front'
        # 特征词命中率高则整段区间都判为前置件
        hit = sum(1 for i in range(first_front) if region[i] == 'front')
        if hit >= 3:
            for i in range(first_front):
                if region[i] != 'toc': region[i] = 'front'
    return region

def cmd_outline(a):
    proj = a.project
    prof = load_profile(state(proj).get('profile', a.profile or 'tfsu-mti'))
    pkg = D.Package(pj(proj, '01_accepted.docx') if os.path.exists(pj(proj, '01_accepted.docx')) else pj(proj, '00_input.docx'))
    root = pkg.tree('word/document.xml'); styles = D.Styles(pkg)
    facts = para_facts(root, styles, prof)
    region = mark_regions(root, styles, prof, facts)
    schemes, hits = discover_schemes([f for f in facts if region[f['idx']] == 'body'])
    strong = any(v['seq_ok'] for v in schemes.values())
    names = prof['outline']['special_titles']
    front = set(prof['outline']['front_matter_titles'])
    neg   = [re.compile(x) for x in prof['outline']['negative_patterns']]
    body_sz = int(prof['body']['size'])
    maxlen = int(prof['outline'].get('max_heading_len', 40))

    rows = []
    for f in facts:
        reg = region[f['idx']]
        if f['empty'] and reg == 'front':
            rows.append((f, 'F', 0.9, ['草稿自带前置件区域内的空段'])); continue
        if f['empty'] and reg == 'toc':
            rows.append((f, 'T', 0.99, ['目录域内的空段'])); continue
        if f['empty']: continue
        t = f['text']; lvl, conf, why = 0, 0.0, []
        if reg == 'toc':
            rows.append((f, 'T', 0.99, ['目录域内容（apply 不动，按 F9 重建）'])); continue
        if reg == 'front':
            rows.append((f, 'F', 0.9, ['草稿自带前置件（将由规范包重新生成）'])); continue
        # 负信号优先
        if any(rx.match(t) for rx in neg):
            rows.append((f, 0, 0.95, ['负信号:标签/案例行'])); continue
        # 特殊名称
        hit_name = next((k for k in names if re.fullmatch(names[k]['pattern'], t)), None)
        if hit_name:
            lvl = names[hit_name]['level']; conf = 0.97
            why.append('特殊名称:' + hit_name)
        else:
            for name, rx, sl in SCHEMES:
                m = rx.match(t)
                if not m: continue
                info = schemes.get(name)
                if not info: break
                lvl = prof['outline']['scheme_levels'].get(name, sl)
                conf = 0.9 if info['seq_ok'] else 0.6
                why.append(f'编号体系:{name}' + ('+序列合法' if info['seq_ok'] else '(序列不连续)'))
                break
        if lvl and f['text'][-1] in '。！？；':
            lvl, conf = 0, 0.93; why.append('以句末标点结尾→正文')
        elif lvl and f['length'] > maxlen:
            lvl, conf = 0, 0.9; why.append(f'过长({f["length"]}字)→正文')
        elif lvl and f['ends_punct']:
            conf -= 0.3; why.append('以标点结尾')
        # 无编号时的兜底信号
        if not lvl and not strong and f['length'] <= maxlen and not f['ends_punct']:
            s = 0.0
            if f['outline'] is not None: s += 0.5; why.append('已有outlineLvl=' + f['outline'])
            if f['size'] > body_sz: s += 0.25; why.append(f'字号{f["size"]}>正文')
            if f['bold']: s += 0.2; why.append('加粗')
            if f['jc'] == 'center': s += 0.15; why.append('居中')
            if s >= 0.65:
                lvl = (int(f['outline']) + 1) if f['outline'] is not None else 2
                conf = min(0.85, s)
        if f['outline'] is not None and lvl:
            why.append('样式outlineLvl=' + f['outline'])
        rows.append((f, lvl, max(0.0, min(conf, 0.99)), why))

    path = pj(proj, '02_outline.tsv')
    with open(path, 'w', encoding='utf-8') as fh:
        fh.write('# 标题识别提案。核对「级别」列，然后把 CONFIRMED 改成 yes，再运行 apply。\n')
        fh.write('# 级别: 1/2/3=标题层级, 0=正文, F=草稿自带前置件(apply 时移除，改由规范包生成), X=删除该段\n')
        fh.write('# CONFIRMED: no\n')
        fh.write('段号\t级别\t置信度\t当前样式\t依据\t文本\n')
        n_head = 0
        for f, lvl, conf, why in rows:
            is_h = isinstance(lvl, int) and lvl > 0
            # 候选：形状像标题但没判为标题的段（短、不以标点结尾），列出来供人工提级
            cand = (lvl == 0 and not f['empty'] and f['length'] <= maxlen
                    and not f['ends_punct'] and f['length'] >= 2)
            if cand and not why: why = ['候选：短行未判定，如是标题请改级别']
            show = is_h or lvl in ('F', 'T') or cand or f['outline'] is not None
            if not show: continue
            if is_h: n_head += 1
            fh.write(f'{f["idx"]}\t{lvl}\t{conf:.2f}\t{f["style"] or "-"}\t{"；".join(why) or "-"}\t{f["text"][:70]}\n')
    low = sum(1 for _, l, c, _ in rows if isinstance(l, int) and l > 0 and c < 0.8)
    print(f'编号体系发现: {json.dumps(schemes, ensure_ascii=False)}')
    print(f'判为标题 {n_head} 条，其中低置信度(<0.80) {low} 条')
    print(f'→ {path}')
    print(f'  请核对后把文件里的 "# CONFIRMED: no" 改为 yes，再运行:')
    print(f'  tfmt.py apply --project {proj}')
    set_state(proj, stage='outline')

def read_outline(proj):
    path = pj(proj, '02_outline.tsv')
    if not os.path.exists(path): sys.exit('缺少 02_outline.tsv，请先运行 outline')
    confirmed = False; plan = {}
    for line in open(path, encoding='utf-8'):
        if line.startswith('#'):
            if re.search(r'CONFIRMED:\s*yes', line, re.I): confirmed = True
            continue
        if line.startswith('段号'): continue
        c = line.rstrip('\n').split('\t')
        if len(c) < 2 or not c[0].strip().isdigit(): continue
        plan[int(c[0])] = c[1].strip().upper()
    if not confirmed:
        sys.exit('02_outline.tsv 尚未确认（把 "# CONFIRMED: no" 改成 yes）。已按严格拦截停止。')
    return plan

# ══════════════════════════ apply：套格式 ══════════════════════════
HDR_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
FTR_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
HDR_RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
FTR_RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'
HF_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
         'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

def hf_part(kind, inner):
    tag = 'hdr' if kind == 'hdr' else 'ftr'
    return (D.DECL + ('<w:%s %s>%s</w:%s>' % (tag, HF_NS, inner, tag)).encode('utf-8'))

def para_xml(runs_xml, jc='center', style=None):
    st = '<w:pStyle w:val="%s"/>' % style if style else ''
    return ('<w:p><w:pPr>%s<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:jc w:val="%s"/></w:pPr>%s</w:p>' % (st, jc, runs_xml))

def run_xml(text, font, sz, bold=False):
    b = '<w:b/><w:bCs/>' if bold else ''
    return ('<w:r><w:rPr><w:rFonts w:hint="eastAsia" w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
            '%s<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
            % (font, font, font, font, b, sz, sz, esc(text)))

def field_xml(instr, font, sz, pre='', post=''):
    rpr = ('<w:rPr><w:rFonts w:hint="eastAsia" w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
           '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (font, font, font, sz, sz))
    t = lambda s: ('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, esc(s))) if s else ''
    return (t(pre) + '<w:r>%s<w:fldChar w:fldCharType="begin"/></w:r>' % rpr
            + '<w:r>%s<w:instrText xml:space="preserve"> %s </w:instrText></w:r>' % (rpr, instr)
            + '<w:r>%s<w:fldChar w:fldCharType="separate"/></w:r>' % rpr
            + '<w:r>%s<w:t>1</w:t></w:r>' % rpr
            + '<w:r>%s<w:fldChar w:fldCharType="end"/></w:r>' % rpr + t(post))

def esc(s): return (s or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def ensure_hf_parts(pkg, prof, meta):
    """建立空白/题目页眉与三种页脚，返回 rId 映射。"""
    f = prof['header']['font']; sz = int(prof['header']['size'])
    title = subst(prof['header']['text'], meta)
    pf = prof['footer']['font']; pz = int(prof['footer']['size'])
    parts = {
        'hdr_blank':  ('header', hf_part('hdr', para_xml('', 'center'))),
        'hdr_title':  ('header', hf_part('hdr', para_xml(run_xml(title, f, sz), 'center'))),
        'ftr_blank':  ('footer', hf_part('ftr', para_xml('', 'center'))),
        'ftr_roman':  ('footer', hf_part('ftr', para_xml(field_xml('PAGE', pf, pz), 'center'))),
        'ftr_arabic': ('footer', hf_part('ftr', para_xml(field_xml('PAGE', pf, pz,
                        prof['footer'].get('arabic_prefix', '- '), prof['footer'].get('arabic_suffix', ' -')), 'center'))),
    }
    ids = {}
    for i, (key, (kind, data)) in enumerate(sorted(parts.items())):
        name = 'word/tfmt_%s.xml' % key
        pkg.add_part(name, data, HDR_CT if kind == 'header' else FTR_CT)
        rid = 'rIdTFMT%d' % (900 + i)
        pkg.add_rel(rid, HDR_RT if kind == 'header' else FTR_RT, os.path.basename(name))
        ids[key] = rid
    return ids

def sectpr_xml(prof, hdr_rid, ftr_rid, pgnum=None, typ='nextPage'):
    pg = prof['page']; m = pg['margin']
    num = ''
    if pgnum:
        num = '<w:pgNumType w:fmt="%s"%s/>' % (pgnum['fmt'],
              (' w:start="%d"' % pgnum['start']) if pgnum.get('start') else '')
    return ('<w:sectPr>'
            '<w:headerReference w:type="default" r:id="%s"/><w:footerReference w:type="default" r:id="%s"/>'
            '<w:type w:val="%s"/><w:pgSz w:w="%d" w:h="%d"/>'
            '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" w:header="%d" w:footer="%d" w:gutter="0"/>'
            '%s<w:cols w:space="425" w:num="1"/><w:docGrid w:linePitch="%d" w:charSpace="0"/></w:sectPr>'
            % (hdr_rid, ftr_rid, typ, pg['w'], pg['h'], m['top'], m['right'], m['bottom'], m['left'],
               pg.get('header', 851), pg.get('footer', 992), num, pg.get('line_pitch', 360)))

def parse_xml(frag):
    return ET.fromstring('<w:root %s>%s</w:root>' % (HF_NS, frag))[0]

def sect_carrier(prof, hdr, ftr, pgnum=None):
    """只承载分节符的极矮段落（避免多出空白页）。"""
    p = ET.Element(q('p')); pr = ET.SubElement(p, q('pPr'))
    sp = ET.SubElement(pr, q('spacing'))
    sp.set(q('before'), '0'); sp.set(q('after'), '0'); sp.set(q('line'), '20'); sp.set(q('lineRule'), 'exact')
    rp = ET.SubElement(pr, q('rPr'))
    ET.SubElement(rp, q('sz')).set(q('val'), '2'); ET.SubElement(rp, q('szCs')).set(q('val'), '2')
    pr.append(parse_xml(sectpr_xml(prof, hdr, ftr, pgnum)))
    return p

def subst(s, meta):
    def rep(m):
        v = meta.get(m.group(1))
        return v if v is not None else m.group(0)
    return re.sub(r'\{\{(\w+)\}\}', rep, s or '')

# ---- 前置件生成（规范包声明式） ----
def build_block(spec, prof, meta):
    """spec: 规范包里的一个前置件块，返回段落列表。"""
    out = []
    for item in spec['paras']:
        kind = item.get('type', 'text')
        if kind == 'blank':
            out.append(D.make_para(line=item.get('line', 240), after=0,
                                   runs=[D.make_run('', sz=item.get('sz', 24))]))
        elif kind == 'pagebreak':
            out.append(D.page_break_para())
        elif kind == 'field_line':
            out.append(build_field_line(item, prof, meta))
        else:
            runs = []
            for r in item['runs']:
                txt = subst(r['t'], meta)
                runs.append(D.make_run(txt, font=r.get('font'), ea=r.get('ea', r.get('font')),
                                       sz=r.get('sz'), bold=r.get('b', False), italic=r.get('i', False),
                                       hint=r.get('hint', 'eastAsia')))
            out.append(D.make_para(runs, jc=item.get('jc', 'center'),
                                   before=item.get('before', 0), after=item.get('after', 0),
                                   line=item.get('line', 360), ind=item.get('ind')))
    return out

def build_field_line(item, prof, meta):
    """封面填空行：制表符带 w:u 画线（行尾空格会被裁掉，不能用带下划线的空格）。"""
    sz = item['sz']; font = item.get('font', prof['body']['font_cn'])
    stops = [c['tab'] for c in item['cells'] if c.get('tab')]
    runs = []
    for c in item['cells']:
        if c.get('label'):
            runs.append(D.make_run(c['label'], font=font, ea=font, sz=sz))
        if c.get('value') is not None:
            v = subst(c['value'], meta)
            if v: runs.append(D.make_run(v, font=font, ea=font, sz=sz, underline=True))
        if c.get('tab'):
            runs.append(D.make_run('', font=font, ea=font, sz=sz,
                                   underline=(c.get('line', True)), tab=True))
    return D.make_para(runs, jc='left', before=0, after=item.get('after', 60),
                       line=item.get('line', 360), ind={'left': 0, 'right': 0, 'firstLine': 0},
                       tabs=[(s, 'none') for s in stops])

STYLE_IDS = {1: 'TFMTHeading1', 2: 'TFMTHeading2', 3: 'TFMTHeading3'}

def ensure_heading_styles(pkg, prof):
    root = pkg.tree('word/styles.xml')
    have = {s.get(q('styleId')) for s in root.findall(q('style'))}
    for h in prof['headings']:
        lvl = h['level']; sid = STYLE_IDS[lvl]
        if sid in have: continue
        b = '<w:b/><w:bCs/>' if h.get('bold', True) else ''
        jc = '<w:jc w:val="%s"/>' % h.get('align', 'left')
        frag = ('<w:style w:type="paragraph" w:styleId="%s"><w:name w:val="heading %d"/>'
                '<w:basedOn w:val="%s"/><w:qFormat/><w:uiPriority w:val="9"/>'
                '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="%d" w:after="%d" w:line="%d" w:lineRule="auto"/>'
                '%s<w:outlineLvl w:val="%d"/></w:pPr>'
                '<w:rPr><w:rFonts w:hint="eastAsia" w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
                '%s<w:color w:val="000000"/><w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr></w:style>'
                % (sid, lvl, D.Styles(pkg).default_pstyle or 'Normal',
                   h.get('space_before', 0), h.get('space_after', 0), prof['body']['line'],
                   jc, lvl - 1, h['font_cn'], h['font_cn'], h['font_cn'], prof['body']['font_en'],
                   b, h['size'], h['size']))
        root.append(parse_xml(frag))

def apply_heading(p, prof, lvl, special=None):
    h = next(x for x in prof['headings'] if x['level'] == lvl)
    size = (special or {}).get('size', h['size'])
    pr = D.pPr(p, create=True)
    for tag in ('pStyle','spacing','jc','ind','outlineLvl','pageBreakBefore','numPr'):
        for e in pr.findall(q(tag)): pr.remove(e)
    ET.SubElement(pr, q('pStyle')).set(q('val'), STYLE_IDS[lvl])
    if h.get('page_break_before') and lvl == 1: ET.SubElement(pr, q('pageBreakBefore'))
    sp = ET.SubElement(pr, q('spacing'))
    sp.set(q('before'), str(h.get('space_before', 0))); sp.set(q('after'), str(h.get('space_after', 0)))
    sp.set(q('line'), str(prof['body']['line'])); sp.set(q('lineRule'), 'auto')
    ET.SubElement(pr, q('jc')).set(q('val'), h.get('align', 'left'))
    for r in p.iter(q('r')):
        rp = D.rPr(r, create=True)
        for tag in ('rFonts','sz','szCs','b','bCs','color','i','iCs'):
            for e in rp.findall(q(tag)): rp.remove(e)
        f = ET.SubElement(rp, q('rFonts'))
        for k, v in (('hint','eastAsia'),('ascii',h['font_cn']),('hAnsi',h['font_cn']),
                     ('eastAsia',h['font_cn']),('cs',prof['body']['font_en'])): f.set(q(k), v)
        if h.get('bold', True): ET.SubElement(rp, q('b')); ET.SubElement(rp, q('bCs'))
        ET.SubElement(rp, q('color')).set(q('val'), '000000')
        ET.SubElement(rp, q('sz')).set(q('val'), str(size))
        ET.SubElement(rp, q('szCs')).set(q('val'), str(size))

def apply_role_title(p, prof, spec):
    """致谢/摘要/Abstract/目录 这类前置件标题：自己的字体字号，不走章标题样式。"""
    pr = D.pPr(p, create=True)
    for tag in ('pStyle','spacing','jc','ind','outlineLvl','numPr'):
        for e in pr.findall(q(tag)): pr.remove(e)
    sp = ET.SubElement(pr, q('spacing'))
    sp.set(q('before'), '120'); sp.set(q('after'), '360')
    sp.set(q('line'), str(prof['body']['line'])); sp.set(q('lineRule'), 'auto')
    ET.SubElement(pr, q('jc')).set(q('val'), spec.get('align', 'center'))
    if spec.get('toc'): ET.SubElement(pr, q('outlineLvl')).set(q('val'), '0')
    font = spec.get('font_cn', prof['body']['font_cn'])
    for r in p.iter(q('r')):
        rp = D.rPr(r, create=True)
        for tag in ('rFonts','sz','szCs','b','bCs','color'):
            for e in rp.findall(q(tag)): rp.remove(e)
        f = ET.SubElement(rp, q('rFonts'))
        for k, v in (('hint','eastAsia'),('ascii',font),('hAnsi',font),('eastAsia',font)): f.set(q(k), v)
        if spec.get('bold', True): ET.SubElement(rp, q('b')); ET.SubElement(rp, q('bCs'))
        ET.SubElement(rp, q('sz')).set(q('val'), str(spec['size']))
        ET.SubElement(rp, q('szCs')).set(q('val'), str(spec['size']))

def apply_body(p, prof):
    b = prof['body']
    pr = D.pPr(p, create=True)
    sp = D.sub(pr, 'spacing', line=str(b['line']), lineRule='auto', after='0')
    ind = pr.find(q('ind'))
    left = ind.get(q('left')) if ind is not None else None
    if not left or left in ('0',):
        if ind is None: ind = ET.SubElement(pr, q('ind'))
        ind.set(q('firstLineChars'), str(b['first_line_chars']))
        ind.set(q('firstLine'), str(int(b['first_line_chars']) * int(b['size']) // 10))
    if b.get('justify', True) and pr.find(q('jc')) is None:
        ET.SubElement(pr, q('jc')).set(q('val'), 'both')

def hygiene(pkg, prof, report):
    """通用卫生：无效字体族名、w14:textFill、全角空格填空、缩进异值。"""
    n_font = n_tf = 0
    for name in [x for x in pkg.names if re.match(r'word/(document|styles|header\S*|footer\S*|footnotes|endnotes|numbering)\.xml$', x)]:
        root = pkg.tree(name); pm = None
        for f in root.iter(q('rFonts')):
            for at in ('ascii','hAnsi','cs','eastAsia'):
                v = f.get(q(at))
                if not v: continue
                base = next((b for b in prof['hygiene']['font_families'] if v != b and v.startswith(b + ' ')), None)
                if base:
                    rp = None
                    f.set(q(at), base); n_font += 1
        if name == 'word/styles.xml':
            pm = D.parent_map(root)
            for e in [x for x in root.iter() if D.ln(x) == 'textFill']:
                if pm.get(e) is not None: pm[e].remove(e); n_tf += 1
    if n_font: report.append(f'字体族名归一 {n_font} 处（如 "Times New Roman Regular" → "Times New Roman"）')
    if n_tf:   report.append(f'清除 w14:textFill {n_tf} 处（换命名空间才删得掉，会覆盖 w:color）')

def find_role_paras(paras, prof):
    roles = {}
    for i, p in enumerate(paras):
        t = D.para_text(p).strip()
        for role, spec in prof['outline']['special_titles'].items():
            if re.fullmatch(spec['pattern'], t) and role not in roles:
                roles[role] = i
    return roles

def cmd_apply(a):
    proj = a.project
    prof = load_profile(state(proj).get('profile', 'tfsu-mti'))
    meta = load_json(pj(proj, 'metadata.json'), {})
    plan = read_outline(proj)
    pkg = D.Package(pj(proj, '01_accepted.docx'))
    root = pkg.tree('word/document.xml'); body = D.body(root)
    paras = D.paragraphs(root)
    keep_text_before = [D.para_text(p) for i, p in enumerate(paras) if plan.get(i, '0') not in ('X', 'F')]
    report = []

    # 0 清除草稿自带的段内分节符（与新分节冲突；正文末尾的 sectPr 稍后整体替换）
    n_old = 0
    for pp in D.all_paragraphs(root):
        pr = pp.find(q('pPr'))
        if pr is not None and pr.find(q('sectPr')) is not None:
            pr.remove(pr.find(q('sectPr'))); n_old += 1
    if n_old: report.append(f'清除草稿自带分节符 {n_old} 处')

    # 1 删除 X / F 段
    drop = [p for i, p in enumerate(paras) if plan.get(i) in ('X', 'F')]
    for p in drop: body.remove(p)
    if drop: report.append(f'按提案移除草稿段落 {len(drop)} 段（X 删除 / F 前置件将由规范包重建）')
    paras = D.paragraphs(root)
    remap = {}
    k = 0
    for i, p in enumerate(D.paragraphs(D.Package(pj(proj, '01_accepted.docx')).tree('word/document.xml'))):
        if plan.get(i) in ('X', 'F'): continue
        remap[i] = k; k += 1

    # 2 标题与正文
    ensure_heading_styles(pkg, prof)
    specials = prof['outline']['special_titles']
    n_h = 0
    for old_i, lvl in plan.items():
        if lvl not in ('1', '2', '3'): continue
        if old_i not in remap: continue
        p = paras[remap[old_i]]
        t = D.para_text(p).strip()
        sp = next((v for v in specials.values() if re.fullmatch(v['pattern'], t)), None)
        if sp and sp.get('role_class') == 'front':
            apply_role_title(p, prof, sp)
        else:
            apply_heading(p, prof, int(lvl), sp)
        n_h += 1
    heads = set()
    for i, v in plan.items():
        if v not in ('1','2','3') or i not in remap: continue
        heads.add(remap[i])
    toc_rows = {remap[i] for i, v in plan.items() if v == 'T' and i in remap}
    for j, p in enumerate(paras):
        if j in heads or j in toc_rows: continue
        if not D.para_text(p).strip(): continue
        apply_body(p, prof)
    report.append(f'套用标题格式 {n_h} 条；正文段落归一 {len(paras)-len(heads)} 段')

    # 3 前置件
    ids = ensure_hf_parts(pkg, prof, meta)
    blocks = []
    fm = prof.get('front_matter', {})
    order = prof.get('front_matter_order', [])
    for role in order:
        if role in fm: blocks.append((role, build_block(fm[role], prof, meta)))
    inserted = 0
    block_start = {}
    sect_roles = {sp['break_before'] for sp in prof['sections']}
    for bi, (role, ps) in enumerate(blocks):
        block_start[role] = inserted
        if bi > 0 and role not in sect_roles and ps:
            D.sub(D.pPr(ps[0], create=True), 'pageBreakBefore')
        for p in ps:
            body.insert(inserted, p); inserted += 1
    if blocks: report.append('生成前置件: ' + '、'.join(r for r, _ in blocks))

    # 4 分节
    paras = D.paragraphs(root)
    roles = find_role_paras(paras, prof)
    first_h1 = next((j for j, p in enumerate(paras) if D.style_id(p) == STYLE_IDS[1]), None)
    plan_secs = prof['sections']
    anchors = {}
    for spec in plan_secs:
        key = spec['break_before']
        if key == 'FIRST_H1': anchors[key] = first_h1
        elif key in block_start: anchors[key] = block_start[key]
        elif key in roles: anchors[key] = roles[key]
    n_sec = 0
    for spec in plan_secs:
        j = anchors.get(spec['break_before'])
        if j is None: continue
        carrier = sect_carrier(prof, ids[spec['header']], ids[spec['footer']], spec.get('pgnum'))
        body.insert(list(body).index(paras[j]), carrier); n_sec += 1
        paras = D.paragraphs(root); roles = find_role_paras(paras, prof)
        first_h1 = next((k2 for k2, p in enumerate(paras) if D.style_id(p) == STYLE_IDS[1]), None)
        bs2 = {r: next((k2 for k2, pp2 in enumerate(paras) if pp2 is blocks[bi2][1][0]), None)
               for bi2, (r, _) in enumerate(blocks) if blocks[bi2][1]}
        for s2 in plan_secs:
            kk = s2['break_before']
            if kk == 'FIRST_H1': anchors[kk] = first_h1
            elif kk in bs2 and bs2[kk] is not None: anchors[kk] = bs2[kk]
            elif kk in roles: anchors[kk] = roles[kk]
    # 末节
    last = body.find(q('sectPr'))
    if last is not None: body.remove(last)
    tail = prof['tail_section']
    body.append(parse_xml(sectpr_xml(prof, ids[tail['header']], ids[tail['footer']], tail.get('pgnum'), 'nextPage')))
    report.append(f'建立分节 {n_sec+1} 节（页眉页脚与罗马/阿拉伯页码按规范包）')

    # 5 前置件各自另起一页
    paras = D.paragraphs(root); roles = find_role_paras(paras, prof)
    for role in prof['outline'].get('page_break_roles', []):
        j = roles.get(role)
        if j is None or j == 0: continue
        prev = paras[j-1]
        if prev.find(q('pPr')) is not None and prev.find(q('pPr')).find(q('sectPr')) is not None: continue
        if any(b.get(q('type')) == 'page' for b in prev.iter(q('br'))): continue
        D.sub(D.pPr(paras[j], create=True), 'pageBreakBefore')

    kids = list(body); rm = 0
    for i, el in enumerate(kids):
        if el.tag != q('p') or i == 0: continue
        pr = el.find(q('pPr'))
        if pr is None or pr.find(q('sectPr')) is None: continue
        prev = kids[i-1]
        if prev.tag == q('p') and not D.para_text(prev).strip() and any(
                b.get(q('type')) == 'page' for b in prev.iter(q('br'))):
            body.remove(prev); rm += 1
    if rm: report.append(f'清除与分节符重复的分页空段 {rm} 处')

    # 目录域尾部那个只含 fldChar end 的空段会占一行、把分节符挤到下一页 → 压到最小高度
    kids = list(body); shrunk = 0
    for i, el in enumerate(kids):
        if el.tag != q('p') or i == 0: continue
        pr = el.find(q('pPr'))
        if pr is None or pr.find(q('sectPr')) is None: continue
        j = i - 1
        while j >= 0 and kids[j].tag == q('p') and not D.para_text(kids[j]).strip():
            prev = kids[j]
            ppr = D.pPr(prev, create=True)
            D.sub(ppr, 'spacing', before='0', after='0', line='20', lineRule='exact')
            rp = ppr.find(q('rPr'))
            if rp is None: rp = ET.SubElement(ppr, q('rPr'))
            D.sub(rp, 'sz', val='2'); D.sub(rp, 'szCs', val='2')
            for r in prev.iter(q('r')):
                rr = D.rPr(r, create=True); D.sub(rr, 'sz', val='2'); D.sub(rr, 'szCs', val='2')
            shrunk += 1; j -= 1
    if shrunk: report.append(f'压缩分节符前的空段 {shrunk} 处（目录域尾段会挤出空白页）')

    # 分节符（nextPage）已经换页，紧随其后的段落若还带 pageBreakBefore 就会多出一整页空白
    kids = list(body); dedup = 0
    for i, el in enumerate(kids[:-1]):
        pr = el.find(q('pPr')) if el.tag == q('p') else None
        if pr is None or pr.find(q('sectPr')) is None: continue
        nxt = kids[i+1]
        npr = nxt.find(q('pPr')) if nxt.tag == q('p') else None
        if npr is not None and npr.find(q('pageBreakBefore')) is not None:
            npr.remove(npr.find(q('pageBreakBefore'))); dedup += 1
    if dedup: report.append(f'去掉与分节符重复的 pageBreakBefore {dedup} 处')

    hygiene(pkg, prof, report)
    out = pj(proj, '03_formatted.docx'); pkg.save(out)

    # 6 自证：草稿内容一字不动
    r2 = D.Package(out).tree('word/document.xml')
    new_texts = [D.para_text(p) for p in D.paragraphs(r2)]
    tail_new = [t for t in new_texts if t.strip()]
    keep = [t for t in keep_text_before if t.strip()]
    ok = len(tail_new) >= len(keep) and tail_new[len(tail_new) - len(keep):] == keep
    # 未被 metadata 填上的占位符不能混进成品
    leftover = sorted(set(re.findall(r'\{\{(\w+)\}\}', ''.join(new_texts))))
    if leftover:
        report.append('！metadata.json 里缺这些字段，占位符原样留在了成品里: ' + '、'.join(leftover))
    print('\n'.join('  ' + r for r in report))
    print(f'自证（草稿正文逐段原样保留）: {"✓ " + str(len(keep)) + " 段" if ok else "✗"}')
    if not ok:
        n = next((i for i in range(min(len(keep), len(tail_new)))
                  if tail_new[len(tail_new)-len(keep)+i] != keep[i]), None)
        sys.exit(f'！正文第 {n} 段发生了非预期变化，已停止。')
    if leftover: sys.exit('！占位符未解析，请补全 metadata.json 后重跑 apply。')
    set_state(proj, stage='applied')
    print(f'→ {out}\n  下一步: tfmt.py audit --project {proj} --docx {out}')

# ══════════════════════════ audit：三轮审查 ══════════════════════════
def cmd_audit(a):
    prof = load_profile(state(a.project).get('profile', 'tfsu-mti') if a.project else a.profile)
    src = a.docx or pj(a.project, '03_formatted.docx')
    pkg = D.Package(src); root = pkg.tree('word/document.xml'); st = D.Styles(pkg)
    paras = D.paragraphs(root); T = [D.para_text(p) for p in paras]
    rels = pkg.rels(); issues = []; lines = []
    def bad(cat, msg): issues.append((cat, msg))
    def out(s): lines.append(s)

    out('## 第一轮 结构 / 分节 / 页码 / 页眉')
    pg = prof['page']; m = pg['margin']
    for i, s in enumerate(root.iter(q('sectPr')), 1):
        d = {D.ln(c): {k.split('}')[-1]: v for k, v in c.attrib.items()} for c in s}
        sz = d.get('pgSz', {}); mar = d.get('pgMar', {})
        hdr = rels.get(d.get('headerReference', {}).get('id', ''), '-')
        ftr = rels.get(d.get('footerReference', {}).get('id', ''), '-')
        out(f'- 节{i}: {sz.get("w")}×{sz.get("h")} 边距 L{mar.get("left")} R{mar.get("right")} '
            f'T{mar.get("top")} B{mar.get("bottom")} 页眉={hdr} 页脚={ftr} 页码={d.get("pgNumType")}')
        if (sz.get('w'), sz.get('h')) != (str(pg['w']), str(pg['h'])): bad('结构', f'节{i} 纸张不符')
        if (mar.get('left'), mar.get('right'), mar.get('top'), mar.get('bottom')) != \
           (str(m['left']), str(m['right']), str(m['top']), str(m['bottom'])): bad('结构', f'节{i} 页边距不符')
    for i, p in enumerate(paras):
        if D.style_id(p) != STYLE_IDS[1]: continue
        prev = paras[i-1] if i else None
        okp = (p.find(q('pPr')) is not None and p.find(q('pPr')).find(q('pageBreakBefore')) is not None) or \
              (prev is not None and prev.find(q('pPr')) is not None and prev.find(q('pPr')).find(q('sectPr')) is not None) or \
              (prev is not None and any(b.get(q('type')) == 'page' for b in prev.iter(q('br'))))
        if not okp: bad('结构', f'「{T[i].strip()[:18]}」未另页开始')

    out('\n## 第二轮 字体 / 字号 / 行距')
    for h in prof['headings']:
        lvl = h['level']; ps = [i for i, p in enumerate(paras) if D.style_id(p) == STYLE_IDS[lvl]]
        n = 0
        for i in ps:
            p = paras[i]; t = T[i].strip()
            rs = [r for r in p.iter(q('r')) if D.run_text(r).strip()]
            if not rs: continue
            sp = next((v for v in prof['outline']['special_titles'].values()
                       if re.fullmatch(v['pattern'], t) and v.get('size')), None)
            want = str((sp or {}).get('size', h['size']))
            probs = []
            if {st.rval(p, r, 'sz') for r in rs} != {want}: probs.append('字号')
            if {st.font(p, r, 'eastAsia') for r in rs} != {h['font_cn']}: probs.append('中文字体')
            if h.get('bold', True) and not all(st.bold(p, r) for r in rs): probs.append('加粗')
            if (st.pval(p, 'jc', 'val') or 'left') != h.get('align', 'left'): probs.append('对齐')
            if h.get('space_after') is not None and st.pval(p, 'spacing', 'after') != str(h['space_after']):
                probs.append('段后间距')
            if probs: n += 1; bad('字体', f'{lvl}级「{t[:16]}」: {"、".join(probs)}')
        out(f'- {lvl} 级标题 {len(ps)} 条，问题 {n} 条')
    from collections import Counter
    ln_c, sz_c = Counter(), Counter()
    heads = {i for i, p in enumerate(paras) if D.style_id(p) in STYLE_IDS.values()}
    first_h1 = next((i for i, p in enumerate(paras) if D.style_id(p) == STYLE_IDS[1]), 0)
    toc_rows = {i for i, p in enumerate(paras)
                if (st.name.get(D.style_id(p) or '', '') or '').lower().startswith('toc ')}
    for i, p in enumerate(paras):
        # 只统计"正文"：第一章及以后、非标题、非目录域
        if i < first_h1 or i in heads or i in toc_rows or not T[i].strip(): continue
        ln_c[(st.pval(p, 'spacing', 'line'), st.pval(p, 'spacing', 'lineRule'))] += 1
        for r in p.iter(q('r')):
            if D.run_text(r).strip(): sz_c[st.rval(p, r, 'sz')] += 1
    out(f'- 正文行距分布 {dict(ln_c)}；字号分布 {dict(sz_c)}')
    want_line = (str(prof['body']['line']), 'auto')
    if any(k != want_line for k in ln_c): bad('字体', f'正文存在非规定行距 {dict(ln_c)}')
    if any(k != str(prof['body']['size']) for k in sz_c): bad('字体', f'正文存在非规定字号 {dict(sz_c)}')
    allx = b''.join(pkg.raw(n) for n in pkg.names if n.endswith('.xml'))
    inv = set()
    for base in prof['hygiene']['font_families']:
        inv |= set(re.findall((r'w:(?:ascii|hAnsi|cs|eastAsia)="(%s [^"]+)"' % re.escape(base)).encode(), allx))
    if inv: bad('字体', '无效字体族名: ' + '、'.join(sorted(x.decode() for x in inv)))
    out(f'- 无效字体族名: {sorted(x.decode() for x in inv) or "无"}')

    out('\n## 第三轮 目录 / 内容一致性')
    toc_rows = [i for i, p in enumerate(paras) if (st.name.get(D.style_id(p) or '', '') or '').lower().startswith('toc ')]
    labels = [re.sub(r'[0-9IVXLCⅠ-Ⅹ]+$', '', T[i].strip()).strip() for i in toc_rows]
    head_txt = [T[i].strip() for i in sorted(heads)]
    miss = [h for h in head_txt if h not in labels]
    out(f'- 目次 {len(toc_rows)} 条；正文标题 {len(head_txt)} 条；目录未收录 {len(miss)} 条')
    if toc_rows and miss: bad('目录', f'目录缺 {len(miss)} 条（需在 Word/WPS 里按 F9 更新）')
    if toc_rows: out('- 提醒：目录页码是域的缓存值，排版改动后必须在 WPS/Word 里全选按 F9 重新更新')
    if not toc_rows: bad('目录', '未发现目录域（apply 会插入，或需手动插入并 F9）')
    for name, rx in prof['checks'].get('text_patterns', {}).items():
        hits = [i for i, t in enumerate(T) if re.search(rx, t)]
        if hits: bad('内容', f'{name}: {len(hits)} 处（段 {hits[:6]}）')
    body_cjk = sum(len(CJK.findall(t)) for t in T)
    lim = prof.get('limits', {})
    if lim.get('body_cjk_min') and body_cjk < lim['body_cjk_min']:
        bad('内容', f'全文汉字 {body_cjk} < 规范下限 {lim["body_cjk_min"]}')
    out(f'- 全文汉字数 {body_cjk}')

    out(f'\n## 问题汇总 {len(issues)} 项')
    for c, mm in issues: out(f'- [{c}] {mm}')
    txt = '\n'.join(lines)
    if a.project:
        p = pj(a.project, 'report_audit.md'); open(p, 'w', encoding='utf-8').write('# 规范符合性审查\n\n' + txt + '\n')
        print(txt); print(f'\n→ {p}')
    else:
        print(txt)
    if issues and a.strict: sys.exit(1)

# ══════════════════════════ pdfcheck：核验导出的 PDF ══════════════════════════
def cmd_pdfcheck(a):
    try:
        import fitz
    except ImportError:
        sys.exit('需要 PyMuPDF：pip install pymupdf（这是唯一的可选依赖，核心流程不需要）')
    pkg = D.Package(a.docx); root = pkg.tree('word/document.xml'); st = D.Styles(pkg)
    paras = D.paragraphs(root); T = [D.para_text(p) for p in paras]
    toc = {}
    for i, p in enumerate(paras):
        if (st.name.get(D.style_id(p) or '', '') or '').lower().startswith('toc '):
            s = T[i].strip(); m = re.search(r'([0-9]+|[IVXLC]+)$', s)
            if m: toc[re.sub(r'([0-9]+|[IVXLC]+)$', '', s).strip()] = m.group(1)
    heads = [T[i].strip() for i, p in enumerate(paras) if D.style_id(p) in STYLE_IDS.values()]
    d = fitz.open(a.pdf)
    def footer(i):
        pg = d[i]
        for b in pg.get_text('dict')['blocks']:
            for l in b.get('lines', []):
                if l['bbox'][1] > pg.rect.height - 70:
                    return ''.join(s['text'] for s in l['spans']).strip()
        return ''
    pm = {}
    for i in range(d.page_count):
        v = footer(i)
        m = re.fullmatch(r'[-—]\s*(\d+)\s*[-—]', v) or re.fullmatch(r'(\d+)', v)
        if m: pm[i] = m.group(1)
        elif re.fullmatch(r'[IVXLCⅠ-Ⅹ]+', v): pm[i] = v
    body_start = min([i for i, v in pm.items() if v.isdigit()], default=0)
    found = {}
    for i in range(body_start, d.page_count):
        for b in d[i].get_text('dict')['blocks']:
            for l in b.get('lines', []):
                t = ''.join(s['text'] for s in l['spans']).strip()
                big = max((s['size'] for s in l['spans']), default=0)
                if big >= 13.5 and t in heads and t not in found: found[t] = pm.get(i, '?')
    okn, bad_rows = 0, []
    for h in heads:
        w, g = toc.get(h), found.get(h)
        if w is None: continue
        if w == g: okn += 1
        else: bad_rows.append((h, w, g))
    blanks = [i + 1 for i in range(d.page_count) if not d[i].get_text().strip()]
    fonts = sorted({f[3] for i in range(d.page_count) for f in d[i].get_fonts(full=False)})
    print(f'PDF 页数 {d.page_count}；空白页 {blanks or "无"}')
    print(f'页码序列抽样: ' + ', '.join(f'p{i+1}={pm.get(i,"无")}' for i in list(range(min(12, d.page_count)))))
    print(f'嵌入字体: {fonts}')
    print(f'目录页码 vs 实际分页: 一致 {okn}，不一致 {len(bad_rows)}')
    for h, w, g in bad_rows[:20]: print(f'   {h[:30]:<32} 目录={w} 实际={g}')
    if bad_rows:
        print('\n！不一致说明 PDF 与目录不是同一排版引擎产出，或导出前没按 F9。')
        sys.exit(1)

# ══════════════════════════ status / main ══════════════════════════
def cmd_status(a):
    s = state(a.project)
    print(json.dumps(s, ensure_ascii=False, indent=2))
    for f in ('00_input.docx','01_accepted.docx','02_outline.tsv','03_formatted.docx','report_audit.md'):
        p = pj(a.project, f)
        print(f'  {"✓" if os.path.exists(p) else "·"} {f}')

def main():
    ap = argparse.ArgumentParser(prog='tfmt.py', description='论文草稿 docx → 规范格式 docx')
    sub = ap.add_subparsers(dest='cmd', required=True)
    p = sub.add_parser('init');     p.add_argument('project'); p.add_argument('--draft', required=True); p.add_argument('--profile', default='tfsu-mti'); p.set_defaults(f=cmd_init)
    p = sub.add_parser('accept');   p.add_argument('--project', required=True); p.set_defaults(f=cmd_accept)
    p = sub.add_parser('outline');  p.add_argument('--project', required=True); p.add_argument('--profile'); p.set_defaults(f=cmd_outline)
    p = sub.add_parser('apply');    p.add_argument('--project', required=True); p.set_defaults(f=cmd_apply)
    p = sub.add_parser('audit');    p.add_argument('--project'); p.add_argument('--docx'); p.add_argument('--profile', default='tfsu-mti'); p.add_argument('--strict', action='store_true'); p.set_defaults(f=cmd_audit)
    p = sub.add_parser('pdfcheck'); p.add_argument('--docx', required=True); p.add_argument('--pdf', required=True); p.set_defaults(f=cmd_pdfcheck)
    p = sub.add_parser('status');   p.add_argument('--project', required=True); p.set_defaults(f=cmd_status)
    a = ap.parse_args()
    a.f(a)

if __name__ == '__main__':
    main()
