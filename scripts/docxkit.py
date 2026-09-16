# -*- coding: utf-8 -*-
"""OOXML 内核：只用 Python 3 标准库，可整目录拷给其他 agent 使用。

设计要点（都是踩过的坑）：
- ElementTree 不注册命名空间会把前缀改成 ns0/ns1，mc:Ignorable 会失配 → 写回前先从原始字节里
  把所有 xmlns 前缀注册一遍。
- ElementTree 没有 getparent() → 需要时построй parent map。
- rFonts 的 ascii/hAnsi/eastAsia/cs 是**逐属性**继承的，不能"找到 rFonts 就返回"。
- 同一个 pStyle 下字号可以合法地不同（参考文献/附录标题四号 vs 章标题三号）。
"""
import re, zipfile, shutil
import xml.etree.ElementTree as ET

W   = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML = '{http://www.w3.org/XML/1998/namespace}'
def q(tag):  return W + tag
def ln(el):  return el.tag.split('}')[-1] if isinstance(el.tag, str) else ''

DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

def register_namespaces(raw: bytes):
    """把文档头部声明的所有前缀注册给 ElementTree，避免写回时前缀被改写。"""
    head = raw[:20000]
    for m in re.finditer(rb'xmlns:([A-Za-z0-9_.\-]+)\s*=\s*"([^"]+)"', head):
        try: ET.register_namespace(m.group(1).decode(), m.group(2).decode())
        except Exception: pass

def root_xmlns(raw: bytes):
    """取根元素上声明的全部 xmlns 前缀。ElementTree 只会输出树里实际用到的那些，
    但 mc:Ignorable 等属性值会引用未被元素使用的前缀，丢了 Word 会判定文件损坏。"""
    m = re.search(rb'<[A-Za-z_][^<>]*?>', raw)
    if not m: return {}
    tag = m.group(0).decode('utf-8', 'ignore')
    return dict(re.findall(r'xmlns:([A-Za-z0-9_.\-]+)\s*=\s*"([^"]+)"', tag))

def restore_xmlns(xml_text: str, decls: dict):
    """把 flush 时被 ElementTree 省略掉的 xmlns 声明补回根元素。"""
    if not decls: return xml_text
    m = re.search(r'<[A-Za-z_][^<>]*?>', xml_text)
    if not m: return xml_text
    tag = m.group(0)
    have = set(re.findall(r'xmlns:([A-Za-z0-9_.\-]+)\s*=', tag))
    add = ''.join(' xmlns:%s="%s"' % (k, v) for k, v in decls.items() if k not in have)
    if not add: return xml_text
    close = ' />' if tag.endswith('/>') else '>'
    body = tag[:-3] if tag.endswith('/>') else tag[:-1]
    return xml_text[:m.start()] + body + add + close + xml_text[m.end():]

class Package:
    """整个 .docx 读入内存，改完整体写出。"""
    def __init__(self, path):
        self.path = path
        self.names, self.parts = [], {}
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if n.endswith('/'): continue
                self.names.append(n); self.parts[n] = z.read(n)
        self._trees = {}
        self._decls = {}

    def has(self, name): return name in self.parts
    def raw(self, name): return self.parts[name]
    def set_raw(self, name, data):
        if name not in self.parts: self.names.append(name)
        self.parts[name] = data
        self._trees.pop(name, None)

    def tree(self, name):
        """解析并缓存一个 XML 部件（返回 root）。"""
        if name not in self._trees:
            raw = self.parts[name]
            register_namespaces(raw)
            self._decls[name] = root_xmlns(raw)
            self._trees[name] = ET.fromstring(raw)
        return self._trees[name]

    def flush(self):
        for name, root in self._trees.items():
            data = ET.tostring(root, encoding='unicode')
            data = restore_xmlns(data, self._decls.get(name, {}))
            self.parts[name] = DECL + data.encode('utf-8')
        self._trees = {}

    def save(self, out):
        self.flush()
        tmp = out + '.tmp'
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
            for n in self.names: zo.writestr(n, self.parts[n])
        shutil.move(tmp, out)

    def drop(self, name):
        if name in self.parts:
            self.names.remove(name); self.parts.pop(name); self._trees.pop(name, None)

    # ---- 部件级便捷操作 ----
    def drop_part_refs(self, base):
        """从 [Content_Types].xml 与所有 .rels 里摘掉某个部件的登记。"""
        ct = self.parts['[Content_Types].xml'].decode('utf-8')
        ct = re.sub(r'<Override[^>]*PartName="/word/%s\.xml"[^>]*/>' % re.escape(base), '', ct)
        self.set_raw('[Content_Types].xml', ct.encode('utf-8'))
        for n in list(self.names):
            if n.endswith('.rels'):
                s = self.parts[n].decode('utf-8')
                s2 = re.sub(r'<Relationship[^>]*Target="%s\.xml"[^>]*/>' % re.escape(base), '', s)
                if s2 != s: self.set_raw(n, s2.encode('utf-8'))

    def add_part(self, name, data, content_type):
        self.set_raw(name, data)
        ct = self.parts['[Content_Types].xml'].decode('utf-8')
        if name not in ct:
            ct = ct.replace('</Types>', '<Override PartName="/%s" ContentType="%s"/></Types>' % (name, content_type))
            self.set_raw('[Content_Types].xml', ct.encode('utf-8'))

    def add_rel(self, rid, rtype, target):
        n = 'word/_rels/document.xml.rels'
        s = self.parts[n].decode('utf-8')
        if 'Id="%s"' % rid not in s:
            s = s.replace('</Relationships>',
                '<Relationship Id="%s" Type="%s" Target="%s"/></Relationships>' % (rid, rtype, target))
            self.set_raw(n, s.encode('utf-8'))

    def rels(self):
        s = self.parts['word/_rels/document.xml.rels'].decode('utf-8')
        return dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', s))

# ---------- 遍历与文本 ----------
def body(root): return root.find(q('body'))
def paragraphs(root):
    """正文顺序的段落（不含表格内段落；表格单独处理）。"""
    return [p for p in body(root) if p.tag == q('p')]
def all_paragraphs(root):
    return list(body(root).iter(q('p')))

def para_text(p):
    out = []
    for n in p.iter():
        if n.tag == q('t'): out.append(n.text or '')
        elif n.tag == q('tab'): out.append('\t')
        elif n.tag == q('br'): out.append('\n')
    return ''.join(out)

def run_text(r):
    return ''.join(t.text or '' for t in r.iter(q('t')))

def doc_text(root):
    """全文纯文本，用于"零丢字"自证。"""
    return ''.join(t.text or '' for t in root.iter(q('t')))

def final_view_text(root):
    """修订"最终视图"文本：跳过 w:del / w:moveFrom 子树。"""
    out = []
    def walk(n, dl):
        if n.tag in (q('del'), q('moveFrom')): dl = True
        if n.tag == q('t') and not dl: out.append(n.text or '')
        for c in n: walk(c, dl)
    walk(root, False)
    return ''.join(out)

def parent_map(root):
    return {c: p for p in root.iter() for c in p}

# ---------- 属性读写 ----------
def pPr(p, create=False):
    e = p.find(q('pPr'))
    if e is None and create:
        e = ET.Element(q('pPr')); p.insert(0, e)
    return e

def rPr(r, create=False):
    e = r.find(q('rPr'))
    if e is None and create:
        e = ET.Element(q('rPr')); r.insert(0, e)
    return e

def style_id(p):
    pr = p.find(q('pPr'))
    if pr is None: return None
    s = pr.find(q('pStyle'))
    return s.get(q('val')) if s is not None else None

def sub(parent, tag, **attrs):
    e = parent.find(q(tag))
    if e is None: e = ET.SubElement(parent, q(tag))
    for k, v in attrs.items(): e.set(q(k), v)
    return e

def set_attrs(el, **attrs):
    for k, v in attrs.items():
        if v is None: el.attrib.pop(q(k), None)
        else: el.set(q(k), v)

# ---------- 有效格式解析 ----------
class Styles:
    def __init__(self, pkg):
        self.root = pkg.tree('word/styles.xml')
        self.by_id = {s.get(q('styleId')): s for s in self.root.findall(q('style'))}
        self.name = {}
        for sid, s in self.by_id.items():
            n = s.find(q('name'))
            self.name[sid] = n.get(q('val')) if n is not None else ''
        dd = self.root.find(q('docDefaults'))
        self.dd_rPr = dd.find(q('rPrDefault')).find(q('rPr')) if dd is not None and dd.find(q('rPrDefault')) is not None else None
        self.dd_pPr = dd.find(q('pPrDefault')).find(q('pPr')) if dd is not None and dd.find(q('pPrDefault')) is not None else None
        self.default_pstyle = next((sid for sid, s in self.by_id.items()
                                    if s.get(q('type')) == 'paragraph' and s.get(q('default')) == '1'), None)

    def chain(self, sid):
        out, seen = [], set()
        s = self.by_id.get(sid)
        while s is not None and id(s) not in seen:
            seen.add(id(s)); out.append(s)
            b = s.find(q('basedOn'))
            s = self.by_id.get(b.get(q('val'))) if b is not None else None
        return out

    def _rpr_sources(self, p, r):
        src = []
        if r is not None and r.find(q('rPr')) is not None: src.append(r.find(q('rPr')))
        sid = style_id(p) or self.default_pstyle
        for s in self.chain(sid):
            e = s.find(q('rPr'))
            if e is not None: src.append(e)
        if self.dd_rPr is not None: src.append(self.dd_rPr)
        return src

    def font(self, p, r, attr):
        """rFonts 的单个属性逐级继承。"""
        for x in self._rpr_sources(p, r):
            f = x.find(q('rFonts'))
            if f is not None and f.get(q(attr)): return f.get(q(attr))
        return None

    def rval(self, p, r, tag):
        for x in self._rpr_sources(p, r):
            e = x.find(q(tag))
            if e is not None:
                v = e.get(q('val'))
                return 'on' if v is None else v
        return None

    def bold(self, p, r):
        return self.rval(p, r, 'b') not in (None, '0', 'false')

    def pval(self, p, tag, attr):
        pr = p.find(q('pPr'))
        if pr is not None:
            e = pr.find(q(tag))
            if e is not None and e.get(q(attr)) is not None: return e.get(q(attr))
        sid = style_id(p) or self.default_pstyle
        for s in self.chain(sid):
            pp = s.find(q('pPr'))
            if pp is not None:
                e = pp.find(q(tag))
                if e is not None and e.get(q(attr)) is not None: return e.get(q(attr))
        if self.dd_pPr is not None:
            e = self.dd_pPr.find(q(tag))
            if e is not None and e.get(q(attr)) is not None: return e.get(q(attr))
        return None

# ---------- 构造 ----------
def make_run(text, font=None, ea=None, sz=None, bold=False, italic=False,
             underline=False, hint='eastAsia', color=None, tab=False):
    r = ET.Element(q('r'))
    pr = ET.SubElement(r, q('rPr'))
    f = ET.SubElement(pr, q('rFonts'))
    if hint: f.set(q('hint'), hint)
    if font: f.set(q('ascii'), font); f.set(q('hAnsi'), font); f.set(q('cs'), font)
    if ea:   f.set(q('eastAsia'), ea)
    if bold:   ET.SubElement(pr, q('b')); ET.SubElement(pr, q('bCs'))
    if italic: ET.SubElement(pr, q('i')); ET.SubElement(pr, q('iCs'))
    if underline: ET.SubElement(pr, q('u')).set(q('val'), 'single')
    if color: ET.SubElement(pr, q('color')).set(q('val'), color)
    if sz:
        ET.SubElement(pr, q('sz')).set(q('val'), str(sz))
        ET.SubElement(pr, q('szCs')).set(q('val'), str(sz))
    if tab:
        ET.SubElement(r, q('tab'))
    else:
        t = ET.SubElement(r, q('t')); t.text = text or ''
        t.set(XML + 'space', 'preserve')
    return r

def make_para(runs=(), jc=None, before=None, after=None, line=None, rule='auto',
              ind=None, style=None, page_break_before=False, tabs=None, keep_next=False):
    p = ET.Element(q('p'))
    pr = ET.SubElement(p, q('pPr'))
    if style: ET.SubElement(pr, q('pStyle')).set(q('val'), style)
    if keep_next: ET.SubElement(pr, q('keepNext'))
    if page_break_before: ET.SubElement(pr, q('pageBreakBefore'))
    if tabs:
        tb = ET.SubElement(pr, q('tabs'))
        for pos, leader in tabs:
            e = ET.SubElement(tb, q('tab')); e.set(q('val'), 'left'); e.set(q('pos'), str(pos))
            if leader and leader != 'none': e.set(q('leader'), leader)
    sp = ET.SubElement(pr, q('spacing'))
    sp.set(q('before'), str(before if before is not None else 0))
    sp.set(q('after'),  str(after  if after  is not None else 0))
    if line: sp.set(q('line'), str(line)); sp.set(q('lineRule'), rule)
    if ind:
        e = ET.SubElement(pr, q('ind'))
        for k, v in ind.items():
            if v is not None: e.set(q(k), str(v))
    if jc: ET.SubElement(pr, q('jc')).set(q('val'), jc)
    for r in runs: p.append(r)
    return p

def page_break_para():
    p = make_para(line=240)
    r = ET.SubElement(p, q('r')); ET.SubElement(r, q('br')).set(q('type'), 'page')
    return p

def split_run(p, r, offsets, make_piece):
    """把 run 按段内偏移切开；offsets 为 {run内偏移: 标记}。
    make_piece(text, mark) 返回新的 run。原 run 就地替换。"""
    ts = list(r.iter(q('t')))
    if len(ts) != 1: raise ValueError('run 含多个 w:t，无法安全切分')
    s = ts[0].text or ''
    segs, cur = [], 0
    for k in sorted(offsets):
        if k > cur: segs.append((s[cur:k], None))
        segs.append((s[k], offsets[k])); cur = k + 1
    if cur < len(s): segs.append((s[cur:], None))
    idx = list(p).index(r)
    for j, (text, mark) in enumerate(segs):
        p.insert(idx + j, make_piece(text, mark))
    p.remove(r)

def clone_run(r, text=None):
    import copy
    nr = copy.deepcopy(r)
    if text is not None:
        ts = list(nr.iter(q('t')))
        if ts:
            ts[0].text = text; ts[0].set(XML + 'space', 'preserve')
    return nr
