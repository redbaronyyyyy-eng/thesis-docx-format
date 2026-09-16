#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""内核必须只依赖标准库。唯一豁免：pdfcheck 用的 fitz（PyMuPDF），它是可选功能。"""
import ast, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
EXEMPT = {'fitz'}          # 只允许出现在 pdfcheck 里
LOCAL  = {'docxkit', 'tfmt'}

def stdlib_names():
    names = getattr(sys, 'stdlib_module_names', None)     # 3.10+
    if names: return set(names)
    # 3.9 回退：够用的白名单
    return {'argparse','ast','collections','copy','csv','datetime','difflib','functools','glob',
            'hashlib','html','io','itertools','json','math','os','pathlib','re','shutil','string',
            'subprocess','sys','tempfile','textwrap','time','types','typing','unicodedata',
            'urllib','uuid','xml','zipfile'}

def main():
    ok = stdlib_names() | LOCAL
    bad = []
    for rel in ('scripts/docxkit.py', 'scripts/tfmt.py'):
        path = os.path.join(ROOT, rel)
        tree = ast.parse(open(path, encoding='utf-8').read())
        for node in ast.walk(tree):
            mods = []
            if isinstance(node, ast.Import):
                mods = [a.name.split('.')[0] for a in node.names]
            elif isinstance(node, ast.ImportFrom) and node.module and node.level == 0:
                mods = [node.module.split('.')[0]]
            for m in mods:
                if m in ok or m in EXEMPT: continue
                bad.append('%s:%d 依赖了非标准库 %s' % (rel, node.lineno, m))
    for b in bad: print(b)
    print('非标准库依赖:', len(bad))
    return 1 if bad else 0

if __name__ == '__main__':
    sys.exit(main())
