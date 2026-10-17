#!/usr/bin/env python3
"""Quill suite rename: QuillForge→QuillPen, QuillLattice→QuillGrid, quillbox→QuillPost.

Usage: rename.py <repo-dir> [--dry-run]
Rewrites every tracked text file and renames tracked paths that contain an old name.
Legacy identifiers that must keep working (old storage keys, old file extensions) are handled
by code added separately; this script only does the mechanical part.
"""
import os, re, subprocess, sys

PAIRS = [
    # Product names in every spelling
    (r'Quill ?Lattice', 'QuillGrid'), (r'quill-lattice', 'QuillGrid'), (r'quilllattice', 'quillgrid'), (r'QUILLLATTICE', 'QUILLGRID'),
    (r'Quill ?Forge', 'QuillPen'), (r'quillforge', 'quillpen'), (r'QUILLFORGE', 'QUILLPEN'),
    (r'Quill ?[Bb]ox', 'QuillPost'), (r'quillbox', 'quillpost'), (r'QUILLBOX', 'QUILLPOST'),
    # Native file formats: QuillLattice Workbook (.qlw/.qlwe) → QuillGrid Workbook (.qgw/.qgwe);
    # QuillForge Document (.qfd/.qfde) → QuillPen Document (.qpd/.qpde)
    (r'\bqlwe\b', 'qgwe'), (r'\bqlw\b', 'qgw'), (r'\bQlwe\b', 'Qgwe'), (r'\bQlw\b', 'Qgw'), (r'\bQLWE?\b', lambda m: m.group(0).replace('QL', 'QG')),
    (r'\bqfde\b', 'qpde'), (r'\bqfd\b', 'qpd'), (r'\bQfde\b', 'Qpde'), (r'\bQfd\b', 'Qpd'), (r'\bQFDE?\b', lambda m: m.group(0).replace('QF', 'QP')),
]
COMPILED = [(re.compile(a), b) for a, b in PAIRS]
# Lines carrying this marker are left alone (legacy-compatibility code that must name the old identifiers).
KEEP = 'rename:keep'

def sub(text):
    out = []
    for line in text.splitlines(keepends=True):
        if KEEP in line:
            out.append(line); continue
        for rx, rep in COMPILED:
            line = rx.sub(rep, line)
        out.append(line)
    return ''.join(out)

def main():
    root = sys.argv[1]; dry = '--dry-run' in sys.argv
    files = subprocess.run(['git', '-C', root, 'ls-files', '-z'], capture_output=True, check=True).stdout.decode().split('\0')
    changed = 0
    for f in filter(None, files):
        p = os.path.join(root, f)
        if not os.path.isfile(p) or os.path.islink(p): continue
        try:
            raw = open(p, 'rb').read()
            if b'\0' in raw[:8000]: continue  # binary
            text = raw.decode('utf-8')
        except UnicodeDecodeError:
            continue
        new = sub(text)
        if new != text:
            changed += 1
            if not dry: open(p, 'w', encoding='utf-8', newline='').write(new)
    # Paths: deepest first so parents move after children.
    moves = []
    for f in filter(None, files):
        nf = '/'.join(sub(part) for part in f.split('/'))
        if nf != f: moves.append((f, nf))
    for old, new in moves:
        print(f'mv {old} -> {new}')
        if not dry:
            os.makedirs(os.path.dirname(os.path.join(root, new)) or root, exist_ok=True)
            subprocess.run(['git', '-C', root, 'mv', '-k', old, new], check=True)
    # Remove directories left empty by the moves.
    if not dry:
        for d, _, _ in sorted(os.walk(root), key=lambda x: -len(x[0])):
            if '.git' in d.split(os.sep) or 'node_modules' in d.split(os.sep): continue
            try:
                if not os.listdir(d): os.rmdir(d)
            except OSError:
                pass
    print(f'{changed} files rewritten, {len(moves)} paths moved')

if __name__ == '__main__':
    main()
