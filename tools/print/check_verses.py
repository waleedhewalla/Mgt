"""Check every Quranic quotation ﴿...﴾ in a docx against the Quran text.

usage: python3 check_verses.py <book.docx> <report.tsv>
For each quotation: the reference that follows it (e.g. «(هود: 9–10)», «[هود: 9]», «هود:9»), if any,
and whether every fragment (split on … or ...) occurs in the referenced ayahs (±1 ayah tolerance).
Status: OK | REF_WRONG (text found elsewhere; the found location is given) | NOT_FOUND (text matches no ayah:
a wording error or a non-Quranic phrase) | NOREF_OK (no reference, text found) | SHORT (too short to judge).
"""
import sys, os, re, difflib
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docx
from docx.oxml.ns import qn
from arnorm import skel
import quran
from surahs import SURAHS

AR = str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')
SUR = '|'.join(sorted((re.escape(s) for s in SURAHS), key=len, reverse=True))
REF = re.compile(r'^[\s\)\]»"]*[\(\[]?\s*(?:سورة\s+)?(' + SUR + r')\s*[:،]?\s*(?:الآية\s*|الآيات\s*)?([0-9٠-٩]+)(?:\s*[-–—،و]\s*([0-9٠-٩]+))?')

def texts(d):
    body = d.element.body
    i = 0
    for el in body.iterchildren():
        if el.tag.endswith('}p'):
            yield f'P{i}', ''.join(t.text or '' for t in el.iter(qn('w:t')))
            i += 1
        elif el.tag.endswith('}tbl'):
            for k, tc in enumerate(el.iter(qn('w:tc'))):
                yield f'T@P{i}#{k}', ''.join(t.text or '' for t in tc.iter(qn('w:t')))

SK = None
def whole():
    global SK
    if SK is None:
        keys = sorted(quran.load())
        SK = [(k, skel(quran.ayah(*k))) for k in keys]
    return SK

def find_anywhere(frag):
    hits = []
    sk = whole()
    for i, (k, s) in enumerate(sk):
        win = s + (sk[i + 1][1] if i + 1 < len(sk) and sk[i + 1][0][0] == k[0] else '')
        if frag in win:
            hits.append(k)
    return hits

def main():
    src, out = sys.argv[1], sys.argv[2]
    d = docx.Document(src)
    rows = []
    items = list(texts(d))
    for idx, (pid, t) in enumerate(items):
        for m in re.finditer(r'﴿([^﴾]{2,})﴾', t):
            q = m.group(1)
            after = t[m.end():m.end() + 60]
            r = REF.match(after)
            if not r and not after.strip() and idx + 1 < len(items):
                r = REF.match(items[idx + 1][1][:60])
            ref = None
            if r:
                s = SURAHS.index(r.group(1)) + 1
                a1 = int(r.group(2).translate(AR)); a2 = int(r.group(3).translate(AR)) if r.group(3) else a1
                if a2 < a1: a2 = a1
                ref = (s, a1, a2)
            frags = [skel(f) for f in re.split(r'…|\.\.\.|۝|\s\.\s', q)]
            frags = [f for f in frags if len(f) >= 4]
            if not frags:
                rows.append((pid, 'SHORT', '', q[:80], '')); continue
            if ref:
                s, a1, a2 = ref
                win = ''.join(skel(quran.ayah(s, a) or '') for a in range(max(1, a1 - 1), a2 + 2))
                if all(f in win for f in frags):
                    rows.append((pid, 'OK', f'{SURAHS[s-1]}:{a1}-{a2}', q[:80], '')); continue
                locs = find_anywhere(max(frags, key=len))
                if locs:
                    rows.append((pid, 'REF_WRONG', f'{SURAHS[s-1]}:{a1}-{a2}', q[:80],
                                 'found at ' + ', '.join(f'{SURAHS[a-1]}:{b}' for a, b in locs[:4]))); continue
                tgt = ' '.join(quran.ayah(s, a) or '' for a in range(a1, a2 + 1))
                rows.append((pid, 'NOT_FOUND', f'{SURAHS[s-1]}:{a1}-{a2}', q[:120], 'mushaf: ' + tgt[:160]))
            else:
                locs = find_anywhere(max(frags, key=len))
                if locs and all(any(f in skel(quran.ayah(*k)) + skel(quran.ayah(k[0], k[1] + 1) or '') for k in locs) or len(f) < 6 for f in frags):
                    rows.append((pid, 'NOREF_OK', '', q[:80], ', '.join(f'{SURAHS[a-1]}:{b}' for a, b in locs[:3])))
                elif locs:
                    rows.append((pid, 'NOREF_PARTIAL', '', q[:120], ', '.join(f'{SURAHS[a-1]}:{b}' for a, b in locs[:3])))
                else:
                    best = max(whole(), key=lambda x: difflib.SequenceMatcher(None, max(frags, key=len), x[1]).ratio() if abs(len(x[1]) - len(max(frags, key=len))) < 60 else 0)
                    rows.append((pid, 'NOT_FOUND', '', q[:120], f'closest {SURAHS[best[0][0]-1]}:{best[0][1]}: ' + quran.ayah(*best[0])[:140]))
    with open(out, 'w') as f:
        for r in rows:
            f.write('\t'.join(r) + '\n')
    from collections import Counter
    print(os.path.basename(src), Counter(r[1] for r in rows))

if __name__ == '__main__':
    main()
