import json,re
from collections import defaultdict
U=json.load(open('units.json',encoding='utf-8'))
heads=defaultdict(list)
for ch,sec,s in U:
    m=re.match(r'^(\d+)\.(\d+)\s\s?\S',s)
    if m and len(s)<130 and ch not in('intro','concl'):
        heads[ch].append((int(m.group(1)),int(m.group(2))))
allsec=set()
for ch,h in heads.items():
    nums=[b for a,b in h]; chs={a for a,b in h}
    if not (chs=={int(ch)} and nums==list(range(1,len(nums)+1))): print('SEQ PROBLEM',ch,h)
    allsec|={f'{a}.{b}' for a,b in h}
bad=defaultdict(list)
for ch,sec,s in U:
    if re.match(r'^\d+\.\s',s): continue
    if re.match(r'^\d+\.\d+\s',s) and len(s)<130: continue
    for m in re.finditer(r'(?<![\d.:/])(\d{1,2})\.(\d{1,2})(?![\d.:%])',s):
        ref=f'{m.group(1)}.{m.group(2)}'
        if 1<=int(m.group(1))<=14 and ref not in allsec:
            pre=s[max(0,m.start()-10):m.start()]
            if re.search('جدول|شكل|CLO',pre): continue
            bad[ref].append((ch,sec,s[max(0,m.start()-50):m.end()+20]))
for r,v in sorted(bad.items()):
    print('MISSING SEC',r,len(v),v[:2])
