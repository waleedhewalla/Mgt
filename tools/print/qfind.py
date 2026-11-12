"""Find Quranic text: python3 qfind.py "نص الآية أو جزء منها"   |   python3 qfind.py هود 9 [11]  (print ayahs)"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import quran
from arnorm import skel
from surahs import SURAHS
a = sys.argv[1:]
if len(a) >= 2 and a[0] in SURAHS:
    s = SURAHS.index(a[0]) + 1; a1 = int(a[1]); a2 = int(a[2]) if len(a) > 2 else a1
    for i in range(a1, a2 + 1):
        print(f'{a[0]}:{i}\t{quran.ayah(s, i)}')
else:
    f = skel(' '.join(a)); keys = sorted(quran.load()); n = 0
    for i, k in enumerate(keys):
        w = skel(quran.ayah(*k)) + (skel(quran.ayah(k[0], k[1] + 1) or ''))
        if f in w:
            print(f'{SURAHS[k[0]-1]}:{k[1]}\t{quran.ayah(*k)}'); n += 1
            if n > 15: break
    if not n: print('NOT FOUND')
