"""Search hadith collections (fawazahmed0/hadith-api Arabic editions):
python3 hsearch.py "عبارة من متن الحديث" [bukhari,muslim,...]
Matching ignores diacritics and spelling of hamza/alef/ya. Prints collection, number and full text."""
import sys, os, json, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from arnorm import norm
D = os.environ.get('HADITH_DIR', '/tmp/claude-0/-home-user-Mgt/be2474ec-46f6-5fbc-9b19-f8201f104606/scratchpad/res/hadith')
q = norm(sys.argv[1])
cols = sys.argv[2].split(',') if len(sys.argv) > 2 else ['bukhari', 'muslim', 'abudawud', 'tirmidhi', 'nasai', 'ibnmajah', 'malik', 'nawawi', 'qudsi']
n = 0
for c in cols:
    for h in json.load(open(f'{D}/ara-{c}.json'))['hadiths']:
        if q in norm(h['text']):
            print(f"== {c} {h['hadithnumber']}\n{h['text']}\n"); n += 1
            if n >= 8: sys.exit()
if not n: print('NOT FOUND')
