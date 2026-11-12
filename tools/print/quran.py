"""Load the Quran text (tanzil-derived, from the npm package quran-json 3.1.2) keyed by (surah, ayah)."""
import json, os, sys
sys.path.insert(0, os.path.dirname(__file__))
from arnorm import nospace
from surahs import SURAHS
SRC = os.environ.get('QURAN_JSON', '/tmp/claude-0/-home-user-Mgt/be2474ec-46f6-5fbc-9b19-f8201f104606/scratchpad/res/quran/package/dist/quran.json')
_Q = None
def load():
    global _Q
    if _Q is None:
        data = json.load(open(SRC))
        _Q = {}
        for ch in data:
            for v in ch['verses']:
                _Q[(ch['id'], v['id'])] = v['text']
    return _Q
def ayah(s, a):
    return load().get((s, a))
def surah_no(name):
    name = name.strip()
    alias = {'الإسراء': 'الإسراء', 'بني إسرائيل': 'الإسراء', 'المؤمنين': 'المؤمنون', 'غافر': 'غافر', 'المؤمن': 'غافر'}
    name = alias.get(name, name)
    if name.startswith('سورة '):
        name = name[5:]
    return SURAHS.index(name) + 1 if name in SURAHS else None
def flat():
    """normalised no-space text of each ayah"""
    return {k: nospace(v) for k, v in load().items()}
