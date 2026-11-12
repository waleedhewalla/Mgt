"""Arabic normalisation shared by the verification tools."""
import re
DIAC = re.compile(r'[ؐ-ًؚ-ٰٟۖ-ۭ࣓-ࣿـ]')
def norm(s):
    s = DIAC.sub('', s)
    s = s.replace('ٱ', 'ا').replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا').replace('ٲ', 'ا').replace('ٳ', 'ا')
    s = s.replace('ى', 'ي').replace('ئ', 'ي').replace('ؤ', 'و').replace('ة', 'ه').replace('ء', '')
    s = s.replace('ۥ', '').replace('ۦ', '').replace('ۥ', '').replace('ۦ', '')
    s = re.sub(r'[^ء-ي ]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()
def nospace(s):
    return norm(s).replace(' ', '')


def skel(s):
    """Consonantal skeleton: tolerant of Uthmani vs standard spelling (long vowels, hamza seats, ta marbuta)."""
    s = DIAC.sub('', s)
    s = s.replace('ة', 'ت').replace('ۃ', 'ت')
    s = re.sub('[اأإآٱٲٳءئؤىيوۥۦـٕٔ]', '', s)
    return re.sub(r'[^ء-ي]', '', s)
