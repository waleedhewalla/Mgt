"""Fetch a tafsir passage: python3 tafsir.py <slug> <surah_no> <ayah>
Arabic slugs: ar-tafsir-ibn-kathir, ar-tafseer-tahrir-al-tanwir (Ibn Ashur), ar-tafseer-al-qurtubi, tafsir-al-razi,
ar-tafsir-al-tabari, ar-tafsir-as-saadi, adwa-al-bayan (al-Shanqiti), al-kashshaf-al-zamakhshari, ar-tafsir-al-baghawi,
tafsir-al-baydawi, al-bahr-al-muhit, tafsir-al-alusi, fath-al-qadir-al-shawkani, ar-tafsir-muyassar, tafsir-ibn-al-qayyim.
(Sayyid Qutb's Zilal is NOT available.) Source: github.com/spa5k/tafsir_api. Cached under $TAFSIR_CACHE."""
import sys, os, json, subprocess
slug, s, a = sys.argv[1], sys.argv[2], sys.argv[3]
C = os.environ.get('TAFSIR_CACHE', '/tmp/claude-0/-home-user-Mgt/be2474ec-46f6-5fbc-9b19-f8201f104606/scratchpad/res/tafsir')
p = f'{C}/{slug}/{s}/{a}.json'
if not os.path.exists(p):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    subprocess.run(['curl', '-sf', '--max-time', '60', '-o', p,
                    f'https://raw.githubusercontent.com/spa5k/tafsir_api/main/tafsir/{slug}/{s}/{a}.json'])
if not os.path.exists(p) or os.path.getsize(p) == 0:
    print('NOT AVAILABLE'); sys.exit(1)
print(json.load(open(p))['text'])
