"""Classical Arabic texts from the OpenITI corpus (github.com/OpenITI/<NNNN>AH, one repo per 25-year span of death date).
python3 openiti.py list <repo e.g. 0675AH> [filter]     -> list text files in that repo (filter = author/title substring)
python3 openiti.py get <repo> <path>                     -> download the file, print its local path
Then grep the local file. Death year d (AH) lives in repo ceil(d/25)*25 AH, e.g. Ibn Khaldun 808 -> 0825AH,
al-Qarafi 684 -> 0700AH, al-Shatibi 790 -> 0800AH, Ibn Ashur 1393 -> 1400AH, al-Shafi'i 204 -> 0225AH,
Ibn Faris 395 -> 0400AH, al-Raghib 502 -> 0525AH, al-Suyuti 911 -> 0925AH, Ibn al-Qayyim 751 -> 0775AH,
Ibn Taymiyya 728 -> 0750AH, al-Zarkashi 794 -> 0800AH, al-Dihlawi 1176 -> 1200AH, al-Ghazali 505 -> 0525AH.
Files ending -ara1 / .mARkdown / .completed are the text. Cached under $OPENITI_CACHE."""
import sys, os, subprocess
C = os.environ.get('OPENITI_CACHE', '/tmp/claude-0/-home-user-Mgt/be2474ec-46f6-5fbc-9b19-f8201f104606/scratchpad/res/openiti_cache')
os.makedirs(C, exist_ok=True)
cmd, repo = sys.argv[1], sys.argv[2]
d = f'{C}/{repo}'
if not os.path.exists(d):
    subprocess.run(['git', 'clone', '-q', '--filter=blob:none', '--no-checkout', '--depth', '1',
                    f'https://github.com/OpenITI/{repo}', d], check=True)
if cmd == 'list':
    flt = sys.argv[3] if len(sys.argv) > 3 else ''
    out = subprocess.run(['git', '-C', d, 'ls-tree', '-r', '--name-only', 'HEAD'], capture_output=True, text=True).stdout
    for l in out.splitlines():
        if flt.lower() in l.lower() and not l.endswith(('.yml', '.md')):
            print(l)
elif cmd == 'get':
    path = sys.argv[3]
    local = f'{d}/{path}'
    if not os.path.exists(local):
        subprocess.run(['git', '-C', d, 'checkout', 'HEAD', '--', path], check=True)
    print(local)
