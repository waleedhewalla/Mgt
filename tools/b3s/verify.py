import re,json,glob,sys,docx
S='/tmp/claude-0/-home-user-Mgt/be2474ec-46f6-5fbc-9b19-f8201f104606/scratchpad'
def norm(s): return re.sub(r'[ً-ْـ«»"“”،؛:.؟!()\[\]﴿﴾]',' ',s)
def grams(s,n=5):
    w=norm(s).split(); return {' '.join(w[i:i+n]) for i in range(len(w)-n+1) if all(len(t)>1 for t in w[i:i+n])}
ref=grams(open(f'{S}/style/src/a_n.txt').read())|grams(open(f'{S}/style/src/b_n.txt').read())
AD=str.maketrans('٠١٢٣٤٥٦٧٨٩','0123456789')
def texts(o):
    out=[]
    for k in ('text','new','texts','rows','paras'):
        v=o.get(k)
        if isinstance(v,str): out.append(v)
        elif isinstance(v,list): out+= [json.dumps(v,ensure_ascii=False)]
    return out
def run(edir,ddir):
    tot=0
    for f in sorted(glob.glob(f'{edir}/*.json')):
        p=f.split('/')[-1][:-5]
        d=json.load(open(f)); ops=d.get('edits',d.get('ops')) if isinstance(d,dict) else d
        new=' '.join(t for o in ops for t in texts(o))
        orig=open(f'{ddir}/{p}.txt').read().translate(AD)
        nums=set(re.findall(r'\d+(?:\.\d+)?',new.translate(AD)))-set(re.findall(r'\d+(?:\.\d+)?',orig))
        ov=grams(new)&ref; tot+=len(ov)
        print(p,'newnums',sorted(nums),'overlap',len(ov),list(ov)[:3])
    print('total overlap',tot)
run(sys.argv[1],sys.argv[2])
