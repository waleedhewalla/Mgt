import sys, json, re, zipfile
sys.path.insert(0,'..')
from lib import *
def dump(path, prefix):
    x=etree.fromstring(zipfile.ZipFile(path).read('word/document.xml'))
    out=[]
    for i,e in enumerate(x.find(q('body'))):
        if e.tag==q('p'):
            s=ptext(e).strip()
            if s: out.append((i,f'{prefix}{i}',s))
        elif e.tag==q('tbl'):
            for r,tr in enumerate(e.findall(q('tr'))):
                for c,tc in enumerate(tr.findall(q('tc'))):
                    s=' / '.join(ptext(p).strip() for p in tc.findall(q('p')) if ptext(p).strip())
                    out.append((i,f'T{prefix}{i}.{r}.{c}',s))
    return out
A=dump('all.docx','P'); F=dump('fb.docx','F')
json.dump({'A':A,'F':F},open('dumps.json','w',encoding='utf-8'),ensure_ascii=False)
starts=[4,144,369,749,1182,1618,2076,2521,2960,3367,3754,4110,4467,4817,5120,5587]
names=['ch01','ch02','ch03','ch04','ch05','ch06','ch07','ch08','ch09','ch10','ch11','ch12','ch13','ch14','appA','appB']
# include preceding Part heading
adj=[]
for s in starts:
    pre=[i for i,_,t in A if s-4<=i<s and t.startswith('الباب')]
    adj.append(pre[0] if pre else s)
adj[0]=0
import os; os.makedirs('dumps',exist_ok=True)
for k,n in enumerate(names):
    lo=adj[k]; hi=adj[k+1] if k+1<len(adj) else 10**9
    with open(f'dumps/{n}.txt','w',encoding='utf-8') as f:
        for i,id_,t in A:
            if lo<=i<hi: f.write(f'{id_}\t{t}\n')
with open('dumps/front_back.txt','w',encoding='utf-8') as f:
    for i,id_,t in F: f.write(f'{id_}\t{t}\n')
for n in names+['front_back']:
    print(n, sum(1 for _ in open(f'dumps/{n}.txt',encoding='utf-8')), os.path.getsize(f'dumps/{n}.txt')//1024,'KB')
