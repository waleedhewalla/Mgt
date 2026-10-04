import zipfile,re,json
from lxml import etree
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
x=etree.fromstring(zipfile.ZipFile('b1.docx').read('word/document.xml'))
body=[e for e in x.find(W+'body') if e.tag!=W+'sectPr']
def pt(e): return ''.join(t.text or '' for t in e.iter(W+'t')).strip()
ORD=['الأول','الثاني','الثالث','الرابع','الخامس','السادس','السابع','الثامن','التاسع','العاشر','الحادي عشر','الثاني عشر','الثالث عشر','الرابع عشر']
marks={}
seen_intro=0
for i,e in enumerate(body):
    if e.tag!=W+'p': continue
    t=pt(e)
    if t=='المقدمة':
        seen_intro+=1
        if seen_intro==2 or (seen_intro==1 and i>50): marks['intro']=i
    if t in ['الفصل '+o for o in ORD] and i>250: marks['ch%02d'%(ORD.index(t[6:])+1)]=i
    if t=='خاتمة الكتاب الأول' and i>250: marks['concl']=i
    if t=='الملاحق' and i>250: marks.setdefault('apps',i)
    if t=='الفهارس' and i>250: marks['idx']=i
print(marks)
order=sorted(marks.items(),key=lambda kv:kv[1])
json.dump(order,open('marks.json','w'))
def dump(a,b):
    out=[]
    for i in range(a,b):
        e=body[i]
        if e.tag==W+'p':
            t=pt(e)
            if t: out.append(f'P{i}\t{t}')
        elif e.tag==W+'tbl':
            out.append(f'TBL{i}\t[جدول]')
            for r,tr in enumerate(e.findall(W+'tr')):
                for c,tc in enumerate(tr.findall(W+'tc')):
                    ct=' / '.join(pt(p) for p in tc.findall(W+'p') if pt(p))
                    out.append(f'TP{i}.{r}.{c}\t{ct}')
    return out
for k,(name,a) in enumerate(order):
    b=order[k+1][1] if k+1<len(order) else len(body)
    if name=='idx': continue
    open(f'dumps/{name}.txt','w').write('\n'.join(dump(a,b)))
