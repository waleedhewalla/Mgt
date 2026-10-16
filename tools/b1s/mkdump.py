import zipfile,re
from lxml import etree
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
x=etree.fromstring(zipfile.ZipFile('b1u.docx').read('word/document.xml'))
body=[e for e in x.find(W+'body') if e.tag!=W+'sectPr']
pt=lambda e:''.join(t.text or '' for t in e.iter(W+'t')).strip()
ORD=['الأول','الثاني','الثالث','الرابع','الخامس','السادس','السابع','الثامن','التاسع','العاشر','الحادي عشر','الثاني عشر','الثالث عشر','الرابع عشر']
hasTC=lambda e:any('TC' in (it.text or '') for it in e.iter(W+'instrText'))
marks={}
for i,e in enumerate(body):
    if e.tag!=W+'p' or not hasTC(e): continue
    t=pt(e)
    if t=='المقدمة': marks['intro']=i
    m=re.match(r'^الفصل ('+'|'.join(sorted(ORD,key=len,reverse=True))+r')$',t)
    if m:
        k=i; j=i-1
        while j>0 and body[j].tag==W+'p' and not pt(body[j]): j-=1
        if pt(body[j]).startswith('الباب'): k=j
        marks['ch%02d'%(ORD.index(m.group(1))+1)]=k
    if t=='خاتمة الكتاب الأول': marks['concl']=i
    if t=='الملاحق': marks['apps']=i
order=sorted(marks.items(),key=lambda kv:kv[1]); print(order)
def dump(a,b):
    out=[]
    for i in range(a,b):
        e=body[i]
        if e.tag==W+'p':
            t=pt(e)
            if e.find('.//'+W+'drawing') is not None: out.append(f'P{i}\t[صورة الشكل التوضيحي]')
            elif t: out.append(f'P{i}\t{t}')
        elif e.tag==W+'tbl':
            out.append(f'TBL{i}\t[جدول]')
            for r,tr in enumerate(e.findall(W+'tr')):
                for c,tc in enumerate(tr.findall(W+'tc')):
                    out.append(f'TP{i}.{r}.{c}\t'+' / '.join(pt(p) for p in tc.findall(W+'p') if pt(p)))
    return out
for k,(name,a) in enumerate(order):
    if name=='apps': continue
    b=order[k+1][1]
    d=dump(a,b); open(f'dumps/{name}.txt','w').write('\n'.join(d))
    print(name, a, b, len(' '.join(l.split('\t',1)[1] for l in d).split()))
