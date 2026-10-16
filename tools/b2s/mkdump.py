import zipfile,re,json
from lxml import etree
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
x=etree.fromstring(zipfile.ZipFile('b2u.docx').read('word/document.xml'))
body=[e for e in x.find(W+'body') if e.tag!=W+'sectPr']
pt=lambda e:''.join(t.text or '' for t in e.iter(W+'t')).strip()
ORD=['الأول','الثاني','الثالث','الرابع','الخامس','السادس','السابع','الثامن','التاسع','العاشر','الحادي عشر','الثاني عشر','الثالث عشر','الرابع عشر']
marks={}
started=False
for i,e in enumerate(body):
    if e.tag!=W+'p': continue
    t=pt(e)
    if t=='الإهداء' and any('TC' in (it.text or '') for it in e.iter(W+'instrText')): marks['front']=i
    m=re.match(r'^الفصل ('+'|'.join(sorted(ORD,key=len,reverse=True))+r')$',t)
    if m and any('TC' in (it.text or '') for it in e.iter(W+'instrText')):
        k=i
        j=i-1
        while j>0 and body[j].tag==W+'p' and not pt(body[j]): j-=1
        if pt(body[j]).startswith('الباب'): k=j
        marks['ch%02d'%(ORD.index(m.group(1))+1)]=k
    if t.startswith('الخاتمة: المؤسسة القرآنية') and any('TC' in (it.text or '') for it in e.iter(W+'instrText')): marks['concl']=i
    if t=='ملحق (ب)': marks['appB']=i
    if t=='الفهارس' and any('TC' in (it.text or '') for it in e.iter(W+'instrText')): marks['idx']=i
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
    if name=='idx': continue
    b=order[k+1][1]
    open(f'dumps/{name}.txt','w').write('\n'.join(dump(a,b)))
