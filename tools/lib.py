from lxml import etree
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS={'w':W}
def q(t): return '{%s}%s'%(W,t)
def text(el): return ''.join(el.itertext()) if el is not None else ''
def ptext(el): return ''.join(t.text or '' for t in el.iter(q('t')))
