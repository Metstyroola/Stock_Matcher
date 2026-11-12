import sys, re, collections, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L_
src,dst=sys.argv[1],sys.argv[2]
wb=openpyxl.load_workbook(src,read_only=True,data_only=True)
LST=[r[:4] for r in wb['Sheet1'].iter_rows(min_row=2,values_only=True) if r and r[0]]
EH=list(next(wb['Sheet2'].iter_rows(max_row=1,values_only=True)))
E=[r for r in wb['Sheet2'].iter_rows(min_row=2,values_only=True) if r and r[0]]
ix={h:i for i,h in enumerate(EH)}
OE={'*':'BMW','MO':'Mercedes','MO1':'Mercedes','MO2':'Mercedes','MOE':'Mercedes','MO-V':'Mercedes',
 'N0':'Porsche','N1':'Porsche','N2':'Porsche','N3':'Porsche','N4':'Porsche','NA0':'Porsche','NC0':'Porsche','ND0':'Porsche','ND1':'Porsche','NE0':'Porsche','NF0':'Porsche','NG0':'Porsche',
 'AO':'Audi','AO1':'Audi','AO2':'Audi','R0':'Audi','R01':'Audi','T0':'Tesla','T1':'Tesla','VOL':'Volvo','VOLDF':'Volvo','VOLFR':'Volvo',
 'J':'Jaguar','JLR':'Jaguar Land Rover','LR':'Land Rover','K1':'Ferrari','K2':'Ferrari','K3':'Ferrari','AML':'Aston Martin','MGT':'Maserati','POL':'Polestar',
 'A39':'Alpine','A11':'Alpine','HN':'Hyundai'}
VAR={'A','AC','ACOUSTIC','ST','S1','DT','DT1','DT2','DT3','CNT','CONNECT','E','G1','VG','VM','VB','VQ','VC','VI','XL','ZP','SELFSEAL','GRNX','EL','TL','GOE','FRV'}
JOIN=[('CNTN0','CNT N0'),('AODT1','AO DT1'),('STVOL','ST VOL'),('DTFRV','DT FRV'),('JLRDT','JLR DT'),('PCY','PRIMACY'),
      ('CROSSCLIMATE2','CROSSCLIMATE 2'),('XZA2','XZA 2'),('XZE2','XZE 2'),('N0PORSCHE','N0'),('I*','*'),('AGILIS51','AGILIS 51')]
def toks(s):
    s=' '+str(s or '').upper().replace('+',' PLUS ')+' '
    s=re.sub(r'\*',' * ',s)
    for a,b in JOIN: s=re.sub(r'(?<= )'+re.escape(a)+r'(?= )',b,s)
    return [t for t in re.split(r'[ ,()]+',s) if t]
def eprel_base(name):
    t=toks(name); oe=set(); var=set()
    while t and (t[-1] in OE or t[-1] in VAR):
        x=t.pop(); (oe if x in OE else var).add(x)
    return ' '.join(t), oe, var
ALIAS={'PILOT SPORT PS3':'PILOT SPORT 3','PRIMACY 3ST':'PRIMACY 3','PRIMACY 4 ST':'PRIMACY 4','E PRIMACY ST':'E PRIMACY'}
def list_base(pt): b=' '.join(toks(pt)); return ALIAS.get(b,b)
def pmarks(p): t=toks(p); return {x for x in t if x in OE}, ('ACOUSTIC' in t or 'BLE-ACOUSTIC' in t)
nn=lambda s: re.sub(r'[^A-Z0-9]','',s)
def nd(t):
    p=str(t or '').upper().split()
    if not p: return ''
    size=re.sub(r'[^0-9/R.]','',p[0]); sv=p[1] if len(p)>1 else ''
    return f"{size} {re.sub(r'[^0-9]','',sv.split('/')[0])}{re.sub(r'[^A-Z]','',sv[-1:])}"
idx=collections.defaultdict(list)
for e in E:
    b,oe,var=eprel_base(e[ix['Commercial Name']]); idx[nd(e[ix['DIMENSIONS']])].append(dict(r=e,base=b,oe=oe,var=var))
out=[]
for b,p,pt,dim in LST:
    k=nd(dim); cands=idx.get(k,[]); lb=nn(list_base(pt)); poe,pac=pmarks(p)
    same=[c for c in cands if nn(c['base'])==lb]
    succ=[c for c in cands if nn(c['base'])==lb+'PLUS'] if not same else []
    def sc(c):
        s=0
        if c['oe']==poe: s+=20
        elif c['oe']&poe: s+=10
        if bool(c['var']&{'A','AC','ACOUSTIC'})==pac: s+=5
        return s
    pool=same or succ
    best=[]
    if pool:
        top=max(sc(c) for c in pool); best=[c for c in pool if sc(c)==top]
    if same:
        mt='Matched - pattern + OE' if best[0]['oe']==poe else ('Matched - pattern, different OE variant' if poe or best[0]['oe'] else 'Matched - pattern')
        if best[0]['oe']==poe and not poe: mt='Matched - pattern'
    elif succ: mt='Review - only PLUS (successor) version in EPREL'
    elif cands: mt='No match - pattern not registered at this size'
    else: mt='No match - size not in EPREL'
    out.append(dict(b=b,p=p,pt=pt,dim=dim,k=k,mt=mt,best=best,cands=cands,poe=poe))
cnt=collections.Counter(o['mt'] for o in out); print(cnt)

# ---- write workbook
F='Arial'; H=Font(name=F,bold=True,color='FFFFFF'); fill=PatternFill('solid',fgColor='1F3864'); B=Font(name=F)
COL={'Matched - pattern + OE':'C6EFCE','Matched - pattern':'C6EFCE','Matched - pattern, different OE variant':'FFEB9C',
     'Review - only PLUS (successor) version in EPREL':'FFEB9C','No match - pattern not registered at this size':'F8CBAD','No match - size not in EPREL':'F8CBAD'}
wo=openpyxl.Workbook(); ws=wo.active; ws.title='Matching Results'
hdr=['Brand','Profile','Profile Text','dimension','Match Type','Profile OE Mark','EPREL ID','All Matching EPREL IDs','No. of Matching IDs',
     'EAN (GTIN)','EPREL Commercial Name','EPREL OE Mark','EPREL Tyre Designation','Energy Class','Wet Grip Class','Noise Class','Noise Value (dB)',
     'EPREL Patterns at this Size']
ws.append(hdr)
for o in out:
    bst=o['best']; r=bst[0]['r'] if bst else None
    g=lambda h: r[ix[h]] if r is not None else None
    eans=sorted({str(c['r'][ix['EAN (GTIN)']]) for c in bst if c['r'][ix['EAN (GTIN)']]})
    pats='' if o['mt'].startswith('Matched') else ', '.join(sorted({c['r'][ix['Commercial Name']].strip() for c in o['cands']}))
    ws.append([o['b'],o['p'],o['pt'],o['dim'],o['mt'],', '.join(sorted(o['poe'])) or None,
        str(g('EPREL ID')) if r else None, ', '.join(str(c['r'][ix['EPREL ID']]) for c in bst) or None, len(bst) or None,
        ', '.join(eans) or None, g('Commercial Name'), ', '.join(sorted(bst[0]['oe'])) if bst and bst[0]['oe'] else None,
        g('Tyre Designation'), g('Energy Class'), g('Wet Grip Class'), g('Noise Class'), g('Noise Value (dB)'), pats or None])
for i,w in enumerate([10,44,22,20,40,12,10,24,9,16,26,10,22,8,8,8,8,60],1): ws.column_dimensions[L_(i)].width=w
for c in ws[1]: c.font=H; c.fill=fill; c.alignment=Alignment(wrap_text=True,vertical='center')
for row in ws.iter_rows(min_row=2):
    for c in row: c.font=B
    row[4].fill=PatternFill('solid',fgColor=COL[row[4].value])
ws.freeze_panes='C2'; ws.auto_filter.ref=ws.dimensions
n=ws.max_row
s=wo.create_sheet('Summary',0)
s['A1']='Michelin list vs EPREL - matching summary'; s['A1'].font=Font(name=F,bold=True,size=14)
s.append([]); s.append(['Match Type','Products','Share','What it means'])
for c in s[3]: c.font=H; c.fill=fill
mean={'Matched - pattern + OE':'Same size, same pattern, same OE mark (e.g. MO, N0, *).',
 'Matched - pattern':'Same size and pattern; neither side has an OE mark.',
 'Matched - pattern, different OE variant':'Same size and pattern, but EPREL only has a different OE version (or the standard one). Check before using the EPREL ID.',
 'Review - only PLUS (successor) version in EPREL':'EPREL has the PLUS/+ successor at this size (e.g. ENERGY XM2 +, AGILIS+), not this pattern.',
 'No match - pattern not registered at this size':'Size exists for Michelin in EPREL, but not this pattern. See "EPREL Patterns at this Size".',
 'No match - size not in EPREL':'No Michelin EPREL registration for this size. Often non-EU / older ranges (XAS, X, LTX FORCE, PILOT LTX).'}
order=list(mean)
for i,m in enumerate(order,4):
    s.cell(i,1,m); s.cell(i,2,f"=COUNTIF('Matching Results'!$E$2:$E${n},A{i})"); s.cell(i,3,f'=B{i}/$B${4+len(order)}'); s.cell(i,4,mean[m])
    s.cell(i,1).fill=PatternFill('solid',fgColor=COL[m]); s.cell(i,3).number_format='0.0%'
t=4+len(order); s.cell(t,1,'Total'); s.cell(t,2,f'=SUM(B4:B{t-1})'); s.cell(t,3,f'=B{t}/B{t}'); s.cell(t,3).number_format='0.0%'
for c in s[t]: c.font=Font(name=F,bold=True)
notes=['Notes',
 'Source: Michelin_eu_label.xlsx - Sheet1 (your 498 products) matched against Sheet2 (4,525 Michelin EPREL records).',
 'Size key: width/profile/rim + first load index + speed (195/80R15 108/106S = 195/80R15 108S; ZR = R).',
 'Pattern: EPREL OE/variant suffixes (MO, MO1, N0, *, AO, VOL, A=acoustic, ST, S1, DT, VM, VG ...) are removed before comparing to Profile Text.',
 'Pattern aliases used: PILOT SPORT PS3 = PILOT SPORT 3; PRIMACY 3ST = PRIMACY 3; PRIMACY 4 ST = PRIMACY 4; E PRIMACY ST = E PRIMACY.',
 'OE variant: picked by comparing the OE mark in your Profile (e.g. "MO Mercedes", "N0 Porsche", "* BMW") with the EPREL suffix.',
 'EAN: blank in every Sheet2 record, so no EAN can be returned.']
for i,x in enumerate(notes,t+2):
    s.cell(i,1,x).font=Font(name=F,bold=(i==t+2),italic=(i>t+2))
for row in s.iter_rows(min_row=4,max_row=t):
    for c in row:
        if c.font.bold is False or c.font.bold is None: c.font=B
for col,w in zip('ABCD',[46,10,8,100]): s.column_dimensions[col].width=w
wo.save(dst)
