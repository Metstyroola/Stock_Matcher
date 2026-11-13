import sys, re, collections, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L_
src,dst=sys.argv[1],sys.argv[2]
wb=openpyxl.load_workbook(src,read_only=True,data_only=True)
LST=[(r[2],r[4],r[5],r[6],r[0],r[1],r[3]) for r in wb['Sheet1'].iter_rows(min_row=2,values_only=True) if r and r[2]]
EH=list(next(wb['Sheet2'].iter_rows(max_row=1,values_only=True)))
E=[r for r in wb['Sheet2'].iter_rows(min_row=2,values_only=True) if r and r[0] and r[4]]
ix={h:i for i,h in enumerate(EH)}
OE={'*':'BMW','MO':'Mercedes','MO1':'Mercedes','MO2':'Mercedes','MOE':'Mercedes','MO-V':'Mercedes',
 'N0':'Porsche','N1':'Porsche','N2':'Porsche','N3':'Porsche','N4':'Porsche','NA0':'Porsche','NC0':'Porsche','ND0':'Porsche','ND1':'Porsche','NE0':'Porsche','NF0':'Porsche','NG0':'Porsche',
 'AO':'Audi','AO1':'Audi','AO2':'Audi','R0':'Audi','R01':'Audi','T0':'Tesla','T1':'Tesla','VOL':'Volvo','VOLDF':'Volvo','VOLFR':'Volvo',
 'J':'Jaguar','JLR':'Jaguar Land Rover','LR':'Land Rover','K1':'Ferrari','K2':'Ferrari','K3':'Ferrari','AML':'Aston Martin','MGT':'Maserati','POL':'Polestar',
 'A39':'Alpine','A11':'Alpine','HN':'Hyundai'}
VAR={'RFT','SSR','ROF','RUNFLAT','ELT','NCS','PNCS','SEAL','CONTISEAL','CS','SIL','KS','A','AC','ACOUSTIC','ST','S1','DT','DT1','DT2','DT3','CNT','CONNECT','E','G1','VG','VM','VB','VQ','VC','VI','XL','ZP','SELFSEAL','GRNX','EL','TL','GOE','FRV'}
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
LVAR={'RWL','OWL','ORWL','BSW','XL','ZP','RFT','SSR','ROF','RUNFLAT','SEAL','ELT'}
def list_base(pt):
    t=toks(pt)
    while len(t)>1 and t[-1] in LVAR: t.pop()
    b=' '.join(t); return ALIAS.get(b,b)
def pmarks(p): t=toks(p); return {x for x in t if x in OE}, ('ACOUSTIC' in t or 'BLE-ACOUSTIC' in t)
nn=lambda s: re.sub(r'[^A-Z0-9]','',s)
def nd(t):
    p=str(t or '').upper().replace('(',' ').replace(')',' ').split()
    if not p: return ''
    size=re.sub(r'[^0-9/R.]','',p[0]); sv=p[1] if len(p)>1 else ''
    return f"{size} {re.sub(r'[^0-9]','',sv.split('/')[0])}{re.sub(r'[^A-Z]','',sv[-1:])}"
def parse(t):
    p=str(t or '').upper().replace('(',' ').replace(')',' ').split()
    if not p: return '',set(),''
    size=re.sub(r'[^0-9/R.]','',p[0]); lis=set(); sp=''
    for i,tk in enumerate(p[1:],1):
        m=re.fullmatch(r'(\d{2,3}(?:/\d{2,3})*)([A-Z]?)',tk)
        if m and not lis:
            lis=set(m.group(1).split('/')); sp=m.group(2)
            if not sp and i+1<len(p) and re.fullmatch(r'[A-Z]',p[i+1]): sp=p[i+1]
        elif m and lis and not sp and m.group(2): sp=m.group(2)
        elif not lis and not sp and tk in ('Z','ZR'): sp='Z'
    return size,lis,sp
def keys(t):
    size,lis,sp=parse(t); return [f"{size} {li}{sp}" for li in sorted(lis)] or [f"{size} "]
idx=collections.defaultdict(list); idx2=collections.defaultdict(list)
for e in E:
    b,oe,var=eprel_base(e[ix['Commercial Name']]); _c=dict(r=e,base=b,oe=oe,var=var); _br=str(e[ix['Brand Clean']]).upper()
    _ks=set(keys(e[ix['Tyre Designation']])) | set(keys(e[ix['DIMENSIONS']]))
    for _k in _ks: idx[(_br,_k)].append(_c)
    idx2[(_br,_ks.pop().split(' ')[0])].append(_c)

ROMAN={'II':'2','III':'3','IV':'4'}
GENERIC={'4X4','2WD','4WD','SUV','AT','HT','MT','AS','4S','EV','HP','UHP','LT','VAN'}
def tokl(s):
    t=[]
    for x in toks(s):
        x=re.sub(r'^(\d+)(III|II|IV)$',lambda m:m.group(1)+' '+ROMAN[m.group(2)],x)
        x=re.sub(r'^([A-Z/]{2,})(III|II)$',lambda m:m.group(1)+ROMAN[m.group(2)],x)
        for y in x.split():
            y=ROMAN.get(y,y)
            if y.startswith('CONTI') and len(y)>8: y=y[5:]
            y=nn(y)
            if y: t.append(y)
    return t
iscode=lambda w: len(w)>=3 and re.search(r'[A-Z]',w) and re.search(r'\d',w) and w not in GENERIC
out=[]
for b,p,pt,dim,spid,dpid,seg in LST:
    kk=keys(dim); k=' / '.join(kk); _seen=set(); cands=[]
    for _k in kk:
        for c in idx.get((str(b).strip().upper(),_k),[]):
            if id(c) not in _seen: _seen.add(id(c)); cands.append(c)
    lbase=list_base(pt); lb=nn(lbase); lt=tokl(lbase); lall=''.join(lt); poe,pac=pmarks(p)
    def sc(c):
        s=0
        if c['oe']==poe: s+=20
        elif c['oe']&poe: s+=10
        if bool(c['var']&{'A','AC','ACOUSTIC'})==pac: s+=5
        return s
    BN=nn(str(b))
    def ct(c):
        if 'bt' not in c:
            raw=str(c['r'][ix['Commercial Name']] or '')
            noisy=bool(re.search(r'\d{3}/\d{2}|\dR\d{2}',raw))
            base=re.sub(r'\(.*?\)',' ',c['base'])
            t=tokl(base)
            if t and t[0]==BN and len(t)>1: t=t[1:]
            if noisy:
                t=[w for w in t if iscode(w) and not re.fullmatch(r'P?\d{2,3}[A-Z]{0,2}|\d+[A-Z]\d+|.*\d{3}.*\d{2}.*',w)]
            c['bt']=t; c['noisy']=noisy
        return c['bt']
    same=[c for c in cands if nn(c['base'])==lb or (ct(c) and sorted(ct(c))==sorted(lt))]
    units=set(lt)|{lt[i]+lt[i+1] for i in range(len(lt)-1)}
    code=[] if same else [c for c in cands if any(iscode(w) for w in ct(c)) and (
          (c.get('noisy') and any(w in units for w in ct(c)) and not ('PLUS' in tokl(c['base']) and 'PLUS' not in lt))
          or (not c.get('noisy') and all(w in units for w in ct(c) if iscode(w)) and all(w in lt for w in ct(c) if not iscode(w))))]
    DIFF={'PLUS','XTREME','EV','SUV','PRO','EVO','VAN','ULTRA','MAX'}
    codex=[c for c in code if any(w in DIFF and w not in ct(c) for w in lt)]
    if code and len(codex)==len(code): code=[]
    succ=[] if (same or code) else [c for c in cands if nn(c['base'])==lb+'PLUS']
    if not (same or code or succ) and codex: pref_code=codex
    else: pref_code=[]
    pref=[] if (same or code or succ or codex) else [c for c in cands if ct(c) and len(ct(c))<len(lt) and lt[:len(ct(c))]==ct(c)]
    pool=same or code or succ or pref_code or pref; best=[]
    if pool:
        top=max(sc(c) for c in pool); best=[c for c in pool if sc(c)==top]
    if same or code:
        oeok=best[0]['oe']==poe
        mt=('Matched - pattern + OE' if poe else 'Matched - pattern') if oeok else 'Matched - pattern, different OE variant'
        if code and oeok: mt='Matched - model code'
    elif succ: mt='Review - only PLUS (successor) version in EPREL'
    elif pref_code: mt='Review - model code matches, but list has extra variant word (PLUS/SUV/EV...)'
    elif pref: mt='Review - EPREL name is shorter (e.g. PZERO vs PZERO PZ4)'
    else:
        alt=[c for c in idx2.get((str(b).strip().upper(),kk[0].split(' ')[0]),[]) if nn(c['base']) in (lb,lb+'PLUS') or (ct(c) and sorted(ct(c))==sorted(lt))]
        if alt:
            top=max(sc(c) for c in alt); best=[c for c in alt if sc(c)==top]; mt='Review - same pattern & size, different load/speed'
        elif cands: mt='No match - pattern not registered at this size'
        else: mt='No match - size not in EPREL'
    out.append(dict(spid=spid,dpid=dpid,seg=seg,b=b,p=p,pt=pt,dim=dim,k=k,mt=mt,best=best,cands=cands,poe=poe))
cnt=collections.Counter(o['mt'] for o in out); print(cnt)

TITLE='Other suppliers (26 brands) vs EPREL - matching summary'
NOTES=['Notes',
 'Source: Other_SUP_for_matching.xlsx - Sheet1 (2,424 products, 26 brands) vs Sheet2 (62,315 EPREL records for the same brands; blank rows ignored). Brand must match (Brand Clean).',
 'Size key: width/profile/rim + load index + speed. Dual load indexes match either value (94W = 96/94W). ZR = R; brackets and LT/C ignored.',
 'Pattern: EPREL OE/variant suffixes (MO, N0, *, AO, VOL, A, ST, RFT, SSR, ELT, NCS, SEAL ...) removed; words compared ignoring punctuation, order, CONTI- prefix and roman numerals (II = 2).',
 'Letter markings at the end of your Profile Text (RWL, OWL, XL, ZP, RFT, SSR, ROF, SEAL) are ignored.',
 'EPREL names that embed the size (e.g. Nankang "P235/75R15 108T N-605 L EUHA") are reduced to their pattern code; bracketed internal codes and a leading brand word are ignored.',
 'Model codes (letters+digits, e.g. PS71, G062, R623) must appear as a whole word in your name; PLUS / SUV / EV / XTREME / PRO / EVO / VAN differences go to Review.',
 'EAN: blank in every Sheet2 record, so no EAN can be returned.']
# ---- write workbook (generic, any brands)
F_='Arial'; H=Font(name=F_,bold=True,color='FFFFFF'); fill=PatternFill('solid',fgColor='1F3864'); B=Font(name=F_)
G,Y,Rd='C6EFCE','FFEB9C','F8CBAD'
MEAN={
 'Matched - pattern + OE':(G,'Same size, same pattern, same OE mark (MO, N0, *, AO ...).'),
 'Matched - pattern':(G,'Same size and pattern (spelling, word order, CONTI- prefix and roman numerals ignored).'),
 'Matched - model code':(G,'EPREL only gives the model code and your name contains it as a whole word, e.g. Kumho ECSTA PS71 = PS71, Bridgestone DURAVIS R 623 = R623.'),
 'Matched - pattern, different OE variant':(Y,'Same size and pattern, but EPREL has another OE version (or the standard one). Check before using the EPREL ID.'),
 'Review - model code matches, but list has extra variant word (PLUS/SUV/EV...)':(Y,'Code matches but your name has PLUS / SUV / EV / XTREME etc. that EPREL lacks, e.g. ECSTA PS71 SUV vs PS71.'),
 'Review - EPREL name is shorter (e.g. PZERO vs PZERO PZ4)':(Y,'EPREL registers the family name only (PZERO, ULTRACONTACT, SCORPION). Often the same tyre - check.'),
 'Review - only PLUS (successor) version in EPREL':(Y,'EPREL has the PLUS/+ successor at this size, not this pattern.'),
 'Review - same pattern & size, different load/speed':(Y,'EPREL has this pattern and size but another load index / speed rating. Different spec.'),
 'No match - pattern not registered at this size':(Rd,'Brand has this size in EPREL, but not this pattern. See "EPREL Patterns at this Size".'),
 'No match - size not in EPREL':(Rd,'No EPREL registration for this brand + size (often non-EU ranges / sizes).')}
wo=openpyxl.Workbook(); ws=wo.active; ws.title='Matching Results'
hdr=['Spinach ID','DP ID','Brand','Brand Segment','Profile','Profile Text','dimension','Match Type','Profile OE Mark','EPREL ID','All Matching EPREL IDs','No. of Matching IDs',
     'EAN (GTIN)','EPREL Commercial Name','EPREL OE Mark','EPREL Tyre Designation','Energy Class','Wet Grip Class','Noise Class','Noise Value (dB)','EPREL Patterns at this Size']
ws.append(hdr)
for o in out:
    bst=o['best']; r=bst[0]['r'] if bst else None
    g_=lambda h: r[ix[h]] if r is not None else None
    eans=sorted({str(c['r'][ix['EAN (GTIN)']]) for c in bst if c['r'][ix['EAN (GTIN)']]})
    pats='' if o['mt'].startswith('Matched') else ', '.join(sorted({str(c['r'][ix['Commercial Name']]).strip() for c in o['cands']}))
    ws.append([o['spid'],o['dpid'],o['b'],o['seg'],o['p'],o['pt'],o['dim'],o['mt'],', '.join(sorted(o['poe'])) or None,
        str(g_('EPREL ID')) if r else None, ', '.join(str(c['r'][ix['EPREL ID']]) for c in bst) or None, len(bst) or None,
        ', '.join(eans) or None, (str(g_('Commercial Name')).strip() if r else None), ', '.join(sorted(bst[0]['oe'])) if bst and bst[0]['oe'] else None,
        g_('Tyre Designation'), g_('Energy Class'), g_('Wet Grip Class'), g_('Noise Class'), g_('Noise Value (dB)'), pats[:2000] or None])
for i,w in enumerate([15,11,13,10,40,30,20,44,10,10,24,9,14,28,10,24,8,8,8,8,60],1): ws.column_dimensions[L_(i)].width=w
for c in ws[1]: c.font=H; c.fill=fill; c.alignment=Alignment(wrap_text=True,vertical='center')
for row in ws.iter_rows(min_row=2):
    for c in row: c.font=B
    row[7].fill=PatternFill('solid',fgColor=MEAN[row[7].value][0])
ws.freeze_panes='F2'; ws.auto_filter.ref=ws.dimensions
n=ws.max_row; R="'Matching Results'"; MT=f"{R}!$H$2:$H${n}"; BR=f"{R}!$C$2:$C${n}"
s=wo.create_sheet('Summary',0)
s['A1']=TITLE; s['A1'].font=Font(name=F_,bold=True,size=14)
s['A3']='By match type'; s['A3'].font=Font(name=F_,bold=True)
s.append(['Match Type','Products','Share','What it means'])
for c in s[4]: c.font=H; c.fill=fill
r0=5
for i,(m,(col,txt)) in enumerate(MEAN.items(),r0):
    s.cell(i,1,m).fill=PatternFill('solid',fgColor=col); s.cell(i,2,f'=COUNTIF({MT},A{i})'); s.cell(i,3,f'=B{i}/$B${r0+len(MEAN)}'); s.cell(i,4,txt)
    s.cell(i,3).number_format='0.0%'
t=r0+len(MEAN); s.cell(t,1,'Total'); s.cell(t,2,f'=SUM(B{r0}:B{t-1})'); s.cell(t,3,f'=B{t}/B{t}'); s.cell(t,3).number_format='0.0%'
for c in s[t]: c.font=Font(name=F_,bold=True)
b0=t+3; s.cell(b0-1,1,'By brand').font=Font(name=F_,bold=True)
for j,h in enumerate(['Brand','Products','Matched','Review','No match','Matched %'],1):
    c=s.cell(b0,j,h); c.font=H; c.fill=fill
brands=[b for b,_ in collections.Counter(o['b'] for o in out).most_common()]
for i,bname in enumerate(brands,b0+1):
    s.cell(i,1,bname); s.cell(i,2,f'=COUNTIF({BR},A{i})')
    s.cell(i,3,f'=COUNTIFS({BR},A{i},{MT},"Matched*")'); s.cell(i,4,f'=COUNTIFS({BR},A{i},{MT},"Review*")'); s.cell(i,5,f'=COUNTIFS({BR},A{i},{MT},"No match*")')
    s.cell(i,6,f'=IF(B{i}=0,0,C{i}/B{i})'); s.cell(i,6).number_format='0.0%'
e=b0+1+len(brands); s.cell(e,1,'Total')
for col in 'BCDE': s[f'{col}{e}']=f'=SUM({col}{b0+1}:{col}{e-1})'
s[f'F{e}']=f'=IF(B{e}=0,0,C{e}/B{e})'; s[f'F{e}'].number_format='0.0%'
for c in s[e]: c.font=Font(name=F_,bold=True)
for i,x in enumerate(NOTES,e+2): s.cell(i,1,x).font=Font(name=F_,bold=(i==e+2),italic=(i>e+2))
for row in s.iter_rows(min_row=5,max_row=e-1):
    for c in row:
        if not c.font.bold: c.font=B
for col,w in zip('ABCDEF',[58,10,10,90,10,10]): s.column_dimensions[col].width=w
wo.save(dst)
