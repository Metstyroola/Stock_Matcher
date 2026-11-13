import sys, re, collections, openpyxl
PD, SUP, RES, DST = sys.argv[1:5]
clean=lambda v: (str(v).strip() if v not in (None,'') else None)
def cid(v):
    v=clean(v)
    if v is None: return None
    v=re.sub(r'\.0$','',v); return v if v.upper() not in ('FALSE','NONE','0','#N/A') else None

# ---- product data
ws=openpyxl.load_workbook(PD,read_only=True,data_only=True).active
it=ws.iter_rows(values_only=True); hdr=None
for r in it:
    if r and r[0]=='dp_id': hdr=r; break
ci={h:i for i,h in enumerate(hdr) if h and h not in ()}
iD,iE,iP=0,hdr.index('ean'),hdr.index('eprel_id')
pd={}; ep_used=collections.defaultdict(set); ean_used=collections.defaultdict(set)
for r in it:
    if not r or not r[0]: continue
    dp=str(r[0]).strip(); e=cid(r[iE]); p=cid(r[iP])
    pd[dp]=(e,p,r[hdr.index('brand')],r[hdr.index('profile')],r[hdr.index('dimensions')])
    if p: ep_used[p].add(dp)
    if e: ean_used[e].add(dp)
print('product data rows',len(pd),'with EPREL',sum(1 for v in pd.values() if v[1]),'with EAN',sum(1 for v in pd.values() if v[0]))

# ---- EPREL records (EAN per EPREL ID)
ws=openpyxl.load_workbook(SUP,read_only=True,data_only=True)['Sheet2']
eprel={}
for r in ws.iter_rows(min_row=2,values_only=True):
    if r and r[0] and r[4]: eprel[cid(r[0])]=dict(ean=cid(r[1]),name=clean(r[2]),desig=clean(r[13]))

# ---- matching results
wr=openpyxl.load_workbook(RES,read_only=True,data_only=True)['Matching Results']
rows=list(wr.iter_rows(values_only=True)); H=list(rows[0]); rows=rows[1:]
h={c:i for i,c in enumerate(H)}

# first pass: proposed EPREL per product (so we can spot the same ID proposed twice)
recs=[]
for r in rows:
    dp=str(r[h['DP ID']]).strip(); mt=r[h['Match Type']]
    cands=[x.strip() for x in str(r[h['All Matching EPREL IDs']] or '').split(',') if x.strip()]
    e_pd,p_pd,*_=pd.get(dp,(None,None,None,None,None)); inpd=dp in pd
    usable=mt.startswith('Matched -') and 'different OE' not in mt
    sug=None; why=''
    if cands:
        if p_pd and p_pd in cands: sug=p_pd; why='Product data EPREL ID is one of the matches'
        elif e_pd and [c for c in cands if eprel.get(c,{}).get('ean')==e_pd]:
            sug=[c for c in cands if eprel[c]['ean']==e_pd][0]; why='Chosen by EAN in product data'
        else:
            free=[c for c in cands if not (ep_used.get(c,set())-{dp})]
            sug=(free or cands)[0]; why='Only match' if len(cands)==1 else ('First match not used elsewhere' if free else 'All matches already used elsewhere')
    recs.append(dict(r=r,dp=dp,mt=mt,cands=cands,e_pd=e_pd,p_pd=p_pd,inpd=inpd,usable=usable,sug=sug,why=why))
prop=collections.defaultdict(set)
for x in recs:
    if x['sug'] and x['usable'] and not x['p_pd']: prop[x['sug']].add(x['dp'])

# second pass: actions + flags
for x in recs:
    cands,sug,dp=x['cands'],x['sug'],x['dp']
    others=[c for c in cands if c!=sug]
    oth=[]
    for c in others:
        tag=[]
        if x['p_pd']: tag.append('PD already has '+x['p_pd'])
        u=ep_used.get(c,set())-{dp}
        if u: tag.append('used by '+', '.join(sorted(u))[:60])
        oth.append(f"{c} ({'; '.join(tag) or 'alternative - do not add'})")
    x['others']=' | '.join(oth) or None
    if not x['inpd']: a='DP ID not in product data'
    elif not cands: a='Existing EPREL ID kept (no EPREL match)' if x['p_pd'] else 'No EPREL match'
    elif x['p_pd'] and x['p_pd']==sug: a='Already in product data - no change'
    elif x['p_pd']: a='Product data has a different EPREL ID - check'
    elif not x['usable']: a='Review match first - do not add yet'
    else:
        dup=ep_used.get(sug,set())-{dp}
        if dup: a='Do not add - EPREL ID already on another DP ID'; x['conf']=sorted(dup)
        elif len(prop[sug])>1: a=f'Check - same EPREL ID proposed for {len(prop[sug])} products'
        else: a='Add EPREL ID'
    x['ep_action']=a
    se=eprel.get(sug,{}).get('ean') if sug else None; x['sug_ean']=se
    if not se: ea='No EAN in EPREL' if sug else None
    elif x['e_pd']==se: ea='Same as product data'
    elif x['e_pd']: ea='Different from product data - check'
    elif ean_used.get(se,set())-{dp}: ea='Do not add - EAN already on another DP ID'; x['conf_e']=sorted(ean_used[se]-{dp})
    elif not x['usable'] or not a.startswith('Add'): ea='Hold - follow EPREL action'
    else: ea='Add EAN'
    x['ean_action']=ea
    def desc(ids):
        out=[]
        for d in ids[:3]:
            v=pd.get(d); out.append(f"{d}: {v[2]} {v[3]} {v[4]}" if v else d)
        return ' | '.join(out) or None
    cf=sorted(set(x.get('conf',[]))|set(x.get('conf_e',[])))
    x['conf_ids']=', '.join(cf) or None; x['conf_desc']=desc(cf)
    if cf:
        me=pd.get(dp)
        tk=lambda v: set(re.sub(r'[^A-Z0-9 ]',' ',str(v or '').upper()).split())
        sz=lambda v: re.sub(r'[^0-9R/.]','',str(v or '').upper().replace('ZR','R').split(' ')[0])
        def related(d):
            o=pd.get(d)
            if not (o and me) or sz(o[4])!=sz(me[4]): return False
            a,b=tk(me[3]),tk(o[3]); return a<=b or b<=a
        rel=[d for d in cf if related(d)]
        x['dupflag']=('Same tyre already listed under another DP ID (variant / EV listing)' if rel
                      else 'Different product holds this ID - check')
    else: x['dupflag']=None

# ---- write workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L_
F='Arial'; HF=Font(name=F,bold=True,color='FFFFFF'); HFILL=PatternFill('solid',fgColor='1F3864'); B=Font(name=F)
def col(a):
    if a is None: return None
    if a.startswith(('Add','Already','Same as')): return 'C6EFCE'
    if a.startswith(('Do not','Different','Product data has','Possible duplicate')): return 'F8CBAD'
    if a.startswith(('Check','Review','Hold')): return 'FFEB9C'
    return 'EDEDED'
wo=openpyxl.Workbook(); ws=wo.active; ws.title='Results + Product Data'
keep=['Spinach ID','DP ID','Brand','Brand Segment','Profile','Profile Text','dimension','Match Type']
new=['PD EPREL ID','PD EAN','Suggested EPREL ID','Why this ID','EPREL Action','Suggested EAN','EAN Action','Conflicting DP ID','Conflicting DP ID details','Conflict Note','Other EPREL IDs (flagged)','No. of Matching IDs']
tail=['EPREL Commercial Name','EPREL Tyre Designation','Energy Class','Wet Grip Class','Noise Class','Noise Value (dB)','EPREL Patterns at this Size']
ws.append(keep+new+tail)
for x in recs:
    r=x['r']
    ws.append([r[h[k]] for k in keep]+[x['p_pd'],x['e_pd'],x['sug'],x['why'] or None,x['ep_action'],x['sug_ean'],x['ean_action'],x['conf_ids'],x['conf_desc'],x['dupflag'],x['others'],len(x['cands']) or None]+[r[h[k]] for k in tail])
for i,w in enumerate([15,11,13,10,36,26,20,40, 12,16,14,30,44,16,34,14,50,30,60,9, 26,24,8,8,8,8,50],1): ws.column_dimensions[L_(i)].width=w
for c in ws[1]: c.font=HF; c.fill=HFILL; c.alignment=Alignment(wrap_text=True,vertical='center')
for row in ws.iter_rows(min_row=2):
    for c in row: c.font=B
    for j in (12,14,17):
        k=col(row[j].value)
        if k: row[j].fill=PatternFill('solid',fgColor=k)
ws.freeze_panes='C2'; ws.auto_filter.ref=ws.dimensions
n=ws.max_row; R="'Results + Product Data'"
s=wo.create_sheet('Summary',0)
s['A1']='Other suppliers - EPREL matches vs product data'; s['A1'].font=Font(name=F,bold=True,size=14)
def table(r0,title,colL,values):
    s.cell(r0,1,title).font=Font(name=F,bold=True)
    for j,t in enumerate(['Action','Products'],1): c=s.cell(r0+1,j,t); c.font=HF; c.fill=HFILL
    for i,v in enumerate(values,r0+2):
        s.cell(i,1,v).font=B; s.cell(i,2,f'=COUNTIF({R}!${colL}$2:${colL}${n},A{i})').font=B
        k=col(v); 
        if k: s.cell(i,1).fill=PatternFill('solid',fgColor=k)
    t=r0+2+len(values); s.cell(t,1,'Total').font=Font(name=F,bold=True); s.cell(t,2,f'=SUM(B{r0+2}:B{t-1})').font=Font(name=F,bold=True)
    return t+2
order=lambda key: [v for v,_ in collections.Counter(x[key] for x in recs if x[key]).most_common()]
nxt=table(3,'EPREL ID action','M',order('ep_action'))
nxt=table(nxt,'EAN action','O',order('ean_action'))
nxt=table(nxt,'Conflicts with product data','R',order('dupflag'))
notes=['How the EPREL ID is chosen when a product has several matches',
 '1. If product data already has an EPREL ID and it is one of the matches, that ID is kept; every other match is flagged "PD already has ...".',
 '2. Otherwise, if the product data EAN equals the EAN of one EPREL match, that match is chosen.',
 '3. Otherwise the first match not already used by another DP ID in product data is chosen.',
 'Duplicate guard: an EPREL ID (or EAN) already stored on another DP ID in product data is never suggested as "Add". The other DP ID and its profile/size are shown; if it is the same profile and size, it is flagged as a possible duplicate product.',
 'If the same EPREL ID would be added to more than one product from this list, those rows are marked "Check - same EPREL ID proposed for N products".',
 'Only "Matched - pattern" and "Matched - model code" rows can be "Add". "Different OE variant" and all Review rows are held until checked.',
 f'Product data: {len(pd):,} DP IDs, {sum(1 for v in pd.values() if v[1]):,} with an EPREL ID, {sum(1 for v in pd.values() if v[0]):,} with an EAN.']
for i,t in enumerate(notes,nxt): s.cell(i,1,t).font=Font(name=F,bold=(i==nxt),italic=(i>nxt))
s.column_dimensions['A'].width=70; s.column_dimensions['B'].width=12
wo.save(DST)
print(collections.Counter(x['ep_action'][:45] for x in recs)); print(collections.Counter(x['ean_action'] if not x['ean_action'] or not x['ean_action'].startswith('Do not') else 'Do not add EAN (dup)' for x in recs))
