#!/usr/bin/env python3
"""Derivador+validador fail-closed del conjunto privado Gramps.
Uso: priv_derive.py <ged_canon> <base_url>   (token via GRAMPS_USER/GRAMPS_PASS)
Salida stdout: JSON {targets:[{type,gramps_id,handle,fingerprint?,content_sha256}], counts, sha256_por_tipo, total}
FAIL CLOSED: cualquier mismatch GED vs API -> exit 1, sin targets.
Solo LECTURA de la API.
"""
import json, sys, re, time, hashlib, urllib.request, urllib.error, os
from collections import Counter

def norm_text(s):
    return re.sub(r'\s+',' ',s).strip()

def parse_ged(path):
    ged=open(path,encoding='utf-8').read()
    recs=re.split(r'\n(?=0 @)', ged)
    persons={}; fams={}; sources={}
    for rec in recs:
        m=re.match(r'0 @([PSF])(\w*)@ (\w+)', rec)
        if not m: continue
        typ,pid=m.group(1),m.group(2)
        if typ=='P': persons[pid]=rec
        elif typ=='F': fams[pid]=rec
        elif typ=='S':
            t=re.search(r'\n1 TITL (.+)',rec)
            nts=[norm_text(nm.group(1).replace('\n2 CONC ','')) for nm in re.finditer(r'(?:^|\n)1 NOTE ((?:[^\n]*)((?:\n2 CONC [^\n]*)*))',rec)]
            sources[pid]={'title':t.group(1).strip() if t else '','notes':nts}
    return persons,fams,sources

def ged_expected(persons,fams,sources=None):
    resn={pid for pid,r in persons.items() if '1 RESN privacy' in r}
    # guarda: citas fuera de nivel persona no soportadas por esta version.
    # nivel-1 SOUR en INDI = cita de persona (soportada); cualquier '2 SOUR' dentro de
    # un registro INDI/FAM cuelga de una subestructura (evento u otra) -> no soportado.
    for pid,r in persons.items():
        if re.search(r'\n2 SOUR ',r):
            raise SystemExit('FAIL_CLOSED: cita bajo subestructura de persona no soportada (objeto %s)'%pid)
    for fid,r in fams.items():
        if re.search(r'(?:^|\n)1 SOUR ',r):
            raise SystemExit('FAIL_CLOSED: cita a nivel familia no soportada (objeto %s)'%fid)
        if re.search(r'\n2 SOUR ',r):
            raise SystemExit('FAIL_CLOSED: cita bajo subestructura de familia no soportada (objeto %s)'%fid)
    exp={'people':set(),'events':set(),'families':set(),'notes':[],'citations':[],'sources':set()}
    note_texts=[]
    exp['notes_bo']={}
    def addn(owner,txt):
        exp['notes_bo'].setdefault(owner,Counter()); exp['notes_bo'][owner][txt]+=1
        note_texts.append(txt)
    exp['ev_notes']={}   # clave evento GED -> Counter notas del evento
    for pid in resn:
        exp['people'].add(pid); r=persons[pid]
        for em in re.finditer(r'(?:^|\n)1 (BIRT|DEAT)((?:\n[2-9] .*)*)',r):
            check_ev_subtags(em.group(2),'P'+pid+':'+em.group(1))
            exp['events'].add(pid+':'+em.group(1))
            ek=pid+':'+em.group(1)
            for nm in re.finditer(r'\n2 NOTE ((?:[^\n]*)((?:\n3 CONC [^\n]*)*))',em.group(2)):
                t=norm_text(nm.group(1).replace('\n3 CONC ',''))
                addn('P'+pid,t); exp['ev_notes'].setdefault(ek,Counter()); exp['ev_notes'][ek][t]+=1
        for nm in re.finditer(r'(?:^|\n)1 NOTE ((?:[^\n]*)((?:\n2 CONC [^\n]*)*))',r):
            addn('P'+pid,norm_text(nm.group(1).replace('\n2 CONC ','')))
        for sm in re.finditer(r'(?:^|\n)1 SOUR @S(\w+)@((?:\n[2-9] .*)*)',r):
            pg=re.search(r'\n2 PAGE ((?:[^\n]*)((?:\n3 CONC [^\n]*)*))',sm.group(2))
            cid=re.search(r'\| (C\d+) \|',pg.group(1)) if pg else None
            lbl=cid.group(1) if cid else 'NOLBL:'+pid+':'+sm.group(1)
            exp['citations'].append(lbl)
            exp['sources'].add(sm.group(1))
            if cid and pg:
                exp.setdefault('cit_ged',{})[cid.group(1)]=(pg.group(1).replace('\n3 CONC ','').strip(), sm.group(1))
            for nm in re.finditer(r'\n2 NOTE ((?:[^\n]*)((?:\n3 CONC [^\n]*)*))',sm.group(2)):
                addn('C:'+lbl,norm_text(nm.group(1).replace('\n3 CONC ','')))
    for fid,r in fams.items():
        members=set(re.findall(r'1 (?:HUSB|WIFE|CHIL) @P(\w+)@',r))
        if members & resn:
            exp['families'].add(fid)
            okey='F:'+','.join(sorted('P'+m for m in members))
            for em in re.finditer(r'(?:^|\n)1 (MARR)((?:\n[2-9] .*)*)',r):
                check_ev_subtags(em.group(2),fid+':MARR')
                exp['events'].add(fid+':MARR')
                ek=fid+':MARR'
                for nm in re.finditer(r'\n2 NOTE ((?:[^\n]*)((?:\n3 CONC [^\n]*)*))',em.group(2)):
                    t=norm_text(nm.group(1).replace('\n3 CONC ',''))
                    addn(okey,t); exp['ev_notes'].setdefault(ek,Counter()); exp['ev_notes'][ek][t]+=1
            for nm in re.finditer(r'(?:^|\n)1 NOTE ((?:[^\n]*)((?:\n2 CONC [^\n]*)*))',r):
                addn(okey,norm_text(nm.group(1).replace('\n2 CONC ','')))
    # v17 (hallazgo BO): las notas de las FUENTES seleccionadas son objetos API
    # independientes: privatizar la fuente NO privatiza su nota. Se incluyen todas.
    if sources:
        for sid in sorted(exp['sources']):
            st_=sources.get(sid,{})
            for t in st_.get('notes',[]):
                addn('S:'+st_.get('title',''),t)
    exp['notes']=Counter(note_texts)
    return exp

_AUTH={}
def http(method,url,payload=None,token=None):
    if token is not None and _AUTH: token=_AUTH.get('tok')
    data=json.dumps(payload).encode() if payload is not None else None
    def build():
        req=urllib.request.Request(url,data=data,method=method)
        if data: req.add_header('Content-Type','application/json')
        if token: req.add_header('Authorization','Bearer '+token)
        return req
    for _ in range(8):
        try:
            with urllib.request.urlopen(build(),timeout=60) as r: return r.status,json.loads(r.read())
        except urllib.error.HTTPError as e:
            if e.code==401 and token and _AUTH:
                _AUTH['tok']=get_token(_AUTH['base'],_AUTH['u'],_AUTH['p']); token=_AUTH['tok']; continue
            if e.code==429: time.sleep(4); continue
            try: return e.code,json.loads(e.read())
            except Exception: return e.code,None
        except Exception: time.sleep(2)
    return 599,None

def get_token(base,u,p):
    st,res=http('POST',base+'/api/token/',{'username':u,'password':p})
    if isinstance(res,dict) and 'access_token' in res: return res['access_token']
    raise SystemExit('ERROR token: '+str(st))

def get_all(base,kind,tok):
    st,res=http('GET','%s/api/%s/?pagesize=2000'%(base,kind),token=tok)
    if st!=200 or not isinstance(res,list): raise SystemExit('ERROR GET %s: %s'%(kind,st))
    if len(res)>=2000: raise SystemExit('FAIL_CLOSED: lista %s alcanza pagesize=2000; posible truncamiento por paginacion no seguida (guarda v16.3)'%kind)
    return res



def check_ev_subtags(body, ctx):
    """v16 (gate BO): contenido completo o aborto. Cualquier subtag de evento no cubierto
    por la clave -> fail-closed. El nombre del tag no es contenido: se reporta tal cual."""
    for l in body.split(chr(10)):
        m=re.match(r'(\d) (\w+)', l)
        if not m: continue
        lvl,tag=int(m.group(1)),m.group(2)
        if lvl==2 and tag not in ('DATE','PLAC','NOTE','SOUR','OBJE'):
            raise ValueError('subtag evento no cubierto: '+tag+' ('+ctx+')')
        if lvl>=3 and tag not in ('CONC','CONT'):
            raise ValueError('subtag nivel %d de evento no cubierto: %s (%s)'%(lvl,tag,ctx))

def ged_event_groups(persons,fams,priv_people,priv_fams):
    """v15: agrupa TODOS los eventos GED por clave de contenido (sin propietario).
    Devuelve (grupos, tratamiento) donde tratamiento[okey]=bool privado.
    Clave identica en forma a la del detector API en main()."""
    tmap={'BIRT':'birth','DEAT':'death','MARR':'marriage'}
    def nd(s):
        if not s: return None
        return s.strip().lower().replace('about','abt')
    groups={}
    treat={}
    def add(okey,tag,body,priv):
        check_ev_subtags(body,okey+':'+tag)
        d=re.search(r'\n2 DATE (.+)',body)
        pl=re.search(r'\n2 PLAC ((?:[^\n]*)((?:\n3 CONC [^\n]*)*))',body)
        nts=[]
        for nm in re.finditer(r'\n2 NOTE ((?:[^\n]*)((?:\n3 CONC [^\n]*)*))',body):
            nts.append(norm_text(nm.group(1).replace('\n3 CONC ','')))
        nc=len(re.findall(r'\n2 SOUR ',body)); nm2=len(re.findall(r'\n2 OBJE',body))
        ck=(tmap[tag],nd(d.group(1)) if d else None,(norm_text(pl.group(1).replace('\n3 CONC ','')).strip().lower() if pl else None),tuple(sorted(Counter(nts).items())),nc,nm2)
        groups.setdefault(ck,[]).append(okey); treat[okey]=priv
    for pid,r in persons.items():
        for em in re.finditer(r'(?:^|\n)1 (BIRT|DEAT)((?:\n[2-9] .*)*)',r):
            add('P'+pid,em.group(1),em.group(2),pid in priv_people)
    for fid,r in fams.items():
        okey='F:'+','.join(sorted('P'+m for m in re.findall(r'1 (?:HUSB|WIFE|CHIL) @P(\w+)@',r)))
        for em in re.finditer(r'(?:^|\n)1 (MARR)((?:\n[2-9] .*)*)',r):
            add(okey,'MARR',em.group(2),fid in priv_fams)
    return groups,treat

def crossing_shapes(groups,treat):
    """Formas redactadas de grupos multi-propietario: (sha clave, n miembros, cruza frontera)."""
    out=[]
    for ck,owners in groups.items():
        if len(set(owners))<2: continue
        sides={treat.get(o,False) for o in owners}
        out.append((hashlib.sha256(repr(ck).encode()).hexdigest(),len(owners),len(sides)>1))
    return sorted(out)

def main():
    ged_path,base=sys.argv[1],sys.argv[2].rstrip('/')
    persons,fams,sources_ged=parse_ged(ged_path)
    exp=ged_expected(persons,fams,sources_ged)
    tok=get_token(base,os.environ['GRAMPS_USER'],os.environ['GRAMPS_PASS'])
    _AUTH.update(base=base,u=os.environ['GRAMPS_USER'],p=os.environ['GRAMPS_PASS'],tok=tok)
    api_people=get_all(base,'people',tok)
    st,api_events=http('GET','%s/api/events/?profile=all&pagesize=2000'%base,token=tok)
    if st!=200 or not isinstance(api_events,list): raise SystemExit('ERROR GET events profile: %s'%st)
    if len(api_events)>=2000: raise SystemExit('FAIL_CLOSED: lista events(profile=all) alcanza pagesize=2000; posible truncamiento por paginacion no seguida (guarda v16.4)')
    api_places=get_all(base,'places',tok)
    api_fams=get_all(base,'families',tok)
    api_notes=get_all(base,'notes',tok)
    api_cits=get_all(base,'citations',tok)
    api_srcs=get_all(base,'sources',tok)
    fails=[]
    # 1) personas RESN: GED vs API (atributo RESN importado)
    def is_resn(a):
        t=a.get('type')
        s=t.get('string') if isinstance(t,dict) else t
        return s=='RESN'
    api_resn=[p for p in api_people if any(is_resn(a) for a in p.get('attribute_list',[]))]
    if len(api_resn)!=len(exp['people']): fails.append('personas RESN: GED %d vs API %d'%(len(exp['people']),len(api_resn)))
    # identidad personas: conjunto de gramps_ids
    api_resn_ids={p.get('gramps_id','') for p in api_resn}
    ged_resn_ids={'P'+pid for pid in exp['people']}
    if api_resn_ids!=ged_resn_ids:
        fails.append('personas RESN identidad: GED-API %s API-GED %s'%(sorted(ged_resn_ids-api_resn_ids)[:3],sorted(api_resn_ids-ged_resn_ids)[:3]))
    handle2gid={p['handle']:p.get('gramps_id','') for p in api_people}
    resn_handles={p['handle'] for p in api_resn}
    # 2) familias con >=1 miembro RESN
    fam_t=[]
    for f in api_fams:
        mem=set()
        if f.get('father_handle'): mem.add(f['father_handle'])
        if f.get('mother_handle'): mem.add(f['mother_handle'])
        mem|={c['ref'] for c in f.get('child_ref_list',[])}
        if mem & resn_handles: fam_t.append(f)
    if len(fam_t)!=len(exp['families']): fails.append('familias: GED %d vs API %d'%(len(exp['families']),len(fam_t)))
    # identidad familias: multiset de conjuntos de miembros (por gramps_id)
    ged_fam_members=[]
    for fid in exp['families']:
        r=fams[fid]
        ged_fam_members.append(tuple(sorted('P'+m for m in re.findall(r'1 (?:HUSB|WIFE|CHIL) @P(\w+)@',r))))
    api_fam_members=[]
    for f in fam_t:
        mem=set()
        if f.get('father_handle'): mem.add(handle2gid.get(f['father_handle'],'?'))
        if f.get('mother_handle'): mem.add(handle2gid.get(f['mother_handle'],'?'))
        mem|={handle2gid.get(c['ref'],'?') for c in f.get('child_ref_list',[])}
        api_fam_members.append(tuple(sorted(mem)))
    if Counter(api_fam_members)!=Counter(ged_fam_members):
        fails.append('familias identidad (miembros) mismatch')
    # guarda API: citas en familias/eventos del cluster no soportadas
    for f in fam_t:
        if f.get('citation_list'): fails.append('familia con citation_list no soportada: '+f.get('gramps_id','?')); break
    fam_handles={f['handle'] for f in fam_t}
    # 3) eventos: referenciados por personas RESN + eventos de familias RESN
    ev_t=set()
    for p in api_resn:
        for er in p.get('event_ref_list',[]): ev_t.add(er['ref'] if isinstance(er,dict) else er)
    for f in fam_t:
        for er in f.get('event_ref_list',[]): ev_t.add(er['ref'] if isinstance(er,dict) else er)
    if len(ev_t)!=len(exp['events']): fails.append('eventos: GED %d vs API %d'%(len(exp['events']),len(ev_t)))
    for h in ev_t:
        e=next((x for x in api_events if x['handle']==h),None)
        if e and e.get('citation_list'): fails.append('evento con citation_list no soportado: '+e.get('gramps_id','?')); break
        if e and e.get('media_list'): fails.append('evento con media_list no soportado: '+e.get('gramps_id','?')); break
    # identidad eventos: multiset (tipo, fecha, lugar) GED vs API
    plname={p['handle']:((p.get('name') or {}).get('value','') if isinstance(p.get('name'),dict) else str(p.get('name',''))).lower() for p in api_places}
    MONTHS={'JAN':1,'FEB':2,'MAR':3,'APR':4,'MAY':5,'JUN':6,'JUL':7,'AUG':8,'SEP':9,'OCT':10,'NOV':11,'DEC':12}
    def nd(s):
        if not s: return None
        return s.strip().lower().replace('about','abt')
    ged_ev_ms=[]
    for pid in exp['people']:
        r=persons[pid]
        for em in re.finditer(r'(?:^|\n)1 (BIRT|DEAT)((?:\n[2-9] .*)*)',r):
            d=re.search(r'\n2 DATE (.+)',em.group(2)); pl=re.search(r'\n2 PLAC (.+)',em.group(2))
            nk=tuple(sorted(exp['ev_notes'].get(pid+':'+em.group(1),{}).items()))
            nc=len(re.findall(r'\n2 SOUR ',em.group(2))); nm2=len(re.findall(r'\n2 OBJE',em.group(2)))
            ged_ev_ms.append(('P'+pid,em.group(1),nd(d.group(1)) if d else None,(pl.group(1).strip().lower() if pl else None),nk,nc,nm2))
    for fid in exp['families']:
        r=fams[fid]
        # propietario por CONTENIDO (miembros), no por FID: los IDs de familia pueden renumerarse en el import
        okey='F:'+','.join(sorted('P'+m for m in re.findall(r'1 (?:HUSB|WIFE|CHIL) @P(\w+)@',r)))
        for em in re.finditer(r'(?:^|\n)1 (MARR)((?:\n[2-9] .*)*)',r):
            d=re.search(r'\n2 DATE (.+)',em.group(2)); pl=re.search(r'\n2 PLAC (.+)',em.group(2))
            nk=tuple(sorted(exp['ev_notes'].get(fid+':MARR',{}).items()))
            nc=len(re.findall(r'\n2 SOUR ',em.group(2))); nm2=len(re.findall(r'\n2 OBJE',em.group(2)))
            ged_ev_ms.append((okey,em.group(1),nd(d.group(1)) if d else None,(pl.group(1).strip().lower() if pl else None),nk,nc,nm2))
    tmap={'BIRT':'birth','DEAT':'death','MARR':'marriage'}
    # identidad eventos CON PROPIETARIO (persona/familia que lo referencia): cierra el intercambio
    # entre eventos gemelos (mismo tipo/fecha/lugar) de personas distintas.
    ev_by_handle={e['handle']:e for e in api_events}
    ntext={n['handle']:norm_text(n.get('text',{}).get('string','') if isinstance(n.get('text'),dict) else str(n.get('text',''))) for n in api_notes}
    def ev_nk(e):
        cn=Counter(ntext.get(hh,'') for hh in e.get('note_list',[]))
        return tuple(sorted(cn.items()))
    api_ev_ms=[]
    for per in api_resn:
        for er in per.get('event_ref_list',[]):
            e=ev_by_handle[er['ref'] if isinstance(er,dict) else er]
            d=e.get('date') or {}
            ds=nd(d.get('text')) if isinstance(d,dict) else None
            api_ev_ms.append((per.get('gramps_id',''),str(e.get('type')).lower(),ds,plname.get(e.get('place')),ev_nk(e),len(e.get('citation_list',[])),len(e.get('media_list',[]))))
    for f in fam_t:
        mem=set()
        if f.get('father_handle'): mem.add(handle2gid.get(f['father_handle'],'?'))
        if f.get('mother_handle'): mem.add(handle2gid.get(f['mother_handle'],'?'))
        mem|={handle2gid.get(c['ref'],'?') for c in f.get('child_ref_list',[])}
        okey='F:'+','.join(sorted(mem))
        for er in f.get('event_ref_list',[]):
            e=ev_by_handle[er['ref'] if isinstance(er,dict) else er]
            d=e.get('date') or {}
            ds=nd(d.get('text')) if isinstance(d,dict) else None
            api_ev_ms.append((okey,str(e.get('type')).lower(),ds,plname.get(e.get('place')),ev_nk(e),len(e.get('citation_list',[])),len(e.get('media_list',[]))))
    ged_ms=Counter((o,tmap[t],d,pl,nk,nc,nm2) for o,t,d,pl,nk,nc,nm2 in ged_ev_ms)
    api_ms=Counter(api_ev_ms)
    if ged_ms!=api_ms:
        h=lambda items:[(hashlib.sha256(repr(k).encode()).hexdigest(),c) for k,c in items]
        fails.append('eventos identidad multiset (sha256): GED-API %s API-GED %s'%(h(list((ged_ms-api_ms).items())[:3]),h(list((api_ms-ged_ms).items())[:3])))
    # 4) citas de personas RESN + mapeo por etiqueta C en page
    cit_t=[c for c in api_cits if any(ch==c['handle'] for p in api_resn for ch in p.get('citation_list',[]))]
    ged_labels=[c for c in exp['citations'] if c.startswith('C')]
    api_labels=[]
    for c in cit_t:
        m=re.search(r'\| (C\d+) \|',c.get('page','') or '')
        api_labels.append(m.group(1) if m else 'NOLBL:'+c['handle'])
    if Counter(api_labels)!=Counter(ged_labels):
        fails.append('citas mapeo C-label: GED %d vs API %d (diff %s)'%(len(ged_labels),len(api_labels),list((Counter(list(ged_labels))-Counter(api_labels)).items())[:5]))
    if len(cit_t)!=len(exp['citations']): fails.append('citas: GED %d vs API %d'%(len(exp['citations']),len(cit_t)))
    # 5) fuentes: todas las citadas por personas RESN (incl compartidas) via source_handle
    src_handles={c.get('source_handle') for c in cit_t if c.get('source_handle')}
    src_t=[s for s in api_srcs if s['handle'] in src_handles]
    if len(src_t)!=len(exp['sources']): fails.append('fuentes: GED %d vs API %d'%(len(exp['sources']),len(src_t)))
    ged_titles=Counter((sources_ged.get(s) or {}).get('title','') for s in exp['sources'])
    api_titles=Counter(s.get('title','') for s in src_t)
    if ged_titles!=api_titles:
        h=lambda items:[(hashlib.sha256(k.encode()).hexdigest(),c) for k,c in items]
        fails.append('fuentes titulos mismatch (sha256): GED-API %s API-GED %s'%(h(list((ged_titles-api_titles).items())[:3]),h(list((api_titles-ged_titles).items())[:3])))
    # 6) notas: universo por contenido (multiset hash) entre objetos API alcanzables desde el cluster y textos GED
    note_handles=set()
    for p in api_resn:
        note_handles|=set(p.get('note_list',[]))
        for er in p.get('event_ref_list',[]):
            h=er['ref'] if isinstance(er,dict) else er
            ev=next((e for e in api_events if e['handle']==h),None)
            if ev: note_handles|=set(ev.get('note_list',[]))
        for ch in p.get('citation_list',[]):
            c=next((x for x in api_cits if x['handle']==ch),None)
            if c: note_handles|=set(c.get('note_list',[]))
    for f in fam_t:
        note_handles|=set(f.get('note_list',[]))
        for er in f.get('event_ref_list',[]):
            h=er['ref'] if isinstance(er,dict) else er
            ev=next((e for e in api_events if e['handle']==h),None)
            if ev: note_handles|=set(ev.get('note_list',[]))
    # v17 (hallazgo BO): notas de las fuentes SELECCIONADAS (objetos independientes:
    # privatizar la fuente no privatiza su nota)
    for s2 in src_t:
        note_handles|=set(s2.get('note_list',[]))
    api_note_texts=Counter(norm_text(n.get('text',{}).get('string','') if isinstance(n.get('text'),dict) else str(n.get('text',''))) for n in api_notes if n['handle'] in note_handles)
    exp_all=Counter(); exp_all.update(exp['notes'])
    # 6b) notas POR PROPIETARIO (persona/familia por contenido/cita por etiqueta): cierra
    # intercambios de notas gemelas entre propietarios distintos.
    api_bo={}
    def addbo(owner,txt):
        api_bo.setdefault(owner,Counter()); api_bo[owner][txt]+=1
    for p in api_resn:
        ok_=p.get('gramps_id','')
        for hh in p.get('note_list',[]): addbo(ok_,ntext.get(hh,''))
        for er in p.get('event_ref_list',[]):
            ev=ev_by_handle.get(er['ref'] if isinstance(er,dict) else er)
            if ev:
                for hh in ev.get('note_list',[]): addbo(ok_,ntext.get(hh,''))
        for ch in p.get('citation_list',[]):
            c=next((x for x in api_cits if x['handle']==ch),None)
            if c:
                mm=re.search(r'\| (C\d+) \|',c.get('page','') or '')
                lbl=mm.group(1) if mm else 'NOLBL:'+c['handle']
                for hh in c.get('note_list',[]): addbo('C:'+lbl,ntext.get(hh,''))
    for f in fam_t:
        mem=set()
        if f.get('father_handle'): mem.add(handle2gid.get(f['father_handle'],'?'))
        if f.get('mother_handle'): mem.add(handle2gid.get(f['mother_handle'],'?'))
        mem|={handle2gid.get(c2['ref'],'?') for c2 in f.get('child_ref_list',[])}
        ok_='F:'+','.join(sorted(mem))
        for hh in f.get('note_list',[]): addbo(ok_,ntext.get(hh,''))
        for er in f.get('event_ref_list',[]):
            ev=ev_by_handle.get(er['ref'] if isinstance(er,dict) else er)
            if ev:
                for hh in ev.get('note_list',[]): addbo(ok_,ntext.get(hh,''))
    for s2 in src_t:
        ok_='S:'+(s2.get('title','') or '')
        for hh in s2.get('note_list',[]): addbo(ok_,ntext.get(hh,''))
    ged_bo=exp['notes_bo']
    if set(ged_bo)!=set(api_bo):
        fails.append('notas propietarios: GED-API %s API-GED %s'%(sorted(set(ged_bo)-set(api_bo))[:5],sorted(set(api_bo)-set(ged_bo))[:5]))
    else:
        for ok_ in ged_bo:
            if ged_bo[ok_]!=api_bo[ok_]:
                g,a=ged_bo[ok_],api_bo[ok_]
                hh=lambda items:[(hashlib.sha256(k.encode()).hexdigest(),c) for k,c in items]
                fails.append('notas propietario %s mismatch (sha256): GED-API %s API-GED %s'%(ok_,hh(list((g-a).items())[:3]),hh(list((a-g).items())[:3])))
    # notas de cita ya incluidas en exp['notes'] via bucles anteriores
    if api_note_texts!=exp_all:
        fails.append('notas texto mismatch: GED %d vs API %d; diff GED-API(sha256) %s; API-GED(sha256) %s'%(sum(exp_all.values()),sum(api_note_texts.values()),[(hashlib.sha256(k.encode()).hexdigest(),c) for k,c in list((exp_all-api_note_texts).items())[:3]],[(hashlib.sha256(k.encode()).hexdigest(),c) for k,c in list((api_note_texts-exp_all).items())[:3]]))
    if len(note_handles)!=sum(exp_all.values()): fails.append('notas count: GED %d vs API %d'%(sum(exp_all.values()),len(note_handles)))
    if fails:
        print(json.dumps({'status':'FAIL_CLOSED','fails':fails},ensure_ascii=False)); sys.exit(1)
    # 7) ambigüedad cross-owner (v15). Sin provenance independiente GED->handle en la API
    #    (verificado empíricamente: objetos sin atributos xref; solo RESN en personas y etiquetas
    #    C en pages de citas), un intercambio de handles entre objetos contenido-idénticos es
    #    indetectable por binding. Si el grupo es MISMO-LADO (todos priv o todos pub) el swap es
    #    INCONSECUENTE para privacidad (mismo tratamiento): informativo. Si CRUZA la frontera,
    #    tratamiento conservador: privatizar TODOS los miembros del grupo (+notas dependientes de
    #    los eventos añadidos). Evento sin propietario en grupo multi: fail-closed.
    resn_ids={p.get('gramps_id','') for p in api_resn}
    def fam_okey(f):
        mem=set()
        if f.get('father_handle'): mem.add(handle2gid.get(f['father_handle'],'?'))
        if f.get('mother_handle'): mem.add(handle2gid.get(f['mother_handle'],'?'))
        mem|={handle2gid.get(c['ref'],'?') for c in f.get('child_ref_list',[])}
        return 'F:'+','.join(sorted(mem))
    fam_priv_okeys={fam_okey(f) for f in fam_t}
    ev_owners={}
    for p in api_people:
        for er in p.get('event_ref_list',[]):
            ev_owners.setdefault(er['ref'] if isinstance(er,dict) else er,set()).add('P:'+p.get('gramps_id',''))
    for f in api_fams:
        ok2=fam_okey(f)
        for er in f.get('event_ref_list',[]):
            ev_owners.setdefault(er['ref'] if isinstance(er,dict) else er,set()).add('F:'+ok2)
    def owner_priv(o):
        k,v=o.split(':',1)
        return (k=='P' and v in resn_ids) or (k=='F' and v in fam_priv_okeys)
    ev_groups={}
    for e in api_events:
        d=e.get('date') or {}
        ds=nd(d.get('text')) if isinstance(d,dict) else None
        ck=(str(e.get('type')).lower(),ds,plname.get(e.get('place')),ev_nk(e),len(e.get('citation_list',[])),len(e.get('media_list',[])))
        ev_groups.setdefault(ck,[]).append(e['handle'])
    amb={'events_crossing':[],'events_same_side':0,'notes_crossing':[],'cits_shared_crossing':0,'srcs_crossing':0}
    ev_added=[]; note_added=[]
    for ck,hs in ev_groups.items():
        if len(hs)<2: continue
        own=set(); orphan=False
        for h in hs:
            os2=ev_owners.get(h,set())
            if not os2: orphan=True
            own|=os2
        if orphan:
            fails.append('evento sin propietario en grupo multi (sha clave): '+hashlib.sha256(repr(ck).encode()).hexdigest()[:16]); continue
        if len(own)<2: continue  # duplicado mismo propietario: mismo tratamiento, inconsecuente
        if len({owner_priv(o) for o in own})==1:
            amb['events_same_side']+=1; continue
        nuevos=[h for h in hs if h not in ev_t]
        if any(ev_by_handle[h].get('citation_list') or ev_by_handle[h].get('media_list') for h in nuevos):
            fails.append('grupo cross-frontera con dependientes no soportados (citas/media) en miembro añadido (sha clave): '+hashlib.sha256(repr(ck).encode()).hexdigest()[:16]); continue
        for h in nuevos:
            ev_t.add(h); ev_added.append(h)
            note_handles|=set(ev_by_handle[h].get('note_list',[]))
        amb['events_crossing'].append({'key_sha256':hashlib.sha256(repr(ck).encode()).hexdigest(),'members':len(hs)})
    # notas: mismo texto en objetos con tratamiento distinto (frontera de seleccion).
    # v17: universo INCLUYE notas de fuentes (politica BO 18/10: escaneo cross uniforme;
    # una nota placeholder identica en fuente NO seleccionada entra por conservative_ambiguity
    # y su fuente permanece publica)
    note_universe=set()
    for p in api_people: note_universe|=set(p.get('note_list',[]))
    for f in api_fams: note_universe|=set(f.get('note_list',[]))
    for e in api_events: note_universe|=set(e.get('note_list',[]))
    for c in api_cits: note_universe|=set(c.get('note_list',[]))
    for s3 in api_srcs: note_universe|=set(s3.get('note_list',[]))
    n_groups={}
    for h in note_universe: n_groups.setdefault(ntext.get(h,''),[]).append(h)
    for txt,hs in n_groups.items():
        if len(hs)<2: continue
        if len({h in note_handles for h in hs})==1: continue
        for h in hs:
            if h not in note_handles:
                note_handles.add(h); note_added.append(h)
        amb['notes_crossing'].append({'text_sha256':hashlib.sha256(txt.encode()).hexdigest(),'members':len(hs)})
    # citas compartidas cruzando (objeto unico, sobre-privatizacion forzada): informativo
    cit_ref={}
    for p in api_people:
        for ch in p.get('citation_list',[]): cit_ref.setdefault(ch,set()).add(p.get('gramps_id','') in resn_ids)
    amb['cits_shared_crossing']=sum(1 for s2 in cit_ref.values() if len(s2)>1)
    # fuentes seleccionadas tambien citadas por personas pub (objeto unico): informativo
    src_cross=set()
    for p in api_people:
        if p.get('gramps_id','') in resn_ids: continue
        for ch in p.get('citation_list',[]):
            c=next((x for x in api_cits if x['handle']==ch),None)
            if c and c.get('source_handle') in src_handles: src_cross.add(c['source_handle'])
    amb['srcs_crossing']=len(src_cross)
    # guarda GED-side: las clases cross-frontera deben coincidir GED vs API (redactado)
    gg,gt=ged_event_groups(persons,fams,set(exp['people']),set(exp['families']))
    g_cross=[x for x in crossing_shapes(gg,gt) if x[2]]
    a_cross=sorted((x['key_sha256'],x['members']) for x in amb['events_crossing'])
    if [(hsh,n) for hsh,n,_ in g_cross]!=[(hsh,n) for hsh,n in a_cross]:
        fails.append('ambigüedad cross-frontera GED vs API mismatch: GED %s API %s'%(g_cross,a_cross))
    # contenido completo o aborto (gate BO): campos de evento fuera de la clave deben estar vacios
    # tanto en seleccionados como en cualquier miembro de grupo multi
    for h in list(ev_t):
        e=ev_by_handle.get(h)
        if e and (e.get('description') or e.get('attribute_list') or e.get('tag_list')):
            fails.append('evento seleccionado con contenido no cubierto por la clave (description/atributos/tags): '+e.get('gramps_id','?'))
    for ck,hs in ev_groups.items():
        if len(hs)<2: continue
        for h in hs:
            e=ev_by_handle[h]
            if e.get('description') or e.get('attribute_list') or e.get('tag_list'):
                fails.append('evento en grupo multi con contenido no cubierto (description/atributos/tags) (sha clave): '+hashlib.sha256(repr(ck).encode()).hexdigest()[:16]); break
    if fails:
        print(json.dumps({'status':'FAIL_CLOSED','fails':fails},ensure_ascii=False)); sys.exit(1)
    src_by_handle={s['handle']:s for s in api_srcs}
    def obj_sha(o):
        return hashlib.sha256(json.dumps({k:v for k,v in o.items() if k not in ('change','private')},sort_keys=True).encode()).hexdigest()
    def full_get(kind,h):
        st,o=http('GET','%s/api/%s/%s'%(base,kind,h),token=tok)
        if st!=200 or not isinstance(o,dict): raise SystemExit('ERROR GET obj %s/%s: %s'%(kind,h,st))
        return o
    # expectativa GED por cita: validar page/título API contra GED (fail-closed) y emitir fingerprint GED
    cit_ged=exp.get('cit_ged',{})
    targets=[]
    for p in api_resn: targets.append({'type':'people','gramps_id':p['gramps_id'],'handle':p['handle']})
    for h in ev_t:
        e=next(x for x in api_events if x['handle']==h); targets.append({'type':'events','gramps_id':e['gramps_id'],'handle':h,'origin':'conservative_ambiguity' if h in ev_added else 'ged_resn'})
    for f in fam_t: targets.append({'type':'families','gramps_id':f['gramps_id'],'handle':f['handle']})
    for h in note_handles:
        n=next(x for x in api_notes if x['handle']==h); targets.append({'type':'notes','gramps_id':n.get('gramps_id',''),'handle':h,'origin':'conservative_ambiguity' if h in note_added else 'ged_resn'})
    for c in cit_t:
        stit=(src_by_handle.get(c.get('source_handle')) or {}).get('title','')
        m=re.search(r'\| (C\d+) \|',c.get('page','') or '')
        lbl=m.group(1) if m else None
        g=cit_ged.get(lbl)
        if not (lbl and g):
            fails.append('cita sin etiqueta C mapeable al GED: '+c.get('handle','?')); continue
        gpage,gsl=g
        gtit=(sources_ged.get(gsl) or {}).get('title','')
        if (c.get('page','') or '').strip()!=gpage: fails.append('cita %s: page API != GED'%lbl)
        if stit!=gtit: fails.append('cita %s: título fuente API != GED'%lbl)
        targets.append({'type':'citations','gramps_id':c['gramps_id'],'handle':c['handle'],'fingerprint':gpage+'|'+gtit})
    if fails:
        print(json.dumps({'status':'FAIL_CLOSED','fails':fails},ensure_ascii=False)); sys.exit(1)
    for s2 in src_t: targets.append({'type':'sources','gramps_id':s2['gramps_id'],'handle':s2['handle'],'fingerprint':s2.get('title','') or ''})
    for t in targets:
        t['content_sha256']=obj_sha(full_get(t['type'],t['handle']))
    by={}
    for t in targets: by.setdefault(t['type'],[]).append(t['handle'])
    sha={t:hashlib.sha256('\n'.join(sorted(v)).encode()).hexdigest() for t,v in by.items()}
    counts={t:len(v) for t,v in by.items()}
    print(json.dumps({'status':'OK','counts':counts,'sha256':sha,'total':len(targets),'ambiguity':amb,'targets':targets},ensure_ascii=False))
main()
