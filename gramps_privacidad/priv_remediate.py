#!/usr/bin/env python3
"""Remediación privacidad Gramps Web API v3 (fail-closed, BO 4 bloques + manifiesto).
stdin JSON: {"base_url","sha256":{tipo:sha256_handles_ordenados},"targets":[{"type","gramps_id","handle","fingerprint"}]}
 - cfg['sha256'] es OBLIGATORIO (lo emite priv_derive.py): se recomputa por tipo sobre los
   handles de entrada y cualquier diferencia aborta ANTES de cualquier escritura.
 - 'handle' OBLIGATORIO por target y debe coincidir con el lookup por gramps_id (len==1).
 - fingerprint (obligatorio en citations/sources): citations = page+'|'+source_title; sources = title.
   Se verifica ANTES de PUT; ID solo no basta.
 - Cada objeto: content_sha256 obligatorio (JSON menos 'change'/'private') verificado contra el
   objeto vivo ANTES de ALREADY/PUT (congela el contenido derivado); PUT roundtrip, readback:
   private==True y resto de campos idénticos en comparación estructural JSON (excl. change/private).
 - Detalle por objeto JSONL en stderr (type/gramps_id/handle/estado/preimage_sha).
 - exit 1 en la PRIMERA anomalía (FIELDS_CHANGED es anomalía, nunca éxito).
Credenciales SOLO por entorno GRAMPS_USER/GRAMPS_PASS. Nunca argv, nunca impresas.
"""
import json, sys, time, hashlib, urllib.request, urllib.error, os

def log(d): print(json.dumps(d,ensure_ascii=False),file=sys.stderr)
def anomaly(reason,**kw):
    log(dict({'state':'ANOMALY','reason':reason},**kw)); sys.exit(1)

_AUTH={}
def http(method,url,payload=None,token=None):
    if token is not None and _AUTH: token=_AUTH.get('tok')
    data=json.dumps(payload).encode() if payload is not None else None
    def build():
        req=urllib.request.Request(url,data=data,method=method)
        if data: req.add_header('Content-Type','application/json')
        if token: req.add_header('Authorization','Bearer '+token)
        return req
    for _ in range(3):
        try:
            with urllib.request.urlopen(build(),timeout=60) as r: return r.status,json.loads(r.read())
        except urllib.error.HTTPError as e:
            if e.code==401 and token and _AUTH:
                _AUTH['tok']=get_token(_AUTH['base'],_AUTH['u'],_AUTH['p']); token=_AUTH['tok']; continue
            try: return e.code,json.loads(e.read())
            except Exception: return e.code,None
    return 599,None

def get_token(base,u,p):
    for _ in range(8):
        st,res=http('POST',base+'/api/token/',{'username':u,'password':p})
        if isinstance(res,dict) and 'access_token' in res: return res['access_token']
        time.sleep(4)
    raise SystemExit('ERROR: token agotado (rate limit)')

def obj_sha(o):
    return hashlib.sha256(json.dumps({k:v for k,v in o.items() if k not in ('change','private')},sort_keys=True).encode()).hexdigest()

def fingerprint(base,kind,obj,tok):
    if kind=='citations':
        sh_=obj.get('source_handle')
        if not sh_: anomaly('citation sin source_handle',handle=obj.get('handle'))
        st,s=http('GET','%s/api/sources/%s'%(base,sh_),token=tok)
        if st!=200 or not isinstance(s,dict): anomaly('source de cita ilegible',handle=obj.get('handle'))
        return (obj.get('page','') or '')+'|'+(s.get('title','') or '')
    if kind=='sources':
        return obj.get('title','') or ''
    return None

def main():
    cfg=json.load(sys.stdin)
    base=cfg['base_url'].rstrip('/')
    # Integridad del conjunto ANTES de cualquier escritura:
    if 'sha256' not in cfg: anomaly('manifest_sha256_ausente')
    by={}
    seen=set()
    for t in cfg['targets']:
        if 'handle' not in t: anomaly('handle_ausente',kind=t.get('type'),gramps_id=t.get('gramps_id'))
        key=(t['type'],t['handle'])
        if key in seen: anomaly('target_duplicado',kind=t['type'],handle=t['handle'])
        seen.add(key)
        if 'content_sha256' not in t: anomaly('content_sha256_ausente',kind=t.get('type'),gramps_id=t.get('gramps_id'))
        by.setdefault(t['type'],[]).append(t['handle'])
    calc={k:hashlib.sha256('\n'.join(sorted(v)).encode()).hexdigest() for k,v in by.items()}
    if calc!=cfg['sha256']:
        anomaly('manifest_sha256_mismatch',tipos_calc=sorted(calc),tipos_cfg=sorted(cfg['sha256']))
    tok=get_token(base,os.environ['GRAMPS_USER'],os.environ['GRAMPS_PASS'])
    _AUTH.update(base=base,u=os.environ['GRAMPS_USER'],p=os.environ['GRAMPS_PASS'],tok=tok)
    flagged=0; already=0
    # FASE 1: validar TODO (lookup, handle, contenido, fingerprint) antes de cualquier escritura.
    pending=[]
    for t in cfg['targets']:
        kind,gid=t['type'],t['gramps_id']
        st,lst=http('GET','%s/api/%s/?gramps_id=%s'%(base,kind,gid),token=tok)
        if st!=200 or not isinstance(lst,list): anomaly('lookup_http',kind=kind,gramps_id=gid,status=st)
        if len(lst)!=1: anomaly('lookup_len!=1',kind=kind,gramps_id=gid,n=len(lst))
        if lst[0].get('gramps_id')!=gid: anomaly('lookup_id_mismatch',kind=kind,gramps_id=gid)
        h=lst[0]['handle']
        if t['handle']!=h: anomaly('handle_mismatch',kind=kind,gramps_id=gid)
        st,obj=http('GET','%s/api/%s/%s'%(base,kind,h),token=tok)
        if st!=200 or not isinstance(obj,dict): anomaly('get_obj',kind=kind,gramps_id=gid,status=st)
        pre_sha=obj_sha(obj)
        if pre_sha!=t['content_sha256']: anomaly('content_mismatch',kind=kind,gramps_id=gid,handle=h)
        if kind in ('citations','sources'):
            fp=fingerprint(base,kind,obj,tok)
            if 'fingerprint' not in t: anomaly('fingerprint_requerido_ausente',kind=kind,gramps_id=gid)
            if fp!=t['fingerprint']: anomaly('fingerprint_mismatch',kind=kind,gramps_id=gid)
        if obj.get('private') is True:
            already+=1; log({'type':kind,'gramps_id':gid,'handle':h,'state':'ALREADY','preimage_sha256':pre_sha}); continue
        pending.append((t,h,obj,pre_sha))
    # FASE 2: todo validado; ahora escribir. Revalida contenido justo antes de cada PUT (TOCTOU).
    for t,h,obj,pre_sha in pending:
        kind,gid=t['type'],t['gramps_id']
        st,cur=http('GET','%s/api/%s/%s'%(base,kind,h),token=tok)
        if st!=200 or not isinstance(cur,dict): anomaly('get_obj_revalida',kind=kind,gramps_id=gid,status=st)
        if obj_sha(cur)!=pre_sha: anomaly('content_drift',kind=kind,gramps_id=gid,handle=h)
        before={k:v for k,v in cur.items() if k not in ('change','private')}
        cur['private']=True
        st,res=http('PUT','%s/api/%s/%s'%(base,kind,h),payload=cur,token=tok)
        if st!=200: anomaly('put_fail',kind=kind,gramps_id=gid,status=st)
        st,rb=http('GET','%s/api/%s/%s'%(base,kind,h),token=tok)
        if st!=200 or not isinstance(rb,dict): anomaly('readback_http',kind=kind,gramps_id=gid,status=st)
        if rb.get('private') is not True: anomaly('readback_private_false',kind=kind,gramps_id=gid)
        after={k:v for k,v in rb.items() if k not in ('change','private')}
        if json.dumps(before,sort_keys=True)!=json.dumps(after,sort_keys=True):
            anomaly('FIELDS_CHANGED',kind=kind,gramps_id=gid,handle=h)
        flagged+=1; log({'type':kind,'gramps_id':gid,'handle':h,'state':'FLAGGED','preimage_sha256':pre_sha})
    print(json.dumps({'status':'OK','flagged':flagged,'already':already,'total':flagged+already},ensure_ascii=False))

if __name__=='__main__':
    main()
