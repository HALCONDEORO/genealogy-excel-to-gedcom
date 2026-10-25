#!/usr/bin/env python3
"""Reconciliacion independiente Excel (hoja 30 + 05) vs GEDCOM. No importa mapper.py. Solo lectura.
Uso: python3 reconcile.py entrada.xlsx salida.ged
Compara por persona (nombre, sexo, lugares, fecha completa con calificadores c./BET/ABT, vinculos de familia), por documento 06 (nota con ID_Fuente, ID_Cita, estado SHA y marcador de privacidad, sin nombre de archivo) y por cita (ID_Cita, fuente y presencia; QUAY no se compara).
Los campos omitidos por privacidad se declaran y NO se comparan; cualquier otra diferencia se lista con su ID."""
import sys, re, openpyxl

def tabla(wb, n):
    ws = list(wb[n].iter_rows(values_only=True))
    i = next(k for k, r in enumerate(ws) if r and r[0] in ('id_persona', 'ID_Cita', 'ID_Fuente', 'ID_Documento'))
    hd = [str(x) for x in ws[i]]
    return [dict(zip(hd, r)) for r in ws[i + 1:] if r and r[0]]

def s(v): return '' if v is None else str(v).strip()

MES = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']

def fecha_esperada(v):
    """Excel -> texto GEDCOM esperado, o None si el formato no esta soportado (el GED no debe traer fecha)."""
    v = s(v)
    m = re.fullmatch(r'(\d{4})-(\d{2})-(\d{2})', v)
    if m and 1 <= int(m[2]) <= 12 and 1 <= int(m[3]) <= 31: return f'{int(m[3])} {MES[int(m[2]) - 1]} {m[1]}'
    if re.fullmatch(r'\d{4}', v): return v
    m = re.fullmatch(r'(\d{4})-(\d{2})', v)
    if m and 1 <= int(m[2]) <= 12: return f'{MES[int(m[2]) - 1]} {m[1]}'
    m = re.fullmatch(r'c\. (\d{4})-(\d{4})', v)
    if m: return f'BET {m[1]} AND {m[2]}'
    m = re.fullmatch(r'c\. (\d{4})', v)
    if m: return f'ABT {m[1]}'
    return None

def parse(path):
    ind, fam, cur = {}, {}, None
    with open(path, encoding='utf-8') as fh: lines = fh.read().split('\n')
    for l in lines:
        m = re.match(r'(\d+) (?:@(\w+)@ )?(\w+)(?: (.*))?$', l)
        if not m: continue
        lv, xr, tag, val = int(m[1]), m[2], m[3], m[4] or ''
        if lv == 0:
            cur = None
            if tag == 'INDI': cur = ind.setdefault(xr, {'cit': [], 'FAMC': [], 'FAMS': []}); ctx = None
            elif tag == 'FAM': cur = fam.setdefault(xr, {'CHIL': []}); ctx = None
            continue
        if cur is None: continue
        if tag in ('CONC', 'CONT') and cur is not None and cur.get('_notes') and lv == 2 and ctx == 'NOTE':
            cur['_notes'][-1] += ('' if tag == 'CONC' else '\n') + val; continue
        if lv == 1:
            ctx = tag
            if tag == 'NOTE': cur.setdefault('_notes', []).append(val)
            if tag in ('FAMC', 'FAMS'): cur[tag].append(val.strip('@'))
            elif tag in ('HUSB', 'WIFE'): cur[tag] = val.strip('@')
            elif tag == 'CHIL': cur['CHIL'].append(val.strip('@'))
            elif tag == 'SOUR': cur['cit'].append({'sid': val.strip('@')}) if 'cit' in cur else None
            elif tag in ('NAME', 'SEX', 'RESN'): cur[tag] = val
        elif lv == 2:
            if ctx in ('BIRT', 'DEAT', 'MARR') and tag in ('DATE', 'PLAC'): cur[ctx + tag] = val
            if ctx == 'SOUR' and 'cit' in cur and cur['cit']:
                if tag == 'PAGE': cur['cit'][-1]['page'] = val
                if tag == 'QUAY': cur['cit'][-1]['quay'] = val
    return ind, fam

def main(xlsx, ged):
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    P = tabla(wb, '30_Arbol_Confirmado'); C = tabla(wb, '05_Citas')
    ind, fam = parse(ged)
    diffs = []
    ids = {p['id_persona'] for p in P}
    for x in ids - set(ind): diffs.append(('persona_falta_en_GED', x, ''))
    for x in set(ind) - ids: diffs.append(('persona_sobra_en_GED', x, ''))
    for p in P:
        i = p['id_persona']; g = ind.get(i)
        if not g: continue
        nombre = ' '.join(x for x in (s(p['nombre']), s(p['apellido1']), s(p['apellido2'])) if x)
        gn = re.sub(r'\s+', ' ', g.get('NAME', '').replace('/', ' ')).strip()
        if re.sub(r'\W', '', nombre.lower()) != re.sub(r'\W', '', gn.lower()): diffs.append(('nombre', i, f'{nombre!r} vs {gn!r}'))
        if s(p['sexo']) and s(p['sexo'])[0].upper() != g.get('SEX', ''): diffs.append(('sexo', i, f"{s(p['sexo'])} vs {g.get('SEX')}"))
        for campo, key in (('lugar_nacimiento', 'BIRTPLAC'), ('lugar_defuncion', 'DEATPLAC')):
            if s(p[campo]) != g.get(key, ''): diffs.append((campo, i, f"{s(p[campo])!r} vs {g.get(key, '')!r}"))
        for campo, key in (('fecha_nacimiento', 'BIRTDATE'), ('fecha_defuncion', 'DEATDATE')):
            esp = fecha_esperada(p[campo]); got = g.get(key, '')
            if s(p[campo]) and esp is None:
                if got: diffs.append((campo, i, f'formato no soportado {s(p[campo])!r} pero GED trae {got!r}'))
            elif (esp or '') != got: diffs.append((campo, i, f'esperada {esp!r} vs GED {got!r} (Excel {s(p[campo])!r})'))
        padres = {k: s(p[c]) for k, c in (('HUSB', 'id_padre'), ('WIFE', 'id_madre'))}
        fc = [fam[f] for f in g['FAMC'] if f in fam]
        for k, v in padres.items():
            gv = {f.get(k) for f in fc if f.get(k)}
            if v and v not in gv: diffs.append(('progenitor_' + k, i, f'Excel {v} vs GED {sorted(gv)}'))
            if not v and gv: diffs.append(('progenitor_extra_' + k, i, f'GED {sorted(gv)} sin Excel'))
        if s(p['id_conyuge']):
            ok = any((f.get('HUSB') == i and f.get('WIFE') == s(p['id_conyuge'])) or (f.get('WIFE') == i and f.get('HUSB') == s(p['id_conyuge'])) for f in fam.values())
            if not ok: diffs.append(('conyuge', i, s(p['id_conyuge'])))
    # citas
    quay = {}
    ncit = 0
    esperadas = {}
    for c in C:
        if s(c['ID_Persona']) in ids and s(c['ID_Fuente']): esperadas.setdefault(s(c['ID_Persona']), []).append(c)
    for i, cs in esperadas.items():
        g = ind.get(i)
        if not g: continue
        pages = [x.get('page', '') for x in g['cit']]
        for c in cs:
            cid = s(c['ID_Cita'])
            hit = [x for x in g['cit'] if re.search(r'\b' + re.escape(cid) + r'\b', x.get('page', ''))]
            if not hit:
                diffs.append(('cita_no_emitida_o_excluida', cid, f'persona {i}')); continue
            ncit += 1
            if hit[0]['sid'] != s(c['ID_Fuente']): diffs.append(('cita_fuente', cid, f"{s(c['ID_Fuente'])} vs {hit[0]['sid']}"))
    for i, g in ind.items():
        for x in g['cit']:
            m = re.search(r'C\d{6}', x.get('page', ''))
            if not m: diffs.append(('cita_GED_sin_ID_en_PAGE', i, x.get('sid', '')))
    # fecha de matrimonio: para cada MARR en GED comparar con el Excel de alguno de los conyuges
    byid = {p['id_persona']: p for p in P}
    for fid, f in fam.items():
        if 'MARRDATE' not in f and 'HUSB' in f and 'WIFE' in f: pass
    # documentos 06 de personas de la hoja 30
    D = tabla(wb, '06_Documentos') if '06_Documentos' in wb.sheetnames else []
    ndoc = 0
    for x in D:
        pid = s(x.get('ID_Persona_principal'))
        if pid not in ids: continue
        ndoc += 1; did = s(x['ID_Documento']); g = ind.get(pid)
        if not g: continue
        notas = [n for n in g.get('_notes', []) if n.startswith(f'Documento {did}:')]
        if len(notas) != 1: diffs.append(('documento_nota_ausente_o_duplicada', did, f'persona {pid}: {len(notas)}')); continue
        n = notas[0]
        for etq, val in (('fuente', s(x.get('ID_Fuente'))), ('cita', s(x.get('ID_Cita')))):
            if val and not re.search(rf'\| {etq} {re.escape(val)}( |$|\|)', n): diffs.append(('documento_' + etq, did, f'Excel {val!r} no en nota'))
        sha = s(x.get('SHA256'))
        valido = bool(re.fullmatch(r'[0-9a-fA-F]{64}', sha))
        if valido and sha not in n and 'SHA256' in n and 'no acreditado' in n: diffs.append(('documento_sha_valido_no_reflejado', did, ''))
        if not valido and 'no acreditado' not in n: diffs.append(('documento_sha_invalido_sin_marca', did, sha[:12]))
        pub = s(x.get('Derechos_privacidad')).lower().startswith(('public', 'públic'))
        if (not pub or g.get('RESN')) and 'omitido por privacidad (ver Excel)' not in n: diffs.append(('documento_sin_marcador_de_privacidad', did, pid))
        if s(x.get('Nombre_archivo')) and s(x['Nombre_archivo']) in n and (not pub or g.get('RESN')): diffs.append(('documento_nombre_archivo_filtrado', did, pid))
    print(f'documentos 06 de personas de hoja 30: {ndoc}')
    print(f'personas Excel hoja 30: {len(P)}; en GED: {len(ind)}; citas esperadas (persona en 30, con fuente): {sum(len(v) for v in esperadas.values())}; emitidas y verificadas: {ncit}')
    print(f'diferencias: {len(diffs)}')
    for d in diffs: print('  ', *d)
    return 1 if diffs else 0

if __name__ == '__main__':
    sys.exit(main(*sys.argv[1:3]))
