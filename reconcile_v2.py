#!/usr/bin/env python3
"""Reconciliacion V2: Excel (01, 02, 05, 06, 07 + overlay + orientacion validada) vs GEDCOM. NO importa mapper.py ni mapper_v2.py.
Comparte con el mapper solo la ESPECIFICACION de seleccion (v2_selector.py). Verifica conversion, NO la verdad genealogica: la orientacion hijo/progenitor sale del CSV validado y se toma como dada.
Uso: python3 reconcile_v2.py entrada.xlsx salida.ged [overlay.csv|-] [orientacion.csv]
Cubre: conjunto de personas, nombre, sexo, fechas y lugares, progenitores (solo los del CSV validado, sin extras), pares Conyuge/Esposo/Esposa, MARR si hay una sola fila con fecha, citas a persona (ID y fuente), notas de documentos 06.
NO cubre: citas por dato, QUAY, medios/hash, campos omitidos por privacidad, ni que el CSV sea cierto."""
import sys, re, csv, datetime, openpyxl
import reconcile, v2_selector
from reconcile import parse, fecha_esperada, s, MES

SPOUSE = ('Cónyuge', 'Esposo', 'Esposa')


def tab(wb, n, key):
    ws = list(wb[n].iter_rows(values_only=True))
    i = next(k for k, r in enumerate(ws) if r and r[0] == key)
    hd = [str(x) for x in ws[i]]
    return [dict(zip(hd, r)) for r in ws[i + 1:] if r and r[0]]


def fx(v, prec):
    """fecha Excel + precision -> texto GEDCOM esperado o None (sin fecha GEDCOM)."""
    if v is None or s(v) == '':
        return None
    t = v.strftime('%Y-%m-%d') if isinstance(v, (datetime.datetime, datetime.date)) else s(v)
    p = s(prec)
    if p.startswith('Exacta'):
        return fecha_esperada(t)
    if p.startswith('Aproximada') and re.fullmatch(r'\d{4}', t):
        return 'ABT ' + t
    return None


def main(xlsx, ged, overlay=None, orient=None, rel_overlay=None):
    res = v2_selector.main(xlsx, overlay, quiet=True, rel_overlay=rel_overlay)
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    L = {r['ID_Lugar']: ', '.join(s(r[k]) for k in ('Nombre_canonico', 'Provincia_estado', 'Pais') if s(r[k])) for r in tab(wb, '07_Lugares', 'ID_Lugar')}
    ids, rels = set(res['inc']), res['rel_ok']
    P = {p['ID_Persona']: p for p in res['P'] if p['ID_Persona'] in ids}
    ori, n_csv = {}, 0
    if orient:   # lector PROPIO, posicional por cabecera: no comparte codigo con mapper_v2
        with open(orient, encoding='utf-8', newline='') as fh:
            filas = list(csv.reader(fh))
        hd = [h.strip() for h in filas[0]] if filas else []
        if not all(c in hd for c in ('ID_Relacion', 'hijo', 'progenitor')):
            raise SystemExit(f'ERROR: el CSV de orientacion no tiene las cabeceras ID_Relacion, hijo, progenitor (leidas: {hd})')
        i0, i1, i2 = hd.index('ID_Relacion'), hd.index('hijo'), hd.index('progenitor')
        for r in filas[1:]:
            if len(r) > max(i0, i1, i2) and r[i1].strip() and r[i2].strip():
                ori[r[i0].strip()] = (r[i1].strip(), r[i2].strip()); n_csv += 1
    ind, fam = parse(ged)
    diffs = []
    for x in ids - set(ind): diffs.append(('persona_falta_en_GED', x, ''))
    for x in set(ind) - ids: diffs.append(('persona_sobra_en_GED', x, ''))
    esp_par = {}
    for r in rels:
        t = s(r['Tipo_relacion'])
        if t in ('Padre biológico', 'Madre biológica') and r['ID_Relacion'] in ori:
            h, pr = ori[r['ID_Relacion']]
            esp_par.setdefault(h, {})['HUSB' if t.startswith('Padre') else 'WIFE'] = pr
    for i, p in P.items():
        g = ind.get(i)
        if not g: continue
        nombre = ' '.join(x for x in (s(p['Nombre']), s(p['Apellidos_nacimiento'])) if x)
        gn = re.sub(r'\s+', ' ', g.get('NAME', '').replace('/', ' ')).strip()
        if re.sub(r'\W', '', nombre.lower()) != re.sub(r'\W', '', gn.lower()): diffs.append(('nombre', i, f'{nombre!r} vs {gn!r}'))
        if s(p['Sexo']) and s(p['Sexo'])[0].upper() != g.get('SEX', ''): diffs.append(('sexo', i, f"{s(p['Sexo'])} vs {g.get('SEX')}"))
        for lc, key in (('ID_Lugar_nacimiento', 'BIRTPLAC'), ('ID_Lugar_defuncion', 'DEATPLAC')):
            e = L.get(s(p[lc]), '') if s(p[lc]) else ''
            if e != g.get(key, ''): diffs.append((lc, i, f'{e!r} vs {g.get(key, "")!r}'))
        for fc, pc, key in (('Fecha_nacimiento', 'Precision_nacimiento', 'BIRTDATE'), ('Fecha_defuncion', 'Precision_defuncion', 'DEATDATE')):
            e = fx(p[fc], p[pc]) or ''
            if e != g.get(key, ''): diffs.append((fc, i, f'esperada {e!r} vs GED {g.get(key, "")!r}'))
        fc = [fam[f] for f in g['FAMC'] if f in fam]
        for k in ('HUSB', 'WIFE'):
            ev = esp_par.get(i, {}).get(k)
            gv = {f.get(k) for f in fc if f.get(k)}
            if ev and ev not in gv: diffs.append(('progenitor_' + k, i, f'CSV {ev} vs GED {sorted(gv)}'))
            if not ev and gv: diffs.append(('progenitor_extra_' + k, i, f'GED {sorted(gv)} sin fila validada'))
    por_par = {}
    for r in rels:
        if s(r['Tipo_relacion']) in SPOUSE:
            por_par.setdefault(frozenset((r['ID_Persona_1'], r['ID_Persona_2'])), []).append(r)
    for par, rs in por_par.items():
        a, b = sorted(par)
        fl = [f for f in fam.values() if {f.get('HUSB'), f.get('WIFE')} == {a, b}]
        if len(fl) != 1: diffs.append(('pareja_familia', f'{a},{b}', f'{len(fl)} familias en GED')); continue
        fechas = {(fx(r['Fecha_desde'], 'Exacta') or '') for r in rs if r['Fecha_desde']}
        if len(fechas) == 1 and len(rs) >= 1 and next(iter(fechas)):
            if fl[0].get('MARRDATE', '') != next(iter(fechas)): diffs.append(('MARR_fecha', f'{a},{b}', f"{next(iter(fechas))!r} vs {fl[0].get('MARRDATE', '')!r}"))
    # hijos del GED sin respaldo validado
    for fid, f in fam.items():
        for c in f['CHIL']:
            for k in ('HUSB', 'WIFE'):
                if f.get(k) and esp_par.get(c, {}).get(k) != f[k]: diffs.append(('hijo_sin_fila_validada', c, f'{k}={f[k]}'))
    C = tab(wb, '05_Citas', 'ID_Cita')
    esperadas = {}
    for c in C:
        if s(c['ID_Persona']) in ids and s(c['ID_Fuente']): esperadas.setdefault(s(c['ID_Persona']), []).append(c)
    ncit = 0
    for i, cs in esperadas.items():
        g = ind.get(i)
        if not g: continue
        for c in cs:
            cid = s(c['ID_Cita'])
            hit = [x for x in g['cit'] if re.search(r'\b' + re.escape(cid) + r'\b', x.get('page', ''))]
            if not hit: diffs.append(('cita_no_emitida', cid, f'persona {i}')); continue
            ncit += 1
            if hit[0]['sid'] != s(c['ID_Fuente']): diffs.append(('cita_fuente', cid, f"{s(c['ID_Fuente'])} vs {hit[0]['sid']}"))
    for i, g in ind.items():
        for x in g['cit']:
            if not re.search(r'C\d{6}', x.get('page', '')): diffs.append(('cita_GED_sin_ID_en_PAGE', i, x.get('sid', '')))
    ndoc = 0
    for x in tab(wb, '06_Documentos', 'ID_Documento'):
        pid = s(x.get('ID_Persona_principal'))
        if pid not in ids or pid not in ind: continue
        ndoc += 1; did = s(x['ID_Documento']); g = ind[pid]
        notas = [n for n in g.get('_notes', []) if n.startswith(f'Documento {did}:')]
        if len(notas) != 1: diffs.append(('documento_nota_ausente_o_duplicada', did, f'persona {pid}: {len(notas)}')); continue
        n = notas[0]
        sha = s(x.get('SHA256')); valido = bool(re.fullmatch(r'[0-9a-fA-F]{64}', sha))
        if valido and sha not in n: diffs.append(('documento_sha_valido_no_reflejado', did, ''))
        if not valido and 'no acreditado' not in n: diffs.append(('documento_sha_invalido_sin_marca', did, sha[:12]))
        pub = s(x.get('Derechos_privacidad')).lower().startswith(('public', 'públic'))
        if (not pub or g.get('RESN')) and 'omitido por privacidad (ver Excel)' not in n: diffs.append(('documento_sin_marcador_de_privacidad', did, pid))
        if s(x.get('Nombre_archivo')) and s(x['Nombre_archivo']) in n and (not pub or g.get('RESN')): diffs.append(('documento_nombre_archivo_filtrado', did, pid))
    usadas = sum(len(v) for v in esp_par.values())
    if orient and n_csv and usadas == 0:
        diffs.append(('csv_con_filas_validadas_pero_0_usadas', orient, f'{n_csv} filas validadas, ninguna aplicable a relaciones del selector'))
    fuera_sel = sorted(set(ori) - {r['ID_Relacion'] for r in rels})
    print(f'filas validadas en el CSV (lectura propia): {n_csv}; dentro del selector: {n_csv - len(fuera_sel)}; fuera del selector: {len(fuera_sel)} {fuera_sel[:20]}')
    print(f'personas esperadas (selector): {len(ids)}; en GED: {len(ind)}; filiaciones validadas en CSV usadas: {sum(len(v) for v in esp_par.values())}; parejas esperadas: {len(por_par)}')
    print(f'citas a persona esperadas: {sum(len(v) for v in esperadas.values())}; verificadas: {ncit}; documentos 06 de esas personas: {ndoc}')
    print(f'diferencias: {len(diffs)}')
    for d in diffs: print('  ', *d)
    return 1 if diffs else 0


if __name__ == '__main__':
    a = [x for x in sys.argv if not x.startswith('--overlay-relaciones=')]
    ro = next((x.split('=', 1)[1] for x in sys.argv if x.startswith('--overlay-relaciones=')), None)
    sys.exit(main(a[1], a[2], a[3] if len(a) > 3 and a[3] != '-' else None, a[4] if len(a) > 4 else None, rel_overlay=ro))
