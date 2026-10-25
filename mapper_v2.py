#!/usr/bin/env python3
"""Mapper V2 (nucleo): personas y relaciones elegidas por v2_selector (01+02), salida GEDCOM por el mismo nucleo y reglas de privacidad que mapper.py.
Uso: python3 mapper_v2.py master.xlsx salida.ged [overlay.csv|-] [orientacion.csv]
Citas a PERSONA (no por dato). Solo Documentado segun el selector; Probable/Hipotesis/Descartado no entran. Sin datos reales fuera de la maquina."""
import sys, re, datetime, collections, openpyxl
import mapper, v2_selector
from mapper import nm

PARENT_TAG = {'Padre biológico': 'id_padre', 'Madre biológica': 'id_madre'}
SPOUSE = ('Cónyuge', 'Esposo', 'Esposa')


def fecha(v, prec):
    if v is None or nm(v) == '':
        return ''
    t = v.strftime('%Y-%m-%d') if isinstance(v, (datetime.datetime, datetime.date)) else nm(v)
    pr = nm(prec)
    if pr.startswith('Exacta'):
        return t
    if pr.startswith('Aproximada'):
        return 'c. ' + t if re.fullmatch(r'\d{4}', t) else '[aprox] ' + t  # no inventar rango: queda como nota
    return f'[{pr[:40] or "sin precision"}] {t}'  # precision desconocida u otra: nota, no fecha GEDCOM


def lugares(wb):
    ws = list(wb['07_Lugares'].iter_rows(values_only=True))
    k = next(i for i, r in enumerate(ws) if r and r[0] == 'ID_Lugar')
    out = {}
    for r in ws[k + 1:]:
        if r and r[0]:
            out[r[0]] = ', '.join(nm(x) for x in (r[1], r[4], r[5]) if nm(x))
    return out


ORI_COLS = ('ID_Relacion', 'hijo', 'progenitor')


def cargar_orientacion(path, rel_ids=None):
    """CSV validado por Genealogia. Cabeceras OBLIGATORIAS: ID_Relacion, hijo, progenitor (otras columnas se ignoran; cualquier otro nombre, p. ej. ID_Persona_1_hijo, es error).
    Fila con hijo y progenitor vacios = no validada. Fila con uno solo, IDs mal formados, relacion inexistente en 02, tipo que no es Padre/Madre o duplicada con valores distintos = error.
    Devuelve (orientacion, estadisticas). Sin fila validada la arista no se importa (fail-closed). El sexo NO decide."""
    import csv
    with open(path, encoding='utf-8', newline='') as fh:
        rows = list(csv.reader(fh))
    if not rows:
        raise ValueError(f'{path}: CSV de orientacion vacio')
    hd = [h.strip() for h in rows[0]]
    falta = [c for c in ORI_COLS if c not in hd]
    if falta:
        raise ValueError(f'{path}: faltan cabeceras obligatorias {falta}; cabeceras leidas: {hd}')
    ix = {c: hd.index(c) for c in ORI_COLS}
    out, st = {}, {'filas': 0, 'validadas': 0, 'sin_validar': 0}
    for n, r in enumerate(rows[1:], 2):
        if not any(x.strip() for x in r):
            continue
        st['filas'] += 1
        g = lambda c: (r[ix[c]].strip() if ix[c] < len(r) else '')
        rid, h, p = g('ID_Relacion'), g('hijo'), g('progenitor')
        if not re.fullmatch(r'R\d{6}', rid):
            raise ValueError(f'{path} linea {n}: ID_Relacion no valido {rid!r}')
        if not h and not p:
            st['sin_validar'] += 1; continue
        if not (re.fullmatch(r'P\d{6}', h) and re.fullmatch(r'P\d{6}', p)) or h == p:
            raise ValueError(f'{path} linea {n} ({rid}): hijo/progenitor no validos {h!r}/{p!r}')
        if rel_ids is not None and rid not in rel_ids:
            raise ValueError(f'{path} linea {n}: {rid} no existe en 02_Relaciones como Padre/Madre biologico')
        if rid in out and out[rid] != (h, p):
            raise ValueError(f'{path} linea {n}: {rid} duplicada con valores distintos')
        out[rid] = (h, p); st['validadas'] = len(out)
    return out, st


def orientar(r, ori):
    a, b = r['ID_Persona_1'], r['ID_Persona_2']
    v = ori.get(r['ID_Relacion'])
    if not v:
        return None, None, 'indeterminada'
    if set(v) != {a, b}:
        raise ValueError(f"{r['ID_Relacion']}: la orientacion validada {v} no coincide con los extremos {a},{b}")
    return v[0], v[1], 'validada'


def orientacion_ids(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    return {r['ID_Relacion'] for r in tab02(wb) if nm(r['Tipo_relacion']) in PARENT_TAG}


EV_COLS = ('R-ID', 'hijo_P1', 'tipo_relacion', 'progenitor_P2', 'C-ID', 'D-ID', 'tipo_documento', 'alcance_validacion', 'literal_leido_exacto', 'incertidumbre_salvedades')


def cargar_evidencia(path, ori, rel_tipos):
    """CSV canonico de evidencia (Genealogia). Cabeceras obligatorias EV_COLS; opcional fila_origen_literal (R cuyo literal vale para esta fila cuando el literal es 'idem').
    Fail closed: R fuera de la orientacion validada, hijo/progenitor/tipo distintos, duplicados, literal o alcance vacios, 'idem' sin origen valido (mismo hijo y D, origen no idem).
    Devuelve {hijo: [fila...]}. Las C-ID son CANDIDATAS (metadatos de 02), nunca cita-prueba."""
    import csv
    with open(path, encoding='utf-8', newline='') as fh:
        rd = csv.DictReader(fh)
        miss = [c for c in EV_COLS if c not in (rd.fieldnames or [])]
        if miss:
            raise mapper.MapperError(f'CSV de evidencia: faltan columnas {miss}')
        rows = list(rd)
    by = {}
    for x in rows:
        r = nm(x['R-ID'])
        if r in by:
            raise mapper.MapperError(f'CSV de evidencia: {r} duplicada')
        by[r] = x
    out = collections.defaultdict(list)
    for r, x in by.items():
        if r not in ori:
            raise mapper.MapperError(f'CSV de evidencia: {r} no esta en la orientacion validada')
        h, p = nm(x['hijo_P1']), nm(x['progenitor_P2'])
        if ori[r] != (h, p):
            raise mapper.MapperError(f'CSV de evidencia: {r} hijo/progenitor {h},{p} distinto de la orientacion {ori[r]}')
        if rel_tipos.get(r) != nm(x['tipo_relacion']):
            raise mapper.MapperError(f"CSV de evidencia: {r} tipo {nm(x['tipo_relacion'])!r} distinto del 02 ({rel_tipos.get(r)!r})")
        lit, alc = nm(x['literal_leido_exacto']), nm(x['alcance_validacion'])
        if not lit or not alc:
            raise mapper.MapperError(f'CSV de evidencia: {r} literal o alcance vacio')
        origen = ''; propio = ''
        if lit.lower().startswith('idem'):
            origen = nm(x.get('fila_origen_literal'))
            o = by.get(origen)
            if not o or nm(o['hijo_P1']) != h or nm(o['D-ID']) != nm(x['D-ID']) or nm(o['literal_leido_exacto']).lower().startswith('idem'):
                raise mapper.MapperError(f"CSV de evidencia: {r} literal 'idem' sin fila_origen_literal valida (mismo hijo y D, origen no idem)")
            lit = nm(o['literal_leido_exacto'])
            if not re.fullmatch(r'idem[\s;:,.]*', nm(x['literal_leido_exacto']), re.I):
                propio = nm(x['literal_leido_exacto'])      # la fila trae texto propio ademas de 'idem': se conserva, no se pierde
        out[h].append({'rel': r, 'hijo': h, 'progenitor': p, 'tipo': nm(x['tipo_relacion']), 'doc': nm(x['D-ID']), 'tipo_doc': nm(x['tipo_documento']), 'alcance': alc,
                       'literal': lit, 'origen': origen, 'propio': propio, 'salvedades': nm(x['incertidumbre_salvedades']), 'citas': nm(x['C-ID'])})
    return dict(out)


def tab02(wb):
    ws = list(wb['02_Relaciones'].iter_rows(values_only=True))
    i = next(k for k, r in enumerate(ws) if r and r[0] == 'ID_Relacion')
    hd = [str(x) for x in ws[i]]
    return [dict(zip(hd, r)) for r in ws[i + 1:] if r and r[0]]


def build(path, overlay=None, orientacion=None, rel_overlay=None, evidencia=None):
    ori, ori_st = (cargar_orientacion(orientacion, orientacion_ids(path)) if orientacion else ({}, {'filas': 0, 'validadas': 0, 'sin_validar': 0}))
    res = v2_selector.main(path, overlay, quiet=True, rel_overlay=rel_overlay)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    L = lugares(wb)
    inc, rel_ok = res['inc'], res['rel_ok']
    conf = []
    dirs = []
    norm = []
    sexo = {p['ID_Persona']: nm(p['Sexo']) for p in res['P']}
    padres = collections.defaultdict(lambda: collections.defaultdict(set))
    conyuges = collections.defaultdict(list)
    asoc = collections.defaultdict(list)
    for r in rel_ok:
        a, b, t = r['ID_Persona_1'], r['ID_Persona_2'], nm(r['Tipo_relacion'])
        if t in PARENT_TAG:
            hijo, prog, modo = orientar(r, ori)
            if modo == 'indeterminada':
                dirs.append(f"{r['ID_Relacion']} ({a},{t},{b})"); continue
            padres[hijo][PARENT_TAG[t]].add(prog)
            norm.append(dict(r, ID_Persona_1=hijo, ID_Persona_2=prog)); continue
        elif t in SPOUSE:
            fd = nm(fecha(r['Fecha_desde'], 'Exacta')) if r['Fecha_desde'] else ''
            lg = L.get(r['ID_Lugar'], '') if r['ID_Lugar'] else ''
            conyuges[a].append((b, fd, lg)); conyuges[b].append((a, fd, lg))
        else:
            asoc[a].append(f"{t} de {b} ({r['ID_Relacion']})"); asoc[b].append(f"{t} de {a} ({r['ID_Relacion']}, asociacion literal, no genera familia)")
    P = []
    P = []
    for p in res['P']:
        i = p['ID_Persona']
        if i not in inc:
            continue
        d = {'id_persona': i, 'nombre': nm(p['Nombre']), 'apellido1': nm(p['Apellidos_nacimiento']), 'apellido2': '', 'sexo': nm(p['Sexo']),
             'fecha_nacimiento': fecha(p['Fecha_nacimiento'], p['Precision_nacimiento']), 'lugar_nacimiento': L.get(p['ID_Lugar_nacimiento'], '') if p['ID_Lugar_nacimiento'] else '',
             'fecha_defuncion': fecha(p['Fecha_defuncion'], p['Precision_defuncion']), 'lugar_defuncion': L.get(p['ID_Lugar_defuncion'], '') if p['ID_Lugar_defuncion'] else '',
             'id_padre': '', 'id_madre': '', 'id_conyuge': '', 'fecha_matrimonio': '', 'lugar_matrimonio': '', 'fuente': '', 'conyuges': []}
        for tag in ('id_padre', 'id_madre'):
            v = padres[i].get(tag, set())
            if len(v) > 1:
                conf.append(f'{i}: varios {tag} en 02 {sorted(v)}; no se asigna')
            elif v:
                d[tag] = next(iter(v))
        seen = set()
        for c in conyuges[i]:
            if c not in seen and c[0] in inc:
                seen.add(c); d['conyuges'].append(c)
        nota = nm(p['Notas'])
        if asoc[i]:
            nota = (nota + ' | ' if nota else '') + 'Asociaciones literales en Excel: ' + '; '.join(asoc[i])
        d['notas'] = nota
        P.append(d)
    res['evidencia'] = None
    if evidencia:
        tipos = {r['ID_Relacion']: nm(r['Tipo_relacion']) for r in rel_ok if nm(r['Tipo_relacion']) in PARENT_TAG}
        res['evidencia'] = cargar_evidencia(evidencia, ori, tipos)
    res['direccion_indeterminada'] = dirs
    res['orientacion_csv'] = ori_st
    rel_norm = norm + [r for r in rel_ok if nm(r['Tipo_relacion']) not in PARENT_TAG]
    return P, rel_norm, res, conf


def main(argv):
    ro = next((x.split('=', 1)[1] for x in argv if x.startswith('--overlay-relaciones=')), None)
    argv = [x for x in argv if not x.startswith('--overlay-relaciones=')]
    ev = next((x.split('=', 1)[1] for x in argv if x.startswith('--evidencia=')), None)
    argv = [x for x in argv if not x.startswith('--evidencia=')]
    if len(argv) < 3:
        print('uso: python3 mapper_v2.py master.xlsx salida.ged [overlay.csv|-] [orientacion.csv]'); return 2
    P, R, res, conf = build(argv[1], argv[3] if len(argv) > 3 and argv[3] != '-' else None, argv[4] if len(argv) > 4 else None, ro, ev)
    st = res['orientacion_csv']
    print(f"CSV de orientacion: {st['filas']} filas, {st['validadas']} validadas, {st['sin_validar']} sin validar")
    print(f"aristas filiacion sin orientacion validada (no importadas): {len(res['direccion_indeterminada'])}")
    if conf:
        print('CONFLICTOS (no se asignan, a corregir en el Excel):'); [print('  ', c) for c in conf]
    try:
        summary, warn, rep, unsup = mapper.run(argv[1], argv[2], persons=P, rels=R, evidencia=res['evidencia'])
    except mapper.MapperError as ex:
        print('ERROR (fail closed):', ex); return 1
    for k, v in summary.items():
        print(f'{k}: {v}')
    print('advertencias:', len(warn)); [print('  ', w) for w in warn[:400]]
    print('cruce con 02:'); [print('  ', r) for r in rep] if rep else print('   sin discrepancias')
    print('familias con union no acreditada:', unsup)
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
