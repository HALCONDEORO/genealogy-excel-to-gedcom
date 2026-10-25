#!/usr/bin/env python3
"""Dry-run del selector V2 (solo informe, no genera GEDCOM, solo lectura). Uso: python3 v2_selector.py master.xlsx
Reglas (project rules): persona incluida si Estado_hechos empieza por 'Documentado' y Nivel in {Alta, Media, Muy alta};
persona con Nivel Baja solo si es extremo de una relacion incluida; Probable y En prueba siempre excluidas aunque la relacion sea Alta;
relacion incluida si Estado empieza por 'Documentado' y Confianza in {Alta, Media, Muy alta} y ambos extremos incluidos.
Hermano/Primo/Cunado/Pareja/Pariente son asociaciones literales, no familia. Valores fuera de esos dominios: fail-closed, listados por ID."""
import sys, re, collections, openpyxl
SHIFT = {'P000198', 'P000205', 'P000209', 'P000210', 'P000211', 'P000212'}  # solo estas 6 son desplazamiento de columna; P040/P117/P118 son validacion por actas
OK_NIVEL = {'Alta', 'Media', 'Muy alta'}
PARENT = {'Padre biológico', 'Madre biológica'}
SPOUSE = {'Cónyuge', 'Esposo', 'Esposa'}

def tabla(wb, n):
    ws = list(wb[n].iter_rows(values_only=True))
    k = next(i for i, r in enumerate(ws) if r and r[0] in ('ID_Persona', 'ID_Relacion'))
    hd = [str(x) for x in ws[k]]
    return [dict(zip(hd, r)) for r in ws[k + 1:] if r and r[0]]

def s(v): return '' if v is None else str(v).strip()

def cargar_overlay(path):
    import csv
    ov = {}
    with open(path, encoding='utf-8') as fh:
      for r in csv.DictReader(fh):
        e = [v for k, v in r.items() if k.startswith('Estado real')][0].split('=', 1)[1].strip()
        n = [v for k, v in r.items() if k.startswith('Nivel real')][0].split('=', 1)[1].strip()
        ov[r['ID']] = (e, n)
    return ov

def aplicar_overlay_relaciones(R, path):
    """Overlay formal de relaciones (Genealogia): rellena SOLO campos vacios de Fecha_desde/ID_Lugar y anade ID_Cita_adicional; un conflicto con un valor existente es error. Cabeceras obligatorias."""
    import csv
    with open(path, encoding='utf-8', newline='') as fh:
        filas = list(csv.DictReader(fh, restkey='__extra__'))
    need = ('ID_Relacion', 'ID_Persona_1', 'ID_Persona_2', 'Fecha_desde', 'Lugar', 'ID_Cita_adicional')
    if not filas:
        raise ValueError(f'{path}: overlay de relaciones vacio')
    falta = [c for c in need if c not in filas[0]]
    if falta:
        raise ValueError(f'{path}: faltan cabeceras {falta}')
    byid = {r['ID_Relacion']: r for r in R}
    traza = []
    for n, f in enumerate(filas, 2):
        if '__extra__' in f or any(v is None for k, v in f.items() if k in need):
            raise ValueError(f'{path} linea {n}: fila con campos desbordados o ausentes')
        r = byid.get(f['ID_Relacion'])
        if not r:
            raise ValueError(f'{path} linea {n}: {f["ID_Relacion"]} no existe en 02')
        if (r['ID_Persona_1'], r['ID_Persona_2']) != (f['ID_Persona_1'], f['ID_Persona_2']):
            raise ValueError(f'{path} linea {n}: extremos del overlay no coinciden con 02 para {f["ID_Relacion"]}')
        for campo, col in (('Fecha_desde', 'Fecha_desde'), ('ID_Lugar', 'Lugar')):
            nuevo = f[col].strip()
            if not nuevo:
                continue
            if s(r[campo]) and s(r[campo]) != nuevo and not (hasattr(r[campo], 'strftime') and r[campo].strftime('%Y-%m-%d') == nuevo):
                raise ValueError(f'{path} linea {n}: {f["ID_Relacion"]} ya tiene {campo}={s(r[campo])!r}, el overlay trae {nuevo!r}')
            if not s(r[campo]):
                r[campo] = nuevo; traza.append(f'{f["ID_Relacion"]}.{campo}={nuevo}')
        adic = f['ID_Cita_adicional'].strip()
        if adic and adic not in s(r['ID_Cita']):
            r['ID_Cita'] = (s(r['ID_Cita']) + '; ' if s(r['ID_Cita']) else '') + adic; traza.append(f'{f["ID_Relacion"]}.ID_Cita+={adic}')
    return traza


def main(path, overlay=None, quiet=False, rel_overlay=None):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    P = tabla(wb, '01_Personas'); R = tabla(wb, '02_Relaciones')
    pid = {p['ID_Persona'] for p in P}
    # remap con rastro (acordado): duplicados/fusionados -> destino; las aristas se remapean ANTES de clasificar extremos
    REMAP = {}
    for p in P:
        m = re.search(r'(?:Fusionad[ao] en|Duplicado de)\s+(P\d{6})', s(p['Estado_hechos']))
        if m and m.group(1) in pid: REMAP[p['ID_Persona']] = m.group(1)
    # resolucion transitiva: P1->P2->P3 acaba en P3; ciclos o autorreferencias = error ruidoso (nunca una arista colgante)
    FINAL = {}
    for o in REMAP:
        seen, x = [o], REMAP[o]
        while x in REMAP:
            if x in seen: raise ValueError(f"remap con ciclo: {' -> '.join(seen + [x])}")
            seen.append(x); x = REMAP[x]
        if x == o: raise ValueError(f'remap autorreferente: {o}')
        FINAL[o] = x
    trace = []
    for r in R:
        for k in ('ID_Persona_1', 'ID_Persona_2'):
            if r[k] in FINAL: trace.append(f"{r['ID_Relacion']}:{r[k]}->{FINAL[r[k]]}"); r[k] = FINAL[r[k]]
    rep = collections.defaultdict(list)
    traza_rel = aplicar_overlay_relaciones(R, rel_overlay) if rel_overlay else []
    if traza_rel: rep['overlay_relaciones_aplicado'] = traza_rel
    pst = {}
    ov = cargar_overlay(overlay) if overlay else {}
    for p in P:
        if p['ID_Persona'] in ov:
            rep['overlay_desplazamiento_columna_aplicado' if p['ID_Persona'] in SHIFT else 'overlay_validacion_actas_aplicado'].append(p['ID_Persona']); p['Estado_hechos'], p['Nivel_documentacion'] = ov[p['ID_Persona']]
    for p in P:
        e, n = s(p['Estado_hechos']), s(p['Nivel_documentacion'])
        if p['ID_Persona'] in REMAP: pst[p['ID_Persona']] = 'remapeada'; rep['persona_remapeada_a_destino'].append(f"{p['ID_Persona']}->{FINAL[p['ID_Persona']]}"); continue
        doc = e.startswith('Documentad')
        if doc and not e.startswith('Documentado'): rep['persona_variante_Documentada_aceptada_a_validar'].append(f"{p['ID_Persona']} ({e[:30]})")
        if e in ('Probable', 'En prueba'): pst[p['ID_Persona']] = 'excluida_estado'; rep['persona_excluida_probable_o_en_prueba'].append(p['ID_Persona'])
        elif not doc: pst[p['ID_Persona']] = 'incoherente'; rep['persona_estado_fuera_de_dominio'].append(f"{p['ID_Persona']} ({e[:30]})")
        elif n in OK_NIVEL: pst[p['ID_Persona']] = 'incluida'
        elif n == 'Baja': pst[p['ID_Persona']] = 'baja'
        else: pst[p['ID_Persona']] = 'incoherente'; rep['persona_nivel_fuera_de_dominio'].append(f"{p['ID_Persona']} ({n[:30]})")
    rel_ok, ext_baja = [], set()
    for r in R:
        e, c = s(r['Estado_hechos']), s(r['Confianza'])
        a, b = r['ID_Persona_1'], r['ID_Persona_2']
        if a not in pid or b not in pid: rep['relacion_con_persona_inexistente'].append(r['ID_Relacion']); continue
        if not e.startswith('Documentad'): rep['relacion_excluida_estado_' + (e or 'vacio')[:20]].append(r['ID_Relacion']); continue
        if c not in OK_NIVEL: rep['relacion_excluida_confianza_' + (c or 'vacia')].append(r['ID_Relacion']); continue
        sa, sb = pst[a], pst[b]
        if 'excluida_estado' in (sa, sb): rep['relacion_Alta_con_extremo_Probable_o_EnPrueba_excluida'].append(f"{r['ID_Relacion']} ({a},{b})"); continue
        if 'incoherente' in (sa, sb): rep['relacion_con_extremo_incoherente_no_importada'].append(f"{r['ID_Relacion']} ({a},{b})"); continue
        rel_ok.append(r)
        for x, sx in ((a, sa), (b, sb)):
            if sx == 'baja': ext_baja.add(x)
    inc = {i for i, v in pst.items() if v == 'incluida'} | ext_baja
    for i, v in pst.items():
        if v == 'baja' and i not in ext_baja: rep['persona_Baja_sin_relacion_incluida_fuera'].append(i)
    # dedupe de aristas identicas (tipo + extremos) con rastro; se conservan los IDs de cita de ambas
    vistos, dup = {}, []
    uniq = []
    for r in rel_ok:
        k = (s(r['Tipo_relacion']), r['ID_Persona_1'], r['ID_Persona_2'])
        if k in vistos: dup.append(f"{r['ID_Relacion']}=>{vistos[k]['ID_Relacion']} (citas {s(r['ID_Cita'])} + {s(vistos[k]['ID_Cita'])})"); continue
        vistos[k] = r; uniq.append(r)
    rel_ok = uniq
    tipos = collections.Counter(s(r['Tipo_relacion']) for r in rel_ok)
    fam = [r for r in rel_ok if r['Tipo_relacion'] in PARENT]; esp = [r for r in rel_ok if r['Tipo_relacion'] in SPOUSE]
    aso = [r for r in rel_ok if r['Tipo_relacion'] not in PARENT | SPOUSE]
    sin_rel = [i for i in inc if not any(i in (r['ID_Persona_1'], r['ID_Persona_2']) for r in rel_ok)]
    res = {'P': P, 'inc': inc, 'rel_ok': rel_ok, 'REMAP': REMAP, 'pst': pst, 'rep': rep, 'dup': dup, 'trace': trace, 'sin_rel': sin_rel}
    if quiet: return res
    print(f'personas 01: {len(P)}; relaciones 02: {len(R)}')
    print(f'PERSONAS incluidas: {len(inc)} (de ellas Baja como extremo de relacion incluida: {len(ext_baja)}); excluidas/incoherentes/otras: {len(P) - len(inc)}')
    print(f'RELACIONES incluidas: {len(rel_ok)} (filiacion {len(fam)}, conyuge/esposo/esposa {len(esp)}, asociacion literal {len(aso)}); tipos: {dict(tipos)}')
    print(f'personas incluidas sin ninguna relacion incluida: {len(sin_rel)}')
    print(f'- aristas duplicadas deduplicadas con rastro: {len(dup)}: ' + '; '.join(dup))
    print(f'- aristas remapeadas con rastro: {len(trace)}: ' + ' '.join(trace))
    for k in sorted(rep): print(f'- {k}: {len(rep[k])}: {" ".join(map(str, rep[k][:60]))}{" ..." if len(rep[k]) > 60 else ""}')
    print('- asociaciones literales (no crean familia): ' + ' '.join(f"{r['ID_Relacion']}:{r['Tipo_relacion']}" for r in aso))
    print('- incluidas sin relacion: ' + ' '.join(sorted(sin_rel)))
    return res

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
