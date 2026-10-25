#!/usr/bin/env python3
"""Comprobador INDEPENDIENTE de bloques de evidencia por relacion (no importa mapper*.py, ni v2_selector).
Uso: python3 check_evidencia.py evidencia.csv salida.ged --excel=libro.xlsx        (modo completo)
     python3 check_evidencia.py evidencia.csv salida.ged --solo-ged                 (LIMITE: privacidad tomada del propio GED; no es gate)
Modo completo: la privacidad esperada de cada persona se calcula desde el Excel (01_Personas y 06_Documentos) con regla propia y se compara con el RESN del GED persona a persona.
Para cada fila del CSV compara el texto COMPLETO del bloque contra el texto esperado, calculado aqui (incluido el saneado de texto libre, sin exencion para URL, correo o numeros de 8+ cifras).
Ademas: 1 bloque por fila, en la FAM que une ese hijo con ese progenitor; sin bloques huerfanos ni duplicados ni en INDI. Salida 1 si hay cualquier fallo.
Compara reconstruyendo la nota segun GEDCOM (CONC une sin separador, CONT = salto) y normalizando solo la longitud de los espacios (varios seguidos = uno); un espacio perdido o pegado ('igual'+'que') es fallo."""
import sys, csv, re, datetime
from urllib.parse import urlsplit

LIMITE_ANIO = 1906                       # nacidos despues y sin defuncion registrada: posiblemente vivos (politica del proyecto)
PUBLICO = ('publico', 'público', 'public')
HOSTS_PRIVADOS = ('drive.google.com', 'docs.google.com', 'mail.google.com', 'sheets.google.com', 'photos.google.com', 'photos.app.goo.gl', 'goo.gl', 'forms.gle',
                  'dropbox.com', 'dropboxusercontent.com', '1drv.ms', 'onedrive.live.com', 'sharepoint.com', 'icloud.com', 'box.com', 'wetransfer.com', 'we.tl', 'mega.nz', 'mega.io',
                  'notion.so', 'notion.site', 'whatsapp.com', 'wa.me', 't.me', 'telegram.me', 'bit.ly', 'tinyurl.com', 'localhost')
CONSULTA_SECRETA = re.compile(r'(?i)(^|[?&;])(token|access_token|key|api_?key|sig|signature|auth|authuser|share|sharing|usp|invite|session|sid|password|pwd)=')
IDENTIDAD = re.compile(r'\b(DNI|NIE|pasaporte|passport|carnet|licencia de conducir|permiso de conducir|libro de familia|documento nacional de identidad)\b', re.I)
ENLACE = re.compile(r'(?i)\b(?:https?|ftp|file)://[^\s<>"\')\]]+|\bwww\.[^\s<>"\')\]]+')
CORREO = re.compile(r'[\w.+-]+@[\w-]+(?:\.[\w-]+)+')
DIGITOS = re.compile(r'(?<![\w])\d{8,}(?![\w])')
MARCA_ID = '[campo omitido por privacidad: referencia a documento de identidad (ver Excel)]'
MARCA_LIT = '[literal omitido por privacidad: hijo o progenitor privado (ver sidecar CSV de Genealogia)]'


def n(s):
    return re.sub(r'\s+', ' ', str(s if s is not None else '')).strip()


def enlace_privado(u):
    bruto = u if re.match(r'(?i)[a-z][a-z0-9+.-]*://', u) else 'https://' + u
    try:
        sp = urlsplit(bruto); host = (sp.hostname or '').lower()
    except ValueError:
        return True
    if sp.scheme.lower() != 'https' or not host or sp.username or sp.password:
        return True
    if re.fullmatch(r'[\d.]+|[0-9a-f:]+', host) or '.' not in host:
        return True
    if any(host == h or host.endswith('.' + h) for h in HOSTS_PRIVADOS):
        return True
    return bool(CONSULTA_SECRETA.search(sp.query) or CONSULTA_SECRETA.search(sp.fragment))


def esperado(texto):
    """Saneado esperado de un campo de texto libre NO privado, calculado aqui."""
    s = n(texto)
    if not s:
        return s
    if IDENTIDAD.search(s):
        return MARCA_ID
    s = ENLACE.sub(lambda m: '[enlace omitido por privacidad]' if enlace_privado(m.group(0)) else m.group(0), s)
    s = CORREO.sub('[correo omitido por privacidad]', s)
    return DIGITOS.sub('[numero omitido por privacidad]', s)


SOLO_POR_FORMATO = []


def privacidad_excel(xlsx):
    import openpyxl
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    def tabla(hoja, clave):
        filas = list(wb[hoja].iter_rows(values_only=True)); k = next(i for i, r in enumerate(filas) if r and r[0] == clave)
        return [dict(zip(filas[k], r)) for r in filas[k + 1:] if r and r[0]]
    docs = tabla('06_Documentos', 'ID_Documento')
    lugares = {l['ID_Lugar'] for l in tabla('07_Lugares', 'ID_Lugar')}      # un lugar de defuncion solo cuenta si es un ID de 07_Lugares (texto libre no se resuelve)
    priv = {}
    for p in tabla('01_Personas', 'ID_Persona'):
        i = p['ID_Persona']
        nac = p.get('Fecha_nacimiento'); prec = n(p.get('Precision_nacimiento'))
        t = nac.strftime('%Y-%m-%d') if isinstance(nac, (datetime.date, datetime.datetime)) else n(nac)
        # texto de fecha tal como lo formatea la conversion: solo 'Exacta*' y 'Aproximada' de un solo anio dejan un anio legible; el resto queda como '[precision] fecha'
        if not t: fmt = ''
        elif prec.startswith('Exacta'): fmt = t
        elif prec.startswith('Aproximada'): fmt = 'c. ' + t if re.fullmatch(r'\d{4}', t) else '[aprox] ' + t
        else: fmt = f'[{prec[:40] or "sin precision"}] {t}'
        anio = (re.match(r'(?:c\. )?(\d{4})', fmt) or [None, None])[1]
        sin_muerte = not (n(p.get('Fecha_defuncion')) or (n(p.get('ID_Lugar_defuncion')) if n(p.get('ID_Lugar_defuncion')) in lugares else ''))
        viva = sin_muerte and (not anio or int(anio) > LIMITE_ANIO)
        doc_priv = any(n(d.get('Derechos_privacidad')).lower() not in PUBLICO for d in docs if d.get('ID_Persona_principal') == i)
        anio_real = (re.search(r'\d{4}', t) or [None])[0]
        if viva and not doc_priv and anio_real and int(anio_real) <= LIMITE_ANIO and sin_muerte:
            SOLO_POR_FORMATO.append(i)      # privada solo porque la fecha no es de un unico anio: sobre-restriccion de la conversion
        priv[i] = bool(viva or doc_priv)
    return priv


def lee_ged(path):
    fams, indi, cur, kind, nota = {}, {}, None, None, None
    for ln in open(path, encoding='utf-8').read().split('\n'):
        m = re.match(r'(\d+) (@\S+@ )?(\S+)(?: (.*))?$', ln)
        if not m:
            continue
        lv, xr, tag, val = int(m.group(1)), (m.group(2) or '').strip().strip('@'), m.group(3), m.group(4) or ''
        if lv == 0:
            nota = None; cur = None
            if tag == 'FAM': cur = fams.setdefault(xr, {'husb': '', 'wife': '', 'chil': [], 'notes': []}); kind = 'FAM'
            elif tag == 'INDI': cur = indi.setdefault(xr, {'resn': False, 'notes': []}); kind = 'INDI'
            continue
        if cur is None:
            continue
        if lv == 1:
            nota = None
            if tag == 'HUSB' and kind == 'FAM': cur['husb'] = val.strip('@')
            elif tag == 'WIFE' and kind == 'FAM': cur['wife'] = val.strip('@')
            elif tag == 'CHIL' and kind == 'FAM': cur['chil'].append(val.strip('@'))
            elif tag == 'RESN' and kind == 'INDI' and val.strip() == 'privacy': cur['resn'] = True
            elif tag == 'NOTE': cur['notes'].append(val); nota = cur['notes']
        elif lv == 2 and nota is not None and tag in ('CONC', 'CONT'):
            nota[-1] += ('\n' if tag == 'CONT' else '') + val
        else:
            nota = None
    return fams, indi


def main(csvp, gedp, excel=None):
    fams, indi = lee_ged(gedp)
    rows = list(csv.DictReader(open(csvp, encoding='utf-8', newline='')))
    por = {r['R-ID'].strip(): r for r in rows}
    err = []
    if excel:
        esp_priv = privacidad_excel(excel)
        for i, d in indi.items():
            if i in esp_priv and esp_priv[i] != d['resn']:
                err.append(f"{i}: RESN en el GED={d['resn']} pero el Excel implica privado={esp_priv[i]}")
        modo = 'completo (privacidad desde Excel)'
    else:
        esp_priv = {i: d['resn'] for i, d in indi.items()}
        modo = 'SOLO-GED (limite: privacidad tomada del propio GED)'
    bloques = {}
    for fx, f in fams.items():
        for t in f['notes']:
            if t.startswith('EVIDENCIA RELACION '):
                bloques.setdefault(t.split()[2], []).append((fx, f, t))
    for xi, d in indi.items():
        if any('EVIDENCIA RELACION' in t for t in d['notes']): err.append(f'bloque de evidencia en INDI {xi}')
    for r in bloques:
        if r not in por: err.append(f'{r}: bloque sin fila en el CSV')
    omit = 0
    for r, x in por.items():
        h, p = x['hijo_P1'].strip(), x['progenitor_P2'].strip()
        b = bloques.get(r, [])
        if len(b) != 1:
            err.append(f'{r}: {len(b)} bloques (se espera 1)'); continue
        fx, f, t = b[0]
        if h not in f['chil'] or p not in (f['husb'], f['wife']):
            err.append(f'{r}: bloque en FAM {fx} que no une hijo {h} con progenitor {p}'); continue
        if any(q != r and q in por and blq[0][0] == fx and por[q]['hijo_P1'].strip() == h and por[q]['progenitor_P2'].strip() == p for q, blq in bloques.items()):
            err.append(f'{r}: otra relacion con mismo par hijo/progenitor en la FAM')
        if h not in esp_priv or p not in esp_priv:
            err.append(f'{r}: hijo o progenitor sin registro de privacidad'); continue
        priv = esp_priv[h] or esp_priv[p]
        lit_raw = n(x['literal_leido_exacto']); origen = ''; propio = ''
        if lit_raw.lower().startswith('idem'):
            origen = n(x.get('fila_origen_literal')); o = por.get(origen)
            if not o or n(o['hijo_P1']) != h or n(o['D-ID']) != n(x['D-ID']) or n(o['literal_leido_exacto']).lower().startswith('idem'):
                err.append(f'{r}: idem sin origen valido'); continue
            if not re.fullmatch(r'idem[\s;:,.]*', lit_raw, re.I): propio = lit_raw
            lit_raw = n(o['literal_leido_exacto'])
        lit = MARCA_LIT if priv else '«' + esperado(lit_raw) + '»'
        if propio: lit += ' | texto propio de la fila: ' + ('[omitido por privacidad]' if priv else '«' + esperado(propio) + '»')
        if origen: lit = f'igual que {origen} (misma acta, ese literal nombra a ambos progenitores): ' + lit
        sv = n(x['incertidumbre_salvedades'])
        if priv: salv = f'Esta relación tiene salvedades de lectura; consulte sidecar {r}' if sv else 'ninguna declarada en el sidecar'
        else: salv = esperado(sv) or 'ninguna declarada en el sidecar'
        omit += 1 if priv else 0
        esp = (f"EVIDENCIA RELACION {r} | hijo {h} | {n(x['tipo_relacion'])} | progenitor {p} | doc {n(x['D-ID'])} ({esperado(x['tipo_documento'])}) | "
               f"alcance de la validacion: {esperado(x['alcance_validacion'])} | literal leido: {lit} | salvedades: {salv} | "
               f"citas candidatas de 02 (NO cotejadas con el literal, no son prueba de esta relacion): {n(x['C-ID']) or 'ninguna'} | "
               f"Nota de FAMILIA: se atribuye solo a la relacion {r}; lectura incierta no es identidad confirmada.")
        if n(t) != n(esp):
            err.append(f'{r}: el bloque no coincide con el texto esperado')
    if SOLO_POR_FORMATO: print(f'INFO: {len(SOLO_POR_FORMATO)} personas privadas solo por formato de fecha (anio <= {LIMITE_ANIO} pero rango o precision no exacta, sin defuncion): {sorted(set(SOLO_POR_FORMATO))}')
    print(f'modo: {modo}; filas CSV: {len(por)}; bloques: {sum(len(v) for v in bloques.values())}; con literal omitido por privacidad: {omit}; fallos: {len(err)}')
    for e in err: print('  FALLO', e)
    return 1 if err else 0


if __name__ == '__main__':
    a = [x for x in sys.argv[1:] if not x.startswith('--')]
    ex = next((x.split('=', 1)[1] for x in sys.argv if x.startswith('--excel=')), None)
    if len(a) != 2 or not (ex or '--solo-ged' in sys.argv):
        print('uso: check_evidencia.py evidencia.csv salida.ged (--excel=libro.xlsx | --solo-ged)'); sys.exit(2)
    sys.exit(main(a[0], a[1], ex))
