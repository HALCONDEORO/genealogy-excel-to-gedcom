import unittest, os, io, csv, sys, tempfile, contextlib, shutil
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
import mapper, mapper_v2, check_evidencia
from test_v2 import pers, rel, libro, oricsv, H02


def rel_n(*a, notas=None, **k):
    r = rel(*a, **k); r[H02.index('Notas')] = notas; return r

EVH = ['R-ID', 'hijo_P1', 'tipo_relacion', 'progenitor_P2', 'C-ID', 'D-ID', 'tipo_documento', 'alcance_validacion', 'literal_leido_exacto', 'incertidumbre_salvedades', 'fila_origen_literal']


def fila(r, h, t, p, lit='acta n.1: «Fulano Falso hijo de Beto Falso y de Cira Falsa»', salv='', c='C900010; C900011', d='D900001', alc='Solo DIRECCION hijo->progenitor (no identidad ni ortografia)', origen=''):
    return [r, h, t, p, c, d, 'acta', alc, lit, salv, origen]


def evcsv(path, filas, header=EVH):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f); w.writerow(header)
        for x in filas: w.writerow(x)


class Ev(unittest.TestCase):
    """Familia ficticia: padres P2 (Beto), P3 (Cira); hijos P1 (Ana) y P4 (Dani). Cuatro aristas; la incertidumbre se declara en una sola (R000004)."""
    def setUp(self):
        self.d = tempfile.mkdtemp()
        self.x = os.path.join(self.d, 'f.xlsx'); self.o = os.path.join(self.d, 'o.ged'); self.ori = os.path.join(self.d, 'ori.csv'); self.ev = os.path.join(self.d, 'ev.csv')
        self.personas = [pers('P000001', 'Ana', 'Falsa Uno', 'F'), pers('P000002', 'Beto', 'Falso Dos', 'M'), pers('P000003', 'Cira', 'Falsa Tres', 'F'), pers('P000004', 'Dani', 'Falso Cuatro', 'M')]
        self.rels = [rel('R000001', 'P000001', 'Padre biológico', 'P000002'), rel('R000002', 'P000001', 'Madre biológica', 'P000003'),
                     rel('R000003', 'P000004', 'Padre biológico', 'P000002'), rel_n('R000004', 'P000004', 'Madre biológica', 'P000003', conf='Media', notas='apellido Zorro? dudoso')]
        oricsv(self.ori, [('R000001', 'P000001', 'P000002'), ('R000002', 'P000001', 'P000003'), ('R000003', 'P000004', 'P000002'), ('R000004', 'P000004', 'P000003')])
        self.filas = [fila('R000001', 'P000001', 'Padre biológico', 'P000002'), fila('R000002', 'P000001', 'Madre biológica', 'P000003'),
                      fila('R000003', 'P000004', 'Padre biológico', 'P000002', lit='acta n.2: «Dani Falso hijo de Beto Falso»'),
                      fila('R000004', 'P000004', 'Madre biológica', 'P000003', lit='acta n.2: «y de Cira Falsa(?)»', salv='Lectura incierta (?) en el literal; no endurecer')]

    def gen(self, filas=None, personas=None):
        libro(self.x, personas or self.personas, self.rels)
        evcsv(self.ev, filas if filas is not None else self.filas)
        P, R, res, conf = mapper_v2.build(self.x, None, self.ori, None, self.ev)
        s, w, rep, un = mapper.run(self.x, self.o, persons=P, rels=R, evidencia=res['evidencia'])
        return s, open(self.o, encoding='utf-8').read()

    def chk(self):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            rc = check_evidencia.main(self.ev, self.o, self.x)
        return rc, buf.getvalue()

    def fails(self, filas, msg=None, ori=None):
        libro(self.x, self.personas, self.rels); evcsv(self.ev, filas)
        if ori: oricsv(self.ori, ori)
        with self.assertRaises(mapper.MapperError) as cm:
            P, R, res, conf = mapper_v2.build(self.x, None, self.ori, None, self.ev)
        if msg: self.assertIn(msg, str(cm.exception))

    def test_one_block_per_row_literal_once_and_independent_check_passes(self):
        s, t = self.gen()
        self.assertEqual(s['evidencia_bloques'], 4); self.assertEqual(t.count('EVIDENCIA RELACION'), 4)
        self.assertEqual(t.replace('\n2 CONC ', '').count('«y de Cira Falsa(?)»'), 1)
        rc, out = self.chk(); self.assertEqual(rc, 0, out)

    def test_uncertainty_only_on_its_edge_and_not_in_other_blocks(self):
        s, t = self.gen(); t = t.replace('\n2 CONC ', '')
        import re
        blk = {re.search(r'EVIDENCIA RELACION (R\d+)', l).group(1): l for l in t.split('\n') if 'EVIDENCIA RELACION' in l}
        self.assertIn('Lectura incierta', blk['R000004']); self.assertNotIn('Lectura incierta', blk['R000003']); self.assertNotIn('Lectura incierta', blk['R000002'])
        self.assertNotIn('Zorro', t)                                   # texto de Notas de 02 no se copia: Confianza Media sin fila no genera nota

    def test_confianza_media_without_row_gives_no_block(self):
        s, t = self.gen(filas=self.filas[:3])
        self.assertEqual(s['evidencia_bloques'], 3); self.assertNotIn('R000004', t.replace('\n2 CONC ', ''))
        evcsv(self.ev, self.filas); rc, out = self.chk()
        self.assertEqual(rc, 1)                                        # el comprobador ve que falta el bloque de la fila 4 si se le da el CSV completo
        evcsv(self.ev, self.filas[:3]); rc, out = self.chk(); self.assertEqual(rc, 0, out)

    def test_candidate_citations_are_labelled_not_proof(self):
        s, t = self.gen(); t = t.replace('\n2 CONC ', '')
        self.assertIn('NO cotejadas con el literal, no son prueba', t); self.assertIn('C900010; C900011', t)
        self.assertNotIn('1 SOUR', [l for l in t.split('\n') if 'EVIDENCIA' in l][0])

    def test_private_person_literal_omitted_and_checker_ok(self):
        pv = list(self.personas); pv[0] = pers('P000001', 'Ana', 'Falsa Uno', 'F', nac='2005-01-01', muerte=None)
        s, t = self.gen(personas=pv); t = t.replace('\n2 CONC ', '')
        self.assertEqual(s['evidencia_literal_omitido_privacidad'], 2)            # R1 y R2 (hijo privado); R3/R4 siguen con literal
        self.assertNotIn('«Fulano Falso hijo', t); self.assertIn('«y de Cira Falsa(?)»', t)
        rc, out = self.chk(); self.assertEqual(rc, 0, out)

    def test_identity_document_and_private_link_in_literal_scrubbed(self):
        f = list(self.filas); f[0] = fila('R000001', 'P000001', 'Padre biológico', 'P000002', lit='acta con DNI 12345678 de Beto')
        f[1] = fila('R000002', 'P000001', 'Madre biológica', 'P000003', lit='ver https://drive.google.com/x y mail a@b.org')
        s, t = self.gen(filas=f); t = t.replace('\n2 CONC ', '')
        self.assertNotIn('12345678', t); self.assertNotIn('drive.google', t); self.assertNotIn('a@b.org', t)
        self.assertIn('R000001:literal', s['omisiones_privacidad_ids']['libre_identidad'])

    def test_bad_inputs_fail_closed(self):
        f = self.filas
        self.fails(f + [f[0]], 'duplicada')
        self.fails([fila('R000099', 'P000001', 'Padre biológico', 'P000002')] + f, 'no esta en la orientacion')
        self.fails([fila('R000001', 'P000002', 'Padre biológico', 'P000001')] + f[1:], 'distinto de la orientacion')
        self.fails([fila('R000001', 'P000001', 'Madre biológica', 'P000002')] + f[1:], 'tipo')
        self.fails([fila('R000001', 'P000001', 'Padre biológico', 'P000002', lit='')] + f[1:], 'vacio')
        self.fails([fila('R000001', 'P000001', 'Padre biológico', 'P000002', alc='')] + f[1:], 'vacio')
        evcsv(self.ev, f, header=[h for h in EVH if h != 'alcance_validacion'])
        libro(self.x, self.personas, self.rels)
        with self.assertRaises(mapper.MapperError): mapper_v2.build(self.x, None, self.ori, None, self.ev)

    def test_idem_requires_valid_origin(self):
        f = list(self.filas)
        f[1] = fila('R000002', 'P000001', 'Madre biológica', 'P000003', lit='idem')
        self.fails(f, 'idem')
        f[1] = fila('R000002', 'P000001', 'Madre biológica', 'P000003', lit='idem', origen='R000003')    # origen de otro hijo
        self.fails(f, 'idem')
        f[1] = fila('R000002', 'P000001', 'Madre biológica', 'P000003', lit='idem', origen='R000001')
        s, t = self.gen(filas=f); t = t.replace('\n2 CONC ', '')
        self.assertIn('igual que R000001', t); rc, out = self.chk(); self.assertEqual(rc, 0, out)

    def test_checker_catches_tampering(self):
        self.gen(); good = open(self.o, encoding='utf-8').read()
        lines = good.split('\n')
        i = next(k for k, l in enumerate(lines) if 'EVIDENCIA RELACION R000002' in l)
        j = i + 1
        while j < len(lines) and lines[j].startswith('2 CONC'): j += 1
        def w(txt):
            open(self.o, 'w', encoding='utf-8').write(txt); return self.chk()
        rc, out = w('\n'.join(lines[:i] + lines[j:])); self.assertEqual(rc, 1); self.assertIn('R000002', out)                       # falta bloque
        rc, out = w('\n'.join(lines[:j] + lines[i:j] + lines[j:])); self.assertEqual(rc, 1); self.assertIn('R000002', out)         # duplicado
        rc, out = w(good.replace('progenitor P000003 | doc D900001', 'progenitor P000002 | doc D900001', 1)); self.assertEqual(rc, 1)   # progenitor cambiado
        rc, out = w(good.replace('EVIDENCIA RELACION R000004', 'EVIDENCIA RELACION R000077')); self.assertEqual(rc, 1); self.assertIn('R000077', out)  # bloque sin fila
        rc, out = w(good.replace('Falsa(?)', 'Falsa')); self.assertEqual(rc, 1)                                                    # literal alterado
        rc, out = w(good.replace('NO cotejadas con el literal, no son prueba', 'cotejadas')); self.assertEqual(rc, 1)             # C como prueba
        rc, out = w(good); self.assertEqual(rc, 0, out)

    def test_checker_flags_literal_leaked_for_private_person(self):
        pv = list(self.personas); pv[0] = pers('P000001', 'Ana', 'Falsa Uno', 'F', nac='2005-01-01', muerte=None)
        self.gen(personas=pv); t = open(self.o, encoding='utf-8').read().replace('\n2 CONC ', '')
        t = t.replace('[literal omitido por privacidad: hijo o progenitor privado (ver sidecar CSV de Genealogia)]', '«acta n.1: «Fulano Falso hijo de Beto Falso y de Cira Falsa»»', 1)
        open(self.o, 'w', encoding='utf-8').write(t); rc, out = self.chk(); self.assertEqual(rc, 1)

    def test_without_evidence_output_has_no_blocks(self):
        libro(self.x, self.personas, self.rels)
        P, R, res, conf = mapper_v2.build(self.x, None, self.ori)
        s, w, rep, un = mapper.run(self.x, self.o, persons=P, rels=R)
        self.assertNotIn('EVIDENCIA RELACION', open(self.o, encoding='utf-8').read()); self.assertEqual(s['evidencia_bloques'], 0)




class EvAdversarial(unittest.TestCase):
    """Comprobador adversarial: dos familias, filas privadas y no privadas, salvedades, idem. Cada mutacion del GED debe hacer fallar al comprobador."""
    def setUp(self):
        self.d = tempfile.mkdtemp()
        self.x = os.path.join(self.d, 'f.xlsx'); self.o = os.path.join(self.d, 'o.ged'); self.ori = os.path.join(self.d, 'ori.csv'); self.ev = os.path.join(self.d, 'ev.csv')
        P = [pers('P000001', 'Ana', 'Falsa Uno', 'F'), pers('P000002', 'Beto', 'Falso Dos', 'M'), pers('P000003', 'Cira', 'Falsa Tres', 'F'), pers('P000004', 'Dani', 'Falso Cuatro', 'M'),
             pers('P000005', 'Eva', 'Falsa Cinco', 'F'), pers('P000006', 'Fito', 'Falso Seis', 'M'), pers('P000007', 'Gus', 'Falso Siete', 'M', nac='2005-01-01', muerte=None)]
        R = [rel('R000001', 'P000001', 'Padre biológico', 'P000002'), rel('R000002', 'P000001', 'Madre biológica', 'P000003'),
             rel('R000003', 'P000004', 'Padre biológico', 'P000002'), rel('R000004', 'P000004', 'Madre biológica', 'P000003'),
             rel('R000005', 'P000007', 'Padre biológico', 'P000006'), rel('R000006', 'P000007', 'Madre biológica', 'P000005')]
        oricsv(self.ori, [(r[0], r[1], r[3]) for r in R])
        F = [fila('R000001', 'P000001', 'Padre biológico', 'P000002', lit='lit uno ALFA', c='C900010'), fila('R000002', 'P000001', 'Madre biológica', 'P000003', lit='lit dos BETA', salv='duda SALV2', c='C900011'),
             fila('R000003', 'P000004', 'Padre biológico', 'P000002', lit='lit tres GAMA', c='C900012'), fila('R000004', 'P000004', 'Madre biológica', 'P000003', lit='lit cuatro DELTA', c='C900013'),
             fila('R000005', 'P000007', 'Padre biológico', 'P000006', lit='lit cinco EPSI', salv='duda SALV5', c='C900014; C900015', d='D900002'),
             fila('R000006', 'P000007', 'Madre biológica', 'P000005', lit='idem', c='C900014', d='D900002', origen='R000005')]
        libro(self.x, P, R); evcsv(self.ev, F)
        Pp, Rr, res, conf = mapper_v2.build(self.x, None, self.ori, None, self.ev)
        mapper.run(self.x, self.o, persons=Pp, rels=Rr, evidencia=res['evidencia'])
        self.good = open(self.o, encoding='utf-8').read().replace('\n2 CONC ', '')

    def chk(self, txt):
        open(self.o, 'w', encoding='utf-8').write(txt)
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf): rc = check_evidencia.main(self.ev, self.o, self.x)
        return rc, buf.getvalue()

    def test_baseline_ok_and_private_uses_fixed_mark_only(self):
        rc, out = self.chk(self.good); self.assertEqual(rc, 0, out)
        self.assertIn('Esta relación tiene salvedades de lectura; consulte sidecar R000005', self.good)
        self.assertNotIn('SALV5', self.good); self.assertNotIn('EPSI', self.good); self.assertIn('SALV2', self.good); self.assertIn('ALFA', self.good)

    def test_every_mutation_is_caught(self):
        g = self.good; L = g.split('\n')
        def blk(r): return next(i for i, l in enumerate(L) if f'EVIDENCIA RELACION {r} ' in l)
        def sw(a, b, x, y):
            return g.replace(a, '@@T@@').replace(b, a).replace('@@T@@', b)
        muts = {
            'borrar bloque': '\n'.join(L[:blk('R000003')] + L[blk('R000003') + 1:]),
            'duplicar bloque': '\n'.join(L[:blk('R000003')] + [L[blk('R000003')]] + L[blk('R000003'):]),
            'etiquetas R intercambiadas': g.replace('RELACION R000001 ', 'RELACION @@ ').replace('RELACION R000003 ', 'RELACION R000001 ').replace('RELACION @@ ', 'RELACION R000003 '),
            'hijo cambiado': g.replace('R000001 | hijo P000001', 'R000001 | hijo P000004'),
            'progenitor cambiado': g.replace('R000001 | hijo P000001 | Padre biológico | progenitor P000002', 'R000001 | hijo P000001 | Padre biológico | progenitor P000003'),
            'tipo cambiado': g.replace('R000001 | hijo P000001 | Padre biológico', 'R000001 | hijo P000001 | Madre biológica'),
            'doc cambiado': g.replace('doc D900001 (', 'doc D900009 (', 1),
            'literal alterado': g.replace('lit uno ALFA', 'lit uno ALFX'),
            'literal intercambiado': g.replace('lit uno ALFA', '@@').replace('lit tres GAMA', 'lit uno ALFA').replace('@@', 'lit tres GAMA'),
            'literal duplicado': g.replace('«lit uno ALFA»', '«lit uno ALFA» «lit uno ALFA»'),
            'salvedad alterada': g.replace('duda SALV2', 'duda SALV9'),
            'salvedad borrada': g.replace('salvedades: duda SALV2 |', 'salvedades: ninguna declarada en el sidecar |'),
            'salvedad inventada': g.replace('R000001 | hijo P000001 | Padre biológico | progenitor P000002 | doc D900001 (acta) | alcance de la validacion: Solo DIRECCION hijo->progenitor (no identidad ni ortografia) | literal leido: «lit uno ALFA» | salvedades: ninguna declarada en el sidecar', 'R000001 | hijo P000001 | Padre biológico | progenitor P000002 | doc D900001 (acta) | alcance de la validacion: Solo DIRECCION hijo->progenitor (no identidad ni ortografia) | literal leido: «lit uno ALFA» | salvedades: duda INVENTADA'),
            'privado: literal filtrado': g.replace('[literal omitido por privacidad: hijo o progenitor privado (ver sidecar CSV de Genealogia)]', 'lit cinco EPSI', 1),
            'privado: texto de salvedad filtrado': g.replace('Esta relación tiene salvedades de lectura; consulte sidecar R000005', 'duda SALV5'),
            'privado: marca fija con otra R': g.replace('consulte sidecar R000005', 'consulte sidecar R000002'),
            'privado: marca fija ausente': g.replace('Esta relación tiene salvedades de lectura; consulte sidecar R000005', 'ninguna declarada en el sidecar'),
            'idem sin origen': g.replace('igual que R000005 ', ''),
            'C como prueba': g.replace('NO cotejadas con el literal, no son prueba', 'cotejadas'),
            'C candidata perdida': g.replace('C900014; C900015', 'C900014'),
            'bloque sin fila': g.replace('EVIDENCIA RELACION R000004 ', 'EVIDENCIA RELACION R000044 '),
            'bloque en FAM equivocada': '\n'.join(L).replace(L[blk('R000005')], '@@MOVE@@', 1).replace('@@MOVE@@', '').replace('1 HUSB @P000002@', L[blk('R000005')] + '\n1 HUSB @P000002@', 1),
            'bloque colgado de persona': g.replace('0 @P000001@ INDI\n', '0 @P000001@ INDI\n' + L[blk('R000001')] + '\n', 1),
            'RESN quitado a persona privada': g.replace('1 RESN privacy\n', '', 1) if False else None,
        }
        muts = {k: v for k, v in muts.items() if v is not None}
        for k, v in muts.items():
            self.assertNotEqual(v, g, f'mutacion sin efecto: {k}')
            rc, out = self.chk(v); self.assertEqual(rc, 1, f'no detectada: {k}\n{out}')
        self.assertEqual(self.chk(g)[0], 0)

    def test_privacy_flip_is_caught(self):
        g = self.good
        i = g.index('0 @P000007@ INDI'); j = g.index('\n0 ', i + 5)
        no_resn = g[:i] + g[i:j].replace('\n1 RESN privacy', '') + g[j:]
        self.assertNotEqual(no_resn, g); rc, out = self.chk(no_resn); self.assertEqual(rc, 1, out)           # si P7 no fuera privada, el literal faltaria


class EvSondasExternas(unittest.TestCase):
    """Las dos sondas del revisor externo (21:55) y el raise por incoherencia, convertidas en tests."""
    def setUp(self):
        self.base = EvAdversarial('test_baseline_ok_and_private_uses_fixed_mark_only'); self.base.setUp()

    def test_sonda_B_resn_stripped_and_literals_restored_is_caught_by_excel_privacy(self):
        g = self.base.good
        sin = g.replace('\n1 RESN privacy', '')
        sin = sin.replace('[literal omitido por privacidad: hijo o progenitor privado (ver sidecar CSV de Genealogia)]', '«lit cinco EPSI»', 1)
        rc, out = self.base.chk(sin); self.assertEqual(rc, 1); self.assertIn('RESN', out)

    def test_sonda_A1_literal_with_url_email_digits_deleted_is_caught(self):
        b = self.base
        f = [fila('R000001', 'P000001', 'Padre biológico', 'P000002', lit='ver https://www.archives41.fr/ark/1 y https://drive.google.com/x, mail a@b.org, ref 12345678', c='C900010'),
             fila('R000002', 'P000001', 'Madre biológica', 'P000003', lit='lit dos BETA', c='C900011')]
        rows = list(csv.DictReader(open(b.ev, encoding='utf-8')))
        evcsv(b.ev, f + [fila(r['R-ID'], r['hijo_P1'], r['tipo_relacion'], r['progenitor_P2'], lit=r['literal_leido_exacto'], salv=r['incertidumbre_salvedades'], c=r['C-ID'], d=r['D-ID'], origen=r['fila_origen_literal']) for r in rows[2:]])
        Pp, Rr, res, conf = mapper_v2.build(b.x, None, b.ori, None, b.ev)
        mapper.run(b.x, b.o, persons=Pp, rels=Rr, evidencia=res['evidencia'])
        g = open(b.o, encoding='utf-8').read().replace('\n2 CONC ', '')
        self.assertEqual(b.chk(g)[0], 0)
        self.assertIn('https://www.archives41.fr/ark/1', g); self.assertNotIn('drive.google', g); self.assertNotIn('a@b.org', g); self.assertNotIn('12345678', g)
        for k, v in {'enlace publico borrado': g.replace('https://www.archives41.fr/ark/1', ''), 'marca de enlace privado borrada': g.replace('[enlace omitido por privacidad]', ''),
                     'marca de correo borrada': g.replace('[correo omitido por privacidad]', ''), 'marca de numero borrada': g.replace('[numero omitido por privacidad]', ''),
                     'correo sin sanear': g.replace('[correo omitido por privacidad]', 'a@b.org'), 'enlace privado sin sanear': g.replace('[enlace omitido por privacidad]', 'https://drive.google.com/x', 1)}.items():
            self.assertNotEqual(v, g, k); self.assertEqual(b.chk(v)[0], 1, k)

    def test_checker_requires_excel_or_explicit_solo_ged(self):
        import subprocess
        r = subprocess.run([sys.executable, os.path.join(os.path.dirname(__file__), '..', 'check_evidencia.py'), self.base.ev, self.base.o], capture_output=True, text=True)
        self.assertEqual(r.returncode, 2)

    def test_evidence_row_whose_parent_is_not_in_the_family_raises(self):
        b = self.base
        Pp, Rr, res, conf = mapper_v2.build(b.x, None, b.ori, None, b.ev)
        for p in Pp:
            if p['id_persona'] == 'P000001': p['id_padre'] = ''     # incoherencia: el hijo ya no tiene a P000002 como padre
        with self.assertRaises(mapper.MapperError) as cm:
            mapper.run(b.x, b.o, persons=Pp, rels=Rr, evidencia=res['evidencia'])
        self.assertIn('R000001', str(cm.exception))


class GedcomTextRoundtrip(unittest.TestCase):
    def unir(self, lines):
        out = ''
        for l in lines:
            if ' CONT' in l[:9] and l.split(' ', 2)[1] == 'CONT': out += '\n' + (l.split(' ', 2)[2] if len(l.split(' ', 2)) > 2 else '')
            elif l.split(' ', 2)[1] == 'CONC': out += l.split(' ', 2)[2]
            else: out += l.split(' ', 2)[2] if len(l.split(' ', 2)) > 2 else ''
        return out

    def test_space_at_chunk_boundary_is_preserved(self):
        M = mapper.MAXLEN
        for pos in (M - 1, M, M + 1, 2 * M - 1):
            for texto in ('a' * pos + ' bbb ccc', 'x\n' + 'a' * (pos - 2) + ' bbb ccc'):
                self.assertEqual(self.unir(mapper.gedcom_text(1, 'NOTE', texto)), texto, (pos, texto[:5]))


class EvEspacios(unittest.TestCase):
    def test_checker_detects_lost_or_glued_space_and_accepts_long_notes_exactly(self):
        b = EvAdversarial('test_baseline_ok_and_private_uses_fixed_mark_only'); b.setUp()
        raw = open(b.o, encoding='utf-8').read()                      # GED real con CONC (sin unir)
        self.assertIn('2 CONC', raw)                                  # el fixture contiene notas largas partidas

        def rc(txt):
            open(b.o, 'w', encoding='utf-8').write(txt)
            with contextlib.redirect_stdout(io.StringIO()): return check_evidencia.main(b.ev, b.o, b.x)
        self.assertEqual(rc(raw), 0)
        for antes, despues in (('literal leido: ', 'literal leido:'), ('igual que R000005', 'igualque R000005'), ('| alcance de la validacion', '|alcance de la validacion')):
            g = raw.replace(antes, despues, 1)
            if g == raw:                                              # la cadena cae en un limite CONC: se une primero
                u = raw.replace('\n2 CONC ', ''); g = u.replace(antes, despues, 1)
            self.assertNotEqual(g, raw, antes); self.assertEqual(rc(g), 1, antes)
        # esos mismos espacios, perdidos justo en un limite de trozo, tambien se detectan
        u = raw.replace('\n2 CONC ', '')
        self.assertEqual(rc(u.replace('igual que', 'igualque', 1)), 1)


if __name__ == '__main__':
    unittest.main()
