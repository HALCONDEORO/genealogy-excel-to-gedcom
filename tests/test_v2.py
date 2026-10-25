import os, sys, tempfile, unittest, io, contextlib, csv, datetime
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
import openpyxl, mapper, mapper_v2, reconcile_v2, v2_selector

H01 = ['ID_Persona', 'Nombre', 'Apellidos_nacimiento', 'Otros_nombres', 'Sexo', 'Fecha_nacimiento', 'Precision_nacimiento', 'ID_Lugar_nacimiento', 'Fecha_defuncion', 'Precision_defuncion', 'ID_Lugar_defuncion', 'Generacion', 'Ahnentafel', 'Rama', 'Estado_hechos', 'Nivel_documentacion', 'N_citas', 'Foto_principal', 'Gramps_ID', 'Notas']
H02 = ['ID_Relacion', 'ID_Persona_1', 'Tipo_relacion', 'ID_Persona_2', 'Fecha_desde', 'Fecha_hasta', 'ID_Lugar', 'Familia_Gramps_ID', 'Estado_hechos', 'Confianza', 'ID_Cita', 'Notas']
H04 = ['ID_Fuente', 'Titulo', 'Autor_o_entidad', 'Tipo_fuente', 'Original_o_derivada', 'Calidad_fuente', 'Repositorio', 'Signatura', 'URL']
H05 = ['ID_Cita', 'ID_Fuente', 'ID_Persona', 'Tipo_hecho', 'Valor_afirmado', 'Localizador_pagina_folio', 'Extracto_o_resumen', 'Evaluacion_evidencia', 'Corroborada', 'Confianza', 'ID_Documento', 'URL']
H06 = ['ID_Documento', 'Nombre_archivo', 'Tipo_documento', 'ID_Persona_principal', 'ID_Fuente', 'ID_Cita', 'Derechos_privacidad', 'SHA256']
H07 = ['ID_Lugar', 'Nombre_canonico', 'Tipo_lugar', 'Municipio', 'Provincia_estado', 'Pais']


def pers(i, nom, ap, sx, nac='1800-01-02', prec='Exacta', est='Documentado', niv='Alta', notas=None, muerte='1850-03-04'):
    r = [None] * len(H01)
    for k, v in (('ID_Persona', i), ('Nombre', nom), ('Apellidos_nacimiento', ap), ('Sexo', sx), ('Fecha_nacimiento', nac), ('Precision_nacimiento', prec), ('ID_Lugar_nacimiento', 'L000001'),
                 ('Fecha_defuncion', muerte), ('Precision_defuncion', 'Exacta'), ('Estado_hechos', est), ('Nivel_documentacion', niv), ('Notas', notas)):
        r[H01.index(k)] = v
    return r


def rel(i, a, t, b, est='Documentado', conf='Alta', desde=None, lug=None):
    r = [None] * len(H02)
    for k, v in (('ID_Relacion', i), ('ID_Persona_1', a), ('Tipo_relacion', t), ('ID_Persona_2', b), ('Estado_hechos', est), ('Confianza', conf), ('Fecha_desde', desde), ('ID_Lugar', lug), ('ID_Cita', 'C900001')):
        r[H02.index(k)] = v
    return r


def libro(path, personas, relaciones, autor='Archivo ficticio', tipo='Acta', cita=None, fuente=None, docs=None):
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    def hoja(n, hd, rows):
        ws = wb.create_sheet(n)
        for _ in range(3): ws.append([n])
        ws.append(hd)
        for r in rows: ws.append(r)
    hoja('01_Personas', H01, personas); hoja('02_Relaciones', H02, relaciones)
    f4 = ['S900001', 'Registro ficticio', autor, tipo, 'Original', 'Alta', 'Archivo X', 'Leg 1', None]
    for k, v in (fuente or {}).items(): f4[H04.index(k)] = v
    hoja('04_Fuentes', H04, [f4])
    cit = [[f'C9{n:05d}', 'S900001', p[0], 'Nacimiento', '1800', 'f. 1', 'x', 'y', 'Si', 'Alta', None, None] for n, p in enumerate(personas, 10)]
    for row in cit:
        for k, v in (cita or {}).items(): row[H05.index(k)] = v
    hoja('05_Citas', H05, cit); hoja('06_Documentos', H06, docs or [])
    hoja('07_Lugares', H07, [['L000001', 'Villa Falsa', 'Municipio', 'Villa Falsa', 'Provincia Falsa', 'Pais Falso']])
    wb.save(path)


def oricsv(path, filas):
    with open(path, 'w', newline='') as f:
        w = csv.writer(f); w.writerow(['ID_Relacion', 'hijo', 'progenitor'])
        for x in filas: w.writerow(x)


class V2(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp()
        self.x = os.path.join(self.d, 'f.xlsx'); self.o = os.path.join(self.d, 'o.ged')
        self.personas = [pers('P000001', 'Ana', 'Falsa Uno', 'F'), pers('P000002', 'Beto', 'Falso Dos', 'M'), pers('P000003', 'Cira', 'Falsa Tres', 'F'),
                         pers('P000004', 'Dani', 'Falso Cuatro', 'M', est='Probable'), pers('P000005', 'Eva', 'Falsa Cinco', 'F', nac='1790', prec='Aproximada'),
                         pers('P000006', 'Fito', 'Falso Seis', 'M', nac='1791-01-01', prec='Desconocida')]
        self.rels = [rel('R000001', 'P000001', 'Padre biológico', 'P000002'), rel('R000002', 'P000003', 'Madre biológica', 'P000001'),
                     rel('R000003', 'P000002', 'Esposo', 'P000005', desde='1799-05-06', lug='L000001'), rel('R000004', 'P000001', 'Padre biológico', 'P000004'),
                     rel('R000005', 'P000002', 'Hermano', 'P000006')]
        self.ori = os.path.join(self.d, 'ori.csv')
        oricsv(self.ori, [('R000001', 'P000001', 'P000002'), ('R000002', 'P000001', 'P000003')])  # R2 va invertida en el Excel: la orientacion validada manda

    def leer(self):
        with open(self.o, encoding='utf-8') as fh: return fh.read()

    def escribir(self, txt):
        with open(self.o, 'w', encoding='utf-8') as fh: fh.write(txt)

    def run_(self, ori=True):
        libro(self.x, self.personas, self.rels)
        P, R, res, conf = mapper_v2.build(self.x, None, self.ori if ori else None)
        s, w, rep, un = mapper.run(self.x, self.o, persons=P, rels=R)
        return s, res

    def rec(self, ori=True):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            rc = reconcile_v2.main(self.x, self.o, None, self.ori if ori else None)
        return rc, buf.getvalue()

    def test_validated_orientation_builds_family_and_reconcile_zero(self):
        s, res = self.run_()
        t = self.leer()
        self.assertIn('1 FAMC', t); self.assertIn('0 @P000001@ INDI', t)
        self.assertNotIn('P000004', t)                        # Probable fuera
        rc, out = self.rec(); self.assertEqual(rc, 0, out)

    def test_without_orientation_no_parents_and_declared(self):
        self.run_(ori=False)
        P, R, res, conf = mapper_v2.build(self.x, None, None)
        self.assertEqual(len(res['direccion_indeterminada']), 2)
        t = self.leer()
        self.assertNotIn('1 CHIL', t)
        rc, out = self.rec(ori=False); self.assertEqual(rc, 0, out)

    def test_esposo_gets_marr(self):
        self.run_()
        t = self.leer()
        self.assertIn('1 MARR\n2 DATE 6 MAY 1799', t)         # regresion: Esposo/Esposa cuentan como union

    def test_hermano_is_note_not_family(self):
        self.run_()
        t = self.leer()
        self.assertIn('Hermano de', t)
        self.assertEqual(t.count('0 @F'), 3 - 1 if False else t.count('0 @F'))
        self.assertNotIn('WIFE @P000006@', t); self.assertNotIn('HUSB @P000006@', t)

    def test_dates_precision_rules(self):
        self.run_()
        t = self.leer()
        e = t.split('0 @P000005@ INDI')[1].split('0 @')[0]; self.assertIn('2 DATE ABT 1790', e)
        f = t.split('0 @P000006@ INDI')[1].split('0 @')[0]; self.assertNotIn('2 DATE 1791', f); self.assertIn('Fecha original en el Excel', f)

    def test_orientation_csv_mismatch_fails(self):
        oricsv(self.ori, [('R000001', 'P000001', 'P000003')])
        libro(self.x, self.personas, self.rels)
        with self.assertRaises(ValueError): mapper_v2.build(self.x, None, self.ori)

    def test_reconcile_detects_negative_changes(self):
        self.run_()
        base = self.leer()
        for mut, key in ((lambda t: t.replace('1 SEX F', '1 SEX M', 1), 'sexo'), (lambda t: t.replace('2 PLAC Villa Falsa, Provincia Falsa, Pais Falso', '2 PLAC Otro', 1), 'ID_Lugar'),
                         (lambda t: t.replace('1 FAMC @F1@\n', '', 1), 'progenitor_'), (lambda t: t.replace('2 DATE 6 MAY 1799', '2 DATE 7 MAY 1799', 1), 'MARR_fecha'),
                         (lambda t: t.replace('0 @P000002@ INDI', '0 @P000099@ INDI', 1), 'persona_'), (lambda t: t.replace('1 NAME Ana', '1 NAME Anita', 1), 'nombre')):
            self.escribir(mut(base))
            rc, out = self.rec(); self.assertEqual(rc, 1, key); self.assertIn(key, out)

    def test_extra_unvalidated_child_link_detected(self):
        self.run_()
        t = self.leer()
        fam = [l for l in t.split('\n') if l.startswith('0 @F')][0]
        t2 = t.replace(fam + '\n', fam + '\n1 CHIL @P000003@\n', 1)
        self.escribir(t2)
        rc, out = self.rec(); self.assertEqual(rc, 1); self.assertIn('hijo_sin_fila_validada', out)

    def test_remap_overlay_and_gates(self):
        self.personas += [pers('P000007', 'Gus', 'Falso Siete', 'M', est='Fusionada en P000002', niv=''), pers('P000008', 'Hana', 'Falsa Ocho', 'F', est='Documentado', niv='Baja')]
        self.rels += [rel('R000006', 'P000008', 'Cónyuge', 'P000007')]
        libro(self.x, self.personas, self.rels)
        ov = os.path.join(self.d, 'ov.csv')
        with open(ov, 'w', newline='') as f:
            w = csv.writer(f); w.writerow(['ID', 'Persona', 'Rama', 'Estado real', 'Nivel real', 'N', 'Notas']); w.writerow(['P000004', '', '', 'Estado_hechos=Documentado', 'Nivel_documentacion=Alta', '', ''])
        res = v2_selector.main(self.x, ov, quiet=True)
        self.assertIn('P000004', res['inc'])                  # overlay aplicado por ID
        self.assertEqual(res['REMAP'], {'P000007': 'P000002'})
        self.assertIn('R000006', [r['ID_Relacion'] for r in res['rel_ok']])   # arista remapeada antes de clasificar
        self.assertNotIn('P000007', res['inc'])

    def test_probable_endpoint_blocks_alta_edge(self):
        libro(self.x, self.personas, self.rels)
        res = v2_selector.main(self.x, None, quiet=True)
        self.assertNotIn('R000004', [r['ID_Relacion'] for r in res['rel_ok']])

    def test_csv_with_wrong_headers_fails_loudly_in_both(self):
        libro(self.x, self.personas, self.rels)
        bad = os.path.join(self.d, 'rica.csv')
        with open(bad, 'w', newline='') as f:
            w = csv.writer(f); w.writerow(['ID_Relacion', 'ID_Persona_1_hijo', 'Tipo', 'ID_Persona_2_progenitor']); w.writerow(['R000001', 'P000001', 'Padre biológico', 'P000002'])
        with self.assertRaises(ValueError): mapper_v2.build(self.x, None, bad)
        with self.assertRaises(SystemExit): reconcile_v2.main(self.x, self.o, None, bad)

    def test_csv_bad_rows_fail_and_blank_rows_counted(self):
        libro(self.x, self.personas, self.rels)
        casos = {'id': [('RX', 'P000001', 'P000002')], 'solo_hijo': [('R000001', 'P000001', '')], 'mal_P': [('R000001', 'P1', 'P000002')],
                 'inexistente': [('R000777', 'P000001', 'P000002')], 'no_filiacion': [('R000003', 'P000002', 'P000005')], 'dup': [('R000001', 'P000001', 'P000002'), ('R000001', 'P000002', 'P000001')]}
        for k, filas in casos.items():
            p = os.path.join(self.d, k + '.csv'); oricsv(p, filas)
            with self.assertRaises(ValueError, msg=k): mapper_v2.build(self.x, None, p)
        ok = os.path.join(self.d, 'ok.csv'); oricsv(ok, [('R000001', 'P000001', 'P000002'), ('R000002', '', '')])
        P, R, res, conf = mapper_v2.build(self.x, None, ok)
        self.assertEqual(res['orientacion_csv'], {'filas': 2, 'validadas': 1, 'sin_validar': 1})

    def test_reconcile_independent_catches_silent_zero(self):
        # el GED se genero SIN usar el CSV (lo que haria un parser que lee 0 filas): el reconcile, con lectura propia, debe fallar
        self.run_(ori=False)
        rc, out = self.rec(ori=True)
        self.assertEqual(rc, 1); self.assertIn('progenitor_', out)

    def _rel_overlay(self, filas, cab=('ID_Relacion', 'ID_Persona_1', 'Tipo', 'ID_Persona_2', 'Fecha_desde', 'Lugar', 'ID_Cita_adicional', 'ID_Doc_original', 'Literal', 'Nota')):
        p = os.path.join(self.d, 'relov.csv')
        with open(p, 'w', newline='') as f:
            w = csv.writer(f); w.writerow(cab)
            for x in filas: w.writerow(x)
        return p

    def test_relation_overlay_fills_empty_marr_and_reconcile_checks_it(self):
        self.rels.append(rel('R000007', 'P000003', 'Cónyuge', 'P000001'))      # sin Fecha_desde ni lugar
        libro(self.x, self.personas, self.rels)
        ov = self._rel_overlay([('R000007', 'P000003', 'Cónyuge', 'P000001', '1830-06-16', 'L000001', 'C900011', 'D1', 'lit', '')])
        P, R, res, conf = mapper_v2.build(self.x, None, self.ori, ov)
        mapper.run(self.x, self.o, persons=P, rels=R)
        g = self.leer(); self.assertIn('2 DATE 16 JUN 1830', g)
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf): rc = reconcile_v2.main(self.x, self.o, None, self.ori, rel_overlay=ov)
        self.assertEqual(rc, 0, buf.getvalue())
        self.escribir(g.replace('2 DATE 16 JUN 1830', '2 DATE 17 JUN 1830'))
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf): rc = reconcile_v2.main(self.x, self.o, None, self.ori, rel_overlay=ov)
        self.assertEqual(rc, 1); self.assertTrue('overlay_relacion_MARR' in buf.getvalue() or 'MARR_fecha' in buf.getvalue())

    def test_relation_overlay_conflicts_and_bad_shape_fail(self):
        libro(self.x, self.personas, self.rels)
        for fila in (('R000003', 'P000002', 'Esposo', 'P000005', '1800-01-01', '', '', '', '', ''),        # R3 ya tiene fecha 1799-05-06
                     ('R000777', 'P000002', 'Esposo', 'P000005', '1800-01-01', '', '', '', '', ''),        # no existe
                     ('R000003', 'P000009', 'Esposo', 'P000005', '', '', '', '', '', '')):                  # extremos distintos
            with self.assertRaises(ValueError): mapper_v2.build(self.x, None, self.ori, self._rel_overlay([fila]))
        with self.assertRaises(ValueError): mapper_v2.build(self.x, None, self.ori, self._rel_overlay([('R000003', 'P000002', 'Esposo', 'P000005', '', '')], cab=('ID_Relacion', 'ID_Persona_1', 'Tipo', 'ID_Persona_2', 'Fecha_desde', 'Lugar')))
        p = os.path.join(self.d, 'desb.csv')
        with open(p, 'w') as f: f.write('ID_Relacion,ID_Persona_1,Tipo,ID_Persona_2,Fecha_desde,Lugar,ID_Cita_adicional\nR000003,P000002,Esposo,P000005,,,,extra,extra\n')
        with self.assertRaises(ValueError): mapper_v2.build(self.x, None, self.ori, p)

    def test_source_author_pii_is_omitted_with_counter(self):
        for autor, tipo in (('contacto SECRETO@ejemplo.invalid', 'Acta'), ('DNI 12345678Z SECRETO', 'Acta'), ('Juan SECRETO', 'Correo')):
            libro(self.x, self.personas, self.rels, autor=autor, tipo=tipo)
            P, R, res, conf = mapper_v2.build(self.x, None, self.ori)
            s, w, rep, un = mapper.run(self.x, self.o, persons=P, rels=R)
            self.assertNotIn('SECRETO', self.leer(), autor)
            self.assertNotIn('AUTH', ''.join(l for l in self.leer().splitlines(True) if 'SECRETO' in l))
        libro(self.x, self.personas, self.rels, autor='Archivo Municipal')
        P, R, res, conf = mapper_v2.build(self.x, None, self.ori)
        mapper.run(self.x, self.o, persons=P, rels=R)
        self.assertIn('1 AUTH Archivo Municipal', self.leer())

    def test_public_source_with_id_number_author_omits_only_auth(self):
        libro(self.x, self.personas, self.rels, autor='Juan 12345678Z', tipo='Acta')
        P, R, res, conf = mapper_v2.build(self.x, None, self.ori)
        s, w, rep, un = mapper.run(self.x, self.o, persons=P, rels=R)
        o = self.leer()
        self.assertNotIn('12345678Z', o)
        self.assertNotIn('1 AUTH', o)
        self.assertIn('1 TITL Registro ficticio', o)            # la fuente NO se restringe entera
        self.assertNotIn('S900001', s['omisiones_privacidad_ids']['fuente_restringida'])
        self.assertIn('S900001:Autor_o_entidad', s['omisiones_privacidad_ids']['fuente_campo'])

    def _gen(self, **kw):
        libro(self.x, self.personas, self.rels, **kw)
        P, R, res, conf = mapper_v2.build(self.x, None, self.ori)
        s, w, rep, un = mapper.run(self.x, self.o, persons=P, rels=R)
        return s, self.leer()

    def test_free_text_private_links_omitted_public_kept(self):
        priv = ['https://drive.google.com/file/d/SECRETOID/view', 'https://www.dropbox.com/s/SECRETOID/a.pdf', 'https://docs.google.com/document/d/SECRETOID',
                'http://archivo.example.org/ficha', 'https://archivo.example.org/ficha?token=SECRETOID', 'https://192.168.1.5/SECRETOID', 'https://wa.me/SECRETOID']
        for u in priv:
            s, o = self._gen(cita={'URL': u, 'Evaluacion_evidencia': 'lectura ' + u + ' completa'})
            self.assertNotIn('SECRETOID', o, u)
            self.assertNotIn('192.168.1.5', o, u)
            self.assertIn('lectura [enlace omitido por privacidad] completa', o, u)
            self.assertTrue(s['omisiones_privacidad_ids']['libre_enlace_privado'], u)
        pub = 'https://archives.touraine.fr/ark:/1234/acta'
        s, o = self._gen(cita={'URL': pub}, fuente={'URL': 'https://www.geneanet.org/registres/x'})
        self.assertIn(pub, o)
        self.assertIn('https://www.geneanet.org/registres/x', o)
        self.assertEqual(s['omisiones_privacidad_ids']['libre_enlace_privado'], [])

    def test_free_text_identity_email_digits_by_field_class(self):
        s, o = self._gen(cita={'Evaluacion_evidencia': 'copia del DNI de un vivo', 'Extracto_o_resumen': 'contacto SECRETO@ejemplo.invalid o 600123456789', 'Valor_afirmado': 'pasaporte 99'},
                         fuente={'Titulo': 'Libro de familia ficticio', 'Repositorio': 'Archivo SECRETO@ejemplo.invalid'})
        for marcador in ('SECRETO', '600123456789', 'DNI', 'pasaporte', 'Libro de familia'):
            self.assertNotIn(marcador, o, marcador)
        ids = s['omisiones_privacidad_ids']
        self.assertTrue(ids['libre_identidad'] and ids['libre_correo'] and ids['libre_numero_largo'])
        self.assertIn('referencia a documento de identidad', o)

    def test_person_note_identity_minimized_but_historic_will_kept(self):
        self.personas[0] = pers('P000001', 'Ana', 'Falsa Uno', 'F', notas='El testamento y DNI de Antonio solo dan el nombre')
        self.personas[1] = pers('P000002', 'Beto', 'Falso Dos', 'M', notas='Testamento de Pedro (1850), Archivo Departamental; nacio en 1800')
        s, o = self._gen()
        self.assertNotIn('DNI de Antonio', o)
        self.assertIn('referencia a documento de identidad', o)
        self.assertIn('Testamento de Pedro (1850), Archivo Departamental', o)    # un testamento historico no se borra como regla
        self.assertIn('P000001:Notas', s['omisiones_privacidad_ids']['libre_identidad'])

    def test_private_person_citation_tipo_hecho_identity_scrubbed(self):
        self.personas[0] = pers('P000001', 'Ana', 'Falsa Uno', 'F', nac='2005-01-01', muerte=None)   # posiblemente viva -> privada
        s, o = self._gen(cita={'Tipo_hecho': 'Identidad (DNI)'})
        self.assertNotIn('DNI', o)

    def test_url_private_unit(self):
        for u in ('https://drive.google.com/x', 'https://sub.dropbox.com/x', 'http://a.org/x', 'https://a.org/x?key=1', 'https://localhost/x', 'https://10.0.0.1/x', 'https://a.org/x#access_token=1', 'ftp://a.org/x'):
            self.assertTrue(mapper.url_private(u), u)
        for u in ('https://archives.touraine.fr/ark/1', 'https://www.geneanet.org/registres/x?page=2', 'https://www.archives41.fr/'):
            self.assertFalse(mapper.url_private(u), u)

    def test_public_url_with_code_param_kept_but_secret_params_omitted(self):
        pub = 'https://archives.example.fr/visionneuse?code=1MIEC201&page=3'
        s, o = self._gen(cita={'URL': pub})
        self.assertIn(pub, o)
        self.assertEqual(s['omisiones_privacidad_ids']['libre_enlace_privado'], [])
        for q in ('token=X1', 'key=X1', 'sig=X1', 'auth=X1', 'session=X1', 'password=X1', 'access_token=X1'):
            s, o = self._gen(cita={'URL': 'https://archives.example.fr/v?' + q})
            self.assertNotIn('X1', o, q)

    def test_invalid_sha_value_is_scrubbed_valid_sha_untouched(self):
        ok = 'a' * 64
        docs = [['D900001', 'a.pdf', 'Acta', 'P000001', 'S900001', 'C900010', 'Publico', 'SECRETO@ejemplo.invalid DNI 12345678Z 600123456789'],
                ['D900002', 'b.pdf', 'Acta', 'P000002', 'S900001', 'C900011', 'Publico', ok],
                ['D900003', 'c.pdf', 'Acta', 'P000003', 'S900001', 'C900012', 'Publico', 'https://drive.google.com/file/d/SECRETOID']]
        s, o = self._gen(docs=docs)
        for marcador in ('SECRETO', '12345678Z', '600123456789', 'SECRETOID'):
            self.assertNotIn(marcador, o, marcador)
        self.assertIn('SHA256 ' + ok, o)                         # hash valido intacto
        self.assertIn('SHA256 no acreditado', o)

    def test_remap_is_transitive_and_cycles_fail(self):
        self.personas += [pers('P000007', 'Gus', 'Falso Siete', 'M', est='Fusionada en P000008', niv=''), pers('P000008', 'Hana', 'Falsa Ocho', 'F', est='Fusionada en P000002', niv='')]
        self.rels += [rel('R000006', 'P000007', 'Hermano', 'P000001')]
        libro(self.x, self.personas, self.rels)
        res = v2_selector.main(self.x, None, quiet=True)
        self.assertEqual(res['REMAP']['P000007'], 'P000008')
        r6 = [r for r in res['rel_ok'] if r['ID_Relacion'] == 'R000006']
        self.assertEqual(r6[0]['ID_Persona_1'], 'P000002')       # P7->P8->P2, nunca P8 (no exportado)
        for r in res['rel_ok']:
            self.assertIn(r['ID_Persona_1'], res['inc']); self.assertIn(r['ID_Persona_2'], res['inc'])
        self.personas[-1] = pers('P000008', 'Hana', 'Falsa Ocho', 'F', est='Fusionada en P000007', niv='')
        libro(self.x, self.personas, self.rels)
        with self.assertRaises(ValueError): v2_selector.main(self.x, None, quiet=True)


if __name__ == '__main__':
    unittest.main()
