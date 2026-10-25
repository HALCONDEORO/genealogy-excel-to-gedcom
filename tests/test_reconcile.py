import os, sys, tempfile, unittest, io, contextlib
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
import mapper, reconcile, openpyxl


class R(unittest.TestCase):
    def _run(self, mutate=None):
        d = tempfile.mkdtemp()
        import subprocess
        subprocess.check_call([sys.executable, os.path.join(os.path.dirname(__file__), '..', 'make_fixture.py')], cwd=d)
        p = os.path.join(d, 'fixtures', 'fixture_sintetico.xlsx')
        o = os.path.join(d, 'o.ged')
        mapper.run(p, o)
        if mutate:
            t = open(o, encoding='utf-8').read(); open(o, 'w', encoding='utf-8').write(mutate(t))
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            rc = reconcile.main(p, o)
        return rc, buf.getvalue()

    def test_fixture_only_known_exclusion_is_reported(self):
        rc, out = self._run()
        lineas = [l for l in out.splitlines() if l.startswith('   ')]
        self.assertTrue(all('cita_no_emitida_o_excluida' in l for l in lineas), out)   # la unica diferencia es una cita excluida por el mapper, listada con ID
        self.assertEqual(len(lineas), 1, out)

    def test_detects_changed_sex_place_parent_and_citation_source(self):
        rc, out = self._run(lambda t: t.replace('1 SEX F', '1 SEX M', 1).replace('2 PLAC Villa Falsa', '2 PLAC Otro', 1).replace('1 FAMC @F1@\n', '', 1).replace('1 SOUR @S000001@', '1 SOUR @S000777@', 1))
        self.assertEqual(rc, 1)
        for k in ('sexo', 'lugar_nacimiento', 'progenitor_', 'cita_fuente'):
            self.assertIn(k, out)

    def test_detects_full_date_qualifier_and_document_note_changes(self):
        rc, out = self._run(lambda t: t.replace('2 DATE 2 MAY 1980', '2 DATE 3 MAY 1980', 1).replace('tipo omitido por privacidad (ver Excel)', 'tipo DNI', 1).replace('| fuente S000001', '| fuente S999999', 1))
        self.assertEqual(rc, 1)
        for k in ('fecha_nacimiento', 'documento_sin_marcador_de_privacidad', 'documento_fuente'):
            self.assertIn(k, out)

    def test_fixture_documents_are_checked(self):
        rc, out = self._run()
        self.assertIn('documentos 06 de personas de hoja 30: 1', out)


if __name__ == '__main__':
    unittest.main()
