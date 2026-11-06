"""Puertas CLI de extremo a extremo: los puntos de entrada reales por linea
de comandos (parseo de argumentos y codigos de salida), que los tests por
modulo no ejercitan. check_evidencia es el gate de privacidad: un GED
manipulado debe dar salida 1."""
import os, sys, subprocess, tempfile, unittest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
ROOT = os.path.join(os.path.dirname(__file__), '..')
from test_v2 import pers, rel, libro, oricsv, H02
from test_evidencia import fila, evcsv
import mapper, mapper_v2


def rel_n(*a, notas=None, **k):
    r = rel(*a, **k); r[H02.index('Notas')] = notas; return r


def cli(*args):
    return subprocess.run([sys.executable] + list(args), capture_output=True, text=True)


def gen_v2(d):
    """Mismo montaje que test_evidencia: familia ficticia, 4 aristas, pipeline mapper_v2 -> mapper."""
    x = os.path.join(d, 'f.xlsx'); o = os.path.join(d, 'o.ged')
    ori = os.path.join(d, 'ori.csv'); ev = os.path.join(d, 'ev.csv')
    personas = [pers('P000001', 'Ana', 'Falsa Uno', 'F'), pers('P000002', 'Beto', 'Falso Dos', 'M'),
                pers('P000003', 'Cira', 'Falsa Tres', 'F'), pers('P000004', 'Dani', 'Falso Cuatro', 'M')]
    rels = [rel('R000001', 'P000001', 'Padre biológico', 'P000002'), rel('R000002', 'P000001', 'Madre biológica', 'P000003'),
            rel('R000003', 'P000004', 'Padre biológico', 'P000002'),
            rel_n('R000004', 'P000004', 'Madre biológica', 'P000003', conf='Media', notas='apellido Zorro? dudoso')]
    filas = [fila('R000001', 'P000001', 'Padre biológico', 'P000002'), fila('R000002', 'P000001', 'Madre biológica', 'P000003'),
             fila('R000003', 'P000004', 'Padre biológico', 'P000002', lit='acta n.2: «Dani Falso hijo de Beto Falso»'),
             fila('R000004', 'P000004', 'Madre biológica', 'P000003', lit='acta n.2: «y de Cira Falsa(?)»', salv='Lectura incierta (?) en el literal; no endurecer')]
    libro(x, personas, rels); evcsv(ev, filas)
    oricsv(ori, [('R000001', 'P000001', 'P000002'), ('R000002', 'P000001', 'P000003'),
                 ('R000003', 'P000004', 'P000002'), ('R000004', 'P000004', 'P000003')])
    P, R, res, conf = mapper_v2.build(x, None, ori, None, ev)
    mapper.run(x, o, persons=P, rels=R, evidencia=res['evidencia'])
    return x, o, ev


class CliEvidencia(unittest.TestCase):

    def setUp(self):
        self.d = tempfile.mkdtemp()
        self.x, self.o, self.ev = gen_v2(self.d)

    def test_full_mode_clean_ged_exits_0(self):
        r = cli(os.path.join(ROOT, 'check_evidencia.py'), self.ev, self.o, '--excel=' + self.x)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)

    def test_full_mode_tampered_literal_exits_1(self):
        t = open(self.o, encoding='utf-8').read().replace('Fulano Falso hijo', 'Fulano Falsoo hijo', 1)
        open(self.o, 'w', encoding='utf-8').write(t)
        r = cli(os.path.join(ROOT, 'check_evidencia.py'), self.ev, self.o, '--excel=' + self.x)
        self.assertEqual(r.returncode, 1, r.stdout + r.stderr)

    def test_solo_ged_mode_clean_exits_0_and_tamper_exits_1(self):
        p = [os.path.join(ROOT, 'check_evidencia.py'), self.ev, self.o, '--solo-ged']
        self.assertEqual(cli(*p).returncode, 0)
        t = open(self.o, encoding='utf-8').read().replace('Fulano Falso hijo', 'Fulano Falsoo hijo', 1)
        open(self.o, 'w', encoding='utf-8').write(t)
        self.assertEqual(cli(*p).returncode, 1)


class CliReconcile(unittest.TestCase):
    """reconcile.py y reconcile_v2.py como CLI sobre fixtures generados."""

    def test_reconcile_clean_fixture_reports_only_known_exclusion(self):
        d = tempfile.mkdtemp()
        subprocess.check_call([sys.executable, os.path.join(ROOT, 'make_fixture.py')], cwd=d)
        x = os.path.join(d, 'fixtures', 'fixture_sintetico.xlsx'); o = os.path.join(d, 'o.ged')
        mapper.run(x, o)
        r = cli(os.path.join(ROOT, 'reconcile.py'), x, o)
        # el fixture produce una unica diferencia conocida (cita excluida por el mapper): salida 1 con solo esa linea
        self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
        lineas = [l for l in r.stdout.splitlines() if l.startswith('   ')]
        self.assertTrue(lineas and all('cita_no_emitida_o_excluida' in l for l in lineas), r.stdout)

    def test_reconcile_tampered_ged_exits_1_with_real_diff(self):
        d = tempfile.mkdtemp()
        subprocess.check_call([sys.executable, os.path.join(ROOT, 'make_fixture.py')], cwd=d)
        x = os.path.join(d, 'fixtures', 'fixture_sintetico.xlsx'); o = os.path.join(d, 'o.ged')
        mapper.run(x, o)
        t = open(o, encoding='utf-8').read().replace('1 SEX F', '1 SEX M', 1)
        open(o, 'w', encoding='utf-8').write(t)
        r = cli(os.path.join(ROOT, 'reconcile.py'), x, o)
        self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
        self.assertIn('sexo', r.stdout)

    def test_reconcile_v2_clean_exits_0(self):
        d = tempfile.mkdtemp()
        x, o, ev = gen_v2(d)
        ori = os.path.join(d, 'ori.csv')
        r = cli(os.path.join(ROOT, 'reconcile_v2.py'), x, o, '-', ori)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)


if __name__ == '__main__':
    unittest.main()
