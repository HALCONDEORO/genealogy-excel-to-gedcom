#!/usr/bin/env python3
"""Tests unitarios OFFLINE de priv_derive.py (solo stdlib, sin API).
Carga las funciones del script sin ejecutar main(): se lee el fuente del script
(adjunto en la raiz del repo) y se neutraliza la llamada final a main().
"""
import os, sys, tempfile, unittest

SRC = os.path.join(os.path.dirname(__file__), '..', 'priv_derive.py')

def load_module():
    src = open(SRC, encoding='utf-8').read()
    assert src.rstrip().endswith('main()'), 'el script debe terminar llamando a main()'
    ns = {'__name__': 'priv_derive_test'}
    exec(src.rstrip()[:-len('main()')], ns)
    return ns

GED_SINTETICO = '''0 HEAD
0 @P1@ INDI
1 NAME Ana /Falsa/
1 SEX F
1 RESN privacy
1 BIRT
2 DATE 1 JAN 1950
2 NOTE nota de nacimiento
1 NOTE nota personal
1 SOUR @S1@
2 PAGE Declaracion directa | C000001 | Identidad
2 NOTE nota de la cita
0 @P2@ INDI
1 NAME Beto /Falso/
1 SEX M
0 @F1@ FAM
1 HUSB @P1@
1 WIFE @P2@
1 NOTE evidencia familia
1 MARR
2 DATE 1887
2 NOTE Fecha original en el Excel: c.1887
0 @F2@ FAM
1 HUSB @P2@
1 MARR
2 DATE 1900
0 @S1@ SOUR
1 TITL Fuente Falsa
0 TRLR
'''

class TestParse(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ns = load_module()
        with tempfile.NamedTemporaryFile('w', suffix='.ged', delete=False) as f:
            f.write(GED_SINTETICO); cls.ged_path = f.name

    def test_resn_y_conteos(self):
        persons, fams, sources = self.ns['parse_ged'](self.ged_path)
        exp = self.ns['ged_expected'](persons, fams)
        self.assertEqual(exp['people'], {'1'})
        self.assertEqual(len(exp['events']), 2)          # BIRT de P1 + MARR de F1
        self.assertEqual(exp['families'], {'1'})          # F2 no tiene miembro RESN
        self.assertEqual(exp['citations'], ['C000001'])   # lista: preserva multiplicidad
        self.assertEqual(exp['sources'], {'1'})

    def test_regresion_nota_marr(self):
        """La nota nivel 2 bajo el MARR de la familia RESN debe capturarse."""
        persons, fams, _ = self.ns['parse_ged'](self.ged_path)
        exp = self.ns['ged_expected'](persons, fams)
        textos = set(exp['notes'])
        self.assertEqual(len(exp['notes']), 5)   # persona + nacimiento + cita + familia + MARR
        for esperada in ('Fecha original en el Excel: c.1887',   # regresion real
                         'nota de nacimiento', 'nota personal', 'evidencia familia', 'nota de la cita'):
            self.assertIn(esperada, textos)

    def test_norm_text(self):
        self.assertEqual(self.ns['norm_text']('  a\n  b  '), 'a b')

    def test_page_con_conc(self):
        # Regresion real del piloto: PAGE con continuacion CONC debe unirse
        # (el importador Gramps concatena CONC); si no, la comparacion API==GED
        # falla en falso en citas con pagina multilinea.
        ged = GED_SINTETICO.replace(
            '2 PAGE Declaracion directa | C000001 | Identidad',
            '2 PAGE Declaracion directa | C000001 | Identidad, parte uno:\n3 CONC  parte dos')
        with tempfile.NamedTemporaryFile('w', suffix='.ged', delete=False) as f:
            f.write(ged); path = f.name
        persons, fams, _ = self.ns['parse_ged'](path)
        exp = self.ns['ged_expected'](persons, fams)
        gpage, _ = exp['cit_ged']['C000001']
        self.assertEqual(gpage, 'Declaracion directa | C000001 | Identidad, parte uno: parte dos')

if __name__ == '__main__':
    unittest.main(verbosity=2)
