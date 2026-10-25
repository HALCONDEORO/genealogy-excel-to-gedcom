#!/usr/bin/env python3
"""Tests del emisor expected_census de priv_derive v18 (modo --emit-expected-census).
Mockea get_token/get_all: verifica estructura {tipo: {handle: true}}, conteos, sha256
estable y fail-closed heredado de la guarda get_all (SystemExit si la API falla)."""
import json, sys, os, hashlib, importlib.util, tempfile
from unittest import mock

_HERE = os.path.dirname(os.path.abspath(__file__))
_CANDIDATOS = [os.path.join(_HERE, 'priv_derive_v18_emisor.py'),
               os.path.join(os.path.dirname(_HERE), 'priv_derive_v18_emisor.py')]  # plano o tests/ bajo raiz repo
_EMISOR = next((p for p in _CANDIDATOS if os.path.exists(p)), _CANDIDATOS[-1])
spec = importlib.util.spec_from_file_location('pd18', _EMISOR)
pd = importlib.util.module_from_spec(spec)
sys.argv = ['pd']  # v18 (derivado de v17) ejecuta main() al importar (sin guarda __main__); argv corto -> IndexError controlado
try:
    spec.loader.exec_module(pd)
except IndexError:
    pass

FAKE = {
    'people': [{'handle': 'h1'}, {'handle': 'h2'}],
    'families': [{'handle': 'f1'}],
    'events': [], 'notes': [{'handle': 'n1'}, {'handle': 'n2'}, {'handle': 'n3'}],
    'citations': [], 'sources': [{'handle': 's1'}],
}

def run_emit(tmp):
    out = os.path.join(tmp, 'ec.json')
    with mock.patch.object(pd, 'get_token', return_value='t'), \
         mock.patch.object(pd, 'get_all', side_effect=lambda base, kind, tok: FAKE[kind]), \
         mock.patch.dict(os.environ, {'GRAMPS_USER': 'u', 'GRAMPS_PASS': 'p'}):
        sys.argv = ['pd', '--emit-expected-census', 'http://x', out]
        pd.main()
    return out

tmp = tempfile.mkdtemp()
out = run_emit(tmp)
data = open(out, 'rb').read()
j = json.loads(data)
assert set(j) == {'people','families','events','notes','citations','sources'}, 'tipos'
assert j['people'] == {'h1': True, 'h2': True} and j['notes'] == {'n1': True, 'n2': True, 'n3': True}
assert j['events'] == {} and j['citations'] == {}, 'vacios presentes como dict'
sha1 = hashlib.sha256(data).hexdigest()
# determinismo: segunda ejecucion, sha identico
out2 = run_emit(tmp)
assert hashlib.sha256(open(out2, 'rb').read()).hexdigest() == sha1, 'sha determinista'
# fail-closed: get_all aborta (guarda pagesize) => main propaga SystemExit
with mock.patch.object(pd, 'get_token', return_value='t'), \
     mock.patch.object(pd, 'get_all', side_effect=SystemExit('FAIL_CLOSED guarda')), \
     mock.patch.dict(os.environ, {'GRAMPS_USER': 'u', 'GRAMPS_PASS': 'p'}):
    sys.argv = ['pd', '--emit-expected-census', 'http://x', os.path.join(tmp, 'ec2.json')]
    try:
        pd.main(); raise AssertionError('debio abortar')
    except SystemExit:
        pass
    assert not os.path.exists(os.path.join(tmp, 'ec2.json')), 'no fichero parcial'
print('3/3 tests emisor OK; sha fija del fixture:', sha1)
