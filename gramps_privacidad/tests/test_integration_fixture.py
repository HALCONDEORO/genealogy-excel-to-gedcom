#!/usr/bin/env python3
"""Test de integracion: exige un stack Gramps Web desechable (docker compose oficial,
Gramps Web 26.x) con fixtures/adv.ged accesible y credenciales de prueba por entorno:
  GRAMPS_TEST_BASE (p.ej. http://localhost:5000), GRAMPS_USER, GRAMPS_PASS
Se salta limpiamente si falta GRAMPS_TEST_BASE. NUNCA apuntar a datos reales.
"""
import hashlib, json, os, subprocess, sys, unittest, urllib.request

BASE = os.environ.get('GRAMPS_TEST_BASE')
ROOT = os.path.join(os.path.dirname(__file__), '..')

@unittest.skipUnless(BASE, 'sin GRAMPS_TEST_BASE: test de integracion omitido')
def run_tool(script, stdin=None):
    ged = os.environ.get('GED_FIXTURE', os.path.join(ROOT, 'fixtures', 'adv.ged'))
    return subprocess.run([sys.executable, os.path.join(ROOT, script)] + ([ged, BASE] if script == 'priv_derive.py' else []),
                          input=stdin, capture_output=True, text=True, timeout=300)


class TestIntegration(unittest.TestCase):
    def test_derive_remediate_idempotente_y_negativas(self):
        d = run_tool('priv_derive.py')
        self.assertEqual(d.returncode, 0, d.stderr[:300])
        j = json.loads(d.stdout)
        self.assertEqual(j['status'], 'OK')
        self.assertEqual(j['counts'], {'people': 1, 'events': 1, 'families': 1, 'notes': 6, 'citations': 1, 'sources': 1})
        self.assertEqual(j['total'], 11)
        cfg = json.dumps({'base_url': BASE, 'targets': j['targets'], 'sha256': j['sha256']})
        r1 = run_tool('priv_remediate.py', cfg)
        self.assertEqual(r1.returncode, 0, r1.stderr[-300:])
        r2 = run_tool('priv_remediate.py', cfg)
        self.assertEqual(r2.returncode, 0)
        self.assertEqual(json.loads(r2.stdout)['flagged'], 0)          # idempotente
        bad = json.loads(cfg)
        i = next(k for k, t in enumerate(bad['targets']) if t['type'] == 'citations')
        bad['targets'][i]['fingerprint'] = 'CORRUPTA'
        r3 = run_tool('priv_remediate.py', json.dumps(bad))
        self.assertEqual(r3.returncode, 1)                              # fail-closed
        self.assertIn('fingerprint_mismatch', r3.stderr)
        # handle incoherente: recomputar sha para aislar la via handle_mismatch
        bad2 = json.loads(cfg); bad2['targets'][0]['handle'] = 'deadbeef' * 4
        by = {}
        for t in bad2['targets']: by.setdefault(t['type'], []).append(t['handle'])
        bad2['sha256'] = {k: hashlib.sha256(chr(10).join(sorted(v)).encode()).hexdigest() for k, v in by.items()}
        r4 = run_tool('priv_remediate.py', json.dumps(bad2))
        self.assertEqual(r4.returncode, 1)
        self.assertIn('handle_mismatch', r4.stderr)
        # target ajeno al conjunto: el manifiesto (sha viejo) ya no cuadra
        bad3 = json.loads(cfg)
        bad3['targets'].append({'type': 'people', 'gramps_id': 'P000001', 'handle': 'f' * 32, 'content_sha256': '0' * 64})
        r5 = run_tool('priv_remediate.py', json.dumps(bad3))
        self.assertEqual(r5.returncode, 1)
        self.assertIn('manifest_sha256_mismatch', r5.stderr)
        # contenido alterado respecto a la derivación (manifiesto recomputado para aislar la via)
        bad4 = json.loads(cfg); bad4['targets'][0]['content_sha256'] = '0' * 64
        by = {}
        for t in bad4['targets']: by.setdefault(t['type'], []).append(t['handle'])
        bad4['sha256'] = {k: hashlib.sha256(chr(10).join(sorted(v)).encode()).hexdigest() for k, v in by.items()}
        r6 = run_tool('priv_remediate.py', json.dumps(bad4))
        self.assertEqual(r6.returncode, 1)
        self.assertIn('content_mismatch', r6.stderr)
        # type_swap: mover un target a otro tipo (manifiesto recomputado) -> lookup falla fail-closed
        bad5 = json.loads(cfg)
        t0 = bad5['targets'][0]; t0['type'] = 'events' if t0['type'] != 'events' else 'people'
        by = {}
        for t in bad5['targets']: by.setdefault(t['type'], []).append(t['handle'])
        bad5['sha256'] = {k: hashlib.sha256(chr(10).join(sorted(v)).encode()).hexdigest() for k, v in by.items()}
        r7 = run_tool('priv_remediate.py', json.dumps(bad5))
        self.assertEqual(r7.returncode, 1)
        # duplicado: mismo tipo+handle repetido (manifiesto recomputado) -> target_duplicado
        bad6 = json.loads(cfg); bad6['targets'].append(dict(bad6['targets'][0]))
        by = {}
        for t in bad6['targets']: by.setdefault(t['type'], []).append(t['handle'])
        bad6['sha256'] = {k: hashlib.sha256(chr(10).join(sorted(v)).encode()).hexdigest() for k, v in by.items()}
        r8 = run_tool('priv_remediate.py', json.dumps(bad6))
        self.assertEqual(r8.returncode, 1)
        self.assertIn('target_duplicado', r8.stderr)



@unittest.skipUnless(BASE, 'sin GRAMPS_TEST_BASE: test de integracion omitido')
class TestGemelosSwap(unittest.TestCase):
    """Regresion: eventos gemelos (mismo tipo/fecha/lugar) de personas distintas.
    P000007 (RESN) y P000008 (no RESN) comparten BIRT en fecha y lugar; el evento
    del gemelo lleva una nota distinta. Intercambiar la ref real via API debe
    romper el binding (propietario + notas) y abortar fail-closed; con el binding
    correcto la derivacion pasa."""

    def api(self, method, path, payload=None, token=None):
        data = json.dumps(payload).encode() if payload is not None else None
        req = urllib.request.Request(BASE + path, data=data, method=method)
        if data: req.add_header('Content-Type', 'application/json')
        if token: req.add_header('Authorization', 'Bearer ' + token)
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read())

    def token_con_reintentos(self):
        import time, urllib.error
        for _ in range(10):
            try:
                return self.api('POST', '/api/token/', {'username': os.environ['GRAMPS_USER'], 'password': os.environ['GRAMPS_PASS']})['access_token']
            except urllib.error.HTTPError as e:
                if e.code == 429: time.sleep(5); continue
                raise
        raise SystemExit('token: 429 persistente')

    def test_swap_refs_gemelos(self):
        tok = self.token_con_reintentos()
        p7 = self.api('GET', '/api/people/?gramps_id=P000007', token=tok)[0]
        p8 = self.api('GET', '/api/people/?gramps_id=P000008', token=tok)[0]
        h7 = p7['event_ref_list'][0]['ref']
        h8 = p8['event_ref_list'][0]['ref']
        self.assertNotEqual(h7, h8)
        swapped = dict(p7)
        swapped['event_ref_list'] = [dict(p7['event_ref_list'][0], ref=h8)]
        st = self.api('PUT', '/api/people/%s' % p7['handle'], swapped, token=tok)
        try:
            d = run_tool('priv_derive.py')
            self.assertEqual(d.returncode, 1, 'el swap de gemelos NO fue detectado')
            self.assertIn('eventos', d.stdout)
        finally:
            restored = dict(p7)
            restored['event_ref_list'] = [dict(p7['event_ref_list'][0], ref=h7)]
            self.api('PUT', '/api/people/%s' % p7['handle'], restored, token=tok)
        d2 = run_tool('priv_derive.py')
        self.assertEqual(d2.returncode, 0, d2.stderr[:300])

if __name__ == '__main__':
    unittest.main(verbosity=2)


class TestCrossFrontera(unittest.TestCase):
    """v15: gemelos cross-frontera SIN notas (2 RESN + 1 pub, BIRT identicos) -> privatizacion
    conservadora de TODOS los miembros del grupo. Fixture adv2_cross.ged via GED_FIXTURE."""
    def test_privatize_all_grupo_cruzando(self):
        if not BASE: self.skipTest('sin GRAMPS_TEST_BASE')
        if 'adv2_cross' not in os.environ.get('GED_FIXTURE', ''): self.skipTest('requiere fixture adv2_cross (GED_FIXTURE)')
        d = run_tool('priv_derive.py')
        self.assertEqual(d.returncode, 0, d.stderr[-400:] + d.stdout[:400])
        j = json.loads(d.stdout)
        self.assertEqual(j['status'], 'OK')
        self.assertEqual(j['counts'].get('people'), 2)
        self.assertEqual(j['counts'].get('events'), 3, 'el gemelo pub debe entrar por privatize-all')
        self.assertEqual(j['total'], 5)
        amb = j.get('ambiguity', {})
        self.assertEqual(len(amb.get('events_crossing', [])), 1)
        self.assertEqual(amb['events_crossing'][0]['members'], 3)
        origins = [t.get('origin') for t in j['targets'] if t['type'] == 'events']
        self.assertEqual(origins.count('conservative_ambiguity'), 1)
        self.assertEqual(origins.count('ged_resn'), 2)
        cfg = json.dumps({'base_url': BASE, 'targets': j['targets'], 'sha256': j['sha256']})
        # tolerante al orden: el harness puede haber remediado ya en la fase manual
        r1 = run_tool('priv_remediate.py', cfg)
        self.assertEqual(r1.returncode, 0, r1.stderr[-400:])
        o1 = json.loads(r1.stdout)
        self.assertEqual(o1.get('flagged', 0) + o1.get('already', 0), 5)
        r2 = run_tool('priv_remediate.py', cfg)
        self.assertEqual(r2.returncode, 0)
        self.assertIn('"already": 5', r2.stdout)
        tok = token_global()
        ev_priv = {}
        per_priv = {}
        for path, dct in (('/api/events/?pagesize=2000', ev_priv), ('/api/people/?pagesize=2000', per_priv)):
            req = urllib.request.Request(BASE + path); req.add_header('Authorization', 'Bearer ' + tok)
            for o in json.loads(urllib.request.urlopen(req, timeout=60).read()): dct[o['gramps_id']] = o.get('private')
        self.assertTrue(all(ev_priv.values()), 'los 3 eventos gemelos deben quedar private')
        self.assertTrue(per_priv.get('P000001') and per_priv.get('P000002'))
        self.assertFalse(per_priv.get('P000003'), 'la persona pub NO debe privatizarse: solo el objeto ambiguo')

def token_global():
    import time, urllib.error
    for _ in range(10):
        try:
            req = urllib.request.Request(BASE + '/api/token/', data=json.dumps({'username': os.environ['GRAMPS_USER'], 'password': os.environ['GRAMPS_PASS']}).encode(), method='POST')
            req.add_header('Content-Type', 'application/json')
            return json.loads(urllib.request.urlopen(req, timeout=60).read())['access_token']
        except urllib.error.HTTPError as e:
            if e.code == 429: time.sleep(5); continue
            raise
    raise SystemExit('token: 429 persistente')


class TestCrossFamilia(unittest.TestCase):
    """v16 (gate BO): MARR gemelos cross-frontera (familia privada + familia pub, mismo contenido,
    0 dependientes) -> privatize-all: AMBOS MARR quedan private; la familia pub NO se marca."""
    def test_marr_cross_privatize_all(self):
        if not BASE: self.skipTest('sin GRAMPS_TEST_BASE')
        if 'adv4_fam' not in os.environ.get('GED_FIXTURE', ''): self.skipTest('requiere fixture adv4_fam (GED_FIXTURE)')
        d = run_tool('priv_derive.py')
        self.assertEqual(d.returncode, 0, d.stderr[-400:] + d.stdout[:400])
        j = json.loads(d.stdout)
        self.assertEqual(j['status'], 'OK')
        self.assertEqual(j['counts'].get('people'), 1)
        self.assertEqual(j['counts'].get('families'), 1)
        self.assertEqual(j['counts'].get('events'), 2, 'el MARR de la familia pub entra por privatize-all')
        self.assertEqual(j['total'], 4)
        amb = j.get('ambiguity', {})
        self.assertEqual(len(amb.get('events_crossing', [])), 1)
        self.assertEqual(amb['events_crossing'][0]['members'], 2)
        origins = sorted(t.get('origin') for t in j['targets'] if t['type'] == 'events')
        self.assertEqual(origins, ['conservative_ambiguity', 'ged_resn'])
        cfg = json.dumps({'base_url': BASE, 'targets': j['targets'], 'sha256': j['sha256']})
        r1 = run_tool('priv_remediate.py', cfg)
        self.assertEqual(r1.returncode, 0, r1.stderr[-400:])
        o1 = json.loads(r1.stdout)
        self.assertEqual(o1.get('flagged', 0) + o1.get('already', 0), 4)
        r2 = run_tool('priv_remediate.py', cfg)
        self.assertEqual(r2.returncode, 0)
        self.assertIn('"already": 4', r2.stdout)
        tok = token_global()
        def get(path):
            req = urllib.request.Request(BASE + path); req.add_header('Authorization', 'Bearer ' + tok)
            return json.loads(urllib.request.urlopen(req, timeout=60).read())
        evs = get('/api/events/?pagesize=2000')
        self.assertTrue(all(e.get('private') for e in evs if str(e.get('type')).lower() == 'marriage'), 'ambos MARR deben quedar private')
        fams = get('/api/families/?pagesize=2000')
        self.assertEqual(len(fams), 2)
        fam_t = {t['handle'] for t in j['targets'] if t['type'] == 'families'}
        for f in fams:
            if f['handle'] in fam_t:
                self.assertTrue(f.get('private'), 'la familia con miembro RESN debe quedar private')
            else:
                self.assertTrue(f.get('private') is not True, 'la familia pub NO debe marcarse')


class TestCrossFronteraDependientes(unittest.TestCase):
    """v16 (gate BO + Codex hallazgo A): dependientes (citas) en gemelos cross-frontera.
    Parte 1: cita SOLO en el gemelo pub via API -> el arbol vivo DIVERGE del GED canonico:
    la clave extendida (nc) separa al gemelo pub del grupo API-side mientras el GED-side sigue
    viendo el grupo unido -> la comparacion cruzada GED<->API detecta la divergencia y la
    derivacion ABORTA (fail-closed correcto: el GED es la referencia).
    Parte 2: citas en TODOS los gemelos -> los seleccionados tienen citation_list -> FAIL_CLOSED
    antes de cualquier escritura."""
    def api(self, method, path, payload=None, tok=None):
        data = json.dumps(payload).encode() if payload is not None else None
        req = urllib.request.Request(BASE + path, data=data, method=method)
        if data: req.add_header('Content-Type', 'application/json')
        if tok: req.add_header('Authorization', 'Bearer ' + tok)
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read())

    def test_dependientes(self):
        if not BASE: self.skipTest('sin GRAMPS_TEST_BASE')
        if 'adv2_cross' not in os.environ.get('GED_FIXTURE', ''): self.skipTest('requiere fixture adv2_cross (GED_FIXTURE)')
        tok = token_global()
        src = self.api('POST', '/api/sources/', {'title': 'Fuente ficticia dependientes'}, tok)
        if isinstance(src, list): src = src[0]
        people = self.api('GET', '/api/people/?pagesize=2000', tok=tok)
        pub_h = {p['gramps_id']: p['handle'] for p in people}['P000003']
        pub_ev = self.api('GET', '/api/people/%s?profile=all' % pub_h, tok=tok)['event_ref_list'][0]
        pub_ev_h = pub_ev['ref'] if isinstance(pub_ev, dict) else pub_ev
        def add_cit(ev_h, label):
            cit = self.api('POST', '/api/citations/', {'source_handle': src['handle'], 'page': 'f. 9 | %s |' % label}, tok)
            if isinstance(cit, list): cit = cit[0]
            full = self.api('GET', '/api/events/%s' % ev_h, tok=tok)
            full['citation_list'] = list(full.get('citation_list', [])) + [cit['handle']]
            self.api('PUT', '/api/events/%s' % ev_h, full, tok)
        # Parte 1: solo el pub -> divergencia GED<->API -> FAIL_CLOSED
        add_cit(pub_ev_h, 'Cdep1')
        d = run_tool('priv_derive.py')
        self.assertNotEqual(d.returncode, 0, 'una cita API-side no presente en el GED debe ABORTAR (divergencia)')
        self.assertIn('FAIL_CLOSED', d.stdout + d.stderr)
        self.assertIn('mismatch', d.stdout)
        # Parte 2: citas tambien en los privados -> grupo unico con dependientes -> FAIL_CLOSED
        for gid in ('P000001', 'P000002'):
            h = {p['gramps_id']: p['handle'] for p in people}[gid]
            ev = self.api('GET', '/api/people/%s?profile=all' % h, tok=tok)['event_ref_list'][0]
            add_cit(ev['ref'] if isinstance(ev, dict) else ev, 'Cdep-' + gid)
        d2 = run_tool('priv_derive.py')
        self.assertNotEqual(d2.returncode, 0, 'con dependientes en seleccionados debe ABORTAR')
        self.assertIn('FAIL_CLOSED', d2.stdout + d2.stderr)
        self.assertTrue('citation_list' in d2.stdout or 'dependientes' in d2.stdout)
