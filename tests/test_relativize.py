import os, sys, gzip, re, tempfile, unittest
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import gramps_relativize as g


def _write(path, xml):
    with gzip.open(path, 'wb') as fh:
        fh.write(xml.encode('utf-8'))


def _read(path):
    with gzip.open(path, 'rb') as fh:
        return fh.read().decode('utf-8')


class T(unittest.TestCase):
    def test_makes_paths_relative_and_keeps_rest(self):
        d = tempfile.mkdtemp(); base = os.path.join(d, 'lote'); os.makedirs(base)
        xml = '<database><object id="M1"><file src="%s/media/a b.pdf" mime="application/pdf" description="x"/></object><x a="1"/></database>' % base
        _write(os.path.join(d, 'in.gramps'), xml)
        n = g.relativize(os.path.join(d, 'in.gramps'), base, os.path.join(d, 'out.gramps'))
        o = _read(os.path.join(d, 'out.gramps'))
        self.assertEqual(n, 1)
        self.assertIn('<file src="media/a b.pdf"', o)
        self.assertEqual(o.replace('media/a b.pdf', base + '/media/a b.pdf'), xml)

    def test_outside_base_fails_closed_without_output(self):
        d = tempfile.mkdtemp(); base = os.path.join(d, 'lote'); os.makedirs(base)
        _write(os.path.join(d, 'in.gramps'), '<d><file src="/etc/passwd" mime="x"/></d>')
        with self.assertRaises(g.RelError):
            g.relativize(os.path.join(d, 'in.gramps'), base, os.path.join(d, 'out.gramps'))
        self.assertFalse(os.path.exists(os.path.join(d, 'out.gramps')))
        self.assertEqual([f for f in os.listdir(d) if f.startswith('.rel-')], [])

    def test_sibling_prefix_is_outside(self):
        d = tempfile.mkdtemp(); base = os.path.join(d, 'lote'); os.makedirs(base)
        _write(os.path.join(d, 'in.gramps'), '<d><file src="%s2/x.pdf" mime="x"/></d>' % base)
        with self.assertRaises(g.RelError):
            g.relativize(os.path.join(d, 'in.gramps'), base, os.path.join(d, 'out.gramps'))

    def test_relative_paths_are_validated_too(self):
        d = tempfile.mkdtemp(); base = os.path.join(d, 'lote'); os.makedirs(base)
        for src in ('../privado.pdf', 'media/../../privado.pdf', '', '.', 'a/../..'):
            _write(os.path.join(d, 'in.gramps'), '<d><file src="%s" mime="x"/></d>' % src)
            with self.assertRaises(g.RelError, msg=src):
                g.relativize(os.path.join(d, 'in.gramps'), base, os.path.join(d, 'out.gramps'))
            self.assertFalse(os.path.exists(os.path.join(d, 'out.gramps')), src)
        _write(os.path.join(d, 'in.gramps'), '<d><file src="media/./a.pdf" mime="x"/><file src="media/x/../b.pdf" mime="x"/></d>')
        n = g.relativize(os.path.join(d, 'in.gramps'), base, os.path.join(d, 'out.gramps'))
        o = _read(os.path.join(d, 'out.gramps'))
        self.assertEqual(n, 2)
        self.assertIn('src="media/a.pdf"', o); self.assertIn('src="media/b.pdf"', o)

    def test_intermediate_symlink_leaving_base_is_rejected(self):
        d = tempfile.mkdtemp(); base = os.path.join(d, 'lote'); os.makedirs(base)
        ext = os.path.join(d, 'externo'); os.makedirs(ext)
        os.symlink(ext, os.path.join(base, 'enlace'))
        for src in (os.path.join(base, 'enlace', 'x.pdf'), 'enlace/x.pdf'):
            _write(os.path.join(d, 'in.gramps'), '<d><file src="%s" mime="x"/></d>' % src)
            with self.assertRaises(g.RelError, msg=src):
                g.relativize(os.path.join(d, 'in.gramps'), base, os.path.join(d, 'out.gramps'))
        os.makedirs(os.path.join(base, 'real'))
        os.symlink(os.path.join(base, 'real'), os.path.join(base, 'alias_dentro'))
        _write(os.path.join(d, 'in.gramps'), '<d><file src="alias_dentro/x.pdf" mime="x"/></d>')
        self.assertEqual(g.relativize(os.path.join(d, 'in.gramps'), base, os.path.join(d, 'out.gramps')), 1)

    def test_cli_strict(self):
        import subprocess
        d = tempfile.mkdtemp()
        for args in ([], ['a', 'b'], ['a', 'b', 'c', 'd'], ['-x', 'b', 'c']):
            r = subprocess.run([sys.executable, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'gramps_relativize.py')] + args, capture_output=True, text=True)
            self.assertEqual(r.returncode, 2)


    def test_windows_and_unc_paths_fail_closed(self):
        d = tempfile.mkdtemp(); base = os.path.join(d, 'lote'); os.makedirs(base)
        for src in ('C:/x/a.pdf', 'c:\\x\\a.pdf', '\\\\srv\\s\\a.pdf', '//srv/s/a.pdf', 'media\\a.pdf'):
            _write(os.path.join(d, 'in.gramps'), '<d><file src="%s" mime="x"/></d>' % src)
            with self.assertRaises(g.RelError, msg=src):
                g.relativize(os.path.join(d, 'in.gramps'), base, os.path.join(d, 'out.gramps'))
            self.assertFalse(os.path.exists(os.path.join(d, 'out.gramps')), src)


if __name__ == '__main__':
    unittest.main()
