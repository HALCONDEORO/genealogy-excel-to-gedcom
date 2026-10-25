"""Reescribe las rutas <file src=...> de un .gramps (XML, gzip) a rutas relativas a una carpeta base.
Gramps convierte las rutas FILE de un GEDCOM en absolutas al importar; Gramps Web empareja los archivos del ZIP por ruta RELATIVA exacta.
Uso: python3 gramps_relativize.py entrada.gramps carpeta_base salida.gramps
Fail closed: si alguna ruta queda fuera de la carpeta base, no escribe nada. No toca nada mas del XML."""
import sys, os, re, gzip, tempfile


class RelError(Exception):
    pass


def relativize(inp, base, out):
    base = os.path.realpath(base)
    if os.path.realpath(inp) == os.path.realpath(out) or os.path.islink(out):
        raise RelError('la salida no puede ser la entrada ni un enlace')
    with gzip.open(inp, 'rb') as fh:
        xml = fh.read().decode('utf-8')
    n = [0]

    def inside(p):
        rp = os.path.realpath(p)
        return rp == base or (rp + '/').startswith(base + '/')

    def fix(m):
        src = m.group(2)
        if not src:
            raise RelError('ruta de archivo vacia')
        if re.match(r'^[A-Za-z]:', src) or '\\' in src or src.startswith('//'):
            raise RelError(f'ruta absoluta de Windows, UNC o con barra invertida no admitida: {src}')
        full = os.path.normpath(src if src.startswith('/') else os.path.join(base, src))
        rel = os.path.relpath(full, base)
        # contencion logica (tras normalizar, sin '..') Y real (symlinks intermedios que salgan de la carpeta)
        if rel == '.' or rel.startswith('..') or not inside(full):
            raise RelError(f'ruta fuera de la carpeta base: {src}')
        n[0] += 1
        return m.group(1) + rel + m.group(3)
    new = re.sub(r'(<file src=")([^"]*)(")', fix, xml)
    d = os.path.dirname(os.path.abspath(out))
    fd, tmp = tempfile.mkstemp(prefix='.rel-', suffix='.tmp', dir=d)
    try:
        with os.fdopen(fd, 'wb') as fh:
            with gzip.GzipFile(fileobj=fh, mode='wb') as gz:
                gz.write(new.encode('utf-8'))
        os.replace(tmp, out)
    except BaseException:
        if os.path.exists(tmp):
            os.unlink(tmp)
        raise
    return n[0]


def main(argv):
    if len(argv) != 4 or any(a.startswith('-') for a in argv[1:]):
        print('uso: python3 gramps_relativize.py entrada.gramps carpeta_base salida.gramps'); return 2
    try:
        n = relativize(argv[1], argv[2], argv[3])
    except (RelError, OSError) as ex:
        print('ERROR (fail closed):', ex); return 1
    print('rutas de archivo reescritas/verificadas:', n); return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
