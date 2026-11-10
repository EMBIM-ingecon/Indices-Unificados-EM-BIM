"""Genera el sitio web estático de los Índices Unificados (GitHub Pages).

Uso:  python3 -I build_site.py --out <carpeta del repo> [--preview]
  - Lee /mnt/project-files/indices-unificados/<AAAA-MM>/ (CSV más reciente = serie completa)
  - Escribe <out>/index.html (datos embebidos), <out>/descargas/<AAAA-MM>/..., README.md y .nojekyll
  --preview: solo escribe index.html sin el esqueleto HTML ni enlaces de descarga (vista previa)
"""
import os, re, sys, json, shutil, glob, datetime
import pandas as pd

BASE = os.environ.get('IU_BASE', '/mnt/project-files/indices-unificados')
HERE = os.path.dirname(os.path.abspath(__file__))
AREAS = {1:"Lambayeque, Piura, Tumbes",2:"Cajamarca, La Libertad",3:"Áncash",4:"Lima y Callao",5:"Ica",6:"Arequipa, Moquegua, Tacna",7:"Amazonas, San Martín",8:"Huánuco, Junín, Pasco",9:"Ayacucho, Huancavelica",10:"Apurímac, Cusco, Madre de Dios",11:"Puno",12:"Loreto",13:"Ucayali"}

def main():
    out = sys.argv[sys.argv.index('--out') + 1]
    preview = '--preview' in sys.argv
    meses = sorted(d for d in os.listdir(BASE) if re.fullmatch(r'\d{4}-\d{2}', d))
    last = meses[-1]
    df = pd.read_csv(os.path.join(BASE, last, f'IU_{last}_base_dic2025.csv'), dtype={'codigo': str})
    P = sorted(df.periodo.unique())
    # Resoluciones: hoja 'Resoluciones' del Excel de consulta más reciente
    rjs = {}
    try:
        r = pd.read_excel(os.path.join(BASE, last, f'indices_unificados_base_dic2025_{last}.xlsx'), sheet_name='Resoluciones', dtype=str)
        rjs = dict(zip(r.periodo, r.resolucion))
    except Exception:
        pass
    codes = sorted(df.codigo.unique(), key=lambda c: (int(c.split('-')[0]), c))
    el = df.drop_duplicates('codigo').set_index('codigo').elemento.to_dict()
    v = {}
    for c in codes:
        v[c] = {}
        for a in range(1, 14):
            s = df[(df.codigo == c) & (df.area == a)].set_index('periodo').indice.reindex(P)
            v[c][str(a)] = [None if pd.isna(x) else round(float(x), 2) for x in s]
    info = []
    for m in meses:
        src = os.path.join(BASE, m); files = sorted(os.listdir(src))
        dst = os.path.join(out, 'descargas', m)
        if not preview and os.path.abspath(src) != os.path.abspath(dst):
            os.makedirs(dst, exist_ok=True)
            for f in files:
                t = os.path.join(dst, f)
                if not os.path.exists(t) or os.path.getsize(t) != os.path.getsize(os.path.join(src, f)):
                    shutil.copy2(os.path.join(src, f), t)
        rel = lambda f: f'descargas/{m}/{f}'
        pick = lambda pat: next((rel(f) for f in files if re.match(pat, f)), None)
        pdfs = [f for f in files if f.startswith('RJ_') and f.endswith('.pdf')]
        if any('extracto' in f for f in pdfs) and any('firmado' in f for f in pdfs):
            pdfs = [f for f in pdfs if 'extracto' in f or 'firmado' in f]
        res = rjs.get(m) or next((re.match(r'RJ_([\d-]+-INEI)', f).group(1) for f in pdfs), '')
        info.append(dict(periodo=m, rj=res,
            pdf=[dict(ruta=rel(f), nombre='Extracto El Peruano' if 'extracto' in f else 'Original firmado' if 'firmado' in f else 'PDF R.J.') for f in pdfs],
            xlsx_inei=pick(r'INEI_.*\.xlsx$'), csv=pick(r'IU_.*\.csv$'), xlsx_consulta=pick(r'indices_unificados_.*\.xlsx$')))
    data = dict(periodos=P, codes=codes, el=el, areas={str(k): n for k, n in AREAS.items()}, v=v, meses=info,
                descargas=not preview, generado=datetime.date.today().strftime('%d/%m/%Y'))
    tpl = open(os.path.join(HERE, 'web_template.html'), encoding='utf-8').read()
    body = tpl.replace('/*__DATA__*/null', json.dumps(data, ensure_ascii=False, separators=(',', ':')))
    if not preview:
        body = ('<!doctype html>\n<html lang="es">\n<head>\n<meta charset="utf-8">\n'
                '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
                + body.replace('</style>', '</style>\n</head>\n<body>', 1) + '\n</body>\n</html>\n')
        open(os.path.join(out, 'README.md'), 'w', encoding='utf-8').write(
            '# Índices Unificados de Precios de la Construcción (INEI)\n\n'
            'Consulta pública de los Índices Unificados del INEI, base diciembre 2025 = 100, '
            'para el reajuste de valorizaciones de obra pública mediante fórmula polinómica.\n\n'
            f'Último mes publicado: **{last}** (R.J. N° {rjs.get(last, "")}).\n\n'
            '- `index.html`: calculadora Ir/Io, tabla del mes por área y descargas.\n'
            '- `descargas/AAAA-MM/`: CSV (serie completa), Excel original del INEI, PDF de la R.J. y Excel de consulta.\n\n'
            'Fuente: INEI y diario oficial El Peruano. Sitio no oficial; prevalece la R.J. publicada.\n')
    os.makedirs(out, exist_ok=True)
    open(os.path.join(out, 'index.html'), 'w', encoding='utf-8').write(body)
    print('OK', last, out)

if __name__ == '__main__':
    main()
