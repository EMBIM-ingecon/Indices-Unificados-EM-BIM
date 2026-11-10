"""Actualiza los Índices Unificados (IU) del INEI.

Uso:  python3 -I update_iu.py [--force]
Salida: /mnt/project-files/indices-unificados/<AAAA-MM>/ con
  - IU_<AAAA-MM>_base_dic2025.csv          (nuestro CSV largo, toda la serie)
  - INEI_<archivo original>.xlsx           (Excel original del INEI)
  - RJ_<nro>-INEI_IU_<mes>_<año>.pdf       (Resolución Jefatural, extracto El Peruano)
  - indices_unificados_base_dic2025_<AAAA-MM>.xlsx (Excel de consulta)
Imprime 'NUEVO <AAAA-MM>' si hubo mes nuevo, 'SIN CAMBIOS <AAAA-MM>' si no.
"""
import re, sys, os, json, subprocess, tempfile, shutil
import pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE = os.environ.get('IU_BASE', '/mnt/project-files/indices-unificados')
STATE = os.path.join(BASE, 'estado.json')
INEI_PAGE = 'https://www.inei.gob.pe/estadisticas/indice-tematico/price-indexes/'
GOB_PAGE = 'https://www.gob.pe/institucion/inei/informes-publicaciones/4025211-indice-unificados-de-precios-de-la-construccion-para-las-seis-areas-geograficas'
UA = 'Mozilla/5.0'
MESN = {'DICIEMBRE':12,'ENERO':1,'FEBRERO':2,'MARZO':3,'ABRIL':4,'MAYO':5,'JUNIO':6,'JULIO':7,'AGOSTO':8,'SETIEMBRE':9,'SEPTIEMBRE':9,'OCTUBRE':10,'NOVIEMBRE':11}
MN = ['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Set','Oct','Nov','Dic']
MESLARGO = ['','enero','febrero','marzo','abril','mayo','junio','julio','agosto','setiembre','octubre','noviembre','diciembre']
AREAS = {1:"Lambayeque, Piura, Tumbes",2:"Cajamarca, La Libertad",3:"Áncash",4:"Lima y Callao",5:"Ica",6:"Arequipa, Moquegua, Tacna",7:"Amazonas, San Martín",8:"Huánuco, Junín, Pasco",9:"Ayacucho, Huancavelica",10:"Apurímac, Cusco, Madre de Dios",11:"Puno",12:"Loreto",13:"Ucayali"}
lab = lambda p: f"{MN[int(p[5:])]}-{p[:4]}"

def curl(url, out=None):
    cmd = ['curl', '-sSL', '--retry', '4', '--retry-delay', '5', '-A', UA, '-f', url]
    if out: cmd += ['-o', out]
    r = subprocess.run(cmd, capture_output=True, check=True)
    return r.stdout.decode('utf-8', 'replace') if not out else out

def norm(c):
    s = str(c).strip().replace(' ', '')
    if re.fullmatch(r'\d+(\.0)?', s): return f"{int(float(s)):02d}"
    return s

def parse(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    rows, names, meta = [], {}, []
    for ws in wb:
        vals = list(ws.iter_rows(values_only=True))
        if 'Relaci' in ws.title:
            for r in vals:
                for i in (0, 2):
                    if len(r) > i+1 and r[i] is not None and r[i+1] and re.fullmatch(r'\d+(-1)?', str(r[i]).strip().replace(' ', '')):
                        names[norm(r[i])] = str(r[i+1]).strip()
            continue
        txt = " ".join(str(c) for r in vals[:6] for c in r if c)
        m = re.search(r'MES DE (\w+) (?:DEL )?(\d{4})', txt)
        if not m or m.group(1).upper() not in MESN: continue
        per = f"{m.group(2)}-{MESN[m.group(1).upper()]:02d}"
        res = re.search(r'N[º°]\s*([\d-]+-INEI)', txt); res = res.group(1) if res else ''
        hdr = next(((ri, ci) for ri, r in enumerate(vals) for ci, c in enumerate(r)
                    if isinstance(c, str) and c.strip().startswith('Cód')), None)
        if not hdr: continue
        ci = hdr[1]
        for r in vals[hdr[0]+1:]:
            c = r[ci]
            if c is None: continue
            code = norm(c)
            if not re.fullmatch(r'\d{2}(-1)?', code): continue
            for a in range(1, 14):
                v = r[ci+a] if len(r) > ci+a else None
                rows.append(dict(periodo=per, codigo=code, area=a,
                                 indice=None if (v is None or isinstance(v, str)) else float(v)))
        meta.append((per, res, ws.title))
    df = pd.DataFrame(rows).drop_duplicates(['periodo', 'codigo', 'area'])
    return df, pd.Series(names), pd.DataFrame(meta, columns=['periodo', 'resolucion', 'hoja'])

def build_excel(df, names, meta, out, src_name):
    H = Font(bold=True, color="FFFFFF"); F = PatternFill("solid", fgColor="1F4E78")
    def style(ws, row, ncol):
        for c in ws[row][:ncol]: c.font = H; c.fill = F; c.alignment = Alignment(horizontal="center", wrap_text=True)
    pers = sorted(df.periodo.unique()); last = pers[-1]
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        d = df[df.periodo == last].pivot(index='codigo', columns='area', values='indice')
        d.columns = [f"Área {a}" for a in d.columns]; d.insert(0, 'Elemento', d.index.map(names))
        d = d.reset_index().rename(columns={'codigo': 'Código'})
        sn = f"Último ({lab(last)})"; d.to_excel(w, sheet_name=sn, index=False, startrow=3); ws = w.sheets[sn]
        r = meta.set_index('periodo').loc[last, 'resolucion']
        ws['A1'] = f"Índices Unificados de Precios de la Construcción – {lab(last)} – R.J. N° {r} (Base dic 2025 = 100)"; ws['A1'].font = Font(bold=True, size=12)
        ws['A2'] = f"Celda vacía = (*) Sin índice. Fuente: INEI, {src_name}"
        style(ws, 4, 15); ws.freeze_panes = 'C5'; ws.column_dimensions['B'].width = 45
        for i in range(3, 16): ws.column_dimensions[get_column_letter(i)].width = 9
        for row in ws.iter_rows(min_row=5, min_col=3):
            for c in row: c.number_format = '0.00'
        s = df.pivot_table(index=['codigo', 'area'], columns='periodo', values='indice', dropna=False).reset_index()
        s.insert(0, 'Clave', s.codigo + '|' + s.area.astype(str)); s.insert(2, 'Elemento', s.codigo.map(names))
        s.columns = ['Clave', 'Código', 'Elemento', 'Área'] + [lab(p) for p in pers]
        s.to_excel(w, sheet_name='Serie', index=False); ws = w.sheets['Serie']; style(ws, 1, len(s.columns))
        ws.freeze_panes = 'E2'; ws.column_dimensions['C'].width = 40
        for row in ws.iter_rows(min_row=2, min_col=5):
            for c in row: c.number_format = '0.00'
        n = len(s)+1; lc = get_column_letter(len(s.columns))
        ws = w.book.create_sheet('Consulta', 0)
        ws['A1'] = "Consulta rápida de índice y factor de variación"; ws['A1'].font = Font(bold=True, size=12)
        for i, (k, v) in enumerate([("Código (ej. 01, 21, 47-1)", '21'), ("Área (1–13)", 4), ("Mes base (presupuesto)", lab(pers[0])), ("Mes de reajuste", lab(last))], start=3):
            ws[f'A{i}'] = k; ws[f'B{i}'] = v; ws[f'B{i}'].fill = PatternFill("solid", fgColor="FFF2CC")
        rng = f"Serie!$A$2:${lc}${n}"; hdr = f"Serie!$A$1:${lc}$1"; key = f'MATCH(B3&"|"&B4,Serie!$A$2:$A${n},0)'
        ws['A8'] = "Elemento"; ws['B8'] = f'=IFERROR(INDEX(Serie!$C$2:$C${n},{key}),"Código/área no encontrado")'
        ws['A9'] = "Índice mes base (Io)"; ws['B9'] = f'=IFERROR(INDEX({rng},{key},MATCH(B5,{hdr},0)),"(*) / no disponible")'
        ws['A10'] = "Índice mes reajuste (Ir)"; ws['B10'] = f'=IFERROR(INDEX({rng},{key},MATCH(B6,{hdr},0)),"(*) / no disponible")'
        ws['A11'] = "Ir / Io"; ws['B11'] = '=IFERROR(ROUND(B10/B9,3),"")'
        ws['A13'] = "Meses disponibles: " + ", ".join(lab(p) for p in pers)
        ws['A14'] = "Nota: el factor K de la fórmula polinómica se redondea a 3 decimales; aquí Ir/Io es solo un monomio de referencia."
        ws['B9'].number_format = ws['B10'].number_format = '0.00'; ws['B11'].number_format = '0.000'; ws['B3'].number_format = '@'
        ws.column_dimensions['A'].width = 30; ws.column_dimensions['B'].width = 45
        dv = DataValidation(type='list', formula1='"' + ",".join(lab(p) for p in pers) + '"'); ws.add_data_validation(dv); dv.add('B5'); dv.add('B6')
        dv2 = DataValidation(type='whole', operator='between', formula1='1', formula2='13'); ws.add_data_validation(dv2); dv2.add('B4')
        pd.DataFrame({'Código': names.index, 'Elemento': names.values}).sort_values('Código').to_excel(w, sheet_name='Códigos', index=False)
        w.sheets['Códigos'].column_dimensions['B'].width = 55; style(w.sheets['Códigos'], 1, 2)
        pd.DataFrame([{'Área': k, 'Departamentos': v} for k, v in AREAS.items()]).to_excel(w, sheet_name='Áreas', index=False)
        w.sheets['Áreas'].column_dimensions['B'].width = 40; style(w.sheets['Áreas'], 1, 2)
        mm = meta.copy(); mm['mes'] = mm.periodo.map(lab); mm[['mes', 'periodo', 'resolucion']].to_excel(w, sheet_name='Resoluciones', index=False)
        style(w.sheets['Resoluciones'], 1, 3)
        df[['periodo', 'codigo', 'elemento', 'area', 'indice']].to_excel(w, sheet_name='Datos (largo)', index=False)

def fetch_pdfs(last, res, out):
    # PDF de la R.J. del mes (gob.pe)
    pdf_note = ''
    nro = res.split('-')[0].lstrip('0')
    mes_txt = f"{MESLARGO[int(last[5:])]}_{last[:4]}"
    # 2a. R.J. original firmada (gob.pe normas legales del INEI)
    try:
        link = None
        for sh in range(1, 8):
            h = curl(f'https://www.gob.pe/institucion/inei/normas-legales?sheet={sh}')
            mm = re.search(rf'/institucion/inei/normas-legales/\d+-0*{nro}-{last[:4] if last[5:]!="12" else int(last[:4])+1}-inei"', h)
            if mm: link = mm.group(0).rstrip('"'); break
        if link:
            p = curl('https://www.gob.pe' + link)
            u = re.search(r'https://cdn\.www\.gob\.pe/uploads/document/file/\d+/[^"\'?]*\.pdf', p)
            if u: curl(u.group(0), os.path.join(out, f'RJ_{res}_IU_{mes_txt}_original_firmado.pdf'))
    except Exception as e:
        pdf_note += f' (original firmado no disponible: {e})'
    # 2b. Extracto publicado en El Peruano (página de publicaciones IU en gob.pe)
    try:
        g = curl(GOB_PAGE)
        pdfs = sorted(set(re.findall(r'https://cdn\.www\.gob\.pe/uploads/document/file/\d+/[^"\'?]*\.pdf', g)))
        cand = [p for p in pdfs if re.search(rf'r-j-n-0*{nro}-{last[:4]}', p)]
        if cand:
            curl(cand[0], os.path.join(out, f'RJ_{res}_IU_{mes_txt}_extracto_ElPeruano.pdf'))
    except Exception as e:
        pdf_note += f' (extracto El Peruano no disponible: {e})'
    if not any(f.startswith('RJ_') for f in os.listdir(out)):
        pdf_note = f'PDF de la R.J. {res} aún no aparece en gob.pe' + pdf_note
    return pdf_note

def main():
    force = '--force' in sys.argv
    state = json.load(open(STATE)) if os.path.exists(STATE) else {}
    tmp = tempfile.mkdtemp()
    # 1. Excel INEI: el enlace más reciente de la serie base dic 2025 (prefijo n07_)
    html = curl(INEI_PAGE)
    links = sorted(set(re.findall(r'(/media/MenuRecursivo/indices_tematicos/[^"\']*indices_unificados[^"\']*\.xlsx)', html)))
    links = [l for l in links if '/n07_' in l] or links
    if not links: sys.exit('ERROR: no encontré el enlace del Excel en ' + INEI_PAGE)
    url_x = 'https://www.inei.gob.pe' + links[-1]
    xname = os.path.basename(url_x)
    xpath = curl(url_x, os.path.join(tmp, xname))
    df, names, meta = parse(xpath)
    last = sorted(df.periodo.unique())[-1]
    if not force and state.get('ultimo_periodo') == last:
        out = os.path.join(BASE, last)
        if os.path.isdir(out) and not any(f.startswith('RJ_') for f in os.listdir(out)):
            note = fetch_pdfs(last, state.get('resolucion', ''), out)
            if any(f.startswith('RJ_') for f in os.listdir(out)):
                print('NUEVO PDF', last); return
            print('SIN CAMBIOS', last, note); return
        print('SIN CAMBIOS', last); return
    res = meta.set_index('periodo').loc[last, 'resolucion']
    df['elemento'] = df.codigo.map(names)
    out = os.path.join(BASE, last); os.makedirs(out, exist_ok=True)
    shutil.copy(xpath, os.path.join(out, 'INEI_' + xname))
    df[['periodo', 'codigo', 'elemento', 'area', 'indice']].to_csv(os.path.join(out, f'IU_{last}_base_dic2025.csv'), index=False)
    build_excel(df, names, meta, os.path.join(out, f'indices_unificados_base_dic2025_{last}.xlsx'), xname)
    pdf_note = fetch_pdfs(last, res, out)
    state.update(ultimo_periodo=last, resolucion=res, excel=url_x)
    json.dump(state, open(STATE, 'w'), indent=1)
    print('NUEVO', last, res, out, pdf_note)

if __name__ == '__main__':
    main()
