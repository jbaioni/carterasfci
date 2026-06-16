#!/usr/bin/env python3
"""
procesar_carteras.py
Procesamiento de archivos Excel de Detalle de Carteras FCI.

Uso:
    python procesar_carteras.py archivo1.xlsx archivo2.xlsx archivo3.xlsx

El script genera data.js en la misma carpeta que el script.
El nombre del archivo debe tener la fecha en formato YYYYMMDD al inicio.
Columna 'Sociedad Gerente' debe estar en los Excel (la agrega el usuario).
"""

import sys, re, json, zipfile, io
from pathlib import Path
import pandas as pd

def patch_infinity(src):
    """Reemplaza Infinity/-Infinity en el XML del xlsx para poder leerlo."""
    buf = io.BytesIO()
    with zipfile.ZipFile(src, 'r') as zin, zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('xl/worksheets/'):
                text = data.decode('utf-8')
                text = re.sub(r'<v>-?Infinity</v>', '<v>0</v>', text)
                data = text.encode('utf-8')
            zout.writestr(item, data)
    buf.seek(0)
    return buf

def extract_date(filename):
    """Extrae la fecha del nombre del archivo (YYYYMMDD)."""
    m = re.search(r'(\d{8})', Path(filename).name)
    if m:
        d = m.group(1)
        return f"{d[:4]}-{d[4:6]}-{d[6:8]}"
    return Path(filename).stem

def read_file(filepath, date_str):
    buf = patch_infinity(filepath)
    # header en fila 4 (índice 3)
    df = pd.read_excel(buf, sheet_name=0, header=3)
    df = df[df['Fondo'].notna()].copy()
    df['Fecha'] = date_str

    # Columna SG: si existe en el Excel la usamos, si no queda vacía
    if 'Sociedad Gerente' not in df.columns:
        df['Sociedad Gerente'] = ''

    # Normalizar Monto $
    df['Monto $'] = pd.to_numeric(df['Monto $'], errors='coerce').fillna(0)

    keep = ['Fecha', 'Fondo', 'Sociedad Gerente', 'Clasificación', 'Clase',
            'Subclase', 'Activo', 'Moneda', 'Monto $', 'Id Fondo CNV']
    existing = [c for c in keep if c in df.columns]
    return df[existing]

def build_json(files):
    frames = []
    for f in files:
        date = extract_date(f)
        print(f"  Procesando {Path(f).name} → {date}")
        df = read_file(f, date)
        frames.append(df)

    full = pd.concat(frames, ignore_index=True)
    dates = sorted(full['Fecha'].unique().tolist())

    # --- Posición por SG y fecha ---
    sg_fecha = (
        full.groupby(['Fecha', 'Sociedad Gerente', 'Fondo', 'Moneda'])['Monto $']
        .sum()
        .reset_index()
    )

    # --- Totales por SG y fecha ---
    sg_totales = (
        full.groupby(['Fecha', 'Sociedad Gerente'])['Monto $']
        .sum()
        .reset_index()
        .rename(columns={'Monto $': 'Total'})
    )

    # --- Activos únicos por Fondo (para drilldown) ---
    activos = (
        full[full['Monto $'] > 0]
        .groupby(['Fecha', 'Fondo', 'Activo', 'Moneda', 'Clasificación'])['Monto $']
        .sum()
        .reset_index()
    )

    # Serializar a listas de dicts
    data = {
        'fechas': dates,
        'posicion': sg_fecha.to_dict(orient='records'),
        'totales': sg_totales.to_dict(orient='records'),
        'activos': activos.head(50000).to_dict(orient='records'),  # cap para no explotar el browser
    }

    return data

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Uso: python procesar_carteras.py archivo1.xlsx [archivo2.xlsx ...]")
        sys.exit(1)

    files = sys.argv[1:]
    print(f"Procesando {len(files)} archivo(s)...")
    data = build_json(files)

    out = Path(__file__).parent / 'data.js'
    with open(out, 'w', encoding='utf-8') as f:
        f.write('const DATA = ')
        json.dump(data, f, ensure_ascii=False, default=str)
        f.write(';\n')

    print(f"✓ Generado {out}")
    print(f"  Fechas: {data['fechas']}")
    print(f"  Registros posición: {len(data['posicion'])}")
