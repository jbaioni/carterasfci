# 📊 Monitor de Carteras FCI

Dashboard web para visualizar posiciones y evolución de carteras de FCI, agrupadas por Sociedad Gerente.

## Archivos

```
fci_dashboard/
├── index.html            ← La web app (no tocar)
├── data.js               ← Los datos (se regenera con el script)
├── procesar_carteras.py  ← Script para actualizar datos
└── README.md
```

## Publicar en GitHub Pages

1. Crear un repositorio en github.com (puede ser privado o público)
2. Subir los 3 archivos (`index.html`, `data.js`, `procesar_carteras.py`)
3. Ir a **Settings → Pages → Source → Deploy from branch → main / root**
4. El sitio queda en: `https://TU_USUARIO.github.io/NOMBRE_REPO/`

## Actualizar datos cada semana

### Paso 1: Agregar columna Sociedad Gerente al nuevo Excel
- Abrí el xlsx nuevo
- Agregá una columna `Sociedad Gerente` con el nombre de la SG para cada fondo
- Guardá el archivo

### Paso 2: Regenerar data.js

```bash
# Siempre pasá exactamente 3 archivos (los 3 viernes que querés mostrar)
# Sacá el más viejo y sumá el nuevo:

python procesar_carteras.py \
    20260529_Detalle_Carteras.xlsx \
    20260605_Detalle_Carteras.xlsx \
    20260612_Detalle_Carteras.xlsx
```

El script genera `data.js` automáticamente.

### Paso 3: Pushear a GitHub

```bash
git add data.js
git commit -m "Actualizar carteras al 12-06-2026"
git push
```

GitHub Pages se actualiza en ~1 minuto.

## Requisitos del script Python

```bash
pip install pandas openpyxl
```

## Formato esperado de los Excel

- Nombre del archivo debe empezar con la fecha: `YYYYMMDD_...xlsx`
- El archivo debe tener una columna llamada `Sociedad Gerente`
- El header de datos está en la fila 4 (formato estándar CNV)
