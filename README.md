# Consumos Picadora (Streamlit)

## Instalación y ejecución
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Uso
1. Cargue el Excel (.xls o .xlsx; el nombre puede cambiar, las columnas deben ser las mismas).
2. **📊 Reporte**: filtre por fechas, categoría, tipo de fibra, corteza y producto; descargue el reporte en Excel o PDF.
3. **⚙️ Configuración de asignaciones**: edite las reglas de reemplazo de CAR y los rangos/códigos de cada categoría.
   - *Aplicar*: usa la configuración en la sesión actual.
   - *Guardar como predeterminada*: la guarda en `config_asignaciones.json` junto a la app.
   - *Exportar / Importar JSON*: para compartir la configuración.

## Archivos
- `app.py` — interfaz Streamlit
- `procesamiento.py` — lectura, reemplazo de CAR y clasificación
- `reportes.py` — generación de Excel y PDF
# consumos_FSC_controlado
