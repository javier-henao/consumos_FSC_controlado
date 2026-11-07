# Consumos Picadora (Streamlit)

## Instalación y ejecución
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Uso
1. Cargue el Excel (.xls o .xlsx). El nombre del archivo y de las columnas puede cambiar:
   la app reconoce variantes como `#VW_CUBO_DADOS_ITEM_BOLETIM.DCR_PRODUTO_ORIGEM`,
   `Dcr Produto Origem` o `Producto`, y si no reconoce el nombre identifica la columna por su
   contenido. También acepta filas de título, celdas de producto agrupadas, subtotales y fechas
   en texto (dd/mm/aaaa). En la barra lateral, **🔗 Columnas del archivo** muestra qué columna se
   usó para cada dato y permite cambiarla.
2. **📊 Reporte**: filtre por fechas, categoría, tipo de fibra, corteza y producto; descargue el reporte en Excel o PDF.
3. **⚙️ Configuración de asignaciones**: edite las reglas de reemplazo de CAR y los rangos/códigos de cada categoría.
   - *Aplicar*: usa la configuración en la sesión actual.
   - *Guardar como predeterminada*: la guarda en `config_asignaciones.json` junto a la app.
   - *Exportar / Importar JSON*: para compartir la configuración.

## Archivos
- `app.py` — interfaz Streamlit
- `procesamiento.py` — lectura, reemplazo de CAR y clasificación
- `reportes.py` — generación de Excel y PDF
