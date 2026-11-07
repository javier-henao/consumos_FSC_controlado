"""Consumos Picadora — Streamlit.

Ejecutar:  streamlit run app.py
"""
import json
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

from procesamiento import (COLORES_CORTEZA, COLORES_FIBRA, SIN_ASIGNAR, cargar_config,
                           colores_para, config_defecto, detectar_columnas, guardar_config,
                           leer_tabla, normalizar, procesar, resumen_categoria, resumen_cruzado,
                           resumen_mensual, resumen_pila, validar_config)
from reportes import generar_excel, generar_pdf

st.set_page_config(page_title="Consumos Picadora", page_icon="🌲", layout="wide")

if "config" not in st.session_state:
    st.session_state.config = cargar_config()


@st.cache_data(show_spinner="Leyendo archivo…")
def _leer_tabla(contenido: bytes, nombre: str) -> pd.DataFrame:
    import io
    f = io.BytesIO(contenido)
    f.name = nombre
    return leer_tabla(f)


@st.cache_data(show_spinner=False)
def _normalizar(contenido: bytes, nombre: str, mapeo: tuple) -> pd.DataFrame:
    return normalizar(_leer_tabla(contenido, nombre), dict(mapeo))


CAMPOS_UI = {
    "Producto": "Producto / tipo de fibra (DCR_PRODUTO_ORIGEM)",
    "Fecha": "Fecha (DATA_INICIO)",
    "Hora": "Hora (HORA_INICIO) — opcional",
    "Pila": "Pila (NUM_PILHA_ORIGEM)",
    "Peso": "Peso (Peso Líquido)",
}
NINGUNA = "(ninguna)"


def obtener_datos():
    """Lee el archivo, detecta las columnas y permite corregir la asignación a mano."""
    if archivo is None:
        return None
    contenido = archivo.getvalue()
    try:
        tabla = _leer_tabla(contenido, archivo.name)
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}")
        return None

    auto = detectar_columnas(tabla)
    cols = [NINGUNA] + list(tabla.columns)
    faltan_auto = [c for c in ("Producto", "Fecha", "Pila", "Peso") if c not in auto]
    mapeo = {}
    with st.sidebar.expander("🔗 Columnas del archivo", expanded=bool(faltan_auto)):
        st.caption("Detectadas automáticamente. Cámbielas si alguna no corresponde.")
        for campo, etiqueta in CAMPOS_UI.items():
            actual = auto.get(campo, NINGUNA)
            sel = st.selectbox(etiqueta, cols, index=cols.index(actual),
                               key=f"map_{campo}_{archivo.name}_{archivo.size}")
            if sel != NINGUNA:
                mapeo[campo] = sel

    faltan = [c for c in ("Producto", "Fecha", "Pila", "Peso") if c not in mapeo]
    if faltan:
        st.warning("No se identificaron estas columnas: **" + ", ".join(faltan) + "**. "
                   "Selecciónelas en la barra lateral, en *🔗 Columnas del archivo*.")
        st.caption("Primeras filas del archivo tal como se leyó:")
        st.dataframe(tabla.head(15), width="stretch")
        return None
    try:
        return _normalizar(contenido, archivo.name, tuple(sorted(mapeo.items())))
    except Exception as e:
        st.error(f"No se pudieron interpretar los datos con las columnas elegidas: {e}")
        st.dataframe(tabla.head(15), width="stretch")
        return None


def _fmt(n, dec=0):
    return f"{n:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ============================================================================ SIDEBAR
st.sidebar.title("🌲 Consumos Picadora")
pagina = st.sidebar.radio("Menú", ["📊 Reporte", "⚙️ Configuración de asignaciones"])
archivo = st.sidebar.file_uploader("Archivo Excel (.xls / .xlsx)", type=["xls", "xlsx"])


# ============================================================================ CONFIGURACIÓN
def pagina_config():
    st.header("⚙️ Configuración de asignaciones")
    st.caption("Los cambios se aplican al reporte al presionar **Aplicar**. "
               "**Guardar como predeterminada** los conserva en `config_asignaciones.json`.")
    cfg = st.session_state.config

    st.subheader("1. Reemplazo de pila CAR según producto")
    st.write("Si la pila coincide con *Pila original* y el producto con *Producto*, "
             "la pila se cambia por *Nuevo valor*.")
    reps = st.data_editor(
        pd.DataFrame(cfg["reemplazos"], columns=["pila_original", "producto", "nuevo_valor"]),
        num_rows="dynamic", width="stretch", key="ed_reps",
        column_config={
            "pila_original": st.column_config.TextColumn("Pila original", required=True),
            "producto": st.column_config.TextColumn("Producto (DCR_PRODUTO_ORIGEM)", required=True),
            "nuevo_valor": st.column_config.TextColumn("Nuevo valor", required=True),
        })

    st.subheader("2. Asignación de pilas a categorías")
    st.write("Una pila pertenece a la categoría si su número está dentro del rango **o** si "
             "aparece en *Códigos* (separados por coma). Los códigos se evalúan primero.")
    cats = st.data_editor(
        pd.DataFrame(cfg["categorias"],
                     columns=["categoria", "rango_min", "rango_max", "codigos", "notas"]),
        num_rows="dynamic", width="stretch", key="ed_cats",
        column_config={
            "categoria": st.column_config.TextColumn("Categoría", required=True),
            "rango_min": st.column_config.NumberColumn("Rango desde", step=1, format="%d"),
            "rango_max": st.column_config.NumberColumn("Rango hasta", step=1, format="%d"),
            "codigos": st.column_config.TextColumn("Códigos (separados por coma)", width="large"),
            "notas": st.column_config.TextColumn("Notas"),
        })

    nuevo = validar_config({"reemplazos": reps.to_dict("records"),
                            "categorias": cats.to_dict("records")})

    c1, c2, c3, c4 = st.columns(4)
    if c1.button("✅ Aplicar", type="primary", width="stretch"):
        st.session_state.config = nuevo
        st.success("Configuración aplicada.")
    if c2.button("💾 Guardar como predeterminada", width="stretch"):
        st.session_state.config = nuevo
        try:
            guardar_config(nuevo)
            st.success("Guardada en config_asignaciones.json.")
        except OSError as e:
            st.error(f"No se pudo guardar: {e}")
    if c3.button("↩️ Restaurar valores de fábrica", width="stretch"):
        st.session_state.config = config_defecto()
        for k in ("ed_reps", "ed_cats"):
            st.session_state.pop(k, None)
        st.rerun()
    c4.download_button("⬇️ Exportar JSON", json.dumps(nuevo, ensure_ascii=False, indent=2),
                       "config_asignaciones.json", "application/json", width="stretch")

    sub = st.file_uploader("Importar configuración (JSON)", type=["json"], key="up_cfg")
    if sub is not None and st.button("Cargar configuración importada"):
        try:
            st.session_state.config = validar_config(json.load(sub))
            for k in ("ed_reps", "ed_cats"):
                st.session_state.pop(k, None)
            st.rerun()
        except Exception as e:
            st.error(f"Archivo inválido: {e}")

    # Vista previa con los datos cargados
    base = obtener_datos()
    if base is not None:
        st.subheader("Vista previa: pilas por categoría con esta configuración")
        df = procesar(base, nuevo)
        prev = (df.groupby(["Categoría", "Pila"], dropna=False)
                .agg(Registros=("Toneladas", "size"), Toneladas=("Toneladas", "sum"))
                .reset_index().sort_values(["Categoría", "Pila"]))
        st.dataframe(prev, width="stretch", hide_index=True,
                     column_config={"Toneladas": st.column_config.NumberColumn(format="%.2f")})
        sa = prev[prev["Categoría"] == SIN_ASIGNAR]
        if len(sa):
            st.warning("Pilas sin asignar: " + ", ".join(sa["Pila"].fillna("(vacío)").astype(str)))


# ============================================================================ REPORTE
def pagina_reporte():
    st.header("📊 Reporte de Consumos Picadora")
    if archivo is None:
        st.info("Cargue el archivo Excel en la barra lateral. El nombre del archivo y de las "
                "columnas puede variar: la app las detecta automáticamente y, si no puede, "
                "le permite elegirlas.")
        return
    base = obtener_datos()
    if base is None:
        return
    if base.empty:
        st.error("No se encontraron registros con fecha válida. Revise en la barra lateral "
                 "que la columna de *Fecha* esté bien asignada.")
        return
    descartadas = base.attrs.get("filas_descartadas", 0)
    if descartadas:
        st.caption(f"ℹ️ Se omitieron {descartadas:,} filas sin fecha (separadores o subtotales "
                   "del reporte de origen).")
    df = procesar(base, st.session_state.config)

    # ---------------- filtros
    st.sidebar.markdown("---")
    st.sidebar.subheader("Filtros")
    fechas = df["Fecha"].dropna()
    fmin, fmax = fechas.min().date(), fechas.max().date()
    rango = st.sidebar.date_input("Rango de fechas", (fmin, fmax), min_value=fmin, max_value=fmax)
    if isinstance(rango, (list, tuple)) and len(rango) == 2:
        d1, d2 = rango
    else:
        d1 = d2 = rango[0] if isinstance(rango, (list, tuple)) else rango

    todas_cat = [c["categoria"] for c in st.session_state.config["categorias"]] + [SIN_ASIGNAR]
    todas_cat = [c for c in todas_cat if c in set(df["Categoría"])]
    sel_cat = st.sidebar.multiselect("Categoría", todas_cat,
                                     [c for c in todas_cat if c != SIN_ASIGNAR])
    tipos = sorted(df["Tipo de fibra"].unique())
    sel_tipo = st.sidebar.multiselect("Tipo de fibra", tipos,
                                      [t for t in tipos if t != "Otro"] or tipos)
    cortezas = sorted(df["Corteza"].unique())
    sel_crt = st.sidebar.multiselect("Corteza", cortezas,
                                     [c for c in cortezas if c != "N/A"] or cortezas)
    prods = sorted(df["Producto"].dropna().unique())
    sel_prod = st.sidebar.multiselect("Producto (opcional)", prods, [])
    excluir_cero = st.sidebar.checkbox("Excluir registros con peso 0", True)

    m = (df["Fecha"].dt.date.between(d1, d2) & df["Categoría"].isin(sel_cat)
         & df["Tipo de fibra"].isin(sel_tipo) & df["Corteza"].isin(sel_crt))
    if sel_prod:
        m &= df["Producto"].isin(sel_prod)
    if excluir_cero:
        m &= df["Peso Líquido"] != 0
    f = df[m].copy()

    filtros = {
        "Archivo": archivo.name,
        "Fechas": f"{d1:%Y-%m-%d} a {d2:%Y-%m-%d}",
        "Categorías": ", ".join(sel_cat) or "(ninguna)",
        "Tipo de fibra": ", ".join(sel_tipo) or "(ninguno)",
        "Corteza": ", ".join(sel_crt) or "(ninguna)",
        "Productos": ", ".join(sel_prod) or "Todos",
        "Peso 0": "Excluido" if excluir_cero else "Incluido",
    }

    # ---------------- KPIs
    k = st.columns(4)
    k[0].metric("Registros", _fmt(len(f)))
    k[1].metric("Toneladas", _fmt(f["Toneladas"].sum(), 2))
    k[2].metric("Peso total (kg)", _fmt(f["Peso Líquido"].sum()))
    k[3].metric("Reemplazos CAR", _fmt(f["Reemplazo aplicado"].notna().sum()))

    # ---------------- descargas
    st.markdown("#### Descargar reporte")
    b1, b2, b3 = st.columns([1, 1, 2])
    sello = datetime.now().strftime("%Y%m%d_%H%M")
    with b1:
        st.download_button("📗 Reporte Excel", data=generar_excel(f, filtros),
                           file_name=f"reporte_consumos_{sello}.xlsx", type="primary",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width="stretch")
    with b3:
        detalle = st.checkbox("Incluir detalle de todos los registros en el PDF", True)
    with b2:
        st.download_button("📕 Reporte PDF", data=generar_pdf(f, filtros, detalle),
                           file_name=f"reporte_consumos_{sello}.pdf", mime="application/pdf",
                           type="primary", width="stretch")

    if f.empty:
        st.warning("No hay registros con los filtros seleccionados.")
        return

    cmap = colores_para(sorted(df["Categoría"].unique()))
    t1, t2, t3, t4, t5 = st.tabs(["Resumen", "Tendencia", "Por pila", "Reemplazos CAR",
                                  "Datos completos"])

    with t1:
        rc = resumen_categoria(f)
        c1, c2 = st.columns([3, 2])
        fig = px.bar(rc.sort_values("Toneladas"), x="Toneladas", y="Categoría", orientation="h",
                     color="Categoría", color_discrete_map=cmap, text_auto=",.1f",
                     hover_data={"Registros": True, "% del total": ":.2f"},
                     title="Toneladas por categoría")
        fig.update_layout(showlegend=False, bargap=0.35)
        c1.plotly_chart(fig, width="stretch")
        g = f.groupby("Corteza", as_index=False)["Toneladas"].sum()
        fig = px.pie(g, names="Corteza", values="Toneladas", hole=0.55, color="Corteza",
                     color_discrete_map=COLORES_CORTEZA, title="Con / sin corteza")
        fig.update_traces(marker_line_color="white", marker_line_width=2)
        c2.plotly_chart(fig, width="stretch")

        c3, c4 = st.columns([3, 2])
        g = f.groupby(["Categoría", "Corteza"], as_index=False)["Toneladas"].sum()
        fig = px.bar(g, x="Categoría", y="Toneladas", color="Corteza", barmode="group",
                     color_discrete_map=COLORES_CORTEZA, title="Categoría por corteza")
        c3.plotly_chart(fig, width="stretch")
        g = f.groupby("Tipo de fibra", as_index=False)["Toneladas"].sum()
        fig = px.pie(g, names="Tipo de fibra", values="Toneladas", hole=0.55,
                     color="Tipo de fibra", color_discrete_map=COLORES_FIBRA,
                     title="Tipo de fibra")
        fig.update_traces(marker_line_color="white", marker_line_width=2)
        c4.plotly_chart(fig, width="stretch")

        st.markdown("##### Resumen por categoría")
        st.dataframe(rc, width="stretch", hide_index=True,
                     column_config={"Toneladas": st.column_config.NumberColumn(format="%.2f"),
                                    "Peso Líquido (kg)": st.column_config.NumberColumn(format="%d")})
        st.markdown("##### Toneladas por categoría y corteza")
        st.dataframe(resumen_cruzado(f, "Categoría", "Corteza"), width="stretch")
        st.markdown("##### Toneladas por tipo de fibra / corteza y categoría")
        st.dataframe(resumen_cruzado(f, ["Tipo de fibra", "Corteza"], "Categoría"),
                     width="stretch")

    with t2:
        tm = resumen_mensual(f).reset_index().melt(id_vars="Mes", var_name="Categoría",
                                                   value_name="Toneladas")
        fig = px.bar(tm, x="Mes", y="Toneladas", color="Categoría", color_discrete_map=cmap,
                     title="Toneladas por mes")
        fig.update_layout(bargap=0.3)
        st.plotly_chart(fig, width="stretch")
        dia = (f.assign(Día=f["Fecha"].dt.date).groupby(["Día", "Categoría"], as_index=False)
               ["Toneladas"].sum())
        fig = px.line(dia, x="Día", y="Toneladas", color="Categoría", color_discrete_map=cmap,
                      title="Consumo diario")
        fig.update_traces(line_width=2)
        fig.update_layout(hovermode="x unified")
        st.plotly_chart(fig, width="stretch")
        st.dataframe(resumen_mensual(f), width="stretch")

    with t3:
        rp = resumen_pila(f)
        top = rp.sort_values("Toneladas", ascending=False).head(25)
        fig = px.bar(top.sort_values("Toneladas"), x="Toneladas", y=top.sort_values("Toneladas")["Pila"].astype(str),
                     color="Categoría", color_discrete_map=cmap, orientation="h",
                     title="Top 25 pilas por toneladas", labels={"y": "Pila"}, height=650)
        st.plotly_chart(fig, width="stretch")
        st.dataframe(rp, width="stretch", hide_index=True)

    with t4:
        rep = f[f["Reemplazo aplicado"].notna()]
        st.write(f"**{len(rep):,}** registros donde la pila CAR fue reemplazada.")
        st.dataframe(rep.groupby("Reemplazo aplicado").agg(
            Registros=("Toneladas", "size"), Toneladas=("Toneladas", "sum")).reset_index(),
            width="stretch", hide_index=True)
        st.dataframe(rep, width="stretch", hide_index=True)
        sa = df[df["Categoría"] == SIN_ASIGNAR]
        if len(sa):
            with st.expander(f"⚠️ {len(sa)} registros sin asignar en el archivo completo"):
                st.dataframe(sa, width="stretch", hide_index=True)

    with t5:
        st.dataframe(f, width="stretch", hide_index=True,
                     column_config={"Fecha": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm"),
                                    "Toneladas": st.column_config.NumberColumn(format="%.3f")})


if pagina.startswith("📊"):
    pagina_reporte()
else:
    pagina_config()
