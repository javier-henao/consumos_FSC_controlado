"""Generación de reportes Excel y PDF."""
from __future__ import annotations

import io
from datetime import datetime

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

from procesamiento import (COLORES_CORTEZA, COLORES_FIBRA, colores_para, resumen_categoria,
                           resumen_cruzado, resumen_mensual, resumen_pila)

AZUL = "1F4E79"
TINTA = "#52514e"


# ============================================================================ EXCEL
def _estilo_encabezado(ws, fila, ncols, col_ini=1):
    for c in range(col_ini, col_ini + ncols):
        cell = ws.cell(row=fila, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=AZUL)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _escribir_df(ws, df, fila=1, col=1, formato_num="#,##0.00"):
    borde = Border(*(Side(style="thin", color="D9D9D9"),) * 4)
    for i, row in enumerate(dataframe_to_rows(df, index=False, header=True)):
        for j, v in enumerate(row):
            if v is pd.NA or (isinstance(v, float) and pd.isna(v)):
                v = None
            cell = ws.cell(row=fila + i, column=col + j, value=v)
            cell.border = borde
            if i > 0 and isinstance(v, (int, float)):
                cell.number_format = formato_num if isinstance(v, float) else "#,##0"
    _estilo_encabezado(ws, fila, len(df.columns), col)
    return fila + len(df)  # última fila escrita


def _autoancho(ws, minimo=10, maximo=45):
    for col in ws.columns:
        largo = max((len(str(c.value)) for c in col if c.value is not None), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(minimo, min(maximo, largo + 2))


def generar_excel(df: pd.DataFrame, filtros: dict) -> bytes:
    wb = Workbook()

    # ---- Resumen
    ws = wb.active
    ws.title = "Resumen"
    ws["A1"] = "Reporte de Consumos Picadora"
    ws["A1"].font = Font(bold=True, size=16, color=AZUL)
    ws["A2"] = f"Generado: {datetime.now():%Y-%m-%d %H:%M}"
    fila = 4
    ws.cell(row=fila, column=1, value="Filtros aplicados").font = Font(bold=True, size=12)
    for k, v in filtros.items():
        fila += 1
        ws.cell(row=fila, column=1, value=k).font = Font(bold=True)
        ws.cell(row=fila, column=2, value=v)
    fila += 2
    ws.cell(row=fila, column=1, value="Indicadores").font = Font(bold=True, size=12)
    kpis = [("Registros", len(df)), ("Peso total (kg)", float(df["Peso Líquido"].sum())),
            ("Toneladas", float(df["Toneladas"].sum())),
            ("Registros con reemplazo CAR", int(df["Reemplazo aplicado"].notna().sum()))]
    for k, v in kpis:
        fila += 1
        ws.cell(row=fila, column=1, value=k).font = Font(bold=True)
        c = ws.cell(row=fila, column=2, value=v)
        c.number_format = "#,##0.00" if isinstance(v, float) else "#,##0"

    fila += 2
    ws.cell(row=fila, column=1, value="Consumo por categoría").font = Font(bold=True, size=12)
    rc = resumen_categoria(df)
    ini = fila + 1
    fin = _escribir_df(ws, rc, ini)
    if len(rc):
        ch = BarChart()
        ch.type = "bar"
        ch.title = "Toneladas por categoría"
        ch.y_axis.title = "Toneladas"
        ch.legend = None
        ch.add_data(Reference(ws, min_col=4, min_row=ini, max_row=fin), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=ini + 1, max_row=fin))
        ch.series[0].graphicalProperties.solidFill = "2A78D6"
        ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
        ch.height, ch.width = 8, 16
        ws.add_chart(ch, "H4")

        pie = PieChart()
        pie.title = "Participación por categoría"
        pie.add_data(Reference(ws, min_col=3, min_row=ini, max_row=fin), titles_from_data=True)
        pie.set_categories(Reference(ws, min_col=1, min_row=ini + 1, max_row=fin))
        pie.dataLabels = DataLabelList(); pie.dataLabels.showPercent = True
        pie.height, pie.width = 8, 12
        ws.add_chart(pie, "H21")
    _autoancho(ws)
    ws.column_dimensions["A"].width = 32

    # ---- Categoría x Corteza / Fibra
    ws2 = wb.create_sheet("Categoría-Corteza-Fibra")
    ws2["A1"] = "Toneladas por categoría y corteza"; ws2["A1"].font = Font(bold=True, size=12)
    t1 = resumen_cruzado(df, "Categoría", "Corteza").reset_index()
    f1 = _escribir_df(ws2, t1, 2)
    if len(t1) > 1 and len(t1.columns) > 2:
        ch = BarChart(); ch.type = "col"; ch.grouping = "stacked"; ch.overlap = 100
        ch.title = "Toneladas por categoría y corteza"
        ch.add_data(Reference(ws2, min_col=2, max_col=len(t1.columns) - 1, min_row=2, max_row=f1 - 1),
                    titles_from_data=True)
        ch.set_categories(Reference(ws2, min_col=1, min_row=3, max_row=f1 - 1))
        ch.height, ch.width = 8, 16
        ws2.add_chart(ch, f"A{f1 + 2}")
    r2 = f1 + 20
    ws2.cell(row=r2, column=1, value="Toneladas por producto (tipo de fibra y corteza)").font = Font(bold=True, size=12)
    t2 = resumen_cruzado(df, ["Tipo de fibra", "Corteza"], "Categoría").reset_index()
    _escribir_df(ws2, t2, r2 + 1)
    _autoancho(ws2)

    # ---- Mensual
    ws3 = wb.create_sheet("Mensual")
    ws3["A1"] = "Toneladas por mes y categoría"; ws3["A1"].font = Font(bold=True, size=12)
    tm = resumen_mensual(df).reset_index()
    if len(tm):
        tm["Total"] = tm.drop(columns="Mes").sum(axis=1)
        fm = _escribir_df(ws3, tm, 2)
        ch = BarChart(); ch.type = "col"; ch.grouping = "stacked"; ch.overlap = 100
        ch.title = "Toneladas por mes"; ch.y_axis.title = "Toneladas"
        ch.add_data(Reference(ws3, min_col=2, max_col=len(tm.columns) - 1, min_row=2, max_row=fm),
                    titles_from_data=True)
        ch.set_categories(Reference(ws3, min_col=1, min_row=3, max_row=fm))
        ch.height, ch.width = 9, 22
        ws3.add_chart(ch, f"A{fm + 3}")
    _autoancho(ws3)

    # ---- Por pila
    ws4 = wb.create_sheet("Por Pila")
    _escribir_df(ws4, resumen_pila(df))
    _autoancho(ws4)

    # ---- Reemplazos CAR
    ws5 = wb.create_sheet("Reemplazos CAR")
    rep = df[df["Reemplazo aplicado"].notna()].copy()
    _escribir_df(ws5, _para_excel(rep))
    _autoancho(ws5)

    # ---- Sin asignar
    from procesamiento import SIN_ASIGNAR
    ws6 = wb.create_sheet("Sin asignar")
    _escribir_df(ws6, _para_excel(df[df["Categoría"] == SIN_ASIGNAR]))
    _autoancho(ws6)

    # ---- Datos completos
    ws7 = wb.create_sheet("Datos")
    _escribir_df(ws7, _para_excel(df))
    for c in ws7["A"][1:]:
        c.number_format = "yyyy-mm-dd hh:mm"
    ws7.freeze_panes = "A2"
    ws7.auto_filter.ref = ws7.dimensions
    _autoancho(ws7)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _para_excel(df):
    d = df.copy()
    for c in d.columns:
        if str(d[c].dtype) == "string":
            d[c] = d[c].astype(object).where(d[c].notna(), None)
    d["Fecha"] = d["Fecha"].dt.to_pydatetime() if len(d) else d["Fecha"]
    return d


# ============================================================================ PDF
def _fig_png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


def _ejes_limpios(ax):
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color("#c9c8c2"); ax.spines["bottom"].set_color("#c9c8c2")
    ax.tick_params(colors=TINTA, labelsize=8)
    ax.grid(axis="y", color="#ecebe7", linewidth=0.8); ax.set_axisbelow(True)


def _graf_categoria(df, cmap):
    rc = resumen_categoria(df).sort_values("Toneladas")
    fig, ax = plt.subplots(figsize=(7, 3.2))
    ax.barh(rc["Categoría"], rc["Toneladas"], color=[cmap[c] for c in rc["Categoría"]],
            height=0.6, edgecolor="white", linewidth=2)
    for y, v, p in zip(range(len(rc)), rc["Toneladas"], rc["% del total"]):
        ax.text(v, y, f"  {v:,.1f} t ({p:.1f}%)", va="center", fontsize=8, color="#0b0b0b")
    _ejes_limpios(ax); ax.grid(axis="x", color="#ecebe7"); ax.grid(axis="y", visible=False)
    ax.set_xlabel("Toneladas", color=TINTA, fontsize=8)
    ax.set_xlim(0, rc["Toneladas"].max() * 1.3 if len(rc) else 1)
    ax.set_title("Consumo por categoría", loc="left", fontsize=10, color="#0b0b0b")
    return _fig_png(fig)


def _graf_mensual(df, cmap):
    tm = resumen_mensual(df)
    fig, ax = plt.subplots(figsize=(9, 3.4))
    base = None
    import numpy as np
    x = np.arange(len(tm))
    base = np.zeros(len(tm))
    for c in tm.columns:
        ax.bar(x, tm[c].values, bottom=base, color=cmap.get(c, "#9a9993"), label=c,
               width=0.65, edgecolor="white", linewidth=1.5)
        base += tm[c].values
    ax.set_xticks(x, tm.index, rotation=0)
    _ejes_limpios(ax); ax.set_ylabel("Toneladas", color=TINTA, fontsize=8)
    ax.legend(fontsize=7, frameon=False, ncol=min(5, len(tm.columns)), loc="upper left",
              bbox_to_anchor=(0, -0.12))
    ax.set_title("Toneladas por mes y categoría", loc="left", fontsize=10, color="#0b0b0b")
    return _fig_png(fig)


def _graf_dona(df, campo, cmap, titulo):
    g = df.groupby(campo)["Toneladas"].sum()
    g = g[g > 0]
    fig, ax = plt.subplots(figsize=(3.6, 3.2))
    if len(g):
        ax.pie(g.values, labels=[f"{k}\n{v:,.0f} t" for k, v in g.items()],
               colors=[cmap.get(k, "#9a9993") for k in g.index], autopct="%1.1f%%",
               pctdistance=0.78, startangle=90, counterclock=False,
               wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2),
               textprops=dict(fontsize=7, color="#0b0b0b"))
    ax.set_title(titulo, fontsize=10, color="#0b0b0b")
    return _fig_png(fig)


def _tabla(df_tab, anchos=None, fuente=7):
    datos = [list(map(str, df_tab.columns))]
    for row in df_tab.itertuples(index=False):
        fila = []
        for v in row:
            if v is None or v is pd.NA or (isinstance(v, float) and pd.isna(v)):
                fila.append("")
            elif isinstance(v, float):
                fila.append(f"{v:,.2f}")
            elif isinstance(v, int):
                fila.append(f"{v:,}")
            elif isinstance(v, pd.Timestamp):
                fila.append(v.strftime("%Y-%m-%d %H:%M"))
            else:
                fila.append(str(v))
        datos.append(fila)
    t = Table(datos, colWidths=anchos, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#" + AZUL)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), fuente),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f3f6fa")]),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#d9d9d9")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
    ]))
    return t


def generar_pdf(df: pd.DataFrame, filtros: dict, incluir_detalle: bool = True) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=1.2 * cm, rightMargin=1.2 * cm,
                            topMargin=1.2 * cm, bottomMargin=1.2 * cm,
                            title="Reporte de Consumos Picadora")
    ss = getSampleStyleSheet()
    h1 = ParagraphStyle("h1", parent=ss["Title"], textColor=colors.HexColor("#" + AZUL), alignment=0)
    h2 = ParagraphStyle("h2", parent=ss["Heading2"], textColor=colors.HexColor("#" + AZUL))
    normal = ParagraphStyle("n", parent=ss["Normal"], fontSize=9)
    el = [Paragraph("Reporte de Consumos Picadora", h1),
          Paragraph(f"Generado: {datetime.now():%Y-%m-%d %H:%M}", normal), Spacer(1, 6)]

    filt = pd.DataFrame(list(filtros.items()), columns=["Filtro", "Valor"])
    kp = pd.DataFrame([
        ("Registros", f"{len(df):,}"),
        ("Peso total (kg)", f"{df['Peso Líquido'].sum():,.0f}"),
        ("Toneladas", f"{df['Toneladas'].sum():,.2f}"),
        ("Registros con reemplazo CAR", f"{int(df['Reemplazo aplicado'].notna().sum()):,}"),
    ], columns=["Indicador", "Valor"])
    cab = Table([[_tabla(filt, [4 * cm, 9 * cm], 8), _tabla(kp, [5.5 * cm, 4 * cm], 9)]],
                colWidths=[13.5 * cm, 10 * cm])
    cab.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    el += [cab, Spacer(1, 10)]

    if len(df) == 0:
        el.append(Paragraph("No hay datos para los filtros seleccionados.", normal))
        doc.build(el)
        return buf.getvalue()

    cmap = colores_para(sorted(df["Categoría"].unique()))
    rc = resumen_categoria(df)
    graf = Table([[Image(_graf_categoria(df, cmap), width=15 * cm, height=6.9 * cm),
                   Image(_graf_dona(df, "Corteza", COLORES_CORTEZA, "Con / sin corteza"),
                         width=7.6 * cm, height=6.8 * cm)]])
    el += [graf, Spacer(1, 6), Paragraph("Consumo por categoría", h2),
           _tabla(rc, [6 * cm, 3 * cm, 5 * cm, 4 * cm, 3 * cm], 8), PageBreak()]

    el += [Image(_graf_mensual(df, cmap), width=24 * cm, height=9.5 * cm), Spacer(1, 4),
           Table([[Image(_graf_dona(df, "Tipo de fibra", COLORES_FIBRA, "Tipo de fibra"),
                         width=7.6 * cm, height=6.8 * cm),
                   _tabla(resumen_cruzado(df, "Categoría", "Corteza").reset_index(), None, 8)]],
                 style=[("VALIGN", (0, 0), (-1, -1), "MIDDLE")]),
           PageBreak()]

    el += [Paragraph("Toneladas por mes y categoría", h2),
           _tabla(_con_total(resumen_mensual(df)).reset_index(), None, 8), Spacer(1, 10),
           Paragraph("Toneladas por tipo de fibra y corteza", h2),
           _tabla(resumen_cruzado(df, ["Tipo de fibra", "Corteza"], "Categoría").reset_index(), None, 8),
           PageBreak(),
           Paragraph("Consumo por pila", h2),
           _tabla(resumen_pila(df), [6 * cm, 4 * cm, 3 * cm, 5 * cm, 4 * cm], 7)]

    rep = df[df["Reemplazo aplicado"].notna()]
    if len(rep):
        el += [PageBreak(), Paragraph(f"Reemplazos CAR aplicados ({len(rep):,} registros)", h2),
               _tabla(rep.groupby("Reemplazo aplicado").agg(
                   Registros=("Toneladas", "size"), Toneladas=("Toneladas", "sum")).reset_index(),
                   [12 * cm, 3 * cm, 4 * cm], 8)]

    if incluir_detalle:
        cols = ["Fecha", "Producto", "Tipo de fibra", "Corteza", "Pila original", "Pila",
                "Categoría", "Peso Líquido", "Toneladas"]
        det = df.sort_values("Fecha")[cols]
        el += [PageBreak(), Paragraph(f"Detalle de registros ({len(det):,})", h2),
               _tabla(det, [3 * cm, 4.2 * cm, 2.4 * cm, 2.4 * cm, 2.4 * cm, 2.6 * cm, 3.8 * cm,
                            2.6 * cm, 2.2 * cm], 6.5)]

    def pie_pagina(canvas, d):
        canvas.saveState(); canvas.setFont("Helvetica", 7); canvas.setFillColor(colors.grey)
        canvas.drawRightString(d.pagesize[0] - 1.2 * cm, 0.7 * cm, f"Página {d.page}")
        canvas.drawString(1.2 * cm, 0.7 * cm, "Consumos Picadora")
        canvas.restoreState()

    doc.build(el, onFirstPage=pie_pagina, onLaterPages=pie_pagina)
    return buf.getvalue()


def _con_total(t):
    t = t.copy()
    t["Total"] = t.sum(axis=1)
    return t
