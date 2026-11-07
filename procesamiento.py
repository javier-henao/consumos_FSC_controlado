"""Lógica de carga, reemplazo de CAR y clasificación de pilas."""
from __future__ import annotations

import copy
import io
import json
import os
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path

import pandas as pd

# En Railway, defina CONFIG_DIR apuntando a un volumen (ej. /data) para que la
# configuración guardada sobreviva a los redespliegues.
CONFIG_PATH = Path(os.environ.get("CONFIG_DIR", Path(__file__).parent)) / "config_asignaciones.json"

SIN_ASIGNAR = "Sin asignar"

# Colores fijos por categoría (siguen a la entidad, no al orden)
COLORES_CATEGORIA = {
    "Eucalipto Controlado": "#2a78d6",
    "Eucalipto FSC": "#eb6834",
    "Pino Controlado": "#1baf7a",
    "Pino FSC": "#eda100",
}
COLORES_EXTRA = ["#e87ba4", "#008300", "#6250d6", "#e34948"]
COLOR_SIN_ASIGNAR = "#9a9993"
COLORES_CORTEZA = {"Con corteza": "#2a78d6", "Sin corteza": "#eb6834", "N/A": COLOR_SIN_ASIGNAR}
COLORES_FIBRA = {"Fibra Corta": "#2a78d6", "Fibra Larga": "#1baf7a", "Otro": COLOR_SIN_ASIGNAR}

CONFIG_DEFECTO = {
    "reemplazos": [
        {"pila_original": "CAR", "producto": "Fibra Corta CON CRT", "nuevo_valor": "EU CAR MENOR"},
        {"pila_original": "CAR", "producto": "Fibra Corta SIN CRT", "nuevo_valor": "EU CAR MENOR"},
        {"pila_original": "CAR", "producto": "Fibra Larga CON CRT", "nuevo_valor": "PN CAR MEN"},
        {"pila_original": "CAR", "producto": "Fibra Larga SIN CRT", "nuevo_valor": "PN CAR MEN"},
    ],
    "categorias": [
        {"categoria": "Eucalipto Controlado", "rango_min": 100, "rango_max": 199,
         "codigos": "EU CAR CNT, T168, R169, T178, R179", "notas": ""},
        {"categoria": "Eucalipto FSC", "rango_min": 300, "rango_max": 399,
         "codigos": "EU CAR MEN, EU CAR MENOR, EU CAR MAY, T330, R331",
         "notas": "Incluye Patio Oriente (305)"},
        {"categoria": "Pino Controlado", "rango_min": 200, "rango_max": 299,
         "codigos": "PN CAR CNT, PN CAR CONTROL, T208, R209, T218, R219", "notas": ""},
        {"categoria": "Pino FSC", "rango_min": 400, "rango_max": 499,
         "codigos": "PN CAR MEN, PN CAR MAY, T428, R429", "notas": ""},
        {"categoria": "Orillos", "rango_min": None, "rango_max": None,
         "codigos": "2000", "notas": "Pila 2000"},
    ],
}


# ----------------------------------------------------------------------------- config
def config_defecto() -> dict:
    return copy.deepcopy(CONFIG_DEFECTO)


def cargar_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return validar_config(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
        except Exception:
            pass
    return config_defecto()


def guardar_config(cfg: dict) -> None:
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def validar_config(cfg: dict) -> dict:
    if not isinstance(cfg, dict) or "categorias" not in cfg or "reemplazos" not in cfg:
        raise ValueError("La configuración debe tener las claves 'categorias' y 'reemplazos'.")
    cats = []
    for c in cfg["categorias"]:
        nombre = str(c.get("categoria") or "").strip()
        if not nombre:
            continue
        cats.append({
            "categoria": nombre,
            "rango_min": _num_o_none(c.get("rango_min")),
            "rango_max": _num_o_none(c.get("rango_max")),
            "codigos": str(c.get("codigos") or ""),
            "notas": str(c.get("notas") or ""),
        })
    reps = []
    for r in cfg["reemplazos"]:
        o, p, n = (str(r.get(k) or "").strip() for k in ("pila_original", "producto", "nuevo_valor"))
        if o and p and n:
            reps.append({"pila_original": o, "producto": p, "nuevo_valor": n})
    return {"categorias": cats, "reemplazos": reps}


def _num_o_none(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
            return None
        return int(float(v))
    except (TypeError, ValueError):
        return None


def colores_para(categorias: list[str]) -> dict:
    """Asigna color fijo a cada categoría; las nuevas toman los siguientes tonos en orden."""
    out, extra = {}, iter(COLORES_EXTRA)
    for c in categorias:
        if c == SIN_ASIGNAR:
            out[c] = COLOR_SIN_ASIGNAR
        elif c in COLORES_CATEGORIA:
            out[c] = COLORES_CATEGORIA[c]
        else:
            out[c] = next(extra, COLOR_SIN_ASIGNAR)
    return out


# ----------------------------------------------------------------------------- datos
# Campos que necesita la app y los nombres aceptados para cada uno (ya normalizados:
# sin prefijo "TABLA.", sin tildes, sin espacios/guiones/símbolos, en mayúsculas).
CAMPOS = {
    "Producto": ["DCRPRODUTOORIGEM", "PRODUTOORIGEM", "DESCPRODUTOORIGEM", "PRODUCTOORIGEN",
                 "DCRPRODUTO", "PRODUTO", "PRODUCTO", "DESCRIPCIONPRODUCTO", "TIPOFIBRA"],
    "Fecha": ["DATAINICIO", "DATAINICIAL", "FECHAINICIO", "FECHAINICIAL", "DATA", "FECHA"],
    "Hora": ["HORAINICIO", "HORAINICIAL", "HORA"],
    "Pila": ["NUMPILHAORIGEM", "PILHAORIGEM", "NUMPILHA", "PILHA", "NUMPILAORIGEN",
             "PILAORIGEN", "NUMPILA", "PILA"],
    "Peso": ["PESOLIQUIDO", "PESONETO", "PESOLIQ", "PESO"],
}
# Palabras clave para coincidencias parciales, por si el nombre trae texto extra
CLAVES = {
    "Producto": ["PRODUT", "PRODUCT"],
    "Fecha": ["DATA", "FECHA"],
    "Hora": ["HORA"],
    "Pila": ["PILHA", "PILA"],
    "Peso": ["PESO"],
}


def normalizar_nombre(nombre) -> str:
    s = str(nombre).strip()
    if "." in s:                       # #VW_CUBO_DADOS_ITEM_BOLETIM.DCR_PRODUTO_ORIGEM -> DCR_PRODUTO_ORIGEM
        s = s.rsplit(".", 1)[-1]
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9]", "", s).upper()


def _puntaje_encabezado(valores) -> int:
    """Cuántos campos se reconocen en una fila candidata a encabezado."""
    nombres = [normalizar_nombre(v) for v in valores if pd.notna(v)]
    return sum(any(n in alias for n in nombres) for alias in CAMPOS.values())


def leer_tabla(archivo) -> pd.DataFrame:
    """Lee el archivo (.xls, .xlsx, o .xls que en realidad es HTML) y ubica la fila de encabezados."""
    nombre = getattr(archivo, "name", str(archivo)).lower()
    datos = archivo.getvalue() if hasattr(archivo, "getvalue") else Path(archivo).read_bytes()

    hojas = []
    for engine in (("xlrd" if nombre.endswith(".xls") else "openpyxl"), None):
        try:
            hojas = list(pd.read_excel(io.BytesIO(datos), engine=engine, header=None,
                                       sheet_name=None).values())
            break
        except Exception:
            continue
    if not hojas:  # algunos sistemas exportan HTML con extensión .xls
        try:
            hojas = []
            for h in pd.read_html(io.BytesIO(datos)):
                # read_html usa los <th> como nombres de columna: devolverlos a la 1a fila
                # para que la búsqueda de encabezado funcione igual que en Excel.
                nombres = [" ".join(map(str, c)) if isinstance(c, tuple) else c for c in h.columns]
                h = pd.concat([pd.DataFrame([nombres]), pd.DataFrame(h.values)], ignore_index=True)
                hojas.append(h)
        except Exception as e:
            raise ValueError(f"No se reconoce el formato del archivo: {e}") from e

    # Elegir la hoja y la fila (entre las primeras 30) que más parece encabezado
    mejor = (-1, None, 0)
    for h in hojas:
        for i in range(min(30, len(h))):
            p = _puntaje_encabezado(h.iloc[i].tolist())
            if p > mejor[0]:
                mejor = (p, h, i)
    puntaje, hoja, fila = mejor
    if hoja is None:
        raise ValueError("El archivo está vacío.")
    if puntaje == 0:   # sin encabezado reconocible: usar la primera fila
        fila = 0
    tabla = hoja.iloc[fila + 1:].reset_index(drop=True)
    tabla.columns = _nombres_unicos(hoja.iloc[fila].tolist())
    tabla = tabla.dropna(axis=1, how="all").dropna(axis=0, how="all")
    return tabla.reset_index(drop=True)


def _nombres_unicos(nombres):
    vistos, out = {}, []
    for i, n in enumerate(nombres):
        n = str(n).strip() if pd.notna(n) and str(n).strip() else f"Columna {i + 1}"
        if n in vistos:
            vistos[n] += 1
            n = f"{n} ({vistos[n]})"
        else:
            vistos[n] = 0
        out.append(n)
    return out


def detectar_columnas(tabla: pd.DataFrame) -> dict:
    """Devuelve {campo: nombre_de_columna}. Busca por nombre y, si no, por contenido."""
    norm = {c: normalizar_nombre(c) for c in tabla.columns}
    asignadas, res = set(), {}

    # 1) nombre exacto (en orden de preferencia)  2) nombre que contiene la palabra clave
    for campo, alias in CAMPOS.items():
        for a in alias:
            c = next((c for c, n in norm.items() if n == a and c not in asignadas), None)
            if c:
                res[campo] = c; asignadas.add(c); break
    for campo, claves in CLAVES.items():
        if campo in res:
            continue
        c = next((c for c, n in norm.items() if c not in asignadas
                  and any(k in n for k in claves)), None)
        if c:
            res[campo] = c; asignadas.add(c)

    # 3) por contenido
    libres = [c for c in tabla.columns if c not in asignadas]
    muestra = {c: tabla[c].dropna().head(300) for c in libres}

    def frac(c, fn):
        s = muestra[c]
        return fn(s).mean() if len(s) else 0

    pruebas = {
        "Producto": lambda s: s.astype(str).str.upper().str.contains("FIBRA|CRT"),
        "Pila": lambda s: s.astype(str).str.strip().str.upper().str.fullmatch(
            r"\d{3,4}(\.0)?|CAR|[TR]\d{3}|(EU|PN) CAR.*"),
        "Hora": lambda s: s.astype(str).str.strip().str.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?"),
        "Fecha": lambda s: _a_fecha(s).notna()
                           & ~s.astype(str).str.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?|\d+(\.\d+)?"),
        "Peso": lambda s: pd.to_numeric(s.astype(str).str.replace(r"[.,\s]", "", regex=True),
                                        errors="coerce").notna(),
    }
    for campo in ["Producto", "Hora", "Fecha", "Pila", "Peso"]:
        if campo in res:
            continue
        cand = [(frac(c, pruebas[campo]), c) for c in libres if c not in asignadas]
        cand = [x for x in cand if x[0] >= 0.8]
        if cand:
            c = max(cand)[1]
            res[campo] = c; asignadas.add(c)
    return res


def normalizar(tabla: pd.DataFrame, mapeo: dict) -> pd.DataFrame:
    """Convierte la tabla original al formato interno usando el mapeo de columnas."""
    faltan = [c for c in ("Producto", "Fecha", "Pila", "Peso") if not mapeo.get(c)]
    if faltan:
        raise KeyError("Faltan columnas: " + ", ".join(faltan))
    col = lambda k: tabla[mapeo[k]] if mapeo.get(k) else pd.Series(pd.NA, index=tabla.index)

    df = pd.DataFrame({
        "Producto": col("Producto"),
        "Fecha": _a_fecha(col("Fecha")),
        "Hora": col("Hora"),
        "Pila original": col("Pila"),
        "Peso Líquido": _a_numero(col("Peso")),
    })
    df["Producto"] = df["Producto"].astype("string").str.strip().replace("", pd.NA)
    # Reportes agrupados: el producto solo viene en la primera fila del grupo
    con_dato = df["Fecha"].notna()
    if df.loc[con_dato, "Producto"].isna().mean() > 0.2:
        df["Producto"] = df["Producto"].ffill()
    # Filas sin fecha = separadoras o subtotales de reportes agrupados: se descartan
    # para no sumar dos veces. Se guarda cuántas fueron para mostrarlo en la app.
    descartadas = int((~con_dato).sum())
    df = df[con_dato]
    df = df[~df["Producto"].fillna("").str.upper().str.startswith("TOTAL")]
    df["Pila original"] = df["Pila original"].map(_normalizar_pila)
    df["Hora"] = df["Hora"].map(_a_hora).astype("string")
    if df["Hora"].isna().all():
        df["Hora"] = df["Fecha"].dt.strftime("%H:%M").astype("string")
    df = df.reset_index(drop=True)
    df.attrs["filas_descartadas"] = descartadas
    return df


def _a_fecha(s: pd.Series) -> pd.Series:
    """Convierte a fecha aceptando: fechas de Excel, números seriales, texto ISO
    (2026-01-13 00:01) y texto día/mes/año (13/01/2026 00:01)."""
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    es_fecha = s.map(lambda v: isinstance(v, (datetime, date)) and not pd.isna(v))
    if es_fecha.any():
        out[es_fecha] = pd.to_datetime(s[es_fecha].tolist(), errors="coerce")
    resto = s[~es_fecha & s.notna()]
    num = pd.to_numeric(resto, errors="coerce")
    serial = num.between(20000, 80000)                     # fecha serial de Excel
    if serial.any():
        out[serial[serial].index] = pd.to_datetime(num[serial], unit="D", origin="1899-12-30")
    txt = resto[num.isna()].astype(str).str.strip()
    # Solo se interpreta texto con forma de fecha real (evita que "T330" se lea como año 330)
    iso = txt.str.match(r"^\d{4}-\d{1,2}-\d{1,2}([ T]\d{1,2}:\d{2}(:\d{2})?)?(\.\d+)?$")
    dmy = txt.str.match(r"^\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}( \d{1,2}:\d{2}(:\d{2})?)?$")
    for mascara, kw in ((iso, dict(format="ISO8601")),                # 2026-01-13 00:01
                        (dmy, dict(dayfirst=True, format="mixed"))):  # 13/01/2026 00:01
        if mascara.any():
            v = pd.to_datetime(txt[mascara], errors="coerce", **kw)
            out[v.index] = v.astype("datetime64[ns]")
    return out


def _a_numero(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0)

    def conv(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        if isinstance(v, (int, float)):
            return float(v)
        t = re.sub(r"[^\d,.\-]", "", str(v))
        if "," in t and "." in t:                             # 1.234,56 o 1,234.56
            t = t.replace(".", "").replace(",", ".") if t.rfind(",") > t.rfind(".") \
                else t.replace(",", "")
        elif "," in t:
            t = t.replace(",", "") if re.fullmatch(r"-?\d{1,3}(,\d{3})+", t) else t.replace(",", ".")
        elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", t):         # 962.080 -> 962080
            t = t.replace(".", "")
        try:
            return float(t)
        except ValueError:
            return 0.0
    return s.map(conv)


def _a_hora(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return pd.NA
    if hasattr(v, "strftime"):
        return v.strftime("%H:%M")
    if isinstance(v, float) and 0 <= v < 1:                  # fracción de día de Excel
        m = round(v * 24 * 60)
        return f"{m // 60:02d}:{m % 60:02d}"
    s = str(v).strip()
    return s or pd.NA


def leer_excel(archivo) -> pd.DataFrame:
    """Lectura automática completa (sin intervención del usuario)."""
    tabla = leer_tabla(archivo)
    return normalizar(tabla, detectar_columnas(tabla))


def _normalizar_pila(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return pd.NA
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = re.sub(r"\s+", " ", str(v)).strip()
    return s if s else pd.NA


def _parsear_codigos(texto: str) -> list[str]:
    return [re.sub(r"\s+", " ", c).strip().upper() for c in str(texto).split(",") if c.strip()]


def clasificar_pila(pila, categorias: list[dict]) -> str:
    if pila is None or pd.isna(pila):
        return SIN_ASIGNAR
    p = str(pila).strip().upper()
    # 1) códigos explícitos
    for c in categorias:
        if p in _parsear_codigos(c["codigos"]):
            return c["categoria"]
    # 2) rangos numéricos
    if re.fullmatch(r"\d+(\.0+)?", p):
        n = int(float(p))
        for c in categorias:
            lo, hi = c.get("rango_min"), c.get("rango_max")
            if lo is not None and hi is not None and lo <= n <= hi:
                return c["categoria"]
    return SIN_ASIGNAR


def procesar(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    df = df.copy()
    pila = df["Pila original"].astype("string")
    prod = df["Producto"].astype("string").str.upper()
    nueva = pila.copy()
    regla = pd.Series(pd.NA, index=df.index, dtype="string")
    for r in cfg["reemplazos"]:
        m = (pila.str.upper() == r["pila_original"].upper()) & (prod == r["producto"].strip().upper())
        m = m.fillna(False) & regla.isna()
        nueva[m] = r["nuevo_valor"]
        regla[m] = f'{r["pila_original"]} + {r["producto"]} → {r["nuevo_valor"]}'
    df["Pila"] = nueva
    df["Reemplazo aplicado"] = regla
    df["Categoría"] = df["Pila"].map(lambda x: clasificar_pila(x, cfg["categorias"]))

    pu = df["Producto"].fillna("").str.upper()
    df["Tipo de fibra"] = "Otro"
    df.loc[pu.str.contains("FIBRA CORTA"), "Tipo de fibra"] = "Fibra Corta"
    df.loc[pu.str.contains("FIBRA LARGA"), "Tipo de fibra"] = "Fibra Larga"
    df["Corteza"] = "N/A"
    df.loc[pu.str.contains("CON CRT"), "Corteza"] = "Con corteza"
    df.loc[pu.str.contains("SIN CRT"), "Corteza"] = "Sin corteza"
    df["Mes"] = df["Fecha"].dt.to_period("M").astype("string")
    df["Toneladas"] = df["Peso Líquido"] / 1000
    orden = ["Fecha", "Hora", "Mes", "Producto", "Tipo de fibra", "Corteza", "Pila original",
             "Pila", "Reemplazo aplicado", "Categoría", "Peso Líquido", "Toneladas"]
    return df[orden]


# ----------------------------------------------------------------------------- resúmenes
def resumen_categoria(df):
    g = df.groupby("Categoría", dropna=False).agg(
        Registros=("Peso Líquido", "size"), Peso_kg=("Peso Líquido", "sum")).reset_index()
    total = g["Peso_kg"].sum()
    g["Toneladas"] = g["Peso_kg"] / 1000
    g["% del total"] = (g["Peso_kg"] / total * 100).round(2) if total else 0
    return g.sort_values("Peso_kg", ascending=False).rename(columns={"Peso_kg": "Peso Líquido (kg)"})


def resumen_cruzado(df, filas, columnas):
    t = pd.pivot_table(df, index=filas, columns=columnas, values="Peso Líquido",
                       aggfunc="sum", fill_value=0, margins=True, margins_name="Total")
    return (t / 1000).round(2)


def resumen_mensual(df):
    return pd.pivot_table(df, index="Mes", columns="Categoría", values="Peso Líquido",
                          aggfunc="sum", fill_value=0).div(1000).round(2).sort_index()


def resumen_pila(df):
    g = df.groupby(["Categoría", "Pila"], dropna=False).agg(
        Registros=("Peso Líquido", "size"), Peso_kg=("Peso Líquido", "sum")).reset_index()
    g["Toneladas"] = (g["Peso_kg"] / 1000).round(2)
    return g.sort_values(["Categoría", "Peso_kg"], ascending=[True, False]).rename(
        columns={"Peso_kg": "Peso Líquido (kg)"})
