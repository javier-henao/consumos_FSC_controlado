"""Lógica de carga, reemplazo de CAR y clasificación de pilas."""
from __future__ import annotations

import copy
import json
import os
import re
from pathlib import Path

import pandas as pd

# En Railway, defina CONFIG_DIR apuntando a un volumen (ej. /data) para que la
# configuración guardada sobreviva a los redespliegues.
CONFIG_PATH = Path(os.environ.get("CONFIG_DIR", Path(__file__).parent)) / "config_asignaciones.json"

# Sufijos de las columnas originales (el prefijo #VW_CUBO_DADOS_ITEM_BOLETIM. puede variar)
COL_PRODUCTO = "DCR_PRODUTO_ORIGEM"
COL_FECHA = "DATA_INICIO"
COL_HORA = "HORA_INICIO"
COL_PILA = "NUM_PILHA_ORIGEM"
COL_PESO = "Peso Líquido"

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
         "codigos": "PN CAR CNT, T208, R209, T218, R219", "notas": ""},
        {"categoria": "Pino FSC", "rango_min": 400, "rango_max": 499,
         "codigos": "PN CAR MEN, PN CAR MAY, T428, R429", "notas": ""},
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
def _buscar_columna(cols, sufijo):
    for c in cols:
        if str(c).strip().upper().endswith(sufijo.upper()):
            return c
    raise KeyError(f"No se encontró la columna que termina en '{sufijo}'.")


def leer_excel(archivo) -> pd.DataFrame:
    """Lee .xls/.xlsx sin importar el nombre del archivo y normaliza las columnas."""
    nombre = getattr(archivo, "name", str(archivo)).lower()
    engine = "xlrd" if nombre.endswith(".xls") else None
    try:
        raw = pd.read_excel(archivo, engine=engine)
    except Exception:
        if hasattr(archivo, "seek"):
            archivo.seek(0)
        raw = pd.read_excel(archivo)
    cols = list(raw.columns)
    df = pd.DataFrame({
        "Producto": raw[_buscar_columna(cols, COL_PRODUCTO)],
        "Fecha": pd.to_datetime(raw[_buscar_columna(cols, COL_FECHA)], errors="coerce"),
        "Hora": raw[_buscar_columna(cols, COL_HORA)],
        "Pila original": raw[_buscar_columna(cols, COL_PILA)],
        "Peso Líquido": pd.to_numeric(raw[_buscar_columna(cols, COL_PESO)], errors="coerce").fillna(0),
    })
    df["Producto"] = df["Producto"].astype("string").str.strip()
    df["Pila original"] = df["Pila original"].map(_normalizar_pila)
    df["Hora"] = df["Hora"].astype("string").str.strip()
    return df


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
