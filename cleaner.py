"""
Limpieza y normalizacion de datos.

Agrupa las operaciones de limpieza mas habituales que pide un cliente:
nombres de columna consistentes, eliminacion de duplicados y filas vacias,
recorte de espacios y conversion de tipos. Cada paso es opcional y se
controla desde la configuracion.
"""
from __future__ import annotations

import re
import unicodedata

import pandas as pd


def _slugify_column(name: str) -> str:
    """Convierte un nombre de columna a formato consistente.

    'Importe Total (€)' -> 'importe_total'
    """
    text = str(name).strip().lower()
    # Quita acentos.
    text = "".join(
        c for c in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(c)
    )
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza los nombres de todas las columnas."""
    df = df.copy()
    df.columns = [_slugify_column(c) for c in df.columns]
    return df


def strip_whitespace(df: pd.DataFrame) -> pd.DataFrame:
    """Recorta espacios sobrantes en todas las columnas de texto."""
    df = df.copy()
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()
        # Restaura los nulos que se hayan convertido a la cadena 'nan'.
        df[col] = df[col].replace({"nan": pd.NA, "None": pd.NA, "": pd.NA})
    return df


def drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Elimina filas completamente vacias (ignorando la columna _origen)."""
    cols = [c for c in df.columns if c != "_origen"]
    return df.dropna(how="all", subset=cols).reset_index(drop=True)


def drop_duplicates(df: pd.DataFrame, subset: list[str] | None = None) -> pd.DataFrame:
    """Elimina filas duplicadas.

    Args:
        df: DataFrame a procesar.
        subset: Columnas que definen un duplicado. Si es None, usa todas
            (excepto la columna interna _origen).
    """
    if subset is None:
        subset = [c for c in df.columns if c != "_origen"]
    return df.drop_duplicates(subset=subset).reset_index(drop=True)


def coerce_numeric(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Fuerza columnas a tipo numerico, tolerando formatos sucios.

    Maneja separadores de miles, simbolos de moneda y comas decimales
    al estilo espanol ('1.234,56 €' -> 1234.56).
    """
    df = df.copy()
    for col in columns:
        if col not in df.columns:
            continue
        series = (
            df[col]
            .astype(str)
            .str.replace(r"[^\d,.\-]", "", regex=True)  # quita € y letras
            .str.replace(".", "", regex=False)          # quita miles
            .str.replace(",", ".", regex=False)          # coma -> punto decimal
        )
        df[col] = pd.to_numeric(series, errors="coerce")
    return df


def coerce_dates(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Convierte columnas a fecha, tolerando varios formatos."""
    df = df.copy()
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
    return df


def clean(
    df: pd.DataFrame,
    *,
    numeric_columns: list[str] | None = None,
    date_columns: list[str] | None = None,
    dedupe_subset: list[str] | None = None,
) -> tuple[pd.DataFrame, dict]:
    """Ejecuta el pipeline completo de limpieza.

    Returns:
        Una tupla (df_limpio, informe) donde 'informe' resume cuantas
        filas se eliminaron en cada paso. Ese resumen luego se vuelca en
        el Excel para que el cliente vea exactamente que se ha hecho.
    """
    report = {"filas_iniciales": len(df)}

    df = normalize_columns(df)
    df = strip_whitespace(df)

    before = len(df)
    df = drop_empty_rows(df)
    report["filas_vacias_eliminadas"] = before - len(df)

    before = len(df)
    df = drop_duplicates(df, subset=dedupe_subset)
    report["duplicados_eliminados"] = before - len(df)

    if numeric_columns:
        df = coerce_numeric(df, numeric_columns)
    if date_columns:
        df = coerce_dates(df, date_columns)

    report["filas_finales"] = len(df)
    return df, report
