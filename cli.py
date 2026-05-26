"""
Interfaz de linea de comandos del automatizador de informes.

Permite ejecutar todo el pipeline (cargar -> limpiar -> analizar ->
informe) desde la terminal, con opciones por argumentos o desde un
archivo de configuracion YAML. Ejemplos:

    # Uso minimo: un archivo de entrada, un informe de salida.
    python -m excel_automator.cli datos.csv -o informe.xlsx

    # Con agregacion (ventas por categoria) y columnas numericas.
    python -m excel_automator.cli ventas.csv -o informe.xlsx \\
        --numeric importe --group-by categoria --value importe

    # Usando un archivo de configuracion.
    python -m excel_automator.cli --config config.yaml
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import analyzer, cleaner, loader, reporter


def _load_config(path: str) -> dict:
    """Carga un archivo de configuracion YAML (opcional)."""
    try:
        import yaml
    except ImportError:
        sys.exit("Para usar --config necesitas instalar PyYAML: pip install pyyaml")
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def run(
    inputs: list[str],
    output: str,
    *,
    numeric: list[str] | None = None,
    dates: list[str] | None = None,
    group_by: str | None = None,
    value: str | None = None,
    aggfunc: str = "sum",
    dedupe: list[str] | None = None,
) -> Path:
    """Ejecuta el pipeline completo y devuelve la ruta del informe."""
    print(f"[1/4] Cargando {len(inputs)} archivo(s)...")
    df = loader.load_many(inputs)

    print("[2/4] Limpiando datos...")
    df_clean, report = cleaner.clean(
        df, numeric_columns=numeric, date_columns=dates, dedupe_subset=dedupe
    )
    print(
        f"      {report['filas_iniciales']} -> {report['filas_finales']} filas "
        f"({report['duplicados_eliminados']} duplicados, "
        f"{report['filas_vacias_eliminadas']} vacias)"
    )

    print("[3/4] Calculando metricas...")
    num_summary = analyzer.numeric_summary(df_clean)
    clean_report_df = analyzer.build_clean_report_sheet(report)

    aggregation = None
    agg_title = "Agregacion"
    if group_by and value:
        aggregation = analyzer.group_aggregate(df_clean, group_by, value, aggfunc)
        agg_title = f"{value} ({aggfunc}) por {group_by}"

    print("[4/4] Generando informe Excel...")
    path = reporter.build_report(
        output,
        clean_data=df_clean.drop(columns=["_origen"], errors="ignore"),
        clean_report=clean_report_df,
        numeric_summary=num_summary,
        aggregation=aggregation,
        aggregation_title=agg_title,
    )
    print(f"\nInforme generado: {path.resolve()}")
    return path


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        description="Limpia datos y genera un informe Excel con formato."
    )
    parser.add_argument("inputs", nargs="*", help="Archivos de entrada (CSV/Excel)")
    parser.add_argument("-o", "--output", default="informe.xlsx", help="Archivo de salida")
    parser.add_argument("--numeric", nargs="*", help="Columnas a forzar a numero")
    parser.add_argument("--dates", nargs="*", help="Columnas a convertir a fecha")
    parser.add_argument("--group-by", help="Columna por la que agrupar")
    parser.add_argument("--value", help="Columna numerica a agregar")
    parser.add_argument("--aggfunc", default="sum", help="Funcion de agregacion (sum, mean, count...)")
    parser.add_argument("--dedupe", nargs="*", help="Columnas que definen un duplicado")
    parser.add_argument("--config", help="Archivo de configuracion YAML")

    args = parser.parse_args(argv)

    if args.config:
        cfg = _load_config(args.config)
        run(
            inputs=cfg["inputs"],
            output=cfg.get("output", "informe.xlsx"),
            numeric=cfg.get("numeric"),
            dates=cfg.get("dates"),
            group_by=cfg.get("group_by"),
            value=cfg.get("value"),
            aggfunc=cfg.get("aggfunc", "sum"),
            dedupe=cfg.get("dedupe"),
        )
        return

    if not args.inputs:
        parser.error("Debes indicar al menos un archivo de entrada o usar --config.")

    run(
        inputs=args.inputs,
        output=args.output,
        numeric=args.numeric,
        dates=args.dates,
        group_by=args.group_by,
        value=args.value,
        aggfunc=args.aggfunc,
        dedupe=args.dedupe,
    )


if __name__ == "__main__":
    main()
