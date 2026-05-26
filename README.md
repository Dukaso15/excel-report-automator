# 📊 Excel Report Automator

Herramienta de línea de comandos en Python que convierte archivos de datos
desordenados (CSV / Excel) en **informes Excel con formato profesional** de
forma automática: limpia los datos, calcula métricas y genera un libro con
varias hojas, estilos, formato condicional y gráficos.

> Pensada para automatizar la tarea repetitiva de "tengo estos datos sucios y
> necesito un informe presentable cada mes". Lo que a mano lleva horas, aquí es
> un comando.

---

## ✨ Qué hace

- **Carga** uno o varios archivos CSV/Excel y los combina (detecta el separador
  automáticamente).
- **Limpia** los datos: normaliza nombres de columna, elimina duplicados y filas
  vacías, recorta espacios y convierte importes al estilo español
  (`1.234,56 €` → `1234.56`) y fechas a tipo fecha real.
- **Analiza**: resumen estadístico (suma, media, mín., máx.) y agregaciones por
  categoría (p. ej. ventas totales por producto).
- **Genera** un informe `.xlsx` con tres hojas:
  - **Resumen** — log de limpieza + estadísticas.
  - **Datos limpios** — la tabla depurada, con cabecera fija y bandas.
  - **Agregación** — tabla agrupada con formato condicional y gráfico de barras.

---

## 🚀 Instalación

```bash
git clone https://github.com/Dukaso15/excel-report-automator.git
cd excel-report-automator
pip install -r requirements.txt
```

## 🧪 Prueba en 30 segundos

```bash
# 1. Genera un CSV de ejemplo deliberadamente "sucio"
python examples/generate_sample_data.py

# 2. Genera el informe
python -m excel_automator.cli examples/ventas_ejemplo.csv \
    -o examples/informe_demo.xlsx \
    --numeric importe_total --dates fecha \
    --group-by categoria --value importe_total --dedupe id_pedido
```

Abre `examples/informe_demo.xlsx` y tendrás el informe completo.

---

## ⚙️ Uso

### Por argumentos

```bash
python -m excel_automator.cli ENTRADA [ENTRADA ...] -o salida.xlsx [opciones]
```

| Opción         | Descripción                                            |
|----------------|--------------------------------------------------------|
| `-o, --output` | Archivo Excel de salida (por defecto `informe.xlsx`).  |
| `--numeric`    | Columnas a forzar a número (tolera `€`, miles, comas). |
| `--dates`      | Columnas a convertir a fecha.                          |
| `--group-by`   | Columna por la que agrupar en la hoja de agregación.   |
| `--value`      | Columna numérica a agregar.                            |
| `--aggfunc`    | Función de agregación: `sum`, `mean`, `count`, `max`…  |
| `--dedupe`     | Columnas que definen un duplicado.                     |
| `--config`     | Usar un archivo de configuración YAML.                 |

### Por archivo de configuración

```bash
python -m excel_automator.cli --config config.example.yaml
```

Ver [`config.example.yaml`](config.example.yaml) como plantilla.

---

## 🧩 Estructura

```
excel_automator/
├── loader.py     # carga de CSV/Excel
├── cleaner.py    # limpieza y normalización
├── analyzer.py   # métricas y agregaciones
├── reporter.py   # generación del Excel con formato
└── cli.py        # interfaz de línea de comandos
```

Cada módulo es independiente y reutilizable: puedes importar `cleaner.clean()`
o `reporter.build_report()` en tus propios scripts.

```python
from excel_automator import loader, cleaner, analyzer, reporter

df = loader.load_file("datos.csv")
df_clean, report = cleaner.clean(df, numeric_columns=["importe"])
resumen = analyzer.numeric_summary(df_clean)
# ...
```

---

## 🎨 Personalización

Los colores del informe se cambian en dos constantes al inicio de
`reporter.py` (`HEADER_FILL` y `BAND_FILL`) para adaptarlo a la marca de
cada cliente.

---

## 🛠️ Tecnologías

Python · pandas · openpyxl · PyYAML

## 📄 Licencia

MIT
