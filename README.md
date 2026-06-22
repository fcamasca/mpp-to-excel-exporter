# MPP to Excel Exporter

Convierte cronogramas de **Microsoft Project (`.mpp`)** en reportes Excel listos para revisar, compartir y filtrar, sin requerir Microsoft Project instalado.

El proyecto nació de una necesidad operativa real: transformar cronogramas técnicos en un formato accesible para equipos de seguimiento y usuarios de negocio.

## Qué resuelve

- Extrae tareas, fechas, duración, trabajo, avance y recursos asignados.
- Conserva visualmente la jerarquía del cronograma.
- Destaca tareas resumen e hitos.
- Genera una hoja con filtros, encabezado congelado y formatos de fecha y porcentaje.
- Elimina automáticamente el JSON temporal utilizado durante la conversión.

## Flujo de procesamiento

```text
Archivo .mpp → MPXJ → JSON temporal → pandas/openpyxl → Reporte .xlsx
```

La lectura del formato MPP se delega a [MPXJ](https://www.mpxj.org/), una biblioteca de código abierto para archivos de planificación. Python se encarga de transformar los datos y construir el reporte final.

## Tecnologías

- Python 3.10+
- pandas
- openpyxl
- Java 11+
- MPXJ 15.x

## Instalación

1. Clona este repositorio y crea un entorno virtual:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

2. Instala Java y verifica que esté disponible:

```powershell
java -version
```

3. Descarga la distribución binaria de MPXJ desde su [sitio oficial](https://www.mpxj.org/) y descomprímela en `mpxj/`, o define la variable `MPXJ_HOME` con su ubicación.

> MPXJ no se incluye en este repositorio para evitar versionar binarios y código de terceros.

## Uso

```powershell
py mpp_to_excel.py "ruta\cronograma.mpp"
```

El Excel se crea junto al archivo de entrada. También puedes indicar otra salida o instalación de MPXJ:

```powershell
py mpp_to_excel.py "entrada.mpp" -o "reportes\cronograma.xlsx" --mpxj-home "C:\tools\mpxj"
```

## Columnas exportadas

| Columna | Descripción |
|---|---|
| Nombre de tarea | Nombre e indentación según el nivel jerárquico |
| Duración | Minutos, horas o jornadas laborales de 8 horas |
| Trabajo | Esfuerzo expresado en horas |
| Comienzo / Fin | Fechas planificadas |
| % completado | Avance de la tarea |
| Nombres de los recursos | Recursos asignados, sin duplicados |

## Pruebas

Las pruebas unitarias validan las transformaciones sin necesitar Java ni archivos corporativos:

```powershell
py -m pip install -r requirements-dev.txt
py -m pytest -q
```

## Privacidad

Los cronogramas pueden contener nombres de personas, iniciativas y fechas internas. Por eso `.gitignore` excluye archivos MPP, Excel, comprimidos, temporales y la instalación local de MPXJ. Los datos usados en pruebas son completamente ficticios.

## Limitaciones

- La duración se presenta usando jornadas estándar de 8 horas.
- La conversión requiere Java y una distribución local de MPXJ.
- El reporte cubre los campos más útiles para seguimiento; no busca replicar todas las vistas de Microsoft Project.

## Estado del proyecto

Versión actual: v1.0.0  
Estado: Mantenimiento activo.

## Próximas mejoras

- Procesamiento por lotes de una carpeta.
- Configuración de horas por jornada.
- Hoja resumen con indicadores del cronograma.
- Empaquetado como aplicación de escritorio.

## Licencia

Este proyecto se distribuye bajo la licencia MIT. Consulte el archivo [LICENSE](LICENSE) para más detalles.
