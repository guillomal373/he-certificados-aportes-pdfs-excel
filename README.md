# he-certificados-aportes-pdfs-excel
...
# README - Procesamiento de Certificados de Aportes en macOS

Este documento explica cómo ejecutar el script `procesar_certificados_aportes.py` en un MacBook para procesar múltiples archivos PDF desde una carpeta y generar un único archivo Excel consolidado.

## 1) Verificar Python en macOS

En macOS, normalmente el comando correcto es `python3`, no `python`.

Ejecuta en Terminal:

```bash
python3 --version
```

Si también quieres validar `pip`:

```bash
pip3 --version
```

Si `pip3` no responde, usa:

```bash
python3 -m pip --version
```

## 2) Instalar dependencias

Instala las librerías necesarias:

```bash
python3 -m pip install pdfplumber openpyxl
```

## 3) Ubicar el archivo del script

Guarda el archivo Python con este nombre:

```text
procesar_certificados_aportes.py
```

Puedes dejarlo, por ejemplo, en:

```text
/Users/guillermomalagon/Downloads/procesar_certificados_aportes.py
```

## 4) Estructura esperada

Debes tener:

- una carpeta con todos los PDFs
- el script Python
- una ruta donde se guardará el Excel de salida

Ejemplo:

```text
/Users/guillermomalagon/OneDrive - DS SOLUTIONS S.A.S/Documentos/info-certificados-ingresos-retenciones/pdfs
```

## 5) Comando para ejecutar el script

### Opción A: una sola línea

```bash
python3 "/Users/guillermomalagon/Downloads/procesar_certificados_aportes.py" --input "/Users/guillermomalagon/OneDrive - DS SOLUTIONS S.A.S/Documentos/info-certificados-ingresos-retenciones/pdfs" --output "/Users/guillermomalagon/OneDrive - DS SOLUTIONS S.A.S/Documentos/info-certificados-ingresos-retenciones/salida.xlsx"
```

### Opción B: varias líneas en zsh

```bash
python3 "/Users/guillermomalagon/Downloads/procesar_certificados_aportes.py" \
  --input "/Users/guillermomalagon/OneDrive - DS SOLUTIONS S.A.S/Documentos/info-certificados-ingresos-retenciones/pdfs" \
  --output "/Users/guillermomalagon/OneDrive - DS SOLUTIONS S.A.S/Documentos/info-certificados-ingresos-retenciones/salida.xlsx"
```

## 6) Qué hace el script

El script:

- recorre todos los archivos `.pdf` de la carpeta indicada
- toma únicamente la última persona/identificación encontrada en el encabezado
- consolida la información en un solo archivo Excel
- crea estas pestañas:
  - `Liquidaciones Pagadas`
  - `Seguridad Social`
  - `Aportes Parafiscales`
  - `Novedades`
- guarda:
  - números como números
  - fechas como fechas reales de Excel
  - tarifas como números
  - novedades con una fila por cada `X`
- continúa con los demás archivos aunque uno falle

## 7) Mensajes esperados en consola

Durante la ejecución verás mensajes como estos:

```text
Procesando 4 archivos PDF...
OK  - 648590.pdf
OK  - 3954297.pdf
ERROR - 1180890.pdf: No se pudo extraer la identidad del encabezado.
OK  - 5503152.pdf
```

## 8) Ejemplo del resumen final

Al final, la consola mostrará un resumen como este:

```text
================================================================================
RESUMEN FINAL
================================================================================
Total de PDFs encontrados : 4
Procesados OK             : 3
Con error                 : 1
Excel generado en         : /Users/guillermomalagon/OneDrive - DS SOLUTIONS S.A.S/Documentos/info-certificados-ingresos-retenciones/salida.xlsx

Archivos OK:
  - 648590.pdf
  - 3954297.pdf
  - 5503152.pdf

Archivos con ERROR:
  - 1180890.pdf: No se pudo extraer la identidad del encabezado.
```

## 9) Qué hacer si aparece `zsh: command not found: python`

En macOS usa:

```bash
python3
```

No uses:

```bash
python
```

Ejemplo correcto:

```bash
python3 "/Users/guillermomalagon/Downloads/procesar_certificados_aportes.py" --input "/ruta/pdfs" --output "/ruta/salida.xlsx"
```

## 10) Qué hacer si un PDF falla

Si un archivo muestra error, revisa:

- el bloque `DEBUG TEXTO nombre_del_archivo.pdf`
- el mensaje exacto del error
- si el PDF tiene una estructura distinta al resto

El script está diseñado para seguir procesando los demás PDFs aunque uno falle. Eso evita que todo el proceso se caiga por un solo archivo problemático.

## 11) Recomendación práctica

Antes de lanzar 100 archivos de una vez, prueba primero con 3 o 4 PDFs.  
Eso te ayuda a validar que:

- el encabezado se está extrayendo bien
- las tablas se interpretan bien
- el Excel sale con el formato esperado

Primero confianza, después volumen. Así duele menos.
