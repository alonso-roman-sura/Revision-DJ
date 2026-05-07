# Sistema de Revisión de Llenado de Declaraciones Juradas

## Descripción general

Automatización en Excel VBA para validar el llenado de declaraciones juradas a partir del cruce entre una base de colaboradores y un reporte exportado de Talentum. El proceso importa ambas fuentes, normaliza su estructura, calcula indicadores de doble registro y de cumplimiento dentro de un período definido, resalta coincidencias en la base de colaboradores y consolida un conteo de personas pendientes.

## Objetivo

Reducir la revisión manual del llenado de declaraciones juradas, mejorar la trazabilidad del cruce entre fuentes y permitir una validación repetible del estado de cumplimiento para distintos períodos sin necesidad de reconstruir el archivo cada vez.

## Alcance funcional

El proyecto cubre las siguientes operaciones:

- Carga de la base de colaboradores desde archivo externo, priorizando la hoja `ACTIVOS` cuando exista.
- Validación básica de estructura y descarte de archivos incompatibles.
- Filtrado de registros por país Perú o Peru.
- Conversión de la base importada en una tabla estructurada llamada `Colaboradores`.
- Carga del reporte de declaraciones juradas desde exportaciones de Talentum, incluyendo archivos `.xls` y archivos HTML abiertos como libro de Excel.
- Detección automática de encabezados aunque el origen no empiece en `A1`.
- Conversión del reporte importado en una tabla estructurada llamada `ReporteDJ`.
- Cálculo de la columna `Doble Planilla` para identificar personas con dos registros.
- Cálculo de la columna de llenado según un rango de fechas definido por el usuario.
- Construcción de una clave interna para cruzar `Nombre Completo` de colaboradores contra `Nombres + Apellidos` del reporte.
- Resaltado de colaboradores con llenado válido dentro del período evaluado.
- Cálculo del total de personas faltantes.
- Reejecución de la comprobación con nuevas fechas sin necesidad de reimportar.
- Limpieza completa de hojas y tablas de trabajo.

## Estructura del proyecto

### Hojas generadas

- `Colaboradores`
- `ReporteDJ`

### Tablas generadas

- `Colaboradores`
- `ReporteDJ`

### Módulos principales

- `modDJConfig`
- `modDJUtils`
- `modDJImport`
- `modDJValidation`
- `modDJMain`

### Formulario

- `frmDateRangeDJ`

### Clases auxiliares del formulario

- `CDateRangeButtonHandler`
- `CDateRangeTextBoxHandler`

## Flujo de uso

1. Ejecutar `CargarColaboradores`.
2. Seleccionar la base de colaboradores.
3. Ejecutar `CargarReporteDJ`.
4. Seleccionar el reporte de declaraciones juradas.
5. Ejecutar `EjecutarComprobacion`.
6. Ingresar el rango de fechas en el formulario.
7. Revisar resultados en `ReporteDJ` y el resaltado aplicado en `Colaboradores`.
8. Ejecutar nuevamente la comprobación si se requiere otro rango.
9. Ejecutar `EliminarDatos` para reiniciar el proceso.

## Requisitos de entrada

### Base de colaboradores

Columnas esperadas:

- País
- Nombre Completo
- Líder / Colaborador
- Nombre Posición
- Compañía
- Nivel Organizacional 2
- Nivel Organizacional 3
- Nivel Organizacional 4
- Nivel Organizacional 5
- Fecha de ingreso
- Fecha de antigüedad
- Nombre Líder Directo
- Correo Comunicados
- Fecha de Nacimiento
- Género
- Ubicación Física
- Administrativo / Ventas - gasto
- Antigüedad
- Edad
- Generación
- País donde se ejecuta el gasto (dato para finanzas)

Reglas principales:

- Se prioriza la hoja `ACTIVOS` cuando existe.
- Solo se conservan registros con país `Perú` o `Peru`.
- Si existe la columna `Fecha de salida`, el archivo se considera inválido para este proceso.
- La validación de estructura es deliberadamente flexible, pero exige coincidencia mínima de encabezados clave.

### Reporte de declaraciones juradas

Columnas esperadas:

- ID de usuario/empleado de Talentum
- Nombres
- Apellidos
- Compañia
- Fecha de registro
- Adjunto declaración

Reglas principales:

- El origen puede ser `.xlsx`, `.xls`, `.xlsm`, `.htm` o `.html`.
- El encabezado puede empezar fuera de `A1`.
- Las columnas de texto se fuerzan a formato texto, especialmente el ID de usuario.

## Lógica de validación

### Doble Planilla

Se calcula con la fórmula:

```excel
=CONTAR.SI.CONJUNTO([Nombres];[@Nombres];[Apellidos];[@Apellidos])=2
```

### Llenado dentro del período

La comprobación usa una fecha inicial y una fecha final definidas en el formulario. Si una persona tiene doble planilla, ambas filas deben estar dentro del rango para marcar cumplimiento.

### Faltan llenar

Se calcula como el conteo de personas únicas de `Colaboradores` cuyo estado de llenado para el período evaluado es `FALSO`.

## Salidas

### En `ReporteDJ`

Se agregan columnas calculadas a la derecha de la tabla:

- `Doble Planilla`
- `Llenado en AAAA` o `Llenado entre dd-mm-aaaa y dd-mm-aaaa`

También se escribe un bloque resumen:

- `Faltan llenar`
- valor calculado del total pendiente

### En `Colaboradores`

Se agregan columnas auxiliares a la derecha de la tabla:

- `Clave Interna`
- columna de estado para el período evaluado

Las filas con cumplimiento válido se resaltan en verde `#92D050` con texto negro.

## Validaciones y manejo de errores

El proyecto contempla los siguientes controles:

- cancelación segura del formulario de fechas sin mostrar mensaje de éxito;
- bloqueo de ejecución si falta alguna de las tablas requeridas;
- bloqueo de ejecución si una tabla existe pero no tiene filas;
- bloqueo de ejecución si faltan columnas críticas;
- aviso cuando se intenta eliminar datos y no hay hojas cargadas;
- confirmación antes de eliminar hojas de trabajo;
- limpieza del resaltado anterior antes de recalcular un nuevo período;
- limpieza del bloque resumen antes de reescribirlo.

## Consideraciones técnicas

- El proyecto está orientado a Excel de escritorio con soporte para VBA.
- La lógica está pensada para trabajar con tablas estructuradas.
- La comparación entre fuentes se basa en normalización textual de nombres.
- La coincidencia por nombre puede requerir ajuste si en el origen existen diferencias de escritura, omisión de segundos nombres o cambios de formato.
- El archivo debe abrirse con macros habilitadas.

## Limitaciones

- La relación entre `Nombre Completo` y `Nombres + Apellidos` depende de consistencia razonable en la escritura.
- No existe integración directa con Talentum ni con otras plataformas; la carga sigue siendo manual.
- El proceso valida llenado por rango de fechas, no la completitud documental más allá de lo reflejado en las fuentes cargadas.

## Posibles mejoras

- centralizar la ejecución en una hoja de inicio con botones y estado del proceso;
- registrar bitácora de ejecuciones con usuario, fecha y rango evaluado;
- exportar resultados a una hoja de resumen o a PDF;
- incorporar validación más robusta de homónimos;
- reemplazar el cruce nominal por un identificador único cuando exista en ambas fuentes;
- agregar control de versiones del período evaluado.

## Campos sugeridos para el registro del proyecto

### Proyecto

Sistema de Revisión de Llenado de Declaraciones Juradas

### Estado

En funcionamiento

### Descripción

Importación y cruce automatizado entre la base de colaboradores y el reporte de declaraciones juradas de Talentum para validar el cumplimiento de llenado dentro de un período definido, identificar doble planilla, resaltar coincidencias y consolidar el total de personas pendientes.

### Necesidad

Reducir la revisión manual del cumplimiento de declaraciones juradas, evitar errores de cruce entre fuentes, permitir reevaluaciones rápidas para distintos períodos y disponer de una validación más consistente y trazable.

### Fecha

08-abr

### Herramienta

Excel VBA

### Nota

La solución depende de la consistencia de nombres entre ambas fuentes, por lo que conviene evaluar en el futuro el uso de un identificador único común. El proceso está diseñado para reejecutarse con distintos rangos de fechas sin necesidad de reimportar los archivos, y requiere Excel de escritorio con macros habilitadas.
