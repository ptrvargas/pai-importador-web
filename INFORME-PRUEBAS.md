# Informe de pruebas · PAI Importador V1.0

Fecha de ejecución: 18 de octubre de 2026 (America/Denver).

## Resultado comprobado

**21 pruebas automatizadas aprobadas** sobre los dos archivos adjuntos, más una verificación independiente del XLSM generado mediante ZIP y ElementTree de Python. Estado al ejecutar la primera batería: sitio todavía no publicado. El estado posterior se documenta en la actualización V1.1.

| Comprobación | Resultado |
| --- | --- |
| Inspección previa de ambos Excel | Realizada: dos partes de 172 preguntas; corrector con 12 hojas |
| Reconocimiento por estructura y nombres arbitrarios | Aprobado; no depende del nombre del archivo |
| Dos nombres distintos en los adjuntos | Bloqueados sin confirmación; asociación simulada expresamente en la suite |
| Mayúsculas, tildes y espacios | Se normalizan sin cambiar otras letras |
| Nombres repetidos o asociaciones ambiguas | Sin asociación automática; confirmación explícita requerida |
| Falta de nombre | Bloqueada incluso con confirmación |
| Datos de identificación contradictorios | Bloqueados |
| Respuestas vacías u opciones no permitidas | Bloqueadas con números de pregunta |
| Cabeceras incompletas, duplicadas o fuera de rango | Bloqueadas |
| Archivos intercambiados, dañados o demasiado grandes | Bloqueados |
| 344 marcas del archivo de prueba | 344 correctas; ninguna pregunta con más de una X |
| Cuatro opciones de respuesta | Verificadas en F, G, H e I mediante una segunda prueba |
| VBA original | Idéntico byte a byte |
| Fórmulas y hojas | Conservadas; mismas fórmulas y 12 hojas |
| Gráficos, baremos y demás componentes no editados | Idénticos byte a byte dentro del ZIP |
| Componentes internos del archivo | Misma lista de entradas ZIP |
| Datos de identificación | Importados; edad y fecha numéricas; cargo de ejemplo limpiado |
| Celdas fuera del ámbito autorizado | Idénticas en comprobación independiente XML |
| Archivos originales en disco | Huellas SHA-256 sin cambios después de las pruebas |

Huella SHA-256 del proyecto VBA conservado:

```text
2700bc86788b628b4b7f3208ad6b7d44d9f1f004031cf759988c76f50f8e83d0
```

La copia de prueba y los Excel fuente permanecen fuera del código distribuible. No se entregan como evaluaciones clínicas: la asociación entre las dos partes fue simulada para probar la técnica.

## Comprobaciones pendientes

1. **Interfaz real en escritorio y móvil.** No se dispone en esta sesión del mecanismo de control de navegador autorizado por el entorno de desarrollo. La revisión del código responsive no sustituye esta prueba.
2. **Instalación PWA y modo sin conexión** en Android, escritorio y Safari de iPhone/iPad. Manifest, service worker, caché cerrada e iconos están preparados; no se afirma compatibilidad comprobada.
3. **Apertura y recálculo en Excel de escritorio**, ejecución de macros y comparación de resultados con el procedimiento manual previamente validado. No hay Excel de escritorio en este entorno ni resultados de referencia del procedimiento anterior. Se conservaron las fórmulas y los componentes, pero eso no demuestra su ejecución correcta en Excel.
4. **Publicación GitHub Pages.** La conexión identifica al propietario y permite leer repositorios y escribir archivos, árboles, commits y referencias en repositorios existentes. No expone creación de repositorios ni configuración de Pages. El usuario creó el repositorio y habilitó Pages posteriormente; el código está en GitHub y se ha probado el flujo desde su teléfono.

## Lista para prueba de aceptación

El propietario autorizó posteriormente un sitio de prueba en Pages para realizar estas verificaciones. La aceptación clínica en Excel de escritorio sigue pendiente.

| Entorno / caso | Procedimiento y criterio |
| --- | --- |
| Escritorio | Abrir por HTTP; cargar los dos adjuntos; comprobar legibilidad, selectores, mensajes, validación y descarga |
| Nombre diferente | Ver ambos nombres; intentar validar sin casilla: bloqueado; confirmar solo para prueba: permite generar |
| Cambio de selección | Después de validar, cambiar evaluación o archivos: desaparecen validación y descarga anteriores |
| Móvil | Ancho 360–430 px, orientación vertical/horizontal, selectores y botones utilizables sin desbordamiento |
| iPhone/iPad | Safari: elegir desde Archivos, descargar XLSM y guardarlo; añadir a pantalla de inicio |
| PWA sin conexión | Cargar con conexión, cerrar, desconectar red y reabrir; cargar los archivos locales y generar sin conexión |
| Privacidad | En red del navegador solo aparecen recursos estáticos; ningún envío de nombres, respuestas o archivos |
| Excel de escritorio | Abrir copia sin aviso de reparación; comprobar identificación; recalcular completamente; comparar resultados y gráficos contra ingreso manual de las mismas 344 respuestas |

Si las respuestas de origen tienen nombres diferentes, el profesional debe determinar su asociación antes de usar el resultado. Las pruebas técnicas no resuelven esa identidad.

## Actualización V1.1 · 18 de octubre de 2026

El propietario completó el proceso en el sitio de prueba desde su teléfono y descargó el XLSM. Su revisión preliminar fue favorable; la comparación formal en Excel de escritorio sigue pendiente.

Se añadieron 13 pruebas automatizadas de lógica con adaptadores de IndexedDB y elementos de interfaz simulados: guardado binario y recuperación en otra instancia; eliminación; fallos de cuota y almacenamiento no disponible; instrucciones iPhone/iPad con identificación de escritorio, Android y Mac; diálogo nativo solo al pulsar y un único uso por evento; estado instalado; limpieza que conserva maestro; restauración sin respuestas; desactivar guardado; eliminación y continuidad tras fallo de guardado. Se probaron con los mismos Excel originales. No se afirma que estas pruebas sean ejecución real de IndexedDB ni una revisión visual de navegador.

El maestro se guarda tal como se selecciona, incluidos sus contenidos previos. La interfaz lo informa y aconseja usar un maestro sin datos reales. Nunca se guarda el XLSX importado ni el XLSM generado.

Pruebas de aceptación pendientes de V1.1: cargar maestro, recargar y confirmar recuperación; limpiar evaluación y confirmar conservación; eliminar maestro y recargar; pulsar instalación en Android; revisar guía de iPhone/iPad y Mac; comprobar nueva caché y funcionamiento sin conexión.
