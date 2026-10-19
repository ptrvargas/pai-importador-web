# PAI Importador · V1.0

Aplicación estática en español para trasladar dos partes de Google Forms (1–172 y 173–344) al corrector PAI compatible. Procesamiento en memoria del navegador. Sin cuentas, servicios de datos, analítica, CDN ni envíos de archivos. No incluye el corrector, ítems del cuestionario ni respuestas clínicas.

## Estado de entrega

Código implementado y pruebas de importación aprobadas con los archivos aportados. Repositorio creado y código preparado para GitHub. Publicación pendiente de comprobación real de interfaz en navegador de escritorio y móvil. La instalación PWA y la apertura/recálculo en Excel de escritorio no han sido verificadas. No declarar lista para uso clínico hasta completar esas comprobaciones. Véase `INFORME-PRUEBAS.md`.

## Uso para Norhorita

1. Carga el corrector XLSM original y el XLSX de Google Forms que contenga ambas partes.
2. Selecciona la evaluación de la primera parte y revisa la segunda. Solo se asocian automáticamente nombres normalizados iguales y únicos en ambas partes.
3. Cuando difieran las letras o existan varios registros del mismo nombre, selecciona expresamente la segunda parte y confirma que es la misma evaluación. No confirmes si tienes dudas: corrige primero el archivo de origen.
4. Pulsa **Validar las 344 respuestas** y después **Generar y descargar XLSM**.
5. Abre la copia en Excel de escritorio y realiza un recálculo completo antes de usar los resultados. Las vistas previas móviles pueden mostrar valores almacenados de la evaluación anterior. La ejecución de macros necesita una versión compatible de Excel.
6. Pulsa **Limpiar archivos y datos** al terminar. La aplicación no mantiene un historial. La descarga sí permanece en tu dispositivo: protégela como cualquier archivo clínico.

## Desarrollo y despliegue

Proyecto sin compilación, sin dependencias remotas. Sirve esta carpeta por HTTP para probarla; abrir `index.html` con `file://` no admite los módulos y PWA de forma fiable.

Después de pasar la comprobación de interfaz, sube exclusivamente esta carpeta al repositorio `pai-importador-web`. No subas `upload/`, `private/` ni ningún Excel. En GitHub, **Settings → Pages → Deploy from a branch → main → /(root) → Save**. No hay que introducir secretos ni configurar Actions. El repositorio público contiene únicamente código.

La URL esperada es `https://ptrvargas.github.io/pai-importador-web/`, pero no representa una publicación confirmada. Tras desplegar, verifica el sitio real y el alcance del service worker.

## Pruebas reproducibles sin publicar pacientes

Se necesita Node.js moderno, sin instalar paquetes (JSZip va incluido).

```bash
node tests/core.test.mjs "/ruta/privada/corrector.xlsm" "/ruta/privada/respuestas.xlsx" "/ruta/privada/prueba.xlsm"
```

El tercer argumento es opcional. El XLSM de prueba que produce la suite asocia expresamente los dos primeros registros para probar el motor. Esa asociación es una simulación técnica: no confirma la identidad del paciente. La suite con adjuntos se ejecuta en privado y nunca se incorpora al repositorio. Algunas verificaciones de esta suite están ajustadas a las características de los dos archivos aportados (12 hojas, datos demográficos de prueba y nombres diferentes).

## Arquitectura e integridad

- `core.js`: lector limitado a OOXML estándar, reconocimiento estructural, asociación, validación y edición de celdas.
- `app.js`: interfaz y estado efímero. Los cambios de selección invalidan cualquier descarga anterior. Se revocan los enlaces Blob al limpiar o reemplazar datos.
- `vendor/jszip.min.js`: JSZip 3.10.1 distribuido localmente; licencia en `vendor/JSZip-LICENSE.md`.
- `sw.js`: caché por lista cerrada de recursos estáticos de la aplicación. No intercepta ni almacena archivos de pacientes.
- `manifest.webmanifest`: instalación progresiva; iconos locales. En Safari de iPhone/iPad, usar Compartir → Añadir a pantalla de inicio.

El escritor no utiliza la exportación genérica de una biblioteca de hojas de cálculo. Reemplaza las celdas F–I de las 344 preguntas y seis campos de identificación en el XML original, manteniendo sus estilos. Limpia el cargo de ejemplo cuando el formulario no lo contiene. Conserva edad y fecha Excel como valores numéricos cuando corresponde. Solicita recálculo completo al abrir mediante dos atributos de `calcPr`; no calcula, reescribe ni sustituye las fórmulas.

Antes de descargar, vuelve a leer el resultado y comprueba: misma lista de componentes ZIP; igualdad byte a byte de todos los componentes no editados, incluido VBA, gráficos y baremos; fórmulas originales idénticas; 344 marcas correctas, exactamente una por pregunta. Se conservan los contenidos internos, no los bytes comprimidos del ZIP completo.

Límites: 30 MB de archivo y 120 MB de contenido declarado descomprimido. Se comprueba CRC. Se rechazan libros cifrados, estructuras no reconocidas y firmas digitales del documento. La verificación de la firma de macros y la ejecución del código VBA pertenecen a Excel, no a esta aplicación.

## Privacidad y uso profesional

La página pública sirve el código; no recibe pacientes. GitHub puede registrar la visita como cualquier alojamiento. No se utilizan almacenamiento local, IndexedDB, cookies de aplicación ni servicios de terceros para guardar información clínica. Los datos permanecen en memoria de la pestaña hasta limpiar o cerrar; no se promete borrado físico de RAM ni control sobre copias realizadas por el sistema operativo. Las descargas son responsabilidad del profesional.

Usar exclusivamente con un corrector y una prueba para los que se tenga autorización. La herramienta no interpreta resultados ni aporta diagnósticos.

Para futuras versiones, añadir adaptadores explícitos por formato de prueba a `core.js`, sin introducir datos o reglas de puntuación en el código público. Mantener la separación entre interfaz, emparejamiento y edición OOXML.
