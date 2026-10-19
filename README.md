# PAI Importador · V1.1

Aplicación estática en español para trasladar dos partes de Google Forms (1–172 y 173–344) al corrector PAI compatible. Procesamiento en memoria del navegador. Sin cuentas, servicios de datos, analítica, CDN ni envíos de archivos. No incluye el corrector, ítems del cuestionario ni respuestas clínicas.

## Estado de entrega

Código disponible en GitHub y sitio de prueba habilitado por el propietario en GitHub Pages. El usuario completó el flujo de carga y descarga desde su teléfono. V1.1 añade corrector recordado localmente y asistencia de instalación. La instalación PWA y la apertura/recálculo en Excel de escritorio no han sido verificadas. No declarar lista para uso clínico hasta completar esas comprobaciones. Véase `INFORME-PRUEBAS.md`.

## Uso para Norhorita

1. Carga el corrector XLSM original una vez. Con «Recordar el corrector» activado, se guarda una copia exacta en IndexedDB de este navegador y dispositivo; después carga el XLSX de Google Forms que contenga ambas partes. La copia incluye cualquier dato previo del maestro: utiliza uno sin datos reales de pacientes.
2. Selecciona la evaluación de la primera parte y revisa la segunda. Solo se asocian automáticamente nombres normalizados iguales y únicos en ambas partes.
3. Cuando difieran las letras o existan varios registros del mismo nombre, selecciona expresamente la segunda parte y confirma que es la misma evaluación. No confirmes si tienes dudas: corrige primero el archivo de origen.
4. Pulsa **Validar las 344 respuestas** y después **Generar y descargar XLSM**.
5. Abre la copia en Excel de escritorio y realiza un recálculo completo antes de usar los resultados. Las vistas previas móviles pueden mostrar valores almacenados de la evaluación anterior. La ejecución de macros necesita una versión compatible de Excel.
6. Pulsa **Limpiar evaluación** al terminar: retira las respuestas y los enlaces de descarga temporales, conservando el maestro. Para borrarlo del navegador, utiliza **Eliminar corrector guardado**. La aplicación no mantiene un historial de evaluaciones. La descarga sí permanece en tu dispositivo: protégela como cualquier archivo clínico.

## Desarrollo y despliegue

Proyecto sin compilación, sin dependencias remotas. Sirve esta carpeta por HTTP para probarla; abrir `index.html` con `file://` no admite los módulos y PWA de forma fiable.

Después de pasar la comprobación de interfaz, sube exclusivamente esta carpeta al repositorio `pai-importador-web`. No subas `upload/`, `private/` ni ningún Excel. En GitHub, **Settings → Pages → Deploy from a branch → main → /(root) → Save**. No hay que introducir secretos ni configurar Actions. El repositorio público contiene únicamente código.

URL de prueba: `https://ptrvargas.github.io/pai-importador-web/`. El propietario confirmó el acceso y la descarga desde su teléfono; la disponibilidad de cada nueva actualización requiere comprobar Pages. Tras desplegar, verifica el sitio real y el alcance del service worker.

## Pruebas reproducibles sin publicar pacientes

Se necesita Node.js moderno, sin instalar paquetes (JSZip va incluido).

```bash
node tests/core.test.mjs "/ruta/privada/corrector.xlsm" "/ruta/privada/respuestas.xlsx" "/ruta/privada/prueba.xlsm"
```

El tercer argumento es opcional. El XLSM de prueba que produce la suite asocia expresamente los dos primeros registros para probar el motor. Esa asociación es una simulación técnica: no confirma la identidad del paciente. La suite con adjuntos se ejecuta en privado y nunca se incorpora al repositorio. Algunas verificaciones de esta suite están ajustadas a las características de los dos archivos aportados (12 hojas, datos demográficos de prueba y nombres diferentes).

## Arquitectura e integridad

- `core.js`: lector limitado a OOXML estándar, reconocimiento estructural, asociación, validación y edición de celdas.
- `storage.js`: guardado, recuperación y eliminación transaccionales de una única copia del maestro en IndexedDB. Si el almacenamiento está bloqueado o lleno, el importador funciona en la sesión sin guardarlo.
- `install.js`: aviso de instalación, diálogo nativo si existe `beforeinstallprompt`, instrucciones para iOS/iPadOS (incluido iPad con identificación de escritorio), macOS y navegadores sin diálogo. No intenta forzar ventanas nativas sin gesto del usuario.
- `app.js`: interfaz y estado efímero de respuestas. Los cambios de selección invalidan cualquier descarga anterior. Se revocan los enlaces Blob al limpiar o reemplazar datos.
- `vendor/jszip.min.js`: JSZip 3.10.1 distribuido localmente; licencia en `vendor/JSZip-LICENSE.md`.
- `sw.js`: caché por lista cerrada de recursos estáticos de la aplicación. No intercepta ni almacena archivos de pacientes.
- `manifest.webmanifest`: instalación progresiva; iconos locales. En Safari de iPhone/iPad, usar Compartir → Añadir a pantalla de inicio.

El escritor no utiliza la exportación genérica de una biblioteca de hojas de cálculo. Reemplaza las celdas F–I de las 344 preguntas y seis campos de identificación en el XML original, manteniendo sus estilos. Limpia el cargo de ejemplo cuando el formulario no lo contiene. Conserva edad y fecha Excel como valores numéricos cuando corresponde. Solicita recálculo completo al abrir mediante dos atributos de `calcPr`; no calcula, reescribe ni sustituye las fórmulas.

Antes de descargar, vuelve a leer el resultado y comprueba: misma lista de componentes ZIP; igualdad byte a byte de todos los componentes no editados, incluido VBA, gráficos y baremos; fórmulas originales idénticas; 344 marcas correctas, exactamente una por pregunta. Se conservan los contenidos internos, no los bytes comprimidos del ZIP completo.

Límites: 30 MB de archivo y 120 MB de contenido declarado descomprimido. Se comprueba CRC. Se rechazan libros cifrados, estructuras no reconocidas y firmas digitales del documento. La verificación de la firma de macros y la ejecución del código VBA pertenecen a Excel, no a esta aplicación.

## Privacidad y uso profesional

La página pública sirve el código; no recibe pacientes. GitHub puede registrar la visita como cualquier alojamiento. IndexedDB conserva únicamente el corrector maestro que el usuario decide recordar, sin alterar sus bytes. Nunca se persisten el XLSX de respuestas importadas ni las copias generadas. No se utilizan cookies de aplicación ni servicios de terceros para guardar pacientes. Las nuevas respuestas permanecen en memoria de la pestaña hasta limpiar o cerrar; no se promete borrado físico de RAM ni control sobre copias realizadas por el sistema operativo. Las descargas son responsabilidad del profesional.

Usar exclusivamente con un corrector y una prueba para los que se tenga autorización. La herramienta no interpreta resultados ni aporta diagnósticos.

Para futuras versiones, añadir adaptadores explícitos por formato de prueba a `core.js`, sin introducir datos o reglas de puntuación en el código público. Mantener la separación entre interfaz, emparejamiento y edición OOXML.

## Actualización 1.1

La preferencia «Recordar» se aplica al maestro seleccionado. Al desactivarla, se elimina la copia guardada, manteniendo el corrector temporal de la sesión. Cerrar, recargar o cambiar a vista de escritorio ya no exige elegir de nuevo el maestro si el navegador conservó IndexedDB. Se vuelve a validar estructuralmente al restaurarlo. Borrar los datos del navegador, su modo privado o sus políticas de espacio pueden retirar la copia. Safari y la app instalada pueden tener almacenes separados; en ese caso carga el maestro una vez en cada uno.

El aviso propio de instalación se muestra al entrar si la ventana no está en modo instalado. El diálogo nativo solo se abre al pulsar y si el navegador lo ha habilitado. En Mac, Safari compatible utiliza «Añadir al Dock»; iPhone/iPad utiliza «Añadir a pantalla de inicio».

Pruebas adicionales de lógica y transacciones (adaptadores simulados; no sustituyen pruebas reales de navegador):

```bash
node tests/local-features.test.mjs "/ruta/privada/corrector.xlsm" "/ruta/privada/respuestas.xlsx"
```
