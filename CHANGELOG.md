# Changelog

Todos los cambios notables de este proyecto se documentan en este archivo.

El formato está basado en [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/)
y este proyecto adhiere a [Versionamiento Semántico](https://semver.org/lang/es/).

## [No publicado]

### Agregado
- Se crearon tests de preservación de comportamiento (PBT) en `tests/preservation.test.js`: 10 property-based tests que validan comportamiento correcto existente para entradas no-buggy — initializeDataTableRenovaciones con arreglos válidos, mergeJsonColumnaGestion preserva campos y concatena documentosProceso, renderEtapas con índices válidos, AssignLead selecciona menor sortingKey y mayor effectiveness
- Se crearon tests de exploración de condición de bug (PBT) para 8 bugs críticos en `tests/bug-condition-exploration.test.js`: validación de fetch/Headers en GAS, DataTable con undefined, JSON corrupto en gestión, índice de etapa fuera de rango, intervalos OTP sin limpiar, navegación sin confirmación en Emisión de Póliza, fallo parcial de archivos no reportado, y desempate no determinístico en AssignLead

### Cambiado
- Se mejoró UI del formulario unificado de procesos especiales: se reemplazó indicador de pasos con badges por el componente `.progressbar` del sistema (3 columnas), se usó `.step-section` + `.step-header` + `.step-icon` consistente con los demás pasos del modal, y se eliminó botón de guardado duplicado (se usa el botón principal existente del modal)
- Se rediseñó botón "Gestionar Procesos Especiales" como mini-card informativa con título "Correcciones y Procesos Especiales", subtítulo descriptivo y botón outline-primary
- Se refinó flujo de Correcciones: ahora permite tipificar como "Caso Corregido" desde el dropdown de estado. El estado ya no se fuerza a "Caso Especial" — usa el valor seleccionado por el usuario (default "Caso Especial" si no selecciona nada). Se agregó "Caso Corregido" como estado terminal y protegido

### Corregido
- Se corrigió error `Cannot read properties of undefined (reading 'includes')` en `actualizarEtapaFunnel`: el selector `.progressbar li` capturaba los `<li>` del nuevo progressbar del formulario de procesos especiales que no tienen `data-etapa-nombre`. Se cambió a `.progressbar li.progress-step`
- Se corrigió que `mostrarProcesosEspeciales` ahora oculta TODOS los pasos del flujo normal (propietario, sarlaft, cumplimiento, póliza, OTP) — los procesos especiales y el flujo de 4 pasos son mutuamente excluyentes
- Se corrigió zona de carga de archivos no clicable: se agregó `stopPropagation` al input file para evitar loop de eventos con el drop zone
- Se corrigió DataTable mostrando pólizas sin información: se agregó filtro `.filter(item => item.leadData && item.leadData.poliza)` en Código.js para excluir registros con JSON vacío o sin número de póliza

### Eliminado
- Se eliminaron inputs ocultos legacy `db_OTROSI_FILE_URL`, `db_CESION_CEDULA_URL` y `db_CESION_DOC_ADICIONAL_URL` de index.html (ya no referenciados tras migración a formulario unificado)

### Agregado
- Se refactorizó `guardarGestionRenovacion` en Renovaciones.js para usar merge no destructivo via `mergeJsonColumnaGestion`: lee JSON actual de columna 6, aplica merge preservando campos existentes, concatena `documentosProceso` nuevos al arreglo existente, maneja JSON inválido/vacío inicializando objeto nuevo, y registra `procesoEspecial` en historial (columna 7) con campos fecha, usuario, observacion, estadoGestion, seguimiento y procesoEspecial
- Se modificó `guardarGestionUnificada` en main.js.html para soportar flujo unificado de procesos especiales: bypass OTP para Correcciones (asigna "Caso Especial" automáticamente), validación OTP para Otro Sí/Cesión, validación de observaciones y archivos, conversión multi-file a base64 con formato `{documentosProceso: [...]}`, y SweetAlert de éxito/error
- Se refactorizó `guardarArchivosEnCarpeta` en Renovaciones.js para procesar arreglo `documentosProceso` (multi-archivo): itera archivos base64, decodifica a Blob, nombra con formato `{PREFIJO}_{poliza}_{index}.{ext}`, guarda en subcarpeta "Procesos Especiales" y retorna arreglo de `{nombre, url}` en `urlsNuevas.documentosProceso`
- Se creó función `obtenerPrefijoProcesoEspecial` para mapear tipo de proceso a prefijo de nomenclatura (OTRO_SI, CESION, CORRECCION)
- Se creó formulario unificado `#formProcesoEspecial` en index.html con encabezado dinámico, indicador de 3 pasos, textarea de observaciones (máx 2000 chars con contador) y contenedor para multi-file upload
- Se agregó card `#cardCorrecciones` para el nuevo proceso "Correcciones" junto a Otro Sí y Cesión
- Se implementó componente de carga múltiple de archivos en main.js.html: `initMultiFileUpload`, `agregarArchivos`, `eliminarArchivo`, `renderizarListaArchivos`, `convertirArchivosABase64` y `formatearTamanoArchivo`
- Se implementó función `mergeJsonColumnaGestion` en Renovaciones.js para merge no destructivo del JSON de Columna_Gestion, preservando campos existentes y concatenando documentosProceso sin sobrescribir datos previos
- Se agregaron constantes `PROCESO_ESPECIAL_CONFIG` y `PROCESO_TIPOS` en main.js.html para centralizar validación de procesos especiales (máx 10 archivos, 15MB, formatos PDF/JPG/PNG, 2000 chars observaciones)
- Se declararon variables globales `procesoEspecialSeleccionado` y `archivosSeleccionados` con tipado JSDoc

### Eliminado
- Se eliminaron formularios legacy `#formOtroSi` y `#formCesion` de index.html (reemplazados por formulario unificado `#formProcesoEspecial`)
- Se eliminaron campos legacy de Cesión (tipoDoc, numDoc, nombre, teléfono del nuevo propietario) del HTML
- Se eliminaron funciones legacy `handleOtroSiFileUpload` y `handleCesionDocAdicionalUpload` de main.js.html (reemplazadas por componente multi-file upload)
- Se eliminaron referencias a áreas de upload legacy (`otroSiUploadArea`, `cesionCedulaUploadArea`, `cesionDocAdicionalUploadArea`) de `configurarDragAndDrop`

### Corregido
- Se agregó función retryDrive con backoff exponencial (5 reintentos) en Renovaciones.js para todas las llamadas a DriveApp (getFolderById, searchFolders, createFolder, createFile, getUrl, getFoldersByName, hasNext, next, getName)
- Se mejoró el manejo de errores en guardarGestionRenovacion para informar al usuario cuando Drive falla después de agotar los reintentos
- Se corrigió error en main.js.html donde archivosSubidos.sarlaft ya estaba en formato base64 pero se intentaba releer con FileReader, causando fallo inmediato con "No se pudieron procesar los documentos adjuntos"
- Se corrigió error "recalcularKPIs is not defined" moviendo la función fuera del scope de $(document).ready al scope global del script
- Se corrigió persistencia de documentos entre leads agregando limpieza de inputs file y previews en limpiarFormularioRenovaciones
- Se agregó onchange a fechaFinVigencia para recalcular meses y cotización cuando el usuario edita la fecha fin en modo "Póliza Nueva"
- Se agregó recálculo de cotización en onchange de fechaInicioVigencia
- Se corrigió cálculo de displayMeses para "Póliza Nueva": ahora usa la duración real de la póliza anterior en vez de siempre +1 año
- Se corrigió obtenerDestinatarioOTP para broker/inmobiliaria: ahora lee teléfono y email del DOM (paso 1) en vez de currentLeadData en memoria
- Se agregó validación obligatoria del campo "Estado del Contacto" al guardar gestión: ahora muestra Swal de advertencia si no se tipifica la llamada
