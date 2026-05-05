# Changelog

Todos los cambios notables de este proyecto se documentan en este archivo.

El formato está basado en [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/)
y este proyecto adhiere a [Versionamiento Semántico](https://semver.org/lang/es/).

## [No publicado]

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
