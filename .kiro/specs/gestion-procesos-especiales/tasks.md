# Implementation Plan: Gestión de Procesos Especiales

## Overview

Refactorización de la sección "Procesos Especiales" del CRM de renovaciones para unificar los formularios de Otro Sí, Cesión del Contrato y el nuevo proceso Correcciones. Se corrige el mecanismo de carga de archivos a Google Drive (FileReader/base64), se implementa bypass de OTP para Correcciones, y se garantiza merge no destructivo en Columna_Gestion.

## Tasks

- [x] 1. Definir constantes, configuración y utilidades compartidas
  - [x] 1.1 Crear constantes de validación y tipos de proceso en main.js.html
    - Agregar el objeto `PROCESO_ESPECIAL_CONFIG` con MAX_ARCHIVOS (10), MAX_SIZE_BYTES (15MB), FORMATOS_PERMITIDOS, EXTENSIONES_PERMITIDAS, MAX_OBSERVACIONES_CHARS (2000)
    - Agregar el objeto `PROCESO_TIPOS` con las claves OTRO_SI, CESION, CORRECCION y sus labels/prefijos
    - Agregar variable global `procesoEspecialSeleccionado` y `archivosSeleccionados` (array)
    - _Requirements: 1.1, 2.1, 3.2, 5.1, 5.3_

  - [x] 1.2 Implementar función `mergeJsonColumnaGestion` en Renovaciones.js
    - Crear función que recibe `currentJson` y `newData`
    - Aplicar spread/merge superficial: `{...currentJson, ...newData}`
    - Para el campo `documentosProceso`: concatenar `[...(currentJson.documentosProceso || []), ...(newData.documentosProceso || [])]`
    - Manejar caso de JSON inválido o vacío inicializando objeto nuevo
    - _Requirements: 7.3, 7.4, 7.5, 7.6, 7.7_

  - [ ]* 1.3 Write property test for mergeJsonColumnaGestion
    - **Property 8: Merge de JSON preserva campos existentes no actualizados**
    - **Validates: Requirements 7.4, 7.6**

  - [ ]* 1.4 Write property test for concatenación de documentosProceso
    - **Property 7: Concatenación de documentosProceso preserva URLs previas**
    - **Validates: Requirements 6.4, 7.3, 7.5**

- [x] 2. Implementar componente Multi-File Upload en el frontend
  - [x] 2.1 Crear funciones del componente de carga múltiple de archivos en main.js.html
    - Implementar `initMultiFileUpload(containerId)` que renderiza el área de drop/selección
    - Implementar `agregarArchivos(files)` con validación de tipo MIME, tamaño (≤15MB) y cantidad (≤10 total)
    - Implementar `eliminarArchivo(index)` que remueve del array y actualiza el UI
    - Implementar `renderizarListaArchivos()` que muestra nombre truncado (50 chars), tamaño formateado (KB/MB) y botón eliminar
    - Implementar `convertirArchivosABase64()` usando FileReader (NO createObjectURL)
    - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5, 5.6, 5.7, 6.1, 6.7_

  - [ ]* 2.2 Write property test for validación de archivos
    - **Property 2: Validación de archivos acepta/rechaza correctamente según tipo y tamaño**
    - **Validates: Requirements 1.6, 5.3, 5.4, 5.5, 6.6**

  - [ ]* 2.3 Write property test for conteo de archivos
    - **Property 3: El conteo de archivos nunca excede 10 tras cualquier secuencia de operaciones**
    - **Validates: Requirements 5.1, 5.2, 5.7**

  - [ ]* 2.4 Write property test for display de archivos
    - **Property 9: Display de archivos trunca nombre a 50 caracteres y formatea tamaño**
    - **Validates: Requirements 5.6**

- [x] 3. Checkpoint - Verificar componente de upload
  - Ensure all tests pass, ask the user if questions arise.

- [x] 4. Implementar formulario unificado de procesos especiales en el frontend
  - [x] 4.1 Crear HTML del formulario unificado y cards de selección en index.html
    - Agregar card `#cardCorrecciones` junto a las existentes `#cardOtroSi` y `#cardCesion`
    - Reemplazar formularios individuales `#formOtroSi` y `#formCesion` por un único `#formProcesoEspecial`
    - Incluir encabezado dinámico, indicador de pasos (1-2-3), textarea `#observacionesProceso` (maxlength 2000), y contenedor para el componente multi-file upload
    - Eliminar campos legacy de Cesión (tipoDoc, numDoc, nombre, teléfono del nuevo propietario)
    - _Requirements: 1.1, 1.2, 1.3, 1.4, 2.1, 2.2, 2.3, 2.4, 3.1, 3.2, 3.3, 3.4_

  - [x] 4.2 Implementar función `seleccionarProceso(tipo)` en main.js.html
    - Mostrar formulario unificado con encabezado dinámico según tipo ("Gestión de Otro Sí" | "Gestión de Cesión del Contrato" | "Gestión de Correcciones")
    - Actualizar indicador de pasos con texto correspondiente al proceso
    - Inicializar componente multi-file upload
    - Resetear estado del formulario (observaciones vacías, archivos vacíos)
    - Asignar `procesoEspecialSeleccionado` con el tipo seleccionado
    - _Requirements: 1.3, 1.4, 2.3, 2.4, 3.3, 3.4_

  - [ ]* 4.3 Write property test for validación de formulario
    - **Property 1: Validación de formulario bloquea guardado cuando faltan campos obligatorios**
    - **Validates: Requirements 1.5, 2.5, 3.6, 4.4**

- [x] 5. Implementar lógica de guardado con bypass OTP para Correcciones
  - [x] 5.1 Modificar función de guardado `guardarGestionUnificada` en main.js.html
    - Si `procesoEspecialSeleccionado === 'CORRECCION'`: omitir verificación OTP, asignar `estadoGestion = "Caso Especial"` automáticamente
    - Si `procesoEspecialSeleccionado === 'OTRO_SI'` o `'CESION'`: verificar que `otpValidado === true` antes de permitir guardado
    - Validar que observaciones no estén vacías (trim)
    - Validar que haya al menos 1 archivo cargado (para Otro Sí)
    - Recopilar archivos del componente multi-file via `convertirArchivosABase64()`
    - Construir objeto `datos` con `procesoEspecial` correspondiente (OTRO_SI, CESION, CORRECCION)
    - Enviar via `google.script.run.guardarGestionRenovacion(datos, observaciones, archivosBase64)`
    - Mostrar SweetAlert de éxito o error según respuesta
    - _Requirements: 1.5, 2.5, 3.5, 3.6, 4.1, 4.2, 4.3, 4.4, 4.5_

  - [ ]* 5.2 Write property test for requerimiento de OTP
    - **Property 4: Requerimiento de OTP depende del tipo de proceso**
    - **Validates: Requirements 4.1, 4.2**

  - [ ]* 5.3 Write property test for guardado de Correcciones
    - **Property 5: Guardado de Correcciones produce estado y procesoEspecial correctos**
    - **Validates: Requirements 3.5, 4.3, 4.5**

- [x] 6. Checkpoint - Verificar flujo frontend completo
  - Ensure all tests pass, ask the user if questions arise.

- [x] 7. Modificar backend para soporte multi-archivo y merge no destructivo
  - [x] 7.1 Modificar `guardarArchivosEnCarpeta` en Renovaciones.js para procesar arreglo documentosProceso
    - Iterar sobre `archivosBase64.documentosProceso` (arreglo de objetos `{name, mimeType, data}`)
    - Decodificar cada archivo base64 → Blob usando `Utilities.newBlob(Utilities.base64Decode(data), mimeType, nombre)`
    - Nombrar archivos con formato `{PREFIJO}_{poliza}_{index}.{extension}` (ej: `OTRO_SI_12345_1.pdf`)
    - Guardar en subcarpeta "Procesos Especiales" dentro de la carpeta de renovación
    - Si la subcarpeta no existe, crearla
    - Retornar arreglo de objetos `{nombre, url}` en `urlsNuevas.documentosProceso`
    - _Requirements: 6.1, 6.2, 6.3, 6.4, 6.5_

  - [ ]* 7.2 Write property test for nomenclatura de archivos
    - **Property 10: Nomenclatura de archivos en backend sigue el patrón convencional**
    - **Validates: Requirements 6.2**

  - [x] 7.3 Modificar `guardarGestionRenovacion` en Renovaciones.js para usar merge no destructivo
    - Leer JSON actual de Columna_Gestion (columna 6)
    - Usar `mergeJsonColumnaGestion(currentJson, finalJsonData)` en lugar de sobrescritura directa
    - Concatenar `documentosProceso` nuevos al arreglo existente
    - Preservar todos los campos existentes no actualizados (poliza, solicitud, email, celular, detalleFinanciero, etc.)
    - Si JSON actual es inválido o vacío, inicializar objeto nuevo
    - Registrar `procesoEspecial` en el historial (columna 7) con fecha, usuario, observación, estado y procesoEspecial
    - Para Correcciones: asignar estado "Caso Especial" y procesoEspecial "CORRECCION" en historial
    - _Requirements: 7.1, 7.2, 7.3, 7.4, 7.5, 7.6, 7.7, 7.8_

  - [ ]* 7.4 Write property test for append al historial
    - **Property 6: Append al historial preserva todas las entradas previas**
    - **Validates: Requirements 7.2**

- [x] 8. Checkpoint - Verificar backend completo
  - Ensure all tests pass, ask the user if questions arise.

- [x] 9. Integración y wiring final
  - [x] 9.1 Conectar eventos de UI con funciones de lógica en main.js.html
    - Vincular click en `#cardOtroSi` → `seleccionarProceso('otroSi')`
    - Vincular click en `#cardCesion` → `seleccionarProceso('cesion')`
    - Vincular click en `#cardCorrecciones` → `seleccionarProceso('correcciones')`
    - Vincular botón guardar del formulario unificado → `guardarGestionUnificada()`
    - Vincular input file change → `agregarArchivos(files)`
    - Vincular botones eliminar archivo → `eliminarArchivo(index)`
    - Asegurar que no quede código huérfano de los formularios legacy eliminados
    - _Requirements: 1.1, 2.1, 3.1, 1.5, 2.5, 3.6_

  - [ ]* 9.2 Write unit tests for flujo de integración
    - Test: seleccionar cada proceso muestra encabezado correcto
    - Test: guardar sin observaciones muestra error
    - Test: guardar Correcciones sin OTP procede correctamente
    - Test: guardar Otro Sí sin OTP bloquea guardado
    - Test: archivos se envían correctamente al backend
    - _Requirements: 1.3, 1.5, 2.3, 2.5, 3.3, 3.6, 4.1, 4.2_

- [x] 10. Final checkpoint - Verificar integración completa
  - Ensure all tests pass, ask the user if questions arise.

## Notes

- Tasks marked with `*` are optional and can be skipped for faster MVP
- Each task references specific requirements for traceability
- Checkpoints ensure incremental validation
- Property tests validate universal correctness properties from the design document
- Unit tests validate specific examples and edge cases
- El proyecto es Google Apps Script (legacy) — se mantiene jQuery + Bootstrap + SweetAlert2 como stack frontend
- Para property tests se recomienda fast-check con Jest local via clasp
- La función `retryDrive` existente se reutiliza para reintentos de Google Drive

## Task Dependency Graph

```json
{
  "waves": [
    { "id": 0, "tasks": ["1.1", "1.2"] },
    { "id": 1, "tasks": ["1.3", "1.4", "2.1"] },
    { "id": 2, "tasks": ["2.2", "2.3", "2.4", "4.1"] },
    { "id": 3, "tasks": ["4.2", "4.3"] },
    { "id": 4, "tasks": ["5.1", "7.1"] },
    { "id": 5, "tasks": ["5.2", "5.3", "7.2", "7.3"] },
    { "id": 6, "tasks": ["7.4", "9.1"] },
    { "id": 7, "tasks": ["9.2"] }
  ]
}
```
