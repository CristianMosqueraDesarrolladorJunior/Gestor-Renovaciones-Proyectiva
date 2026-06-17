# Requirements Document

## Introduction

Refinamiento de la sección "Gestión de Procesos Especiales" en el sistema CRM de renovaciones de pólizas (Google Apps Script). El objetivo es simplificar los formularios de los procesos "Otro Sí" y "Cesión del Contrato", agregar un tercer proceso "Correcciones", corregir el cargue de documentos que actualmente no funciona, y permitir que "Correcciones" se guarde sin verificación OTP.

## Glossary

- **CRM_Renovaciones**: Sistema de gestión de renovaciones de pólizas construido con Google Apps Script, desplegado mediante clasp.
- **Proceso_Especial**: Trámite adicional que un asesor puede gestionar durante la renovación de una póliza. Los tipos son: Otro Sí, Cesión del Contrato y Correcciones.
- **Asesor**: Usuario del sistema que gestiona la renovación de pólizas y ejecuta procesos especiales.
- **Formulario_Proceso_Especial**: Interfaz de captura de datos para un proceso especial, compuesta por un campo de observaciones y un área de carga de documentos.
- **Columna_Gestion**: Columna en la hoja de cálculo "JSON" donde se almacena la información de gestión de cada póliza en formato JSON.
- **Carpeta_Renovacion**: Carpeta en Google Drive asociada al usuario donde se almacenan los documentos de la renovación.
- **OTP**: Código de verificación de un solo uso enviado al cliente para validar la operación.
- **Observaciones_Proceso**: Campo de texto libre donde el asesor registra el resumen puntual del caso del cliente para el proceso especial.

## Requirements

### Requirement 1: Simplificación del formulario Otro Sí

**User Story:** Como asesor, quiero que el formulario de "Otro Sí" muestre únicamente un campo de observaciones y un área de carga de documentos, para agilizar la gestión sin campos innecesarios.

#### Acceptance Criteria

1. WHEN el asesor selecciona el proceso "Otro Sí", THE CRM_Renovaciones SHALL mostrar un formulario con exactamente dos campos: un textarea para Observaciones_Proceso con un límite máximo de 2000 caracteres y un área de carga de hasta 10 documentos que acepte archivos en formato PDF, JPG o PNG con un tamaño máximo de 15 MB por archivo.
2. WHEN el formulario "Otro Sí" se muestra, THE CRM_Renovaciones SHALL eliminar todos los campos adicionales que existían previamente (campo único de PDF, alertas específicas de contrato firmado).
3. WHEN el formulario "Otro Sí" se muestra, THE CRM_Renovaciones SHALL indicar en el encabezado del formulario el texto "Gestión de Otro Sí" para que el asesor identifique claramente qué proceso está gestionando.
4. WHEN el asesor selecciona "Otro Sí", THE CRM_Renovaciones SHALL mostrar un indicador de pasos numerado (sin animaciones ni elementos decorativos adicionales) que describa el flujo del proceso: (1) Registrar observaciones del caso, (2) Cargar documentos de soporte, (3) Guardar gestión.
5. IF el asesor intenta guardar la gestión "Otro Sí" sin haber ingresado texto en el campo Observaciones_Proceso o sin haber cargado al menos 1 documento, THEN THE CRM_Renovaciones SHALL impedir el guardado y mostrar un mensaje de error indicando los campos obligatorios faltantes.
6. IF el asesor intenta cargar un archivo que excede 15 MB o cuyo formato no es PDF, JPG ni PNG, THEN THE CRM_Renovaciones SHALL rechazar el archivo y mostrar un mensaje de error indicando la restricción de formato o tamaño.

### Requirement 2: Simplificación del formulario Cesión del Contrato

**User Story:** Como asesor, quiero que el formulario de "Cesión del Contrato" muestre únicamente un campo de observaciones y un área de carga de documentos, para eliminar los campos de datos del nuevo propietario que ya no son necesarios en este paso.

#### Acceptance Criteria

1. WHEN el asesor selecciona el proceso "Cesión del Contrato", THE CRM_Renovaciones SHALL mostrar un formulario con exactamente dos campos: un textarea para Observaciones_Proceso (obligatorio, máximo 2000 caracteres) y un área de carga de hasta 10 documentos.
2. WHEN el formulario "Cesión del Contrato" se muestra, THE CRM_Renovaciones SHALL eliminar los campos de tipo de documento, número de documento, nombre y teléfono del nuevo propietario, de modo que dichos campos no sean visibles ni editables.
3. THE CRM_Renovaciones SHALL indicar en el encabezado del formulario el texto "Gestión de Cesión del Contrato" para que el asesor identifique qué proceso está gestionando.
4. WHEN el asesor selecciona "Cesión del Contrato", THE CRM_Renovaciones SHALL mostrar un indicador de pasos que describa el flujo del proceso: (1) Registrar observaciones del caso, (2) Cargar documentos de soporte, (3) Guardar gestión.
5. IF el asesor intenta guardar el formulario "Cesión del Contrato" con el campo Observaciones_Proceso vacío, THEN THE CRM_Renovaciones SHALL impedir el guardado y mostrar un mensaje indicando que las observaciones son obligatorias.

### Requirement 3: Nuevo proceso "Correcciones"

**User Story:** Como asesor, quiero disponer de un tercer botón "Correcciones" en la sección de procesos especiales, para gestionar correcciones de póliza con el mismo formulario simplificado.

#### Acceptance Criteria

1. THE CRM_Renovaciones SHALL mostrar un tercer botón "Correcciones" en la sección de selección de procesos especiales, junto a "Otro Sí" y "Cesión del Contrato".
2. WHEN el asesor selecciona el proceso "Correcciones", THE CRM_Renovaciones SHALL mostrar un formulario con exactamente dos campos: un textarea para Observaciones_Proceso (obligatorio, máximo 2000 caracteres) y un área de carga de hasta 10 documentos.
3. THE CRM_Renovaciones SHALL indicar en el encabezado del formulario el texto "Gestión de Correcciones" para que el asesor identifique claramente qué proceso está gestionando.
4. WHEN el asesor selecciona "Correcciones", THE CRM_Renovaciones SHALL mostrar un indicador de pasos numerado que describa el flujo del proceso: (1) Registrar observaciones de la corrección, (2) Cargar documentos de soporte, (3) Guardar gestión.
5. WHEN el asesor guarda una gestión con proceso "Correcciones", THE CRM_Renovaciones SHALL registrar el valor "CORRECCION" en el campo procesoEspecial del JSON de la Columna_Gestion.
6. IF el asesor intenta guardar la gestión "Correcciones" sin haber ingresado texto en el campo Observaciones_Proceso, THEN THE CRM_Renovaciones SHALL impedir el guardado y mostrar un mensaje indicando que las observaciones son obligatorias.

### Requirement 4: Envío sin verificación OTP para Correcciones

**User Story:** Como asesor, quiero que al guardar una gestión con proceso "Correcciones" no se requiera verificación OTP, para agilizar las correcciones que no necesitan validación del cliente.

#### Acceptance Criteria

1. WHEN el asesor guarda una gestión con proceso especial "Correcciones", THE CRM_Renovaciones SHALL omitir el paso de verificación OTP (no mostrar la interfaz de envío ni validación de código) y proceder directamente al guardado en un tiempo máximo de 30 segundos.
2. WHEN el asesor guarda una gestión con proceso especial "Otro Sí" o "Cesión del Contrato", THE CRM_Renovaciones SHALL requerir la validación OTP completada (estado OTP_VALIDADO) antes de permitir el guardado, manteniendo el flujo de validación existente.
3. WHEN el proceso "Correcciones" se guarda sin OTP, THE CRM_Renovaciones SHALL asignar automáticamente el valor "Caso Especial" en el campo estadoGestion del registro en la Columna_Gestion, sin requerir que el asesor lo seleccione manualmente del dropdown de estado.
4. IF el asesor intenta guardar una gestión con proceso "Correcciones" sin haber ingresado texto en el campo Observaciones_Proceso, THEN THE CRM_Renovaciones SHALL bloquear el guardado y mostrar un mensaje indicando que las observaciones son obligatorias.
5. WHEN el proceso "Correcciones" se guarda exitosamente sin OTP, THE CRM_Renovaciones SHALL registrar en la entrada del historial de observaciones (columna 7) el campo procesoEspecial con valor "CORRECCION" y el campo estado con valor "Caso Especial".

### Requirement 5: Carga de hasta 10 documentos por proceso

**User Story:** Como asesor, quiero poder cargar hasta 10 documentos en cada proceso especial, para adjuntar toda la evidencia necesaria sin restricciones de cantidad.

#### Acceptance Criteria

1. THE CRM_Renovaciones SHALL permitir al asesor seleccionar hasta 10 archivos en el área de carga de documentos de cada proceso especial.
2. WHEN el asesor intenta cargar archivos que, sumados a los ya seleccionados en la sesión actual, superen un total de 10, THE CRM_Renovaciones SHALL rechazar los archivos excedentes y mostrar un mensaje de advertencia indicando que el máximo permitido es 10 documentos en total por proceso.
3. THE CRM_Renovaciones SHALL aceptar únicamente archivos cuya extensión sea PDF, JPG, JPEG o PNG, con un tamaño máximo de 15 MB por archivo.
4. IF el asesor selecciona un archivo con extensión diferente a PDF, JPG, JPEG o PNG, THEN THE CRM_Renovaciones SHALL rechazar el archivo y mostrar un mensaje indicando los formatos permitidos.
5. WHEN el asesor carga un archivo que excede 15 MB, THE CRM_Renovaciones SHALL rechazar el archivo y mostrar un mensaje indicando que el tamaño máximo permitido es 15 MB por archivo.
6. THE CRM_Renovaciones SHALL mostrar una lista con el nombre (truncado a 50 caracteres si es más largo) y el tamaño en KB o MB de cada archivo seleccionado, con un botón de eliminar por cada archivo individual antes de guardar.
7. WHEN el asesor elimina un archivo de la lista de seleccionados, THE CRM_Renovaciones SHALL actualizar el contador de archivos y permitir la carga de nuevos archivos hasta completar el máximo de 10.

### Requirement 6: Corrección del cargue de documentos a Google Drive

**User Story:** Como asesor, quiero que los documentos cargados se almacenen correctamente en Google Drive, para que la evidencia documental quede disponible y vinculada a la gestión.

#### Acceptance Criteria

1. WHEN el asesor guarda una gestión con documentos adjuntos, THE CRM_Renovaciones SHALL convertir cada archivo a base64 en el frontend y enviarlo al backend mediante google.script.run para su almacenamiento en Google Drive.
2. WHEN el backend recibe los archivos en base64, THE CRM_Renovaciones SHALL decodificarlos y almacenarlos en la subcarpeta "Procesos Especiales" dentro de la Carpeta_Renovacion del usuario en Google Drive, nombrando cada archivo con el formato "{PREFIJO}_{numero_poliza}.{extension}".
3. IF la subcarpeta "Procesos Especiales" no existe en la Carpeta_Renovacion, THEN THE CRM_Renovaciones SHALL crearla antes de almacenar los archivos.
4. WHEN los archivos se almacenan exitosamente en Google Drive, THE CRM_Renovaciones SHALL guardar las URLs de los archivos en el JSON de la Columna_Gestion bajo la clave "documentosProceso" como un arreglo de objetos con nombre y URL.
5. IF ocurre un error al almacenar archivos en Google Drive, THEN THE CRM_Renovaciones SHALL mostrar al asesor un mensaje de error indicando que el servicio de Google Drive falló y que debe reintentar, y no SHALL persistir la gestión en la hoja de cálculo.
6. IF el asesor adjunta un archivo que excede 15 MB de tamaño o cuyo tipo no es PDF, imagen PNG o imagen JPG, THEN THE CRM_Renovaciones SHALL rechazar el archivo antes del envío al backend y mostrar un mensaje indicando el límite de tamaño o los formatos permitidos.
7. WHEN el frontend procesa archivos para envío, THE CRM_Renovaciones SHALL utilizar FileReader para obtener la representación base64 del archivo y NO SHALL generar URLs locales con createObjectURL como resultado final del cargue.

### Requirement 7: Almacenamiento de datos en la Columna Gestion (sin sobrescritura)

**User Story:** Como sistema, quiero que los datos de procesos especiales se agreguen al JSON existente en la Columna_Gestion sin sobrescribir valores previos, para mantener la trazabilidad completa de cada gestión.

#### Acceptance Criteria

1. WHEN el asesor guarda un proceso especial, THE CRM_Renovaciones SHALL actualizar el campo "procesoEspecial" en el JSON de la Columna_Gestion con el valor correspondiente: "OTRO_SI", "CESION" o "CORRECCION".
2. WHEN el asesor guarda un proceso especial, THE CRM_Renovaciones SHALL agregar una nueva entrada al historial de observaciones (columna 7 de la hoja) con los campos fecha (formato dd/MM/yyyy HH:mm:ss en zona America/Bogota), usuario (email de sesión activa), observacion, estado, seguimiento, y procesoEspecial, sin eliminar entradas previas del historial.
3. WHEN el asesor guarda un proceso especial con documentos, THE CRM_Renovaciones SHALL almacenar las URLs de los nuevos documentos en el campo "documentosProceso" del JSON de la Columna_Gestion como un arreglo de objetos con nombre y URL, preservando las URLs previamente almacenadas.
4. THE CRM_Renovaciones SHALL leer el JSON existente de la Columna_Gestion, aplicar un merge superficial (spread) de los nuevos campos sobre los existentes, y escribir el resultado combinado de vuelta a la celda, de modo que los campos no incluidos en la nueva escritura conserven su valor previo.
5. IF el JSON existente en la Columna_Gestion contiene URLs de documentos previos en "documentosProceso", THEN THE CRM_Renovaciones SHALL concatenar los nuevos documentos al final del arreglo existente sin eliminar los previos.
6. THE CRM_Renovaciones SHALL preservar todos los campos existentes del JSON que no son parte de la actualización actual, incluyendo pero no limitado a: poliza, solicitud, email, celular, detalleFinanciero, datosCesion, cotizacionTotal, documento, asegurado, sarlaftArchivoURL y propietarioDocURL.
7. IF el contenido de la Columna_Gestion no es un JSON válido o la celda está vacía, THEN THE CRM_Renovaciones SHALL inicializar un objeto JSON nuevo con los datos del proceso especial actual sin intentar hacer merge con contenido previo.
8. IF ocurre un error al escribir en la Columna_Gestion (fallo de servicio de Google Sheets o Drive), THEN THE CRM_Renovaciones SHALL retornar un objeto con status "error" y un mensaje descriptivo al usuario sin modificar parcialmente los datos de la fila.
