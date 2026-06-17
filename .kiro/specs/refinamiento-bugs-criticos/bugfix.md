# Documento de Requisitos de Corrección de Bugs

## Introducción

Este documento describe 8 bugs críticos identificados en el sistema "Gestor Renovaciones Proyectiva", una aplicación web de Google Apps Script para gestión de renovaciones de pólizas de seguros. Los bugs afectan funcionalidades clave: envío de mensajes WhatsApp, inicialización de tablas de datos, persistencia de gestión, renderizado de etapas, gestión de temporizadores OTP, navegación entre leads, reporte de fallos en carga de archivos, y asignación equitativa de leads.

---

## Bug Analysis

### Current Behavior (Defect)

1.1 WHEN las funciones `RenovaSendWppVencida()` o `RenovaSendWppProxVen()` se ejecutan en el runtime de Google Apps Script THEN el sistema falla con error de referencia porque `fetch()` y `new Headers()` no existen en el entorno servidor de GAS

1.2 WHEN `data.renovaciones` es `undefined` (no solo `null`) y se pasa a `initializeDataTableRenovaciones()` THEN el sistema lanza un TypeError al intentar iterar sobre `undefined` como dataset del DataTable

1.3 WHEN el JSON existente en la columna de gestión de la hoja de cálculo está malformado o corrupto y se invoca `guardarGestionRenovacion()` THEN la operación `JSON.parse()` dentro de `mergeJsonColumnaGestion()` lanza una excepción que puede interrumpir el flujo de guardado sin notificar al usuario

1.4 WHEN `currentEtapaIndex` tiene un valor que excede la longitud del arreglo `etapaNombres` o `etapaNombresRenovaciones` THEN el acceso a `etapaNombres[currentEtapaIndex]` retorna `undefined`, causando errores de renderizado en `renderEtapas()`

1.5 WHEN el usuario cierra el modal de renovaciones sin completar la validación OTP THEN los intervalos de progreso (`setInterval`) asociados a la carga de archivos y barras de progreso continúan ejecutándose en segundo plano, causando fugas de memoria y posibles glitches si el modal se reabre

1.6 WHEN el usuario navega al siguiente/anterior lead sin haber guardado cambios en la etapa "Emisión de Póliza" THEN el sistema navega directamente sin advertencia, perdiendo todos los datos del formulario no guardados

1.7 WHEN la carga de archivos a Google Drive falla parcialmente (por timeout o error de Drive) pero el registro de gestión se guarda exitosamente THEN el sistema no reporta al usuario que los archivos no se subieron, dando la impresión de éxito total

1.8 WHEN todos los asesores disponibles tienen el mismo `sortingKey` calculado y la misma efectividad (`effectiveness`) en la función `AssignLead()` THEN el algoritmo selecciona al primer asesor encontrado en el arreglo sin un criterio determinístico de desempate, generando distribución desigual

### Expected Behavior (Correct)

2.1 WHEN las funciones `RenovaSendWppVencida()` o `RenovaSendWppProxVen()` se ejecutan en el runtime de Google Apps Script THEN el sistema SHALL usar `UrlFetchApp.fetch()` con las opciones apropiadas (`method`, `headers`, `payload`, `muteHttpExceptions`) para enviar mensajes WhatsApp exitosamente a través de la API de Infobip

2.2 WHEN `data.renovaciones` es `undefined` o `null` y se pasa a `initializeDataTableRenovaciones()` THEN el sistema SHALL usar un arreglo vacío `[]` como dataset por defecto, inicializando la tabla sin errores y mostrando un estado vacío al usuario

2.3 WHEN el JSON existente en la columna de gestión está malformado o corrupto y se invoca `guardarGestionRenovacion()` THEN el sistema SHALL capturar el error de parseo, registrar un log de advertencia con el contenido corrupto, iniciar con un objeto vacío `{}` para el merge, y continuar la operación de guardado notificando al usuario que los datos previos de gestión no pudieron recuperarse

2.4 WHEN `currentEtapaIndex` tiene un valor que excede la longitud del arreglo de etapas disponible THEN el sistema SHALL limitar (clamp) el índice al rango válido `[0, etapas.length - 1]` antes de acceder al arreglo, previniendo accesos a posiciones `undefined`

2.5 WHEN el usuario cierra el modal de renovaciones sin completar la validación OTP THEN el sistema SHALL limpiar todos los intervalos activos (`clearInterval`) asociados a barras de progreso y temporizadores, liberando recursos y evitando ejecuciones en segundo plano

2.6 WHEN el usuario navega al siguiente/anterior lead y tiene cambios sin guardar en la etapa "Emisión de Póliza" THEN el sistema SHALL mostrar un diálogo de confirmación preguntando si desea descartar los cambios o cancelar la navegación, de forma consistente con el comportamiento de las demás etapas

2.7 WHEN la carga de archivos a Google Drive falla parcialmente pero el registro de gestión se guarda exitosamente THEN el sistema SHALL retornar un estado de éxito parcial (`status: "partial"`) con un mensaje claro indicando qué archivos fallaron, permitiendo al usuario reintentar la carga

2.8 WHEN todos los asesores disponibles tienen el mismo `sortingKey` y la misma efectividad en `AssignLead()` THEN el sistema SHALL aplicar un criterio de desempate determinístico (orden alfabético por email del asesor) para garantizar distribución predecible y auditable

### Unchanged Behavior (Regression Prevention)

3.1 WHEN las funciones de envío WhatsApp (`sendWhatsAppSarlaft`, `enviarOtpWhatsapp`) que ya usan `UrlFetchApp.fetch()` correctamente se ejecutan THEN el sistema SHALL CONTINUE TO enviar mensajes exitosamente sin cambios en su comportamiento

3.2 WHEN `data.renovaciones` contiene un arreglo válido con datos y se pasa a `initializeDataTableRenovaciones()` THEN el sistema SHALL CONTINUE TO inicializar el DataTable correctamente mostrando todos los registros

3.3 WHEN el JSON existente en la columna de gestión es válido y bien formado y se invoca `guardarGestionRenovacion()` THEN el sistema SHALL CONTINUE TO realizar el merge no destructivo preservando campos existentes y concatenando `documentosProceso` correctamente

3.4 WHEN `currentEtapaIndex` tiene un valor dentro del rango válido del arreglo de etapas THEN el sistema SHALL CONTINUE TO renderizar la etapa correspondiente sin alteraciones en el flujo de navegación

3.5 WHEN el usuario completa la validación OTP exitosamente y cierra el modal THEN el sistema SHALL CONTINUE TO registrar la validación correctamente y mantener el estado `OTP_VALIDADO`

3.6 WHEN el usuario navega entre leads en etapas distintas a "Emisión de Póliza" y tiene cambios sin guardar THEN el sistema SHALL CONTINUE TO mostrar el diálogo de confirmación existente antes de navegar

3.7 WHEN la carga de archivos a Google Drive se completa exitosamente junto con el registro de gestión THEN el sistema SHALL CONTINUE TO retornar `status: "ok"` con las URLs de los archivos guardados

3.8 WHEN los asesores tienen diferentes `sortingKey` o diferente efectividad en `AssignLead()` THEN el sistema SHALL CONTINUE TO seleccionar al asesor con menor `sortingKey` y mayor efectividad como desempate primario

---

## Derivación de la Condición de Bug

### Bug 1: fetch() en GAS

```pascal
FUNCTION isBugCondition(X)
  INPUT: X of type FunctionCall
  OUTPUT: boolean
  
  RETURN X.runtime = "GoogleAppsScript" AND X.function IN {RenovaSendWppVencida, RenovaSendWppProxVen}
END FUNCTION
```

```pascal
// Property: Fix Checking - Uso de UrlFetchApp
FOR ALL X WHERE isBugCondition(X) DO
  result ← X.function'(X.params)
  ASSERT uses_UrlFetchApp(result) AND no_reference_error(result)
END FOR
```

### Bug 2: DataTable con undefined

```pascal
FUNCTION isBugCondition(X)
  INPUT: X of type DataTableInput
  OUTPUT: boolean
  
  RETURN X.dataSet = undefined OR X.dataSet = null
END FUNCTION
```

```pascal
// Property: Fix Checking - Manejo de undefined
FOR ALL X WHERE isBugCondition(X) DO
  result ← initializeDataTableRenovaciones'(X.dataSet)
  ASSERT no_TypeError(result) AND table_initialized_empty(result)
END FOR
```

### Bug 3: JSON corrupto en merge

```pascal
FUNCTION isBugCondition(X)
  INPUT: X of type CellContent
  OUTPUT: boolean
  
  RETURN X.jsonString != "" AND NOT is_valid_json(X.jsonString)
END FUNCTION
```

```pascal
// Property: Fix Checking - Manejo de JSON corrupto
FOR ALL X WHERE isBugCondition(X) DO
  result ← guardarGestionRenovacion'(X.datos, X.observaciones, X.archivos)
  ASSERT no_unhandled_exception(result) AND (result.status = "ok" OR result.status = "partial") AND log_contains_warning(result)
END FOR
```

### Bug 4: Índice fuera de rango en etapas

```pascal
FUNCTION isBugCondition(X)
  INPUT: X of type EtapaNavigation
  OUTPUT: boolean
  
  RETURN X.currentEtapaIndex >= X.etapasArray.length OR X.currentEtapaIndex < 0
END FUNCTION
```

```pascal
// Property: Fix Checking - Clamp de índice
FOR ALL X WHERE isBugCondition(X) DO
  result ← renderEtapas'(X.etapasArray[clamp(X.currentEtapaIndex)])
  ASSERT no_undefined_access(result) AND index_in_valid_range(result)
END FOR
```

### Bug 5: Intervalos sin limpiar al cerrar modal

```pascal
FUNCTION isBugCondition(X)
  INPUT: X of type ModalEvent
  OUTPUT: boolean
  
  RETURN X.action = "close" AND X.hasActiveIntervals = true AND X.otpNotValidated = true
END FUNCTION
```

```pascal
// Property: Fix Checking - Limpieza de intervalos
FOR ALL X WHERE isBugCondition(X) DO
  result ← closeModal'(X)
  ASSERT active_intervals_count(result) = 0
END FOR
```

### Bug 6: Navegación sin guardar en Emisión de Póliza

```pascal
FUNCTION isBugCondition(X)
  INPUT: X of type LeadNavigation
  OUTPUT: boolean
  
  RETURN X.currentEtapa = "Emisión de Póliza" AND X.hasUnsavedChanges = true AND X.forceNavigation = false
END FUNCTION
```

```pascal
// Property: Fix Checking - Confirmación antes de navegar
FOR ALL X WHERE isBugCondition(X) DO
  result ← navigateToLead'(X.direction)
  ASSERT confirmation_dialog_shown(result) AND (user_confirms → navigation_proceeds) AND (user_cancels → navigation_blocked)
END FOR
```

### Bug 7: Fallo parcial de archivos no reportado

```pascal
FUNCTION isBugCondition(X)
  INPUT: X of type SaveOperation
  OUTPUT: boolean
  
  RETURN X.hasFiles = true AND X.driveUploadFails = true AND X.gestionSaveSucceeds = true
END FUNCTION
```

```pascal
// Property: Fix Checking - Reporte de fallo parcial
FOR ALL X WHERE isBugCondition(X) DO
  result ← guardarGestionRenovacion'(X.datos, X.obs, X.archivos)
  ASSERT result.status = "partial" AND result.message CONTAINS "archivos" AND user_notified(result)
END FOR
```

### Bug 8: Desempate en asignación de leads

```pascal
FUNCTION isBugCondition(X)
  INPUT: X of type AssignmentInput
  OUTPUT: boolean
  
  RETURN EXISTS a1, a2 IN X.agents WHERE a1.sortingKey = a2.sortingKey AND a1.effectiveness = a2.effectiveness
END FUNCTION
```

```pascal
// Property: Fix Checking - Desempate determinístico
FOR ALL X WHERE isBugCondition(X) DO
  result1 ← AssignLead'(X)
  result2 ← AssignLead'(X)
  ASSERT result1.email = result2.email AND is_alphabetically_first(result1.email, tied_agents)
END FOR
```

### Propiedad de Preservación Global

```pascal
// Property: Preservation Checking
FOR ALL X WHERE NOT isBugCondition(X) DO
  ASSERT F(X) = F'(X)
END FOR
```

Esto garantiza que para todas las entradas que no activan ninguna de las condiciones de bug, el código corregido se comporta de forma idéntica al código original.
