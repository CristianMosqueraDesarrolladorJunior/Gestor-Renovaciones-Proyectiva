# Refinamiento Bugs Críticos — Bugfix Design

## Overview

Este documento formaliza el diseño de corrección para 8 bugs críticos identificados en el sistema "Gestor Renovaciones Proyectiva". Los bugs abarcan: uso de APIs de navegador en entorno servidor GAS, manejo de datos nulos/undefined, parseo de JSON corrupto, acceso a índices fuera de rango, fugas de memoria por intervalos no limpiados, pérdida de datos por navegación sin guardar, fallos parciales de carga no reportados, y asignación no determinística de leads. La estrategia de corrección es mínima e incremental: cada bug se corrige de forma aislada para evitar regresiones cruzadas.

## Glossary

- **Bug_Condition (C)**: Condición que activa el bug — el conjunto de entradas/estados que producen el comportamiento defectuoso
- **Property (P)**: Comportamiento esperado correcto cuando se activa la condición de bug
- **Preservation**: Comportamiento existente que NO debe cambiar tras aplicar la corrección
- **GAS**: Google Apps Script — runtime servidor donde se ejecutan `Código.js` y `Renovaciones.js`
- **UrlFetchApp**: API nativa de GAS para realizar peticiones HTTP desde el servidor
- **DataTable**: Plugin jQuery DataTables usado para renderizar tablas de leads/renovaciones
- **mergeJsonColumnaGestion()**: Función en `Renovaciones.js` que realiza merge no destructivo de JSON en la columna de gestión
- **renderEtapas()**: Función en `main.js.html` que renderiza el formulario de la etapa actual del funnel
- **AssignLead()**: Función en `Código.js` que asigna leads a asesores según disponibilidad y efectividad
- **currentEtapaIndex**: Variable global en `main.js.html` que indica la etapa visible del funnel
- **OTP**: One-Time Password — código de verificación enviado por WhatsApp/SMS

## Bug Details

### Bug Condition

Los 8 bugs se manifiestan bajo condiciones distintas pero comparten un patrón: ausencia de validación defensiva en los límites del sistema (entradas externas, datos de hoja de cálculo, estado de UI).

**Formal Specification (Compuesta):**

```
FUNCTION isBugCondition(input)
  INPUT: input of type SystemInput
  OUTPUT: boolean
  
  RETURN isBug1(input) OR isBug2(input) OR isBug3(input) OR isBug4(input)
         OR isBug5(input) OR isBug6(input) OR isBug7(input) OR isBug8(input)
END FUNCTION

FUNCTION isBug1(input)
  // fetch() no existe en GAS
  RETURN input.runtime = "GoogleAppsScript"
         AND input.function IN {"RenovaSendWppVencida", "RenovaSendWppProxVen"}
         AND input.usesAPI IN {"fetch", "Headers"}
END FUNCTION

FUNCTION isBug2(input)
  // DataTable con undefined
  RETURN input.function = "initializeDataTableRenovaciones"
         AND (input.dataSet = undefined OR input.dataSet = null)
END FUNCTION

FUNCTION isBug3(input)
  // JSON corrupto en columna gestión
  RETURN input.function = "guardarGestionRenovacion"
         AND input.columnaGestionValue != ""
         AND NOT isValidJson(input.columnaGestionValue)
END FUNCTION

FUNCTION isBug4(input)
  // Índice fuera de rango
  RETURN input.function = "renderEtapas"
         AND (input.currentEtapaIndex >= input.etapasArray.length
              OR input.currentEtapaIndex < 0)
END FUNCTION

FUNCTION isBug5(input)
  // Intervalos sin limpiar
  RETURN input.event = "modalClose"
         AND input.activeIntervals.length > 0
         AND input.otpValidated = false
END FUNCTION

FUNCTION isBug6(input)
  // Navegación sin guardar en Emisión de Póliza
  RETURN input.event IN {"navigateNext", "navigatePrev"}
         AND input.currentEtapa = "Emisión de Póliza"
         AND input.hasUnsavedChanges = true
END FUNCTION

FUNCTION isBug7(input)
  // Fallo parcial de archivos no reportado
  RETURN input.function = "guardarGestionRenovacion"
         AND input.archivosBase64 != empty
         AND input.driveUploadPartialFailure = true
         AND input.gestionSaveSuccess = true
END FUNCTION

FUNCTION isBug8(input)
  // Desempate no determinístico en AssignLead
  RETURN input.function = "AssignLead"
         AND EXISTS a1, a2 IN input.agents
           WHERE a1.sortingKey = a2.sortingKey
           AND a1.effectiveness = a2.effectiveness
           AND a1.email != a2.email
END FUNCTION
```

### Examples

- **Bug 1**: `RenovaSendWppVencida()` ejecuta `fetch("https://qgmx9r.api.infobip.com/...")` → `ReferenceError: fetch is not defined`
- **Bug 2**: Backend retorna `{dataFront: undefined}` → `initializeDataTableRenovaciones(undefined)` → `TypeError: Cannot read properties of undefined`
- **Bug 3**: Celda contiene `{poliza: "5012345", canon: NaN, datos: {incompleto` → `JSON.parse()` lanza `SyntaxError`
- **Bug 4**: Lead con `etapaFunel = "Emisión de Póliza"` en contexto Renovaciones (3 etapas) → `currentEtapaIndex = 5` → `etapaNombresRenovaciones[5]` es `undefined`
- **Bug 5**: Usuario abre modal, inicia OTP con `setInterval` para countdown, cierra modal con X → intervalos siguen ejecutándose
- **Bug 6**: Usuario modifica campos en "Emisión de Póliza", presiona "Siguiente Lead" → datos perdidos sin confirmación
- **Bug 7**: 3 archivos a subir, 1 falla por timeout de Drive → respuesta `{status: "ok"}` sin mención del archivo fallido
- **Bug 8**: Dos asesores con `sortingKey = -2.5` y `effectiveness = 85%` → primer asesor en el array siempre gana (depende del orden de lectura de la hoja)

## Expected Behavior

### Preservation Requirements

**Unchanged Behaviors:**
- Las funciones `sendWhatsAppSarlaft` y `enviarOtpWhatsapp` que ya usan `UrlFetchApp.fetch()` correctamente deben seguir funcionando sin cambios
- Cuando `data.renovaciones` contiene un arreglo válido, `initializeDataTableRenovaciones()` debe seguir inicializando la tabla normalmente
- Cuando el JSON de la columna de gestión es válido, `mergeJsonColumnaGestion()` debe seguir realizando el merge no destructivo preservando campos y concatenando `documentosProceso`
- Cuando `currentEtapaIndex` está dentro del rango válido, `renderEtapas()` debe renderizar la etapa sin alteraciones
- Cuando el usuario completa OTP exitosamente y cierra el modal, el estado `OTP_VALIDADO` debe mantenerse
- Cuando el usuario navega entre leads en etapas distintas a "Emisión de Póliza" con cambios sin guardar, el diálogo de confirmación existente debe seguir apareciendo
- Cuando todos los archivos se suben exitosamente, la respuesta debe seguir siendo `{status: "ok"}` con URLs
- Cuando los asesores tienen diferentes `sortingKey` o diferente efectividad, `AssignLead()` debe seguir seleccionando al de menor `sortingKey` y mayor efectividad

**Scope:**
Todas las entradas que NO activan ninguna de las 8 condiciones de bug deben producir exactamente el mismo resultado que el código original. Esto incluye:
- Flujos normales de gestión de renovaciones
- Navegación estándar entre etapas del funnel
- Operaciones CRUD sobre la hoja de cálculo
- Asignación de leads cuando hay un ganador claro

## Hypothesized Root Cause

### Bug 1: fetch() en GAS
Las funciones `RenovaSendWppVencida` y `RenovaSendWppProxVen` fueron escritas copiando código de un snippet de navegador (Postman/browser). Usan `new Headers()` y `fetch()` que son APIs del Web API estándar, no disponibles en el runtime V8 de Google Apps Script. La API correcta es `UrlFetchApp.fetch(url, options)`.

### Bug 2: DataTable con undefined
El backend (`getDatauserPro` o equivalente para renovaciones) puede retornar un objeto donde `dataFront` es `undefined` si la consulta a la hoja no encuentra datos o si hay un error silencioso. La función `initializeDataTableRenovaciones` no valida el parámetro antes de pasarlo a `DataTable({data: dataSet})`.

### Bug 3: JSON corrupto
La columna de gestión en Google Sheets puede contener JSON corrupto por: edición manual de celdas, truncamiento por límite de caracteres (50,000 chars por celda), o escrituras parciales por timeout. Aunque `mergeJsonColumnaGestion` ya tiene un `try/catch`, la función `guardarGestionRenovacion` lee el JSON de la celda en un punto anterior sin protección adecuada.

### Bug 4: Índice fuera de rango
Cuando un lead tiene una `etapaFunel` que no corresponde al arreglo de etapas del contexto actual (ej: etapa de Vida en contexto de Renovaciones), el `indexOf` retorna -1 o el índice calculado excede el arreglo. No hay clamp ni validación de límites.

### Bug 5: Intervalos sin limpiar
Los `setInterval` para countdown de OTP y barras de progreso de carga se crean dentro del modal pero no se registran en un array centralizado. El evento de cierre del modal (`hidden.bs.modal`) no ejecuta `clearInterval` sobre estos timers.

### Bug 6: Navegación sin guardar
Las funciones de navegación entre leads (`loadLeadByIndex(currentLeadIndexInTable + 1)`) no verifican si hay cambios sin guardar cuando la etapa actual es "Emisión de Póliza". Otras etapas sí tienen esta verificación mediante `showCustomConfirm`.

### Bug 7: Fallo parcial de archivos
La función `guardarGestionRenovacion` en `Renovaciones.js` usa `retryDrive` para subir archivos, pero si un archivo falla después de todos los reintentos, el error se captura silenciosamente y la función continúa retornando `{status: "ok"}` sin incluir información sobre archivos fallidos.

### Bug 8: Desempate no determinístico
En `AssignLead()`, la condición de selección es:
```javascript
if (sortingKey < bestSortingKey || (sortingKey === bestSortingKey && effectiveness > highestEffectiveness))
```
Cuando `sortingKey` y `effectiveness` son iguales, ninguna condición se cumple y el primer asesor encontrado (que depende del orden de filas en la hoja) permanece como `bestAgent`. No hay criterio de desempate terciario.

## Correctness Properties

Property 1: Bug Condition - APIs de GAS para envío WhatsApp

_For any_ input donde las funciones `RenovaSendWppVencida` o `RenovaSendWppProxVen` se ejecutan en el runtime de Google Apps Script, la función corregida SHALL usar `UrlFetchApp.fetch()` con opciones `{method, headers, payload, muteHttpExceptions}` sin lanzar `ReferenceError`.

**Validates: Requirements 2.1**

Property 2: Bug Condition - DataTable con dataset nulo/undefined

_For any_ input donde `dataSet` es `undefined` o `null` al invocar `initializeDataTableRenovaciones`, la función corregida SHALL inicializar la tabla con un arreglo vacío `[]` sin lanzar `TypeError`.

**Validates: Requirements 2.2**

Property 3: Bug Condition - JSON corrupto en columna de gestión

_For any_ input donde el contenido JSON de la columna de gestión es malformado, la función corregida SHALL capturar el error, registrar un log de advertencia, usar un objeto vacío `{}` como base para el merge, y continuar la operación de guardado.

**Validates: Requirements 2.3**

Property 4: Bug Condition - Índice de etapa fuera de rango

_For any_ input donde `currentEtapaIndex` excede los límites del arreglo de etapas, la función corregida SHALL limitar (clamp) el índice al rango `[0, etapas.length - 1]` antes de acceder al arreglo.

**Validates: Requirements 2.4**

Property 5: Bug Condition - Limpieza de intervalos al cerrar modal

_For any_ evento de cierre de modal donde existen intervalos activos y el OTP no fue validado, la función corregida SHALL ejecutar `clearInterval` sobre todos los intervalos activos, dejando el conteo en 0.

**Validates: Requirements 2.5**

Property 6: Bug Condition - Confirmación antes de navegar en Emisión de Póliza

_For any_ intento de navegación entre leads cuando la etapa actual es "Emisión de Póliza" y hay cambios sin guardar, la función corregida SHALL mostrar un diálogo de confirmación antes de proceder.

**Validates: Requirements 2.6**

Property 7: Bug Condition - Reporte de fallo parcial de archivos

_For any_ operación de guardado donde la carga de archivos a Drive falla parcialmente pero la gestión se guarda exitosamente, la función corregida SHALL retornar `{status: "partial"}` con detalle de archivos fallidos.

**Validates: Requirements 2.7**

Property 8: Bug Condition - Desempate determinístico en AssignLead

_For any_ conjunto de asesores donde dos o más tienen el mismo `sortingKey` y la misma `effectiveness`, la función corregida SHALL seleccionar al asesor cuyo email sea alfabéticamente menor, garantizando resultado determinístico.

**Validates: Requirements 2.8**

Property 9: Preservation - Comportamiento sin condición de bug

_For any_ input donde NINGUNA de las 8 condiciones de bug se activa (isBugCondition retorna false), las funciones corregidas SHALL producir exactamente el mismo resultado que las funciones originales, preservando todo el comportamiento existente.

**Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.7, 3.8**

## Fix Implementation

### Changes Required

Asumiendo que el análisis de causa raíz es correcto:

---

**Bug 1 — File**: `Renovaciones.js`
**Functions**: `RenovaSendWppVencida`, `RenovaSendWppProxVen`

**Specific Changes**:
1. **Reemplazar `new Headers()` y `fetch()`** por `UrlFetchApp.fetch(url, options)`
2. **Adaptar estructura de opciones**: Cambiar `{method, headers, body}` por `{method, headers: {}, payload, muteHttpExceptions: true}`
3. **Parametrizar datos**: Recibir nombre, teléfono y dirección como parámetros en lugar de valores hardcodeados
4. **Mover credenciales**: Usar `PropertiesService.getScriptProperties()` para la API key de Infobip

---

**Bug 2 — File**: `main.js.html`
**Function**: `initializeDataTableRenovaciones`

**Specific Changes**:
1. **Validación de entrada**: Agregar `dataSet = Array.isArray(dataSet) ? dataSet : []` al inicio de la función
2. **Validación en el caller**: En los callbacks que invocan `initializeDataTableRenovaciones(response.dataFront)`, usar `response.dataFront || []`

---

**Bug 3 — File**: `Renovaciones.js`
**Function**: `guardarGestionRenovacion`

**Specific Changes**:
1. **Fortalecer try/catch**: En la lectura de `rawColumnaGestion`, envolver el `JSON.parse` en un try/catch que registre el contenido corrupto con `Logger.log`
2. **Continuar con objeto vacío**: Si el parseo falla, usar `{}` como `currentGestionJson` y agregar flag `datosAnterioresCorruptos: true` al resultado
3. **Notificar al usuario**: Incluir en la respuesta un campo `warning` cuando se detecta corrupción

---

**Bug 4 — File**: `main.js.html`
**Function**: `renderEtapas` y lógica de cálculo de `currentEtapaIndex`

**Specific Changes**:
1. **Clamp del índice**: Agregar `currentEtapaIndex = Math.max(0, Math.min(currentEtapaIndex, etapasArray.length - 1))` antes de acceder al arreglo
2. **Validación en `indexOf`**: Cuando `indexOf` retorna -1, usar 0 como valor por defecto

---

**Bug 5 — File**: `main.js.html`
**Functions**: Lógica de OTP y evento de cierre del modal

**Specific Changes**:
1. **Registro centralizado de intervalos**: Crear array `activeIntervalIds = []` y registrar cada `setInterval` en él
2. **Limpieza en cierre de modal**: En el handler `hidden.bs.modal`, iterar `activeIntervalIds` ejecutando `clearInterval` y vaciar el array
3. **Limpieza en completar OTP**: También limpiar intervalos cuando OTP se valida exitosamente

---

**Bug 6 — File**: `main.js.html`
**Functions**: Handlers de navegación entre leads (botones siguiente/anterior)

**Specific Changes**:
1. **Detectar cambios sin guardar**: Implementar función `hasUnsavedChangesInEmision()` que compare el estado actual del formulario con los datos guardados
2. **Agregar confirmación**: Antes de `loadLeadByIndex()`, verificar si `currentEtapaIndex` corresponde a "Emisión de Póliza" y hay cambios sin guardar, mostrando `showCustomConfirm`
3. **Consistencia**: Usar el mismo patrón de confirmación que ya existe en otras etapas

---

**Bug 7 — File**: `Renovaciones.js`
**Function**: `guardarGestionRenovacion`

**Specific Changes**:
1. **Tracking de archivos fallidos**: Crear array `archivosFallidos = []` para registrar nombres de archivos que fallan después de reintentos
2. **Retorno diferenciado**: Si `archivosFallidos.length > 0` pero la gestión se guardó, retornar `{status: "partial", message: "...", archivosFallidos: [...]}`
3. **Frontend**: En `main.js.html`, manejar `status: "partial"` mostrando alerta con detalle de archivos fallidos

---

**Bug 8 — File**: `Código.js`
**Function**: `AssignLead`

**Specific Changes**:
1. **Agregar desempate terciario**: Modificar la condición de selección para incluir comparación alfabética de email cuando `sortingKey` y `effectiveness` son iguales
2. **Condición actualizada**:
   ```javascript
   if (sortingKey < bestSortingKey 
       || (sortingKey === bestSortingKey && effectiveness > highestEffectiveness)
       || (sortingKey === bestSortingKey && effectiveness === highestEffectiveness && email < bestAgent.email)) {
   ```

## Testing Strategy

### Validation Approach

La estrategia de testing sigue un enfoque de dos fases: primero, generar contraejemplos que demuestren los bugs en el código sin corregir, luego verificar que la corrección funciona y preserva el comportamiento existente.

### Exploratory Bug Condition Checking

**Goal**: Generar contraejemplos que demuestren cada bug ANTES de implementar la corrección. Confirmar o refutar el análisis de causa raíz.

**Test Plan**: Escribir tests unitarios que simulen cada condición de bug y ejecutarlos sobre el código sin corregir para observar los fallos.

**Test Cases**:
1. **Bug 1 - fetch en GAS**: Invocar `RenovaSendWppVencida()` en un mock de entorno GAS sin `fetch` global (fallará con ReferenceError)
2. **Bug 2 - DataTable undefined**: Invocar `initializeDataTableRenovaciones(undefined)` (fallará con TypeError)
3. **Bug 3 - JSON corrupto**: Invocar `mergeJsonColumnaGestion("{corrupto", {campo: "valor"})` (verificar que no lanza excepción no capturada)
4. **Bug 4 - Índice fuera de rango**: Setear `currentEtapaIndex = 10` con arreglo de 3 elementos y llamar `renderEtapas()` (fallará con undefined access)
5. **Bug 5 - Intervalos**: Simular apertura de modal, inicio de OTP, cierre sin completar (verificar que intervalos persisten)
6. **Bug 6 - Navegación sin guardar**: Simular cambios en formulario de Emisión y trigger de navegación (verificar que no hay confirmación)
7. **Bug 7 - Fallo parcial**: Simular `retryDrive` que falla para 1 de 3 archivos (verificar que respuesta es "ok" sin warning)
8. **Bug 8 - Desempate**: Crear 2 asesores con mismo sortingKey y effectiveness, ejecutar `AssignLead` múltiples veces (verificar inconsistencia)

**Expected Counterexamples**:
- Bug 1: `ReferenceError: fetch is not defined`
- Bug 2: `TypeError: Cannot read properties of undefined (reading 'length')` o similar
- Bug 3: Excepción no capturada en el flujo de guardado
- Bug 4: `undefined` al acceder a `etapaNombres[10]`
- Bug 8: Resultado diferente dependiendo del orden de filas en la hoja

### Fix Checking

**Goal**: Verificar que para todas las entradas donde la condición de bug se cumple, la función corregida produce el comportamiento esperado.

**Pseudocode:**
```
FOR ALL input WHERE isBugCondition(input) DO
  result := fixedFunction(input)
  ASSERT expectedBehavior(result)
END FOR
```

### Preservation Checking

**Goal**: Verificar que para todas las entradas donde la condición de bug NO se cumple, la función corregida produce el mismo resultado que la función original.

**Pseudocode:**
```
FOR ALL input WHERE NOT isBugCondition(input) DO
  ASSERT originalFunction(input) = fixedFunction(input)
END FOR
```

**Testing Approach**: Property-based testing es recomendado para preservation checking porque:
- Genera muchos casos de prueba automáticamente sobre el dominio de entrada
- Detecta edge cases que tests manuales podrían omitir
- Provee garantías fuertes de que el comportamiento no cambió para entradas no-buggy

**Test Plan**: Observar comportamiento en código sin corregir para entradas normales, luego escribir property-based tests capturando ese comportamiento.

**Test Cases**:
1. **Preservation WhatsApp**: Verificar que `sendWhatsAppSarlaft` sigue funcionando sin cambios
2. **Preservation DataTable**: Verificar que arreglos válidos siguen inicializando la tabla correctamente
3. **Preservation JSON merge**: Verificar que JSON válido sigue haciendo merge no destructivo
4. **Preservation renderEtapas**: Verificar que índices válidos siguen renderizando correctamente
5. **Preservation AssignLead**: Verificar que asesores con diferente sortingKey siguen siendo seleccionados por menor sortingKey

### Unit Tests

- Test de `RenovaSendWppVencida` corregida: verifica que usa `UrlFetchApp.fetch` con parámetros correctos
- Test de `initializeDataTableRenovaciones` con `undefined`, `null`, `[]`, y arreglo válido
- Test de `mergeJsonColumnaGestion` con JSON corrupto, vacío, y válido
- Test de clamp de `currentEtapaIndex` con valores -1, 0, length-1, length, length+10
- Test de limpieza de intervalos al cerrar modal
- Test de confirmación de navegación en "Emisión de Póliza"
- Test de respuesta parcial cuando archivos fallan
- Test de desempate alfabético en `AssignLead` con asesores empatados

### Property-Based Tests

- Generar strings aleatorios y verificar que `mergeJsonColumnaGestion` nunca lanza excepción no capturada (siempre retorna objeto)
- Generar índices aleatorios (negativos, cero, positivos grandes) y verificar que el clamp siempre produce un índice válido
- Generar conjuntos aleatorios de asesores con sortingKeys y effectiveness iguales, verificar que `AssignLead` siempre retorna el mismo resultado para la misma entrada (determinismo)
- Generar arreglos de archivos con fallos aleatorios y verificar que el status retornado refleja correctamente el resultado parcial

### Integration Tests

- Test end-to-end de envío WhatsApp: mock de `UrlFetchApp.fetch` verificando payload correcto a Infobip
- Test de flujo completo de gestión de renovación con JSON corrupto en celda: verificar que el guardado completa y el usuario recibe warning
- Test de navegación entre leads en "Emisión de Póliza" con formulario modificado: verificar diálogo y comportamiento según respuesta del usuario
- Test de carga de archivos con fallo parcial de Drive: verificar respuesta al frontend y mensaje al usuario
