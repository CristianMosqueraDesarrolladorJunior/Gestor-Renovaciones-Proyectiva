# Implementation Plan

- [x] 1. Write bug condition exploration test
  - **Property 1: Bug Condition** - 8 Bugs Críticos en Gestor Renovaciones
  - **CRITICAL**: This test MUST FAIL on unfixed code - failure confirms the bugs exist
  - **DO NOT attempt to fix the test or the code when it fails**
  - **NOTE**: This test encodes the expected behavior - it will validate the fix when it passes after implementation
  - **GOAL**: Surface counterexamples that demonstrate each bug exists
  - **Scoped PBT Approach**: Scope each property to the concrete failing case(s) for reproducibility
  - Bug 1: Invoke `RenovaSendWppVencida()` in a mock GAS environment without global `fetch`/`Headers` — expect `ReferenceError: fetch is not defined`
  - Bug 2: Invoke `initializeDataTableRenovaciones(undefined)` — expect `TypeError` when iterating undefined dataset
  - Bug 3: Invoke `mergeJsonColumnaGestion("{corrupto_sin_cerrar", {campo: "valor"})` — verify no unhandled exception propagates from `guardarGestionRenovacion`
  - Bug 4: Set `currentEtapaIndex = 10` with `etapaNombresRenovaciones` (length 3) — expect `undefined` access on `etapaNombres[10]`
  - Bug 5: Simulate modal open → OTP `setInterval` start → modal close without OTP completion — expect intervals still running (leak)
  - Bug 6: Simulate unsaved changes in "Emisión de Póliza" stage → trigger `loadLeadByIndex(next)` — expect navigation proceeds without confirmation dialog
  - Bug 7: Simulate `retryDrive` failure for 1 of 3 files with successful gestion save — expect response `{status: "ok"}` without file failure info
  - Bug 8: Create 2+ agents with identical `sortingKey` and `effectiveness`, run `AssignLead` multiple times — expect non-deterministic results depending on array order
  - Run tests on UNFIXED code
  - **EXPECTED OUTCOME**: Tests FAIL (this is correct - it proves the bugs exist)
  - Document counterexamples found to understand root cause
  - Mark task complete when tests are written, run, and failure is documented
  - _Requirements: 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8_

- [x] 2. Write preservation property tests (BEFORE implementing fix)
  - **Property 2: Preservation** - Comportamiento Existente Sin Condición de Bug
  - **IMPORTANT**: Follow observation-first methodology
  - **Observe on UNFIXED code**:
  - Observe: `sendWhatsAppSarlaft` y `enviarOtpWhatsapp` que ya usan `UrlFetchApp.fetch()` envían mensajes exitosamente
  - Observe: `initializeDataTableRenovaciones([{poliza: "123", ...}])` con arreglo válido inicializa DataTable correctamente
  - Observe: `mergeJsonColumnaGestion('{"poliza":"5012345","canon":100}', {nuevosCampos: "valor"})` realiza merge no destructivo preservando campos y concatenando `documentosProceso`
  - Observe: `renderEtapas()` con `currentEtapaIndex = 0, 1, 2` (dentro de rango) renderiza etapa correspondiente sin error
  - Observe: Validación OTP exitosa + cierre de modal mantiene estado `OTP_VALIDADO`
  - Observe: Navegación entre leads en etapas distintas a "Emisión de Póliza" con cambios sin guardar muestra diálogo de confirmación
  - Observe: Carga exitosa de todos los archivos a Drive retorna `{status: "ok"}` con URLs
  - Observe: `AssignLead()` con asesores de diferente `sortingKey` selecciona al de menor `sortingKey`; con mismo `sortingKey` pero diferente `effectiveness` selecciona al de mayor `effectiveness`
  - **Write property-based tests**:
  - Property: For all valid arrays passed to `initializeDataTableRenovaciones`, table initializes without error showing all records
  - Property: For all valid JSON strings, `mergeJsonColumnaGestion` preserves existing fields and concatenates `documentosProceso`
  - Property: For all `currentEtapaIndex` in range `[0, etapas.length-1]`, `renderEtapas` accesses valid array position
  - Property: For all agent sets where one agent has strictly lower `sortingKey`, that agent is selected regardless of other factors
  - Property: For all agent sets where `sortingKey` is equal but one has strictly higher `effectiveness`, that agent is selected
  - Run tests on UNFIXED code
  - **EXPECTED OUTCOME**: Tests PASS (this confirms baseline behavior to preserve)
  - Mark task complete when tests are written, run, and passing on unfixed code
  - _Requirements: 3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.7, 3.8_

- [ ] 3. Fix Bug 1 — Reemplazar fetch()/Headers por UrlFetchApp en funciones WhatsApp

  - [-] 3.1 Implementar corrección en `Renovaciones.js`
    - Reemplazar `new Headers()` y `fetch()` por `UrlFetchApp.fetch(url, options)` en `RenovaSendWppVencida` y `RenovaSendWppProxVen`
    - Adaptar estructura de opciones: `{method: "POST", headers: {"Authorization": "...", "Content-Type": "..."}, payload: JSON.stringify(body), muteHttpExceptions: true}`
    - Mover credenciales de Infobip a `PropertiesService.getScriptProperties()`
    - Parametrizar datos de destinatario (nombre, teléfono, dirección) en lugar de valores hardcodeados
    - _Bug_Condition: isBug1(input) where input.runtime = "GoogleAppsScript" AND input.function IN {"RenovaSendWppVencida", "RenovaSendWppProxVen"} AND input.usesAPI IN {"fetch", "Headers"}_
    - _Expected_Behavior: La función usa UrlFetchApp.fetch() con opciones {method, headers, payload, muteHttpExceptions} sin lanzar ReferenceError_
    - _Preservation: sendWhatsAppSarlaft y enviarOtpWhatsapp que ya usan UrlFetchApp siguen funcionando sin cambios_
    - _Requirements: 2.1, 3.1_

  - [ ] 3.2 Verify bug condition exploration test now passes for Bug 1
    - **Property 1: Expected Behavior** - APIs de GAS para envío WhatsApp
    - **IMPORTANT**: Re-run the SAME test from task 1 (Bug 1 case) - do NOT write a new test
    - The test from task 1 encodes the expected behavior: `UrlFetchApp.fetch()` is used without `ReferenceError`
    - **EXPECTED OUTCOME**: Test PASSES (confirms bug is fixed)
    - _Requirements: 2.1_

- [ ] 4. Fix Bug 2 — Validar dataset undefined/null en initializeDataTableRenovaciones

  - [ ] 4.1 Implementar corrección en `main.js.html`
    - Agregar validación de entrada al inicio de `initializeDataTableRenovaciones`: `dataSet = Array.isArray(dataSet) ? dataSet : []`
    - En los callbacks que invocan la función, usar `response.dataFront || []` como fallback
    - _Bug_Condition: isBug2(input) where input.function = "initializeDataTableRenovaciones" AND (input.dataSet = undefined OR input.dataSet = null)_
    - _Expected_Behavior: La tabla se inicializa con arreglo vacío [] sin lanzar TypeError, mostrando estado vacío al usuario_
    - _Preservation: Arreglos válidos con datos siguen inicializando la tabla correctamente mostrando todos los registros_
    - _Requirements: 2.2, 3.2_

  - [ ] 4.2 Verify bug condition exploration test now passes for Bug 2
    - **Property 1: Expected Behavior** - DataTable con dataset nulo/undefined
    - **IMPORTANT**: Re-run the SAME test from task 1 (Bug 2 case) - do NOT write a new test
    - **EXPECTED OUTCOME**: Test PASSES (confirms bug is fixed)
    - _Requirements: 2.2_

- [ ] 5. Fix Bug 3 — Manejo robusto de JSON corrupto en columna de gestión

  - [ ] 5.1 Implementar corrección en `Renovaciones.js`
    - Fortalecer try/catch en la lectura de `rawColumnaGestion` dentro de `guardarGestionRenovacion`
    - Si `JSON.parse` falla, registrar contenido corrupto con `Logger.log` (warning)
    - Usar `{}` como `currentGestionJson` cuando el parseo falla
    - Agregar flag `datosAnterioresCorruptos: true` al resultado y campo `warning` en la respuesta al frontend
    - _Bug_Condition: isBug3(input) where input.function = "guardarGestionRenovacion" AND input.columnaGestionValue != "" AND NOT isValidJson(input.columnaGestionValue)_
    - _Expected_Behavior: Captura error, registra log de advertencia, usa {} como base para merge, continúa guardado notificando al usuario_
    - _Preservation: JSON válido sigue haciendo merge no destructivo preservando campos y concatenando documentosProceso_
    - _Requirements: 2.3, 3.3_

  - [ ] 5.2 Verify bug condition exploration test now passes for Bug 3
    - **Property 1: Expected Behavior** - JSON corrupto en columna de gestión
    - **IMPORTANT**: Re-run the SAME test from task 1 (Bug 3 case) - do NOT write a new test
    - **EXPECTED OUTCOME**: Test PASSES (confirms bug is fixed)
    - _Requirements: 2.3_

- [ ] 6. Fix Bug 4 — Clamp de currentEtapaIndex fuera de rango

  - [ ] 6.1 Implementar corrección en `main.js.html`
    - Agregar clamp del índice antes de acceder al arreglo: `currentEtapaIndex = Math.max(0, Math.min(currentEtapaIndex, etapasArray.length - 1))`
    - Cuando `indexOf` retorna -1 al buscar la etapa del lead, usar 0 como valor por defecto
    - Aplicar validación tanto para `etapaNombres` (6 etapas) como para `etapaNombresRenovaciones` (3 etapas)
    - _Bug_Condition: isBug4(input) where input.function = "renderEtapas" AND (input.currentEtapaIndex >= input.etapasArray.length OR input.currentEtapaIndex < 0)_
    - _Expected_Behavior: El índice se limita al rango [0, etapas.length - 1] antes de acceder al arreglo, previniendo accesos undefined_
    - _Preservation: Índices válidos dentro del rango siguen renderizando la etapa correspondiente sin alteraciones_
    - _Requirements: 2.4, 3.4_

  - [ ] 6.2 Verify bug condition exploration test now passes for Bug 4
    - **Property 1: Expected Behavior** - Índice de etapa fuera de rango
    - **IMPORTANT**: Re-run the SAME test from task 1 (Bug 4 case) - do NOT write a new test
    - **EXPECTED OUTCOME**: Test PASSES (confirms bug is fixed)
    - _Requirements: 2.4_

- [ ] 7. Fix Bug 5 — Limpieza de intervalos al cerrar modal OTP

  - [ ] 7.1 Implementar corrección en `main.js.html`
    - Crear array centralizado `activeIntervalIds = []` para registrar cada `setInterval` creado en el contexto del modal OTP
    - En cada `setInterval` de countdown y barras de progreso, registrar el ID en `activeIntervalIds`
    - En el handler `hidden.bs.modal` del modal de renovaciones, iterar `activeIntervalIds` ejecutando `clearInterval` y vaciar el array
    - También limpiar intervalos cuando OTP se valida exitosamente (flujo normal)
    - _Bug_Condition: isBug5(input) where input.event = "modalClose" AND input.activeIntervals.length > 0 AND input.otpValidated = false_
    - _Expected_Behavior: Todos los intervalos activos se limpian con clearInterval, dejando conteo en 0_
    - _Preservation: Validación OTP exitosa + cierre de modal mantiene estado OTP_VALIDADO correctamente_
    - _Requirements: 2.5, 3.5_

  - [ ] 7.2 Verify bug condition exploration test now passes for Bug 5
    - **Property 1: Expected Behavior** - Limpieza de intervalos al cerrar modal
    - **IMPORTANT**: Re-run the SAME test from task 1 (Bug 5 case) - do NOT write a new test
    - **EXPECTED OUTCOME**: Test PASSES (confirms bug is fixed)
    - _Requirements: 2.5_

- [ ] 8. Fix Bug 6 — Confirmación de navegación en etapa Emisión de Póliza

  - [ ] 8.1 Implementar corrección en `main.js.html`
    - Implementar función `hasUnsavedChangesInEmision()` que compare estado actual del formulario con datos guardados en `activeDatosEtapas`
    - En los handlers de navegación entre leads (botones siguiente/anterior), verificar si `currentEtapaIndex` corresponde a "Emisión de Póliza"
    - Si hay cambios sin guardar, mostrar `showCustomConfirm("Tiene cambios sin guardar. ¿Desea descartarlos?", callback)` antes de `loadLeadByIndex()`
    - Usar el mismo patrón de confirmación que ya existe en otras etapas
    - _Bug_Condition: isBug6(input) where input.event IN {"navigateNext", "navigatePrev"} AND input.currentEtapa = "Emisión de Póliza" AND input.hasUnsavedChanges = true_
    - _Expected_Behavior: Se muestra diálogo de confirmación; si usuario confirma → navega; si cancela → se bloquea navegación_
    - _Preservation: Navegación en etapas distintas a "Emisión de Póliza" con cambios sin guardar sigue mostrando diálogo existente_
    - _Requirements: 2.6, 3.6_

  - [ ] 8.2 Verify bug condition exploration test now passes for Bug 6
    - **Property 1: Expected Behavior** - Confirmación antes de navegar en Emisión de Póliza
    - **IMPORTANT**: Re-run the SAME test from task 1 (Bug 6 case) - do NOT write a new test
    - **EXPECTED OUTCOME**: Test PASSES (confirms bug is fixed)
    - _Requirements: 2.6_

- [ ] 9. Fix Bug 7 — Reporte de fallo parcial en carga de archivos

  - [ ] 9.1 Implementar corrección en `Renovaciones.js`
    - Crear array `archivosFallidos = []` en `guardarGestionRenovacion` para registrar nombres de archivos que fallan después de reintentos con `retryDrive`
    - Si `archivosFallidos.length > 0` pero la gestión se guardó exitosamente, retornar `{status: "partial", message: "Gestión guardada. Los siguientes archivos no se pudieron subir: ...", archivosFallidos: [...]}`
    - _Bug_Condition: isBug7(input) where input.function = "guardarGestionRenovacion" AND input.archivosBase64 != empty AND input.driveUploadPartialFailure = true AND input.gestionSaveSuccess = true_
    - _Expected_Behavior: Retorna {status: "partial"} con detalle de archivos fallidos, permitiendo al usuario reintentar_
    - _Preservation: Carga exitosa de todos los archivos sigue retornando {status: "ok"} con URLs_
    - _Requirements: 2.7, 3.7_

  - [ ] 9.2 Implementar manejo de `status: "partial"` en frontend (`main.js.html`)
    - En el callback de `google.script.run` que recibe la respuesta de `guardarGestionRenovacion`, manejar `status: "partial"`
    - Mostrar alerta con `showWarningAlertbasic()` indicando qué archivos fallaron
    - Mantener comportamiento actual para `status: "ok"` (éxito total) y `status: "error"` (fallo total)
    - _Requirements: 2.7_

  - [ ] 9.3 Verify bug condition exploration test now passes for Bug 7
    - **Property 1: Expected Behavior** - Reporte de fallo parcial de archivos
    - **IMPORTANT**: Re-run the SAME test from task 1 (Bug 7 case) - do NOT write a new test
    - **EXPECTED OUTCOME**: Test PASSES (confirms bug is fixed)
    - _Requirements: 2.7_

- [ ] 10. Fix Bug 8 — Desempate determinístico en AssignLead

  - [ ] 10.1 Implementar corrección en `Código.js`
    - Modificar la condición de selección en `AssignLead()` para incluir desempate terciario por orden alfabético de email
    - Condición actualizada: `if (sortingKey < bestSortingKey || (sortingKey === bestSortingKey && effectiveness > highestEffectiveness) || (sortingKey === bestSortingKey && effectiveness === highestEffectiveness && email < bestAgent.email))`
    - Asegurar que `bestAgent` se inicializa con un email que permita la primera comparación (ej: inicializar con email vacío o manejar caso null)
    - _Bug_Condition: isBug8(input) where EXISTS a1, a2 IN input.agents WHERE a1.sortingKey = a2.sortingKey AND a1.effectiveness = a2.effectiveness AND a1.email != a2.email_
    - _Expected_Behavior: Se selecciona al asesor cuyo email es alfabéticamente menor, garantizando resultado determinístico y auditable_
    - _Preservation: Asesores con diferente sortingKey siguen siendo seleccionados por menor sortingKey; con mismo sortingKey pero diferente effectiveness, se selecciona al de mayor effectiveness_
    - _Requirements: 2.8, 3.8_

  - [ ] 10.2 Verify bug condition exploration test now passes for Bug 8
    - **Property 1: Expected Behavior** - Desempate determinístico en AssignLead
    - **IMPORTANT**: Re-run the SAME test from task 1 (Bug 8 case) - do NOT write a new test
    - **EXPECTED OUTCOME**: Test PASSES (confirms bug is fixed)
    - _Requirements: 2.8_

- [ ] 11. Verify preservation tests still pass after all fixes
  - **Property 2: Preservation** - Comportamiento Existente Sin Condición de Bug
  - **IMPORTANT**: Re-run the SAME tests from task 2 - do NOT write new tests
  - Run all preservation property tests from step 2
  - **EXPECTED OUTCOME**: Tests PASS (confirms no regressions introduced)
  - Confirm all preservation tests still pass after all 8 fixes are applied
  - Verify: Valid arrays still initialize DataTable correctly
  - Verify: Valid JSON still merges non-destructively
  - Verify: Valid etapa indices still render correctly
  - Verify: AssignLead with clear winner still selects correctly
  - Verify: Existing WhatsApp functions (sendWhatsAppSarlaft, enviarOtpWhatsapp) still work

- [ ] 12. Checkpoint - Ensure all tests pass
  - Run full test suite to confirm all exploration tests (Property 1) now pass
  - Run full test suite to confirm all preservation tests (Property 2) still pass
  - Verify no regressions in existing functionality
  - Ensure all 8 bugs are resolved and verified
  - Ask the user if questions arise
