# Auditoria UI/UX y API - Proceso de Renovaciones

## Objetivo
Este documento presenta una revision completa del flujo de renovaciones desde inicio a fin, enfocada en:
- Experiencia de usuario del asesor (UI/UX).
- Riesgos tecnicos funcionales en frontend y backend.
- Sobrecarga y uso ineficiente de consultas API.
- Funciones que deben evolucionar a manejo asincrono controlado.

Archivos auditados:
- `main.js.html` (flujo UI del asesor, validaciones, llamadas al backend).
- `Renovaciones.js` (logica de renovaciones, SARLAFT, OTP, guardado, integraciones externas).
- `Codigo.js` (servicios backend consumidos por `google.script.run`).

## Resumen ejecutivo
- Hay **errores funcionales criticos** que pueden romper el flujo del asesor (funciones invocadas no definidas y codigo de cliente dentro de backend Apps Script).
- Existen **inconsistencias UX** que generan friccion operativa (modales mezclados con `alert/confirm`, mensajes duplicados, simulaciones de exito sin backend real).
- Se detecta **sobrecosto de operaciones** por recargas globales repetidas y escrituras fila a fila.
- El flujo OTP tiene **riesgo alto de seguridad y trazabilidad** al validarse en frontend sin verificacion servidor.
- Hay funciones que deben adoptar patron asincrono para evitar bloqueos, estados visuales inconsistentes y dobles envios.

---

## Hallazgos quirurgicos por severidad

## Criticos (impacto alto, riesgo inmediato)

1) Funcion no definida en flujo de carga de archivos
- Funcion implicada: `subirArchivoADrive` (invocada desde `handlePropietarioDocUpload`, `handleOtroSiUpload`, `handleCesionDocAdicionalUpload` en `main.js.html`).
- Problema: se llama en tres puntos pero no existe implementacion encontrada en el proyecto.
- Impacto UX: el asesor cree que cargo documentos, pero la promesa falla y bloquea continuidad de procesos especiales.
- Riesgo tecnico: perdida de evidencia documental para expedicion.
- Accion: implementar `subirArchivoADrive` o reemplazar por lectura local + envio en `guardarGestionUnificada`.

2) Codigo de frontend en archivo backend de Apps Script
- Funciones implicadas en `Renovaciones.js`: `sendWhatsAppSarlaft`, `RenovaSendWppVencida`, `RenovaSendWppProxVen`, `sendMailInmbiliariaBrocker`.
- Problema: usan `document`, `fetch`, `Headers` (API de navegador), no validas en entorno Apps Script server.
- Impacto UX: botones/acciones asociadas pueden fallar silenciosamente o no ejecutarse.
- Riesgo tecnico: falsa sensacion de envio de comunicaciones.
- Accion: migrar estas funciones a frontend o reescribir en backend con `UrlFetchApp.fetch`.

3) Validacion OTP vulnerable (solo frontend)
- Funciones implicadas: `enviarOTP`, `validarOTP`, `reenviarOTP` en `main.js.html`.
- Problema: el OTP se genera y valida en cliente (`codigoOTPGenerado`), sin validacion en servidor.
- Impacto UX: estado "confirmado" puede marcarse localmente sin evidencia robusta.
- Riesgo tecnico/compliance: suplantacion de validacion del cliente.
- Accion: generar OTP en backend, almacenar con expiracion, validar OTP en backend con intento maximo y auditoria.

4) Credenciales expuestas en codigo fuente
- Funciones implicadas: `enviarOtpWhatsapp`, `requestSarlaft`, `sendWhatsappVPInteractive` y otras con `Authorization Basic` y/o llaves.
- Problema: secretos incrustados en codigo.
- Impacto UX indirecto: bloqueos de servicio por rotacion forzada de credenciales.
- Riesgo tecnico: fuga de seguridad y abuso de API.
- Accion: mover secretos a `PropertiesService` y rotar credenciales.

## Altos (impacto operativo continuo)

5) Sobreconsulta/redundancia de carga principal de datos
- Funciones implicadas: llamada inicial `getDatauserPro` + `reloadDataTable` (que vuelve a llamar `getDatauserPro`).
- Problema: doble carga de dataset completo en inicializacion y en multiples puntos post-guardado.
- Impacto UX: pantallas con loader frecuente, latencia visible, sensacion de lentitud.
- Riesgo tecnico: consumo innecesario de cuota Apps Script/API.
- Accion: cache local incremental y recarga parcial por lead, no recargar tabla completa en cada accion.

6) Escrituras celda por celda en procesos masivos
- Funciones implicadas:
  - `UpdateSarlaft` (setValue por fila).
  - `UpdateInqulinoSAI` (setValue por fila dentro de lotes).
  - `guardarGestionRenovacion` (multiples `setValue` por cada gestion).
- Problema: I/O fragmentado a hoja, alto costo por llamada.
- Impacto UX: guardados lentos y bloqueo percibido.
- Riesgo tecnico: timeout por volumen.
- Accion: consolidar y usar `setValues` por bloque.

7) Filtro logico incorrecto de recuperacion
- Funcion implicada: `getDatauserPro` en `Codigo.js`.
- Problema: condicion `(row[4] !== "caso especial" || row[4] !== "enviar a expedicion")` siempre evalua verdadero.
- Impacto UX: bandeja de recuperacion contaminada con casos que no deberian estar.
- Riesgo tecnico: mala priorizacion y gestion de leads.
- Accion: cambiar `||` por `&&`.

## Medios (friccion UX y deuda tecnica)

8) Inconsistencia de componentes de alerta
- Funciones implicadas: multiples en `main.js.html` (`alert`, `confirm`, `Swal.fire` mezclados).
- Problema: experiencia visual no uniforme, sin contexto de accion en algunos mensajes.
- Impacto UX: interrupciones bruscas, baja confianza.
- Accion: estandarizar en `Swal` + patron comun de confirmacion/error/success.

9) Doble listener en el mismo campo
- Funcion implicada: listeners de `#frecuenciaDeporte` en `main.js.html`.
- Problema: dos eventos `change` activos, puede duplicar mensajes y acciones.
- Impacto UX: alertas repetidas y confusion.
- Accion: consolidar en un solo listener con flujo claro.

10) Error tipografico que rompe flujo visual
- Funcion implicada: `forzarRenovacionSarlaft`.
- Problema: selector `##sarlaft-pregunta-inicial` (doble `#`).
- Impacto UX: seccion no reaparece, asesor queda bloqueado o desorientado.
- Accion: corregir a `#sarlaft-pregunta-inicial`.

11) Simulaciones no productivas en pasos criticos
- Funcion implicada: `procesarEnvioLinkSarlaft` (usa `setTimeout` de exito simulado).
- Problema: muestra confirmacion sin ejecutar backend real.
- Impacto UX: estado falso de gestion completada.
- Accion: integrar `google.script.run` real + resultado verificable.

12) Posible null no controlado en asignacion de lead
- Funcion implicada: `getNewLeadAssignment` -> `AssignLead("Renovations")`.
- Problema: si no hay asesor disponible, `asignacion` puede ser `null` y romper `asignacion.email`.
- Impacto UX: fallo en asignacion automatica.
- Accion: manejo defensivo con fallback de cola/no disponible.

---

## Analisis de exceso de consultas API / I-O

## Donde hay sobrecarga
- `reloadDataTable` recarga todo el universo de datos tras varias acciones que podrian actualizar solo 1 lead.
- `UpdateSarlaft` consulta API por cada registro de forma secuencial.
- `guardarGestionRenovacion` hace busquedas completas (`createTextFinder`, lectura total columna B) por cada guardado.
- Esquema actual privilegia consistencia inmediata total sobre eficiencia incremental.

## Recomendacion de arquitectura de llamadas
- Implementar un **servicio de datos incremental**:
  - Endpoint para actualizar solo el lead modificado.
  - Endpoint de recarga full solo bajo accion explicita (`Refrescar`).
- Aplicar **batch write** en Sheets:
  - Acumular cambios en arrays y persistir por bloque.
- Aplicar **cache corta (30-120s)** para catalogo/estado general del asesor.
- Agregar **idempotencia** para evitar doble click y doble guardado.

---

## Funciones que deben ser asincronas (o reforzar asincronia)

## Frontend (deben ser asincronas con `async/await`)

1) `enviarOTP(metodo)`
- Hoy: callback style con `google.script.run`.
- Debe: `await` a wrapper Promise para control de estados, timeout y reintentos.
- Beneficio: evita dobles clics y estados visuales inconsistentes.

2) `reenviarOTP()`
- Hoy: callback con logica repetida.
- Debe: reutilizar servicio asincrono comun `sendOtpToPhone`.
- Beneficio: menos duplicacion y mejor trazabilidad de errores.

3) `validarOTP()`
- Hoy: validacion local con `setTimeout`.
- Debe: `await validateOtpServer(...)`.
- Beneficio: seguridad + evidencia real.

4) `guardarGestionUnificada(tipoAccion)`
- Ya es `async`, pero debe:
  - centralizar conversion de archivos con `Promise.all`.
  - bloquear UI hasta respuesta definitiva.
  - evitar `reloadDataTable` full cuando no se requiere.

5) `procesarEnvioLinkSarlaft(metodo)`
- Hoy: simulacion con `setTimeout`.
- Debe: `await` a backend real y confirmacion condicionada al status.

## Backend Apps Script (asincronia conceptual por lotes)
Apps Script no trabaja con `async/await` tradicional del navegador en este caso, pero si requiere:
- Paralelismo controlado con `UrlFetchApp.fetchAll` (ya aplicado parcialmente en `UpdateInqulinoSAI`).
- Batch de escritura (`setValues`) en:
  - `UpdateSarlaft`
  - `UpdateInqulinoSAI`
  - `guardarGestionRenovacion` (reducir writes fragmentadas)

---

## Plan de remediacion recomendado (quirurgico)

## Fase 1 - Estabilizacion critica (1-2 dias)
1. Implementar o corregir `subirArchivoADrive`.
2. Corregir `forzarRenovacionSarlaft` (`##` -> `#`).
3. Corregir filtro de recuperacion en `getDatauserPro` (`||` -> `&&`).
4. Eliminar/migrar funciones frontend incrustadas en `Renovaciones.js` backend.
5. Mover credenciales a `PropertiesService`.

## Fase 2 - Seguridad y trazabilidad (2-3 dias)
1. OTP server-side:
   - generar OTP en backend,
   - TTL (5 min),
   - max intentos,
   - hash del OTP,
   - log de validacion.
2. Reescribir `validarOTP` para consumir validacion backend.
3. Registrar eventos de negocio (envio OTP, validacion, guardado final).

## Fase 3 - Rendimiento y UX (3-5 dias)
1. Evitar recargas globales:
   - actualizar solo fila/lead afectado.
2. Migrar escrituras por lote (`setValues`).
3. Unificar alertas en `Swal`.
4. Consolidar listeners duplicados y remover `setTimeout` simulados.
5. Agregar estados visuales estandar: loading, success, recoverable error.

## Fase 4 - Calidad y prevencion (continuo)
1. Checklist de QA por flujo:
   - normal,
   - proceso especial,
   - desistimiento,
   - OTP fallido y reintento.
2. Pruebas de regresion para guardado y estado de embudo.
3. Monitoreo de tiempos:
   - tiempo de carga tabla,
   - tiempo de guardado gestion,
   - tasa de error API.

---

## Riesgos de no corregir
- Aumento de tiempos de gestion por asesor.
- Errores de trazabilidad legal/comercial (OTP y evidencias).
- Costos de cuota y latencia por sobreconsultas.
- Falsos positivos de "gestion exitosa" sin persistencia real.
- Incidentes de seguridad por secretos en codigo.

---

## Resultado esperado tras aplicar mejoras
- Flujo estable y consistente para el asesor.
- Reduccion de latencia percibida en guardado/recarga.
- Menor consumo de API y de operaciones de hoja.
- Mayor seguridad y auditabilidad del proceso de renovacion.

---

## Validacion de hallazgos (Auditoria de la Auditoria)

### Estado de cada hallazgo documentado

| # | Hallazgo | Validado | Notas |
|---|----------|----------|-------|
| 1 | `subirArchivoADrive` no definida | ✅ CONFIRMADO | Se invoca en 3 puntos de `main.js.html` (lineas ~5693, ~6221, ~6253) pero no existe implementacion en ningun archivo del proyecto. |
| 2 | Codigo frontend en backend `Renovaciones.js` | ✅ CONFIRMADO | `RenovaSendWppVencida`, `RenovaSendWppProxVen`, `sendMailInmbiliariaBrocker` usan `new Headers()`, `fetch()` (API navegador). `sendWhatsAppSarlaft` usa `document.getElementById`. Todas invalidas en Apps Script server. |
| 3 | OTP validado solo en frontend | ✅ CONFIRMADO | `codigoOTPGenerado` se genera en cliente (linea ~6620), se valida con `setTimeout` local (linea ~6685-6686). No hay llamada a backend para validar. |
| 4 | Credenciales expuestas | ✅ CONFIRMADO | `Authorization: Basic THVpc2FfU2FudG9zX01rdDpCb2xpMjAyMnZhci4=` aparece en al menos 5 funciones de `Renovaciones.js` y 1 en `Codigo.js`. |
| 5 | Sobreconsulta `getDatauserPro` + `reloadDataTable` | ✅ CONFIRMADO | `reloadDataTable` llama a `getDatauserPro` y se invoca en al menos 6 puntos post-guardado (lineas ~3694, ~3788, ~3808, ~3911, ~5356, ~7186). |
| 6 | Escrituras celda por celda | ✅ CONFIRMADO | `UpdateSarlaft` usa `setValue` por fila (linea ~640). `UpdateInqulinoSAI` usa `setValue` por fila dentro de lotes (linea ~296). `guardarGestionRenovacion` hace multiples `setValue` individuales. |
| 7 | Filtro logico incorrecto `\|\|` vs `&&` | ✅ CONFIRMADO | Linea ~1123 de `Codigo.js`: `(row[4] !== "caso especial" \|\| row[4] !== "enviar a expedicion")` siempre es `true`. Debe ser `&&`. |
| 8 | Inconsistencia de alertas | ✅ CONFIRMADO | Se encontraron 6 usos de `alert()` nativo y 2 usos de `confirm()` nativo en `main.js.html`, mezclados con `Swal.fire`. |
| 9 | Doble listener `#frecuenciaDeporte` | ✅ CONFIRMADO | Hay 3 listeners `change` registrados: uno con `.off().on()` (linea ~3279), y dos adicionales sin `.off()` previo (lineas ~4620 y ~4626). Los dos ultimos son duplicados exactos. |
| 10 | Doble `#` en selector | ✅ CONFIRMADO | Linea ~5464: `$('##sarlaft-pregunta-inicial')` — selector invalido. |
| 11 | `procesarEnvioLinkSarlaft` simulada | ✅ CONFIRMADO | Usa `setTimeout` de 2 segundos con exito simulado (linea ~5717). El `google.script.run` real esta comentado (linea ~5733). |
| 12 | Null no controlado en `getNewLeadAssignment` | ✅ CONFIRMADO | Si `AssignLead("Renovations")` retorna `null`, `asignacion.email` lanza error. No hay manejo defensivo. |

### Hallazgo del README con nombre incorrecto

- El README menciona la funcion `sendWhatsAppSarlaft` como parte del hallazgo 2 (codigo frontend en backend). La funcion SI existe en `Renovaciones.js` (linea ~810) y SI usa `document.getElementById` (API de navegador), por lo que el hallazgo es correcto. Sin embargo, el README la lista como `sendWhatsAppSarlaft` pero no la incluye explicitamente en la lista de funciones del hallazgo 2. Se confirma que tambien debe ser migrada.

---

## Hallazgos adicionales NO documentados en la auditoria original

### Criticos

13) Doble inclusion de `kaiadmin.min.js` en `index.html`
- Archivo: `index.html`, lineas 33-34.
- Problema: `<?!= include("kaiadmin.min.js");?>` aparece dos veces consecutivas.
- Impacto UX: carga duplicada de JS, posibles conflictos de inicializacion, listeners duplicados, mayor tiempo de carga.
- Accion: eliminar la linea duplicada.

14) Variable global implicita `userData` en `AssignLead` (`Codigo.js`)
- Funcion: `AssignLead` en `Codigo.js` (linea ~811).
- Problema: `userData = userDataleads.filter(...)` se asigna sin `let`/`const`/`var`, creando una variable global implicita.
- Impacto tecnico: contaminacion del scope global, posibles colisiones si otra funcion usa el mismo nombre, comportamiento impredecible en ejecuciones concurrentes.
- Accion: declarar con `let` o `const`.

15) `console.log` del codigo OTP en produccion
- Archivo: `main.js.html`, linea ~6621.
- Problema: `console.log('Código OTP generado:', codigoOTPGenerado)` imprime el OTP en la consola del navegador en produccion.
- Impacto seguridad: cualquier persona con acceso a DevTools puede ver el OTP sin necesidad de recibirlo por WhatsApp, anulando completamente la seguridad del flujo OTP.
- Accion: eliminar el `console.log` o condicionarlo a un flag de debug.

### Altos

16) IDs de Spreadsheets hardcodeados en codigo fuente
- Archivos: `Codigo.js` (lineas ~16-20) y `Renovaciones.js` (lineas ~2-14).
- Problema: multiples IDs de Google Sheets estan directamente en el codigo (`1uQO1x5npOpVBwV8IJw3s-J2hrQ38RdY4QU6BDB6xfY4`, `1wxqoUCggSYXE0vOUHdgLnwDYfBnlATeyfTZ8CgQQEY4`, etc.) y un ID de carpeta de Drive (`1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC`).
- Impacto: dificulta migracion entre ambientes (dev/staging/prod), riesgo de modificar datos de produccion en pruebas.
- Accion: mover todos los IDs a `PropertiesService` (algunos ya lo usan, como `idPolizas`, pero la mayoria no).

17) URLs de API Infobip hardcodeadas
- Archivos: `Renovaciones.js` y `Codigo.js`.
- Problema: la URL base `https://qgmx9r.api.infobip.com` esta repetida en al menos 6 puntos del codigo.
- Impacto: si cambia el endpoint, hay que modificar multiples archivos. Riesgo de inconsistencia.
- Accion: centralizar en una constante o en `PropertiesService`.

### Medios

18) Llamadas `google.script.run` sin `withFailureHandler` en `main.js.html`
- Funcion: `google.script.run.withSuccessHandler(...).getDatauserPro()` en linea ~4659 (flujo de autenticacion).
- Problema: no tiene `withFailureHandler`. Si el backend falla, el usuario queda en un estado de carga infinita sin feedback.
- Impacto UX: spinner eterno, asesor bloqueado sin saber que paso.
- Accion: agregar `withFailureHandler` con mensaje de error y ocultar spinner en todas las llamadas `google.script.run`.

19) Validacion OTP con `setTimeout` artificial de 1.5s
- Funcion: `validarOTP()` en `main.js.html` (linea ~6685).
- Problema: la validacion (que ya es local) se envuelve en un `setTimeout` de 1500ms que no aporta nada funcional, solo simula una espera.
- Impacto UX: latencia artificial innecesaria en cada intento de validacion.
- Accion: eliminar el `setTimeout` (o mejor aun, migrar la validacion al backend como indica el hallazgo 3).

20) Funcion `sendWhatsAppSarlaft` en `Renovaciones.js` usa API de navegador
- Funcion: `sendWhatsAppSarlaft` (linea ~810 de `Renovaciones.js`).
- Problema: usa `document.getElementById` para leer datos del DOM, lo cual es imposible en el entorno server-side de Apps Script.
- Impacto: la funcion nunca puede ejecutarse correctamente desde el backend.
- Nota: esta funcion no fue listada explicitamente en el hallazgo 2 del README original, aunque pertenece al mismo patron. Se agrega como hallazgo independiente para visibilidad.

21) Numero de telefono de prueba hardcodeado en funciones de envio
- Funciones: `RenovaSendWppVencida` y `RenovaSendWppProxVen` en `Renovaciones.js`.
- Problema: ambas funciones tienen el numero `573222340943` y datos de prueba (`Nikol Rodriguez`, `3123123`) hardcodeados. No reciben parametros dinamicos.
- Impacto: estas funciones no son funcionales en produccion; solo envian al mismo numero de prueba.
- Accion: parametrizar destinatario y datos del template, o eliminar si son solo prototipos.
