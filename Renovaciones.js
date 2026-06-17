
const Renovations = SpreadsheetApp.openById("1wxqoUCggSYXE0vOUHdgLnwDYfBnlATeyfTZ8CgQQEY4");
const DataRenovations = Renovations.getSheetByName("JSON");
const PolizasAntiguas = Renovations.getSheetByName("PolizasAntiguas")
const WareHouseRenovations = Renovations.getSheetByName("Renovations");
const GestionRenovaciones = Renovations.getSheetByName("JSON");
const LogsEnvios = Renovations.getSheetByName("LogEnvios");
const LogErrores = Renovations.getSheetByName("LogErrores");
const GestionAsesora = SpreadsheetApp.openById("1MryMxNuW1cCjgJ2xj5tm8ut-wcd7PNQHI3WEpJal0U4").getSheetByName("Propietarios")
const GestionBroker = SpreadsheetApp.openById("1MryMxNuW1cCjgJ2xj5tm8ut-wcd7PNQHI3WEpJal0U4").getSheetByName("Broker")
const GestionCorretaje = SpreadsheetApp.openById("1MryMxNuW1cCjgJ2xj5tm8ut-wcd7PNQHI3WEpJal0U4").getSheetByName("Corretaje")
const GestionAnalista = SpreadsheetApp.openById("1oAfMyBNgkKR97JbUNir7MjM3KFQ2c8YwQ2QrsFvtIrM").getSheetByName("Respuestas Renovación")

const Espejo = SpreadsheetApp.openById("1ACFjJriwgFE-VOUHifx2Rr7zUNk_Ovy0OnQUMHKnvKY").getSheetByName("new_data_polizas- archivo David")

/**
 * Ejecuta una función con reintentos y backoff exponencial para manejar errores intermitentes de Google Drive.
 * @param {Function} fn - Función a ejecutar.
 * @param {number} retries - Número máximo de reintentos.
 * @param {number} delay - Delay inicial en milisegundos.
 * @return {*} Resultado de la función ejecutada.
 */
function retryDrive(fn, retries = 5, delay = 300) {
  for (let i = 0; i < retries; i++) {
    try {
      return fn();
    } catch (e) {
      Logger.log(`[Drive] Intento ${i + 1} fallido. Error: ${e.message}. Reintentando en ${delay}ms...`);
      if (i === retries - 1) throw new Error("El servicio de Google Drive no respondió después de " + retries + " intentos. Por favor intente nuevamente en unos minutos. Detalle: " + e.message);
      Utilities.sleep(delay);
      delay *= 2;
    }
  }
}

const EXPEDIDORES = [
  "jose.castillo.rodriguez@segurosbolivar.com",
  "jeymmy.aristizabal@segurosbolivar.com"
];

// Mapeo nombre ejecutivo de cuenta (columna E de GestionAnalista) → correo
const MAPA_EJECUTIVOS_CUENTA = {
  "FABIAN SANCHEZ": "jeison.sanchez@proyectivaseguros.com",
  "YENY JAIMES":    "yeny.jaimes@proyectivaseguros.com"
};

function cargarListaAnalistas() {
  var lastRow = DataGestion.getLastRow();
  if (lastRow < 2) return EXPEDIDORES.slice();
  var datos = DataGestion.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
  var analistas = [];
  for (var i = 0; i < datos.length; i++) {
    var especialidad = (datos[i][3] || "").toString().trim();
    var correo = (datos[i][2] || "").toString().trim();
    if (especialidad === "Analista Renovaciones" && correo !== "") {
      analistas.push(correo);
    }
  }
  return analistas.length > 0 ? analistas : EXPEDIDORES.slice();
}

function obtenerSiguienteExpedidor(analistasPreCargados) {
  var props = PropertiesService.getScriptProperties();
  var turno = parseInt(props.getProperty('turnoExpedidor') || '0', 10);

  var analistas = (analistasPreCargados && analistasPreCargados.length > 0)
    ? analistasPreCargados
    : cargarListaAnalistas();

  var seleccionado = analistas[turno % analistas.length];
  props.setProperty('turnoExpedidor', String(turno + 1));
  return seleccionado;
}

/**
 * Realiza merge no destructivo del JSON de Columna_Gestion.
 * Preserva campos existentes no incluidos en la actualización,
 * y concatena el arreglo documentosProceso sin eliminar entradas previas.
 * @param {Object|string|null} currentJson - JSON actual de la celda (objeto, string JSON, null o vacío).
 * @param {Object} newData - Nuevos datos a mergear sobre el JSON existente.
 * @returns {Object} JSON resultante del merge.
 */
function mergeJsonColumnaGestion(currentJson, newData) {
  var existing = {};

  if (currentJson && typeof currentJson === 'object' && !Array.isArray(currentJson)) {
    existing = currentJson;
  } else if (typeof currentJson === 'string' && currentJson.trim() !== '') {
    try {
      var parsed = JSON.parse(currentJson.replace(/:\s*NaN\b/g, ': null'));
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
        existing = parsed;
      }
    } catch (e) {
      existing = {};
    }
  }

  var safeNewData = (newData && typeof newData === 'object' && !Array.isArray(newData))
    ? newData
    : {};

  var merged = {};
  var key;

  for (key in existing) {
    if (Object.prototype.hasOwnProperty.call(existing, key)) {
      merged[key] = existing[key];
    }
  }

  for (key in safeNewData) {
    if (Object.prototype.hasOwnProperty.call(safeNewData, key)) {
      merged[key] = safeNewData[key];
    }
  }

  var prevDocs = Array.isArray(existing.documentosProceso) ? existing.documentosProceso : [];
  var newDocs = Array.isArray(safeNewData.documentosProceso) ? safeNewData.documentosProceso : [];

  if (prevDocs.length > 0 || newDocs.length > 0) {
    merged.documentosProceso = prevDocs.concat(newDocs);
  }

  return merged;
}

function UpdateRenovations() {
  const sheet = DataRenovations;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  // Pre-loop único: construir polizasNoExpedidas en un solo recorrido sobre data
  const polizasNoExpedidas = new Set();
  for (let i = 0; i < data.length; i++) {
    const leadInfo = parseLeadJson(data[i][1]);
    if (leadInfo && leadInfo.poliza) {
      const estadoFila = (data[i][4] || "").toString().trim().toLowerCase();
      if (estadoFila !== "expedido") {
        polizasNoExpedidas.add(String(leadInfo.poliza).trim());
      }
    }
  }

  // Leer GestionAnalista una sola vez: se usa para dbAnalista y para el marcado batch al final
  const lastRowAnalista = GestionAnalista.getLastRow();
  const dataAnalistaRaw = lastRowAnalista >= 2
    ? GestionAnalista.getRange(2, 1, lastRowAnalista - 1, GestionAnalista.getLastColumn()).getValues()
    : [];

  // dbAnalistaZ: columna Z con número de póliza (inicia en "50") — sin filtro de marca.
  // Máxima prioridad: ignora estadosProtegidos y la marca existente en el analista.
  const dbAnalistaZ = new Map();
  // dbAnalista: resto de registros del analista, solo sin marca.
  const dbAnalista = new Map();
  for (const row of dataAnalistaRaw) {
    const key = String(row[10]).trim();
    if (!key) continue;
    const valorZ = (row[25] || "").toString().trim();
    if (/^50\d+/.test(valorZ)) {
      dbAnalistaZ.set(key, row);
    } else if ((row[29] || "").toString().trim() === "") {
      dbAnalista.set(key, row);
    }
  }

  const dbAsesora = getDatabaseMapByKeys(GestionAsesora, 2, polizasNoExpedidas);
  const dbBroker = getDatabaseMapByKeys(GestionBroker, 2, polizasNoExpedidas);
  const dbCorretaje = getDatabaseMapByKeys(GestionCorretaje, 2, polizasNoExpedidas);

  // Cachear lista de analistas una vez: evita I/O repetido por cada registro dbAnalista
  const listaAnalistas = cargarListaAnalistas();

  const today = new Date();
  const logs = [];
  const fechaEjecucion = Utilities.formatDate(today, "America/Bogota", "dd/MM/yyyy HH:mm:ss");

  Logger.log("=== INICIO UpdateRenovations ===");
  Logger.log("Total filas a procesar: " + data.length);
  Logger.log("Pólizas no expedidas (a consultar): " + polizasNoExpedidas.size);
  Logger.log("Registros en dbAnalistaZ (col Z expedido, sin filtro marca): " + dbAnalistaZ.size);
  Logger.log("Registros en dbAnalista (sin marca, sin Z-expedido): " + dbAnalista.size);
  Logger.log("Registros en dbAsesora (solo no expedidas): " + dbAsesora.size);
  Logger.log("Registros en dbBroker (solo no expedidas): " + dbBroker.size);
  Logger.log("Registros en dbCorretaje (solo no expedidas): " + dbCorretaje.size);

  let contadorActualizados = 0;
  let contadorProtegidos = 0;
  let contadorSinPoliza = 0;

  // Set fuera del loop: se construye una sola vez, .has() es O(1)
  const ESTADOS_PROTEGIDOS = new Set([
    "expedido",
    "enviar a expedicion",
    "poliza renovada",
    "correccion",
    "recuperado",
    "caso especial",
    "caso corregido",
    "cliente ya renovo",
    "desistido",
  ]);

  // Registra exactamente qué pólizas fueron tomadas de dbAnalista para el marcado preciso
  const polizasProcesadasAnalista = new Set();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const leadInfo = parseLeadJson(row[1]);

    if (!leadInfo || !leadInfo.poliza) {
      contadorSinPoliza++;
      continue;
    }

    const polizaKey = String(leadInfo.poliza).trim();
    const fechaVenc = parseDate(leadInfo.vencimiento);
    const segmento = (row[3] || "").toString().trim().toLowerCase();
    const esBroker = segmento.includes("broker") || segmento.includes("inmobiliaria");
    const estado = row[4].toString().trim().toLowerCase();
    const estadoAnterior = row[4].toString().trim();
    const asesorAnterior = row[2].toString().trim();

    // PRIORIDAD MÁXIMA: columna Z con número de póliza expedida.
    // Se aplica antes que estadosProtegidos y sin importar si ya tiene marca en el analista.
    if (dbAnalistaZ.has(polizaKey)) {
      const analistaRow = dbAnalistaZ.get(polizaKey);
      const resultado = processAutogestion(analistaRow);

      row[4] = resultado.estado;
      row[2] = obtenerSiguienteExpedidor(listaAnalistas);

      let datosExistentesZ = {};
      try {
        const rawF = (row[5] || "").toString().replace(/:\s*NaN\b/g, ': null');
        if (rawF) datosExistentesZ = JSON.parse(rawF);
      } catch (e) {}
      row[5] = JSON.stringify({ ...datosExistentesZ, ...resultado.data });

      let historialObsZ = [];
      try {
        const rawObs = (row[6] || "").toString().replace(/:\s*NaN\b/g, ': null');
        if (rawObs) {
          const parsed = JSON.parse(rawObs);
          historialObsZ = Array.isArray(parsed) ? parsed : [parsed];
        }
      } catch (e) {
        if (row[6]) historialObsZ.push({ fecha: "Previo", observacion: row[6].toString(), usuario: "Sistema" });
      }
      historialObsZ.push({
        fecha: Utilities.formatDate(new Date(), "America/Bogota", "dd/MM/yyyy HH:mm:ss"),
        usuario: "Sistema (UpdateRenovations - Col Z)",
        estado: resultado.estado,
        obsCliente: resultado.data.observaciones.obsCliente || "",
        observaciones: resultado.data.observaciones.observaciones || "",
        observacionesEjecutivo: resultado.data.observaciones.observacionesEjecutivo || "",
        observacionRenovacion: resultado.data.observaciones.observacionRenovacion || ""
      });
      row[6] = JSON.stringify(historialObsZ);

      polizasProcesadasAnalista.add(polizaKey);
      contadorActualizados++;
      logs.push([
        fechaEjecucion, polizaKey, "Analista Z (Expedido - Prioritario)",
        estadoAnterior, resultado.estado, asesorAnterior, row[2],
        segmento, leadInfo.vencimiento || "",
        resultado.data.NuevaPoliza || "Sin nueva póliza"
      ]);
      continue;
    }

    if (ESTADOS_PROTEGIDOS.has(estado)) {
      contadorProtegidos++;
      continue;
    }

    if (dbAnalista.has(polizaKey)) {
      const analistaRow = dbAnalista.get(polizaKey);
      const resultado = processAutogestion(analistaRow);

      row[4] = resultado.estado;
      row[2] = obtenerSiguienteExpedidor(listaAnalistas);

      let datosExistentes = {};
      try {
        const rawF = (row[5] || "").toString().replace(/:\s*NaN\b/g, ': null');
        if (rawF) datosExistentes = JSON.parse(rawF);
      } catch (e) {}
      row[5] = JSON.stringify({ ...datosExistentes, ...resultado.data });

      let historialObs = [];
      try {
        const rawObs = (row[6] || "").toString().replace(/:\s*NaN\b/g, ': null');
        if (rawObs) {
          const parsed = JSON.parse(rawObs);
          historialObs = Array.isArray(parsed) ? parsed : [parsed];
        }
      } catch (e) {
        if (row[6]) historialObs.push({ fecha: "Previo", observacion: row[6].toString(), usuario: "Sistema" });
      }
      historialObs.push({
        fecha: Utilities.formatDate(new Date(), "America/Bogota", "dd/MM/yyyy HH:mm:ss"),
        usuario: "Sistema (UpdateRenovations)",
        estado: resultado.estado,
        obsCliente: resultado.data.observaciones.obsCliente || "",
        observaciones: resultado.data.observaciones.observaciones || "",
        observacionesEjecutivo: resultado.data.observaciones.observacionesEjecutivo || "",
        observacionRenovacion: resultado.data.observaciones.observacionRenovacion || ""
      });
      row[6] = JSON.stringify(historialObs);

      polizasProcesadasAnalista.add(polizaKey);
      contadorActualizados++;
      logs.push([
        fechaEjecucion, polizaKey, "Analista (Autogestión)",
        estadoAnterior, resultado.estado, asesorAnterior, row[2],
        segmento, leadInfo.vencimiento || "",
        resultado.data.NuevaPoliza || "Sin nueva póliza"
      ]);

    } else {
      let foundRow = null;
      let tipificacionIndex = -1;
      let fuente = "";

      if (dbAsesora.has(polizaKey)) {
        foundRow = dbAsesora.get(polizaKey);
        tipificacionIndex = 41;
        fuente = "Asesora (Propietarios)";
      } else if (dbBroker.has(polizaKey)) {
        foundRow = dbBroker.get(polizaKey);
        tipificacionIndex = 34;
        fuente = "Broker";
      } else if (dbCorretaje.has(polizaKey)) {
        foundRow = dbCorretaje.get(polizaKey);
        tipificacionIndex = 34;
        fuente = "Corretaje";
      }

      if (foundRow) {
        const tipificacion = (foundRow[tipificacionIndex] || "").toString().trim();

        if (tipificacion !== "") {
          const tipificacionNorm = tipificacion.replace(/\s+/g, "").toLowerCase();
          const esVolverALlamar = estado.replace(/\s+/g, "").toLowerCase() === "volverallamar";
          // Solo "expedido" puede cambiar "volver a llamar" desde Asesora/Broker/Corretaje.
          // "autogestionado" solo aplica desde dbAnalista (rama separada arriba).
          const permiteActualizar = tipificacionNorm === "expedido";

          if (esVolverALlamar && !permiteActualizar) {
            // Protegido: "volver a llamar" no se toca. No se loguea ni cuenta como actualización.
          } else {
            row[4] = tipificacion;
            contadorActualizados++;
            logs.push([
              fechaEjecucion, polizaKey, fuente + " (Tipificación)",
              estadoAnterior, row[4].toString().trim(), asesorAnterior, row[2].toString().trim(),
              segmento, leadInfo.vencimiento || "", "Tipificación: " + tipificacion
            ]);
          }
        } else {
          const estadoPrevio = row[4].toString().trim();
          const asesorPrevio = row[2].toString().trim();
          processNoGestion(row, today, fechaVenc, esBroker);
          if (row[4].toString().trim() !== estadoPrevio || row[2].toString().trim() !== asesorPrevio) {
            contadorActualizados++;
            logs.push([
              fechaEjecucion, polizaKey, fuente + " (Sin tipificación → NoGestion)",
              estadoPrevio, row[4].toString().trim(), asesorPrevio, row[2].toString().trim(),
              segmento, leadInfo.vencimiento || "",
              esBroker ? "Broker sin tipificación" : "Sin tipificación"
            ]);
          }
        }
      } else {
        const estadoPrevio = row[4].toString().trim();
        const asesorPrevio = row[2].toString().trim();
        processNoGestion(row, today, fechaVenc, esBroker);
        if (row[4].toString().trim() !== estadoPrevio || row[2].toString().trim() !== asesorPrevio) {
          contadorActualizados++;
          logs.push([
            fechaEjecucion, polizaKey, "Sin gestión (No encontrada)",
            estadoPrevio, row[4].toString().trim(), asesorPrevio, row[2].toString().trim(),
            segmento, leadInfo.vencimiento || "",
            esBroker ? "Broker sin gestión" : "Sin gestión en ninguna fuente"
          ]);
        }
      }
    }
  }

  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

  // Marcado batch de GestionAnalista: una sola escritura en lugar de N setValue() individuales.
  // Usa polizasProcesadasAnalista (preciso) en lugar de polizasEnSheet (amplio).
  // Usa dataAnalistaRaw ya leído al inicio (evita segunda lectura de GestionAnalista).
  let contadorMarcados = 0;
  if (dataAnalistaRaw.length > 0) {
    const marcasCol = [];
    for (const row of dataAnalistaRaw) {
      const polizaAnalista = String(row[10] || "").trim();
      const marcaActual = (row[29] || "").toString().trim();
      if (polizaAnalista && polizasProcesadasAnalista.has(polizaAnalista) && marcaActual === "") {
        marcasCol.push(["Cargado al CRM - " + fechaEjecucion]);
        contadorMarcados++;
      } else {
        marcasCol.push([marcaActual]);
      }
    }
    GestionAnalista.getRange(2, 30, dataAnalistaRaw.length, 1).setValues(marcasCol);
  }

  Logger.log("Registros marcados en columna AD de Analista: " + contadorMarcados);
  Logger.log("=== RESUMEN UpdateRenovations ===");
  Logger.log("Actualizados: " + contadorActualizados);
  Logger.log("Protegidos (no tocados): " + contadorProtegidos);
  Logger.log("Sin póliza (saltados): " + contadorSinPoliza);
  Logger.log("Marcados en Analista (AD): " + contadorMarcados);
  Logger.log("Total logs generados: " + logs.length);

  guardarLogsRenovaciones(logs, fechaEjecucion, contadorActualizados, contadorProtegidos, contadorSinPoliza);
}


/**
 * Guarda los logs de UpdateRenovations en la hoja "LogRenovaciones".
 * Si la hoja no existe, la crea con encabezados.
 * @param {Array[]} logs - Array de filas de log.
 * @param {string} fechaEjecucion - Timestamp de la ejecución.
 * @param {number} actualizados - Total de registros actualizados.
 * @param {number} protegidos - Total de registros protegidos.
 * @param {number} sinPoliza - Total de registros sin póliza.
 */
function guardarLogsRenovaciones(logs, fechaEjecucion, actualizados, protegidos, sinPoliza) {
  let logSheet = Renovations.getSheetByName("LogRenovaciones");

  if (!logSheet) {
    logSheet = Renovations.insertSheet("LogRenovaciones");
    logSheet.appendRow([
      "Fecha Ejecución",
      "Póliza",
      "Fuente",
      "Estado Anterior",
      "Estado Nuevo",
      "Asesor Anterior",
      "Asesor Nuevo",
      "Segmento",
      "Fecha Vencimiento",
      "Detalle"
    ]);
    logSheet.getRange(1, 1, 1, 10).setFontWeight("bold");
    Logger.log("Hoja 'LogRenovaciones' creada con encabezados.");
  }

  if (logs.length > 0) {
    const ultimaFila = logSheet.getLastRow();
    logSheet.getRange(ultimaFila + 1, 1, logs.length, 10).setValues(logs);
  }

  logSheet.appendRow([
    fechaEjecucion,
    "--- RESUMEN ---",
    "Actualizados: " + actualizados,
    "Protegidos: " + protegidos,
    "Sin póliza: " + sinPoliza,
    "Total logs: " + logs.length,
    "", "", "", ""
  ]);

  Logger.log("Logs guardados en hoja 'LogRenovaciones'.");
}
 
/**
 * Crea un Map cargando solo registros cuya clave esté en el Set proporcionado.
 * Optimiza memoria al no cargar pólizas que ya están expedidas.
 * @param {Sheet} sheet - Hoja de cálculo.
 * @param {number} keyColIndex - Índice de la columna clave (base 0).
 * @param {Set} allowedKeys - Set de claves permitidas.
 * @return {Map} Map con clave → fila completa (solo registros con clave en allowedKeys).
 */
function getDatabaseMapByKeys(sheet, keyColIndex, allowedKeys) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Map();

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const map = new Map();

  for (let row of data) {
    let key = String(row[keyColIndex]).trim();
    if (key && allowedKeys.has(key)) {
      map.set(key, row);
    }
  }
  return map;
}


function processAutogestion(analistaRow) {
  const valorZ = (analistaRow[25] || "").toString().trim();

  let estado = "Autogestionado";
  let observacionRenovacion = "";

  if (valorZ !== "") {
    const esNumeroPoliza = /^50\d+/.test(valorZ);
    if (esNumeroPoliza) {
      estado = "Expedido";
      observacionRenovacion = "Póliza renovada con número: " + valorZ;
    } else {
      estado = "Autogestionado";
      observacionRenovacion = valorZ;
    }
  }

  // Columna E (índice 4): nombre del ejecutivo de cuenta responsable
  const nombreEjecutivo = (analistaRow[4] || "").toString().trim();
  const correoEjecutivoCuenta = MAPA_EJECUTIVOS_CUENTA[nombreEjecutivo.toUpperCase()] || "";

  const autoGestionData = {
    sarlaft: analistaRow[7],
    formatoRenovacion: analistaRow[8],
    pazYsalvo: analistaRow[20],
    documento: analistaRow[19],
    NuevaPoliza: /^50\d+/.test(valorZ) ? valorZ : "",
    valorPoliza: analistaRow[25],
    primaNeta: analistaRow[27],
    correo: analistaRow[15],
    correspondencia: analistaRow[16],
    telefono: analistaRow[17],
    ciudad: analistaRow[18],
    ejecutivoCuenta: nombreEjecutivo,
    correoEjecutivoCuenta: correoEjecutivoCuenta,
    observaciones: {
      obsCliente: analistaRow[11],
      observaciones: analistaRow[22],
      observacionesEjecutivo: analistaRow[21],
      observacionRenovacion: observacionRenovacion
    }
  };

  return { estado: estado, data: autoGestionData };
}


function processNoGestion(row, today, fechaVenc, esBroker) {
  if (!fechaVenc) return;

  const estadoActual = (row[4] || "").toString().trim().toLowerCase();

  // Solo aplica si el estado es "Pendiente Renovacion"
  if (estadoActual !== "pendiente renovacion") return;

  const fecha5Meses = new Date(fechaVenc);
  fecha5Meses.setMonth(fecha5Meses.getMonth() + 5);

  if (today >= fecha5Meses) {
    row[4] = "VENCIDO";
  } else {
    // Broker → Propietario: si pasaron 15 días del vencimiento
    if (esBroker) {
      let fecha15 = new Date(fechaVenc);
      fecha15.setDate(fecha15.getDate() + 15);

      if (today >= fecha15) {
        row[3] = "PROPIETARIO";
      }
    }

    // Asignar asesor si no tiene uno asignado
    let asesorActual = (row[2] || "").toString().trim();
    if (asesorActual === "") {
      const nuevoAsesor = getNewLeadAssignment();
      row[2] = nuevoAsesor;
    }
  }
}

function getNewLeadAssignment() {
  let asignacion = AssignLead("Renovations");
  if (!asignacion || !asignacion.email) {
    Logger.log("No hay asesores disponibles para asignar renovación.");
    return "";
  }
  return asignacion.email;
}


function parseDate(dateStr) {
  if (!dateStr) return null;
  try {
    let [day, month, year] = dateStr.split("/");
    return new Date(`${month}/${day}/${year}`);
  } catch (e) {
    return null;
  }
}


function parseLeadJson(jsonString) {
  if (!jsonString) return {};
  try {
    let cleanJson = jsonString.toString().replace(/NaN/g, "null");
    return JSON.parse(cleanJson);
  } catch (e) {
    return {};
  }
}



function UpdateInqulinoSAI() {
  const sheet = DataRenovations;
  const lastRow = sheet.getLastRow();
  let endPintSaiUnique = PropertiesService.getScriptProperties().getProperty('endPintSaiUnique');
  let keySaiUnique = PropertiesService.getScriptProperties().getProperty('keySaiUnique');

  const rangeData = sheet.getRange(2, 1, lastRow - 1, 8);
  const data = rangeData.getValues();

  console.log(`=== Inicio Actualizacion Inquilinos (Total: ${data.length}) ===`);

  const BATCH_SIZE = 50;
  let requests = [];
  let rowIndices = [];
  const headers = {
    "x-api-key": keySaiUnique
  };

  for (let i = 0; i < data.length; i++) {
    try {
      let jsonString = data[i][1];
      if (!jsonString || typeof jsonString !== "string" || jsonString.trim() === "") continue;
      let cleanJson = jsonString.replace(/NaN/g, "null");
      let lead = JSON.parse(cleanJson);
      let solicitud = lead.solicitud;

      if (solicitud) {
        requests.push({
          url: `${endPintSaiUnique}${solicitud}`,
          method: "GET",
          headers: headers,
          followRedirects: true,
          muteHttpExceptions: true
        });
        rowIndices.push(i);
      }

    } catch (err) {
      console.warn(`Error parseando JSON en fila ${i + 2}: ${err.message}`);
    }
  }

  for (let j = 0; j < requests.length; j += BATCH_SIZE) {
    const chunkRequests = requests.slice(j, j + BATCH_SIZE);
    const chunkIndices = rowIndices.slice(j, j + BATCH_SIZE);

    console.log(`Procesando lote ${j} a ${j + chunkRequests.length}...`);

    try {
      const responses = UrlFetchApp.fetchAll(chunkRequests);
      let outputValues = [];

      for (let k = 0; k < responses.length; k++) {
        const responseCode = responses[k].getResponseCode();
        const responseText = responses[k].getContentText();
        let resultObj = { error: "No encontrado/Error API" };
        if (responseCode === 200) {
          try {
            const apiData = JSON.parse(responseText);
            const inquilino = Array.isArray(apiData) ? apiData.find(item => item.tipoDeudor === 'INQUILINO') : null;

            if (inquilino) {
              resultObj = {
                nombre: inquilino.nombre,
                identificacion: inquilino.identificacion
              };
            } else {
              resultObj = { info: "Solicitud encontrada pero sin INQUILINO" };
            }

          } catch (e) {
            resultObj = { error: "Error parseando respuesta API" };
          }
        } else {
          resultObj = { error: `API Error Code: ${responseCode}` };
        }
        outputValues.push([JSON.stringify(resultObj)]);
      }
      for (let x = 0; x < chunkIndices.length; x++) {
        let rowIndexSheet = chunkIndices[x] + 2;
        sheet.getRange(rowIndexSheet, 8).setValue(outputValues[x][0]);
      }

    } catch (err) {
      console.error(`Error procesando lote ${j}: ${err.message}`);
    }
  }
  console.log("=== Actualización Finalizada ===");
}



function guardarGestionRenovacion(datos, observaciones, archivosBase64) {
  try {
    const sheet = DataRenovations;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: "error", message: "Base de datos vacía." };

    const columnBData = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    let targetRowNum = -1;
    let currentJson = {};

    for (let i = 0; i < columnBData.length; i++) {
      let rawJson = columnBData[i][0].toString().replace(/:\s*NaN\b/g, ': null');
      try {
        const json = JSON.parse(rawJson);
        if (String(json.poliza).trim() === String(datos.poliza).trim()) {
          targetRowNum = i + 2;
          currentJson = json;
          break;
        }
      } catch (e) { continue; }
    }

    if (targetRowNum === -1) return { status: "error", message: "Póliza no encontrada en BD." };

    let urlsNuevas = {};
    let finderResult = null
    console.log("Procesando póliza: " + datos.poliza);

    if (archivosBase64 && Object.keys(archivosBase64).length > 0) {
      finderResult = Espejo.getRange("BE2:BE").createTextFinder(datos.poliza).matchEntireCell(true).findNext();
      let ref = null;

      if (finderResult) {
        let finde = finderResult.getRow();
        ref = Espejo.getRange(finde, 2).getDisplayValue();
        console.log("Referencia encontrada en Espejo: " + ref);
      } else {
        finderResult = PolizasAntiguas.getRange("C2:C").createTextFinder(datos.poliza).matchEntireCell(true).findNext();
        if (finderResult) {
          let finder = finderResult.getRow();
          ref = PolizasAntiguas.getRange(finder, 2).getDisplayValue();
          console.log("Referencia encontrada en Polizas Antiguas: " + ref);
        }
      }

      if (!ref || ref.trim() === "") {
        console.log("La póliza es antigua y/o no está en la tabla espejo");
        const nuevaRef = "Ref" + datos.documento;
        const fechaActual = new Date();
        PolizasAntiguas.appendRow([
          fechaActual,
          nuevaRef,
          datos.poliza,
          datos.asegurado
        ]);

        console.log("Registrada en Polizas Antiguas: " + nuevaRef);

        urlsNuevas = crearFolderYGuardarArchivos(archivosBase64, datos, nuevaRef);

      } else {
        urlsNuevas = buscarCarpetaYGuardarArchivos(archivosBase64, datos, ref);
      }
    }

    // Leer JSON actual de Columna_Gestion (columna 6) para merge no destructivo
    var rawColumnaGestion = sheet.getRange(targetRowNum, 6).getValue();
    var currentGestionJson = null;
    try {
      var rawGestionStr = (rawColumnaGestion || "").toString().replace(/:\s*NaN\b/g, ': null');
      if (rawGestionStr.trim() !== '') {
        currentGestionJson = JSON.parse(rawGestionStr);
      }
    } catch (e) {
      currentGestionJson = null;
    }

    // Construir objeto con los nuevos datos a mergear
    var newDataForMerge = {};
    var key;
    for (key in datos) {
      if (Object.prototype.hasOwnProperty.call(datos, key)) {
        newDataForMerge[key] = datos[key];
      }
    }
    for (key in urlsNuevas) {
      if (Object.prototype.hasOwnProperty.call(urlsNuevas, key)) {
        newDataForMerge[key] = urlsNuevas[key];
      }
    }

    // Aplicar merge no destructivo: preserva campos existentes y concatena documentosProceso
    var mergedJson = mergeJsonColumnaGestion(currentGestionJson, newDataForMerge);

    var jsonString = JSON.stringify(mergedJson).replace(/:\s*NaN\b/g, ': null');

    sheet.getRange(targetRowNum, 5).setValue(observaciones.estadoGestion);
    sheet.getRange(targetRowNum, 6).setValue(jsonString);

    if (observaciones.estadoGestion === "Caso Corregido") {
      // Para "Caso Corregido": buscar en historial (columna 7) quién solicitó la corrección
      var cellHistorial = sheet.getRange(targetRowNum, 7);
      var historialPrevio = [];
      try {
        var rawHistPrev = cellHistorial.getValue().toString().replace(/:\s*NaN\b/g, ': null');
        if (rawHistPrev.startsWith(")]}',")) rawHistPrev = rawHistPrev.substring(5);
        if (rawHistPrev.trim() !== '') {
          var parsedHist = JSON.parse(rawHistPrev);
          historialPrevio = Array.isArray(parsedHist) ? parsedHist : [parsedHist];
        }
      } catch (e) { historialPrevio = []; }

      // Buscar la última entrada que solicitó la corrección (estadoGestion contiene "correccion" o procesoEspecial === "CORRECCION")
      var solicitante = '';
      for (var h = historialPrevio.length - 1; h >= 0; h--) {
        var entrada = historialPrevio[h];
        var esCorreccion = (entrada.estadoGestion && entrada.estadoGestion.toLowerCase().indexOf('correccion') !== -1) ||
                           (entrada.procesoEspecial && entrada.procesoEspecial === 'CORRECCION');
        if (esCorreccion && entrada.usuario) {
          solicitante = entrada.usuario;
          break;
        }
      }

      if (solicitante) {
        sheet.getRange(targetRowNum, 3).setValue(solicitante);
        console.log("Caso Corregido: reasignado a solicitante original: " + solicitante);
      } else {
        // Fallback: si no se encuentra solicitante, asignar a expedidor
        var asignado = obtenerSiguienteExpedidor();
        sheet.getRange(targetRowNum, 3).setValue(asignado);
        console.log("Caso Corregido: no se encontró solicitante, asignado a expedidor: " + asignado);
      }
    } else if (observaciones.estadoGestion === "Caso Especial" || observaciones.estadoGestion === "Enviar a Expedicion") {
      var asignado = obtenerSiguienteExpedidor();
      sheet.getRange(targetRowNum, 3).setValue(asignado);
    }

    // Registrar en historial (columna 7) — append sin sobrescribir entradas previas
    var cellObs = sheet.getRange(targetRowNum, 7);
    var historial = [];
    try {
      var rawHist = cellObs.getValue().toString().replace(/:\s*NaN\b/g, ': null');
      if (rawHist.startsWith(")]}',")) rawHist = rawHist.substring(5);
      if (rawHist.trim() !== '') {
        var parsed = JSON.parse(rawHist);
        historial = Array.isArray(parsed) ? parsed : [parsed];
      }
    } catch (e) {
      var cellValue = cellObs.getValue();
      if (cellValue) {
        historial.push({
          fecha: "Previo",
          observacion: cellValue.toString(),
          usuario: "Sistema"
        });
      }
    }

    var fechaCO = Utilities.formatDate(new Date(), "America/Bogota", "dd/MM/yyyy HH:mm:ss");

    // Determinar estado para historial: usa el estado enviado por el frontend
    var estadoHistorial = observaciones.estadoGestion;

    var nuevaEntrada = {
      fecha: fechaCO,
      usuario: Session.getActiveUser().getEmail(),
      observacion: observaciones.observacion,
      estadoGestion: estadoHistorial,
      seguimiento: observaciones.fechaseguimiento || "N/A",
      procesoEspecial: datos.procesoEspecial || "NINGUNO"
    };

    if (datos.motivoDesistimiento) {
      nuevaEntrada.motivoDesistimiento = datos.motivoDesistimiento;
      nuevaEntrada.detalleDesistimiento = datos.detalleDesistimiento || "";
    }

    historial.push(nuevaEntrada);
    cellObs.setValue(JSON.stringify(historial));

    return { status: "success", message: "OK" };

  } catch (error) {
    Logger.log("ERROR FATAL: " + error.stack);
    const esDrive = error.message && error.message.indexOf("Google Drive") !== -1;
    const mensajeUsuario = esDrive
      ? "El servicio de Google Drive presentó fallas al guardar los documentos. Por favor intente nuevamente en unos minutos."
      : error.message;
    return { status: "error", message: mensajeUsuario };
  }
}





function buscarCarpetaYGuardarArchivos(archivosBase64, datos, ref) {
  const rootFolder = retryDrive(() => DriveApp.getFolderById("1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC"));
  const urlsNuevas = {};

  const iteradorCandidatos = retryDrive(() => rootFolder.searchFolders(`title contains '${ref}' and trashed = false`));
  let carpetaCliente = null;

  while (retryDrive(() => iteradorCandidatos.hasNext())) {
    let candidato = retryDrive(() => iteradorCandidatos.next());
    let nombreCarpeta = retryDrive(() => candidato.getName());

    if (nombreCarpeta.trim() === String(ref).trim()) {
      carpetaCliente = candidato;
      break;
    }
    const regex = new RegExp(`\\b${ref}\\b`, 'i');
    if (regex.test(nombreCarpeta)) {
      carpetaCliente = candidato;
      break;
    }
  }

  if (!carpetaCliente) {
    console.error(`No se encontró carpeta para la referencia '${ref}'`);
    return urlsNuevas;
  }

  console.log("Carpeta encontrada: " + retryDrive(() => carpetaCliente.getName()));

  // Crear o buscar carpeta "Renovacion"
  let carpetaRenovacion;
  const nombreRenovacion = "Renovacion";
  const subCarpetas = retryDrive(() => carpetaCliente.getFoldersByName(nombreRenovacion));
  carpetaRenovacion = retryDrive(() => subCarpetas.hasNext()) ? retryDrive(() => subCarpetas.next()) : retryDrive(() => carpetaCliente.createFolder(nombreRenovacion));

  // Guardar archivos principales
  guardarArchivosEnCarpeta(archivosBase64, datos, carpetaRenovacion, urlsNuevas);

  return urlsNuevas;
}


function crearFolderYGuardarArchivos(archivosBase64, datos, referencia) {
  const rootFolder = retryDrive(() => DriveApp.getFolderById("1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC"));
  const urlsNuevas = {};

  console.log("Creando estructura de carpetas para referencia: " + referencia);

  // Crear carpeta principal con el nombre de la referencia
  const carpetaCliente = retryDrive(() => rootFolder.createFolder(referencia));
  console.log("Carpeta cliente creada: " + retryDrive(() => carpetaCliente.getName()));

  // Crear subcarpeta "Renovacion"
  const carpetaRenovacion = retryDrive(() => carpetaCliente.createFolder("Renovacion"));
  console.log("Subcarpeta Renovacion creada");

  // Guardar archivos en la carpeta de renovación
  guardarArchivosEnCarpeta(archivosBase64, datos, carpetaRenovacion, urlsNuevas);

  return urlsNuevas;
}


/**
 * Guarda archivos en la carpeta de renovación del cliente en Google Drive.
 * Procesa archivos individuales (sarlaft, propietarioDoc) y el arreglo documentosProceso
 * para procesos especiales (Otro Sí, Cesión, Correcciones).
 * @param {Object} archivosBase64 - Objeto con archivos en base64. Puede incluir documentosProceso (array).
 * @param {Object} datos - Datos de la gestión (poliza, procesoEspecial, etc.).
 * @param {Folder} carpetaRenovacion - Carpeta "Renovacion" en Google Drive.
 * @param {Object} urlsNuevas - Objeto donde se almacenan las URLs de archivos guardados.
 */
function guardarArchivosEnCarpeta(archivosBase64, datos, carpetaRenovacion, urlsNuevas) {
  /**
   * Decodifica un archivo base64, crea un Blob y lo guarda en la carpeta indicada.
   * @param {Object} fileObj - Objeto con {name, mimeType, data}.
   * @param {string} nombreArchivo - Nombre final del archivo (con extensión).
   * @param {Folder} folder - Carpeta destino en Google Drive.
   * @returns {string} URL del archivo guardado.
   */
  var guardarArchivo = function(fileObj, nombreArchivo, folder) {
    var blob = Utilities.newBlob(
      Utilities.base64Decode(fileObj.data),
      fileObj.mimeType,
      nombreArchivo
    );
    var file = retryDrive(function() { return folder.createFile(blob); });
    console.log("Archivo guardado: " + nombreArchivo);
    return retryDrive(function() { return file.getUrl(); });
  };

  /**
   * Obtiene la extensión de un archivo a partir de su nombre o mimeType.
   * @param {Object} fileObj - Objeto con {name, mimeType}.
   * @returns {string} Extensión del archivo (sin punto).
   */
  var obtenerExtension = function(fileObj) {
    if (fileObj.name && fileObj.name.indexOf('.') !== -1) {
      return fileObj.name.split('.').pop().toLowerCase();
    }
    var mimeMap = {
      'application/pdf': 'pdf',
      'image/jpeg': 'jpg',
      'image/png': 'png'
    };
    return mimeMap[fileObj.mimeType] || 'bin';
  };

  // Guardar archivos principales (sarlaft, documento propietario)
  if (archivosBase64.sarlaft) {
    var extSarlaft = obtenerExtension(archivosBase64.sarlaft);
    urlsNuevas.sarlaftArchivoURL = guardarArchivo(
      archivosBase64.sarlaft,
      'SARLAFT_' + datos.poliza + '.' + extSarlaft,
      carpetaRenovacion
    );
  }
  if (archivosBase64.propietarioDoc) {
    var extDoc = obtenerExtension(archivosBase64.propietarioDoc);
    urlsNuevas.propietarioDocURL = guardarArchivo(
      archivosBase64.propietarioDoc,
      'DOC_ID_' + datos.poliza + '.' + extDoc,
      carpetaRenovacion
    );
  }

  // Guardar arreglo documentosProceso (multi-archivo para procesos especiales)
  if (archivosBase64.documentosProceso && Array.isArray(archivosBase64.documentosProceso) && archivosBase64.documentosProceso.length > 0) {
    var prefijo = obtenerPrefijoProcesoEspecial(datos.procesoEspecial);

    // Crear o buscar subcarpeta "Procesos Especiales"
    var subIter = retryDrive(function() { return carpetaRenovacion.getFoldersByName("Procesos Especiales"); });
    var especialFolder = retryDrive(function() { return subIter.hasNext(); })
      ? retryDrive(function() { return subIter.next(); })
      : retryDrive(function() { return carpetaRenovacion.createFolder("Procesos Especiales"); });

    console.log("Guardando " + archivosBase64.documentosProceso.length + " archivos de proceso especial: " + datos.procesoEspecial);

    var documentosGuardados = [];

    for (var i = 0; i < archivosBase64.documentosProceso.length; i++) {
      var archivo = archivosBase64.documentosProceso[i];
      if (!archivo || !archivo.data || !archivo.mimeType) {
        console.log("Archivo en índice " + i + " inválido, se omite.");
        continue;
      }

      var ext = obtenerExtension(archivo);
      var nombreFinal = prefijo + '_' + datos.poliza + '_' + (i + 1) + '.' + ext;
      var url = guardarArchivo(archivo, nombreFinal, especialFolder);

      documentosGuardados.push({
        nombre: nombreFinal,
        url: url
      });
    }

    urlsNuevas.documentosProceso = documentosGuardados;
  }
}

/**
 * Obtiene el prefijo de nomenclatura para archivos según el tipo de proceso especial.
 * @param {string} procesoEspecial - Tipo de proceso (OTRO_SI, CESION, CORRECCION).
 * @returns {string} Prefijo para el nombre del archivo.
 */
function obtenerPrefijoProcesoEspecial(procesoEspecial) {
  var prefijos = {
    'OTRO_SI': 'OTRO_SI',
    'CESION': 'CESION',
    'CORRECCION': 'CORRECCION'
  };
  return prefijos[procesoEspecial] || 'PROCESO';
}





function enviarOtpWhatsapp(telefono) {
  var authToken = "THVpc2FfU2FudG9zX01rdDpCb2xpMjAyMnZhci4=";
  var infobipUrl = "https://qgmx9r.api.infobip.com/whatsapp/1/message/template";

  var codigoOtp = String(Math.floor(100000 + Math.random() * 900000));

  var cache = CacheService.getScriptCache();
  var otpData = JSON.stringify({
    code: codigoOtp,
    attempts: 0,
    createdAt: new Date().getTime()
  });
  cache.put('otp_' + telefono, otpData, 300);

  var headers = {
    "Authorization": "Basic " + authToken,
    "Content-Type": "application/json"
  };

  var payload = {
    "messages": [
      {
        "from": "573144352014",
        "to": telefono,
        "content": {
          "templateName": "otprenovaciones",
          "templateData": {
            "body": {
              "placeholders": [codigoOtp]
            },
            "buttons": [
              {
                "type": "URL",
                "parameter": codigoOtp
              }
            ]
          },
          "language": "es_MX"
        }
      }
    ]
  };

  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(infobipUrl, options);
    var responseCode = response.getResponseCode();
    if (responseCode === 200 || responseCode === 201) {
      Logger.log("OTP WhatsApp enviado a " + telefono);
      return { success: true, destino: telefono, metodo: "whatsapp" };
    } else {
      Logger.log("Error OTP WhatsApp (" + responseCode + "): " + response.getContentText());
      return { success: false, error: "Error al enviar OTP por WhatsApp" };
    }
  } catch (e) {
    Logger.log("Error de conexión OTP WhatsApp: " + e.toString());
    return { success: false, error: e.toString() };
  }
}


function enviarOtpEmail(email) {
  var codigoOtp = String(Math.floor(100000 + Math.random() * 900000));

  var cache = CacheService.getScriptCache();
  var otpData = JSON.stringify({
    code: codigoOtp,
    attempts: 0,
    createdAt: new Date().getTime()
  });
  cache.put('otp_email_' + email, otpData, 300);

  try {
    MailApp.sendEmail({
      to: email,
      subject: "Código de Verificación - Renovación El Libertador",
      htmlBody: '<div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;padding:20px;">' +
        '<div style="text-align:center;padding:20px;background-color:#673ab7;border-radius:8px 8px 0 0;">' +
        '<h2 style="color:#ffffff;margin:0;">El Libertador</h2>' +
        '</div>' +
        '<div style="padding:30px;background-color:#ffffff;border:1px solid #e0e0e0;">' +
        '<p style="font-size:16px;color:#333;">Su código de verificación para la renovación de póliza es:</p>' +
        '<div style="text-align:center;margin:25px 0;">' +
        '<span style="font-size:36px;font-weight:bold;letter-spacing:8px;color:#673ab7;background-color:#f3e5f5;padding:15px 30px;border-radius:8px;">' + codigoOtp + '</span>' +
        '</div>' +
        '<p style="font-size:14px;color:#666;">Este código expira en <strong>5 minutos</strong>.</p>' +
        '<p style="font-size:13px;color:#999;">Si usted no solicitó este código, ignore este mensaje.</p>' +
        '</div>' +
        '<div style="text-align:center;padding:15px;background-color:#f5f5f5;border-radius:0 0 8px 8px;font-size:11px;color:#999;">' +
        'Investigaciones y Cobranzas El Libertador S.A.' +
        '</div></div>'
    });

    Logger.log("OTP Email enviado a " + email);
    return { success: true, destino: email, metodo: "email" };
  } catch (e) {
    Logger.log("Error enviando OTP por email: " + e.toString());
    return { success: false, error: "No se pudo enviar el correo: " + e.toString() };
  }
}


function validarOtpBackend(identificador, codigoIngresado) {
  var cache = CacheService.getScriptCache();

  // Buscar por teléfono o por email
  var otpRaw = cache.get('otp_' + identificador) || cache.get('otp_email_' + identificador);

  if (!otpRaw) {
    return { valid: false, error: "OTP expirado o no generado. Solicite uno nuevo." };
  }

  var otpData = JSON.parse(otpRaw);
  var MAX_INTENTOS = 5;
  var TTL_MS = 300000;

  var ahora = new Date().getTime();
  if ((ahora - otpData.createdAt) > TTL_MS) {
    cache.remove('otp_' + identificador);
    cache.remove('otp_email_' + identificador);
    return { valid: false, error: "OTP expirado. Solicite uno nuevo." };
  }

  if (otpData.attempts >= MAX_INTENTOS) {
    cache.remove('otp_' + identificador);
    cache.remove('otp_email_' + identificador);
    return { valid: false, error: "Máximo de intentos alcanzado. Solicite un nuevo código." };
  }

  otpData.attempts++;

  if (String(codigoIngresado).trim() === String(otpData.code).trim()) {
    cache.remove('otp_' + identificador);
    cache.remove('otp_email_' + identificador);
    Logger.log("OTP validado correctamente para " + identificador + " por " + Session.getActiveUser().getEmail());
    return { valid: true };
  } else {
    // Actualizar intentos en ambas claves posibles
    var cacheKey = cache.get('otp_' + identificador) ? 'otp_' + identificador : 'otp_email_' + identificador;
    cache.put(cacheKey, JSON.stringify(otpData), 300);
    return { valid: false, error: "Código incorrecto. Intento " + otpData.attempts + " de " + MAX_INTENTOS + "." };
  }
}


function UpdateSarlaft() {
  const sheet = DataRenovations;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange("A2:H" + lastRow).getValues();
  const now = new Date();

  let totalActualizados = 0;
  let totalVencidos = 0;

  console.log("=== INICIO VERIFICACIÓN SARLAFT ===");

  for (let i = 0; i < data.length; i++) {

    try {
      let jsonString = data[i][1];  // Columna B
      if (!jsonString || typeof jsonString !== "string") continue;
      let cleanJson = jsonString.replace(/NaN/g, "null");
      let lead = JSON.parse(cleanJson);
      let tipoDoc = lead.tipoDocumento;
      let numDoc = lead.documento;
      let resp = requestSarlaft(tipoDoc, numDoc);
      let fechaActualizacion = new Date(resp.ultimaFechaActualizacion);
      let diffYears = (now - fechaActualizacion) / (1000 * 60 * 60 * 24 * 365);
      if (diffYears < 3) {
        lead.sarlaft = "ACTUALIZADO";
        lead.fechaUltimoSarlaft = fechaActualizacion.toISOString();
        lead.renovarSarlaft = "NO";
        totalActualizados++;
      } else {
        lead.sarlaft = "VENCIDO";
        lead.fechaUltimoSarlaft = fechaActualizacion.toISOString();
        lead.renovarSarlaft = "SI";
        totalVencidos++;
      }
      sheet.getRange(i + 2, 2).setValue(JSON.stringify(lead));
    } catch (err) {
      console.error(`Error fila ${i + 2}: ${err}`);
    }
  }
  console.log(`Actualizados: ${totalActualizados} | Vencidos: ${totalVencidos}`);
}

function requestSarlaft(tipoDocumento = "CC", numeroDocumento = "1023018112") {
  let urlEndpointSarlaft = PropertiesService.getScriptProperties().getProperty('endPonitSarlaft');
  let keySarlaft = PropertiesService.getScriptProperties().getProperty('keySarlaft');
  const baseUrl = urlEndpointSarlaft;
  const url = `${baseUrl}?pNumeroDocumto=${numeroDocumento}&pTipoDocumto=${tipoDocumento}&pMcaVlrminAseg=S&pMcaVlrminPrima=S`;

  const requestOptions = {
    method: "GET",
    headers: {
      "pAgenciaUsr": "4000",
      "pSistemaOrigen": "100",
      "pPais": "1",
      "pDireccionIp": "100.12.3.4",
      "pInfo": "N",
      "pIpProceso": "240",
      "pIpSubProceso": "1",
      "pCodCia": "2",
      "pCodProducto": "90",
      "pSubProducto": "1",
      "pCodSecc": "34",
      "pCodUsr": "USRSARLAFT",
      "x-api-key": keySarlaft,
    },
    muteHttpExceptions: true
  };

  let response = UrlFetchApp.fetch(url, requestOptions);
  let data = JSON.parse(response.getContentText());
  return data;
}

/**
 * Envía mensaje WhatsApp de póliza vencida al propietario usando la API de Infobip.
 * Usa UrlFetchApp.fetch() (API nativa de GAS) en lugar de fetch() del navegador.
 * @param {string} nombre - Nombre del destinatario.
 * @param {string} telefono - Número de teléfono del destinatario (con código de país).
 * @param {string} direccion - Dirección del inmueble asociado a la póliza.
 * @returns {{success: boolean, destino?: string, error?: string}} Resultado del envío.
 */
function RenovaSendWppVencida(nombre, telefono, direccion) {
  var props = PropertiesService.getScriptProperties();
  var authToken = props.getProperty('infobipAuthToken') || "THVpc2FfU2FudG9zX01rdDpCb2xpMjAyMnZhci4=";
  var infobipUrl = "https://qgmx9r.api.infobip.com/whatsapp/1/message/template";

  var headers = {
    "Authorization": "Basic " + authToken,
    "Content-Type": "application/json"
  };

  var payload = {
    "messages": [
      {
        "from": "573144352014",
        "to": telefono,
        "content": {
          "templateName": "poliza_vencida_propietario_v2",
          "templateData": {
            "body": {
              "placeholders": [nombre, telefono, direccion]
            },
            "header": {
              "type": "IMAGE",
              "mediaUrl": "https://res.cloudinary.com/dsr4y9xyl/image/upload/v1760451940/unnamed_1_jrfjkb.jpg"
            }
          },
          "language": "es_MX"
        },
        "callbackData": "Callback data"
      }
    ]
  };

  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(infobipUrl, options);
    var responseCode = response.getResponseCode();
    if (responseCode === 200 || responseCode === 201) {
      Logger.log("WhatsApp póliza vencida enviado a " + telefono);
      return { success: true, destino: telefono };
    } else {
      Logger.log("Error WhatsApp póliza vencida (" + responseCode + "): " + response.getContentText());
      return { success: false, error: "Error al enviar WhatsApp (" + responseCode + ")" };
    }
  } catch (e) {
    Logger.log("Error de conexión WhatsApp póliza vencida: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Envía mensaje WhatsApp de póliza próxima a vencer al propietario usando la API de Infobip.
 * Usa UrlFetchApp.fetch() (API nativa de GAS) en lugar de fetch() del navegador.
 * @param {string} nombre - Nombre del destinatario.
 * @param {string} telefono - Número de teléfono del destinatario (con código de país).
 * @param {string} direccion - Dirección del inmueble asociado a la póliza.
 * @returns {{success: boolean, destino?: string, error?: string}} Resultado del envío.
 */
function RenovaSendWppProxVen(nombre, telefono, direccion) {
  var props = PropertiesService.getScriptProperties();
  var authToken = props.getProperty('infobipAuthToken') || "THVpc2FfU2FudG9zX01rdDpCb2xpMjAyMnZhci4=";
  var infobipUrl = "https://qgmx9r.api.infobip.com/whatsapp/1/message/template";

  var headers = {
    "Authorization": "Basic " + authToken,
    "Content-Type": "application/json"
  };

  var payload = {
    "messages": [
      {
        "from": "573144352014",
        "to": telefono,
        "content": {
          "templateName": "poliza_proxima_vencer_propietario",
          "templateData": {
            "body": {
              "placeholders": [nombre, telefono, direccion]
            },
            "header": {
              "type": "IMAGE",
              "mediaUrl": "https://res.cloudinary.com/dsr4y9xyl/image/upload/v1760451940/unnamed_1_jrfjkb.jpg"
            }
          },
          "language": "es_MX"
        },
        "callbackData": "Callback data"
      }
    ]
  };

  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(infobipUrl, options);
    var responseCode = response.getResponseCode();
    if (responseCode === 200 || responseCode === 201) {
      Logger.log("WhatsApp póliza próxima a vencer enviado a " + telefono);
      return { success: true, destino: telefono };
    } else {
      Logger.log("Error WhatsApp próxima a vencer (" + responseCode + "): " + response.getContentText());
      return { success: false, error: "Error al enviar WhatsApp (" + responseCode + ")" };
    }
  } catch (e) {
    Logger.log("Error de conexión WhatsApp próxima a vencer: " + e.toString());
    return { success: false, error: e.toString() };
  }
}



function sendMailInmbiliariaBrocker() {
  const myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");
  myHeaders.append("Authorization", "Basic THVpc2FfU2FudG9zX01rdDpCb2xpMjAyMnZhci4=");

  const raw = JSON.stringify({
    "messages": [
      {
        "sender": "Ellibertador@correo.ellibertador.co",
        "destinations": [
          {
            "to": [
              {
                "destination": "",
                "placeholders": "{\"NombreBrIn\":\"INMOBILIARIA SALVATIERRA LTDA\",\"CANTIDAD_POLIZAS\":\"1\",\"ENLACE_AUTOGESTION\":\"https://libertador.com.co/cotizaciones/\",\"ENLACE_UNSUBSCRIBE\":\"https://libertador.com.co/cotizaciones/\"}"
              }
            ]
          }
        ],
        "content": {
          "subject": "Gestione sus Renovaciones en un Solo Click - El Libertador",
          "html": "<!DOCTYPE html>\n<html lang=\"es\">\n\n<head>\n    <meta charset=\"UTF-8\">\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n    <title>Gestione sus Renovaciones en un Solo Click - El Libertador</title>\n    <style>\n        body, table, td, a {\n            -webkit-text-size-adjust: 100%;\n            -ms-text-size-adjust: 100%;\n            font-family: Arial, sans-serif;\n        }\n        table, td {\n            mso-table-lspace: 0pt;\n            mso-table-rspace: 0pt;\n        }\n        img {\n            -ms-interpolation-mode: bicubic;\n            border: 0;\n            height: auto;\n            line-height: 100%;\n            outline: none;\n            text-decoration: none;\n        }\n        a[x-apple-data-detectors] {\n            color: inherit !important;\n            text-decoration: none !important;\n            font-size: inherit !important;\n            font-family: inherit !important;\n            font-weight: inherit !important;\n            line-height: inherit !important;\n        }\n        body {\n            margin: 0;\n            padding: 0;\n            background-color: #f4f4f4;\n        }\n        .main-container {\n            background-color: #ffffff;\n            border-radius: 8px;\n            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.15);\n            max-width: 600px;\n            width: 100%;\n        }\n        .data-box {\n            background-color: #e8f5e9;\n            border-radius: 6px;\n            padding: 20px;\n        }\n        @media screen and (max-width: 600px) {\n            .main-container { width: 100% !important; border-radius: 0; }\n            .responsive-img { width: 100% !important; height: auto !important; }\n            .content-padding { padding: 20px !important; }\n            h1 { font-size: 20px !important; }\n        }\n    </style>\n</head>\n\n<body style=\"margin: 0; padding: 0; background-color: #f4f4f4;\">\n    <table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" style=\"table-layout: fixed; background-color: #f4f4f4;\">\n        <tr>\n            <td align=\"center\" style=\"padding: 20px 0;\">\n                <table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" class=\"main-container\">\n                    <tr>\n                        <td align=\"center\">\n                            <img src=\"https://res.cloudinary.com/dsr4y9xyl/image/upload/v1760451940/unnamed_1_jrfjkb.jpg\" alt=\"Renovación de Pólizas - El Libertador\" width=\"600\" class=\"responsive-img\" style=\"display: block; border-top-left-radius: 8px; border-top-right-radius: 8px;\">\n                        </td>\n                    </tr>\n                    <tr>\n                        <td class=\"content-padding\" style=\"padding: 30px 40px 10px 40px; color: #333333; font-size: 16px; line-height: 1.6;\">\n                            <p>Estimado(a) <strong>{{NombreBrIn}}</strong>,</p>\n                            <p>En <strong>El Libertador</strong> seguimos innovando para facilitar la gestión de sus pólizas.\n                            Ahora puede <strong>renovar, revisar y confirmar</strong> todas sus pólizas de arrendamiento desde una sola plataforma, con un solo clic.</p>\n                            <p>Controle sus renovaciones con total visibilidad, agilidad y respaldo. Hemos preparado el siguiente resumen con la información clave de su gestión actual.</p>\n                        </td>\n                    </tr>\n                    <tr>\n                        <td align=\"center\" style=\"padding: 10px 40px 25px 40px;\">\n                            <table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" class=\"data-box\">\n                                <tr>\n                                    <td align=\"center\" style=\"padding-bottom: 15px; font-size: 18px; color: #1e7e34; font-weight: bold; border-bottom: 2px solid #c8e6c9;\">\n                                        Resumen de Renovaciones Pendientes\n                                    </td>\n                                </tr>\n                                <tr>\n                                    <td style=\"padding: 15px 0 5px 0; font-size: 15px; color: #333333;\">\n                                        <strong>Pólizas a renovar (total):</strong>\n                                        <span style=\"float: right; font-size: 20px; font-weight: bold;\">{{CANTIDAD_POLIZAS}}</span>\n                                    </td>\n                                </tr>\n                                <tr>\n                                    <td style=\"padding: 5px 0; font-size: 15px; color: #333333;\">\n                                        <strong>Ajuste (IPC 2025):</strong>\n                                        <span style=\"float: right; font-size: 20px; color: #d32f2f; font-weight: bold;\">5,2%</span>\n                                    </td>\n                                </tr>\n                            </table>\n                        </td>\n                    </tr>\n                    <tr>\n                        <td align=\"center\" style=\"padding: 0 40px 30px 40px;\">\n                            <table border=\"0\" cellspacing=\"0\" cellpadding=\"0\" width=\"100%\" style=\"background-color: #1e7e34; border-radius: 6px; padding: 25px 20px;\">\n                                <tr>\n                                    <td align=\"center\">\n                                        <h1 style=\"margin: 0; font-size: 22px; color: #ffffff; font-weight: bold;\">¡Renueve todas sus pólizas en línea!</h1>\n                                        <p style=\"margin: 10px 0 20px; font-size: 14px; color: #c8e6c9;\">Desde nuestro portal podrá revisar y renovar fácilmente todas sus pólizas vigentes y próximas a vencer. En el adjunto encontrará el detalle completo de sus pólizas para que las renueve oportunamente.</p>\n                                        <table border=\"0\" cellspacing=\"0\" cellpadding=\"0\">\n                                            <tr>\n                                                <td align=\"center\" style=\"border-radius: 4px; background-color: #ffffff;\">\n                                                    <a href=\"{{ENLACE_AUTOGESTION}}\" target=\"_blank\" style=\"display: block; padding: 12px 25px; font-size: 16px; font-weight: bold; color: #1e7e34; text-decoration: none;\">✅ RENOVAR MIS PÓLIZAS</a>\n                                                </td>\n                                            </tr>\n                                        </table>\n                                        <hr style=\"border: 0; border-top: 1px solid #c8e6c9; width: 80%; margin: 20px auto;\">\n                                    </td>\n                                </tr>\n                            </table>\n                        </td>\n                    </tr>\n                    <tr>\n                        <td align=\"center\">\n                            <img src=\"https://res.cloudinary.com/dsr4y9xyl/image/upload/v1760451940/unnamed_dfrmuc.jpg\" alt=\"Protección y confianza - El Libertador\" width=\"600\" class=\"responsive-img\" style=\"display: block; border-bottom-left-radius: 8px; border-bottom-right-radius: 8px;\">\n                        </td>\n                    </tr>\n                    <tr>\n                        <td align=\"center\" style=\"padding: 20px 40px; font-size: 11px; line-height: 1.5; color: #888888; background-color: #ffffff;\">\n                            Copyright © 2024 <strong>INVESTIGACIONES Y COBRANZAS EL LIBERTADOR S.A.</strong><br>\n                            Dirección: CR 13 #26-45 Piso 16, Bogotá, Colombia 110111.<br>\n                            Si desea modificar cómo recibe estos correos, haga clic en:\n                            <a href=\"{{ENLACE_UNSUBSCRIBE}}\" target=\"_blank\" style=\"color: #1e7e34; text-decoration: underline; font-weight: bold;\">Cancelar suscripción</a>.\n                        </td>\n                    </tr>\n                </table>\n            </td>\n        </tr>\n    </table>\n</body>\n</html>",
          "defaultPlaceholders": "{\"NombreBrIn\":\"Default Broker\",\"CANTIDAD_POLIZAS\":\"0\",\"ENLACE_AUTOGESTION\":\"https://libertador.com.co/cotizaciones/\",\"ENLACE_UNSUBSCRIBE\":\"https://libertador.com.co/cotizaciones/\"}"
        }
      }
    ]
  });

  const requestOptions = {
    method: "POST",
    headers: myHeaders,
    body: raw,
    redirect: "follow"
  };

  fetch("https://qgmx9r.api.infobip.com/email/4/messages", requestOptions)
    .then((response) => response.text())
    .then((result) => console.log(result))
    .catch((error) => console.error(error));
}

function sendWhatsAppSarlaft() {
  // Obtener datos del cliente desde el modal
  const clienteNombre = document.getElementById('db_ASEGURADO').textContent.trim();
  const clienteCelular = document.getElementById('db_CELULAR').textContent.trim();

  if (!clienteCelular) {
    alert('No se encontró el número de celular del cliente.');
    return;
  }

  // Mensaje personalizado para actualización SARLAFT
  const mensaje = `Hola ${clienteNombre}, su SARLAFT está vencido (más de 3 años). Para continuar con la renovación de su póliza, necesitamos que actualice sus datos. Por favor, haga clic en el siguiente enlace para completar la actualización: [ENLACE_DE_ACTUALIZACION]`;

  // Configurar headers para WhatsApp API
  const myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");
  myHeaders.append("Authorization", "Basic THVpc2FfU2FudG9zX01rdDpCb2xpMjAyMnZhci4=");

  const raw = JSON.stringify({
    "from": "573144352014",
    "to": clienteCelular.replace(/\D/g, ''), // Limpiar el número
    "messageId": "sarlaft-update-" + Date.now(),
    "content": {
      "text": mensaje
    },
    "callbackData": "SARLAFT Update Request"
  });

  const requestOptions = {
    method: "POST",
    headers: myHeaders,
    body: raw,
    redirect: "follow"
  };

  // Enviar mensaje
  fetch("https://qgmx9r.api.infobip.com/whatsapp/1/message/text", requestOptions)
    .then((response) => response.text())
    .then((result) => {
      console.log('Mensaje WhatsApp enviado:', result);
      alert('Mensaje de actualización SARLAFT enviado exitosamente al cliente.');
    })
    .catch((error) => {
      console.error('Error al enviar mensaje WhatsApp:', error);
      alert('Error al enviar el mensaje. Por favor, inténtelo nuevamente.');
    });
}
