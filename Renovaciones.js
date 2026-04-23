
const Renovations = SpreadsheetApp.openById("1wxqoUCggSYXE0vOUHdgLnwDYfBnlATeyfTZ8CgQQEY4");
const DataRenovations = Renovations.getSheetByName("JSON");
const PolizasAntiguas = Renovations.getSheetByName("PolizasAntiguas")
const WareHouseRenovations = Renovations.getSheetByName("Renovations");
const GestionRenovaciones = Renovations.getSheetByName("JSON");
const LogsEnvios = Renovations.getSheetByName("LogEnvios");
const LogErrores = Renovations.getSheetByName("LogErrores");
const GestionAsesora = SpreadsheetApp.openById("1MryMxNuW1cCjgJ2xj5tm8ut-wcd7PNQHI3WEpJal0U4").getSheetByName("Propietarios")
const GestionBroker = SpreadsheetApp.openById("1MryMxNuW1cCjgJ2xj5tm8ut-wcd7PNQHI3WEpJal0U4").getSheetByName("Corretaje")
const GestionCorretaje = SpreadsheetApp.openById("1MryMxNuW1cCjgJ2xj5tm8ut-wcd7PNQHI3WEpJal0U4").getSheetByName("Broker")
const GestionAnalista = SpreadsheetApp.openById("1oAfMyBNgkKR97JbUNir7MjM3KFQ2c8YwQ2QrsFvtIrM").getSheetByName("Respuestas Renovación")

const Espejo = SpreadsheetApp.openById("1ACFjJriwgFE-VOUHifx2Rr7zUNk_Ovy0OnQUMHKnvKY").getSheetByName("new_data_polizas- archivo David")

const EXPEDIDORES = [
  "jose.castillo.rodriguez@segurosbolivar.com",
  "jeymmy.aristizabal@segurosbolivar.com"
];

function obtenerSiguienteExpedidor() {
  var props = PropertiesService.getScriptProperties();
  var turno = parseInt(props.getProperty('turnoExpedidor') || '0', 10);
  var expedidor = EXPEDIDORES[turno % EXPEDIDORES.length];
  props.setProperty('turnoExpedidor', String(turno + 1));
  return expedidor;
}



 
function UpdateRenovations() {
  const sheet = DataRenovations;  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, 1, lastRow - 1, 8); // A:H
  const data = range.getValues();


  const dbAnalista = getDatabaseMap(GestionAnalista, 10);  
  const dbAsesora = getDatabaseMap(GestionAsesora, 2);    
  const dbBroker = getDatabaseMap(GestionBroker, 2);
  const dbCorretaje = getDatabaseMap(GestionCorretaje, 2);

  const today = new Date();
  const updates = []; 
  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    let jsonString = row[1];
    let leadInfo = parseLeadJson(jsonString);

    if (!leadInfo || !leadInfo.poliza) {
      continue; 
    }

    const polizaKey = String(leadInfo.poliza).trim();
    const fechaVenc = parseDate(leadInfo.vencimiento);
    const segmento = (row[3] || "").toString().trim().toLowerCase();
    const esBroker = segmento.includes("broker") || segmento.includes("inmobiliaria");
    const estado = row[4].toString().trim().toLowerCase();

     const estadosProtegidos = [
      "expedido",
      "enviar a expedicion",
      "poliza renovada",
      "correccion",
      "recuperado",
      "caso especial",
      "desistido",
      "volver a llamar",
      "caso revisado"
    ];

    if (!estadosProtegidos.includes(estado)) {

      if (dbAnalista.has(polizaKey)) {
        const analistaRow = dbAnalista.get(polizaKey);
        const resultado = processAutogestion(analistaRow);

        row[2] = obtenerSiguienteExpedidor();  
        row[4] = resultado.estado;                     
        row[5] = JSON.stringify(resultado.data);    
        row[6] = JSON.stringify(resultado.data.observaciones);

      } else {
         let foundRow = dbAsesora.get(polizaKey) || dbBroker.get(polizaKey) || dbCorretaje.get(polizaKey);

        if (foundRow) {
          const tipificacion = dbAsesora.has(polizaKey) ? foundRow[41] : foundRow[34];

          if (tipificacion.toString().trim() !== "") {
            row[4] = tipificacion;
          } else {
             processNoGestion(row, today, fechaVenc, esBroker);
          }
        } else {
           processNoGestion(row, today, fechaVenc, esBroker);
        }
      }

    }
  }
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
}
 
function getDatabaseMap(sheet, keyColIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Map();

  // Obtenemos todos los datos de una vez
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const map = new Map();

  for (let row of data) {
    let key = String(row[keyColIndex]).trim(); // Normalizamos la llave
    if (key) map.set(key, row);
  }
  return map;
}


function processAutogestion(analistaRow) {
  // Indices basados en tu código original (ajustados a base 0)
  // Col K es indice 10 en la hoja, pero aquí analistaRow es toda la fila.
  // Col 25 (Nueva Poliza) -> Index 24

  const nuevaPoliza = analistaRow[24];
  const estado = nuevaPoliza ? "Expedido" : "Autogestionado";

  const autoGestionData = {
    sarlaft: analistaRow[7],
    formatoRenovacion: analistaRow[8],
    pazYsalvo: analistaRow[20],
    documento: analistaRow[19],
    NuevaPoliza: nuevaPoliza,
    valorPoliza: analistaRow[25],
    primaNeta: analistaRow[27],
    correo: analistaRow[15],
    correspondencia: analistaRow[16],
    telefono: analistaRow[17],
    ciudad: analistaRow[18],
    observaciones: {
      obsCliente: analistaRow[11],
      observaciones: analistaRow[22],
      observacionesEjecutivo: analistaRow[21]
    }
  };

  return { estado: estado, data: autoGestionData };
}


function processNoGestion(row, today, fechaVenc, esBroker) {
  if (!fechaVenc) return;

  const fecha2Meses = new Date(fechaVenc);
  fecha2Meses.setMonth(fecha2Meses.getMonth() + 2); // Ojo: Tu código decía +6 en la variable fecha2Meses, aunque el nombre sugiere 2.

  let asesorActual = row[2];
  let estadoActual = row[4];

  if (today >= fecha2Meses) {
    row[4] = "VENCIDO";
  } else {
    // Lógica Broker
    if (esBroker) {
      let fecha15 = new Date(fechaVenc);
      fecha15.setDate(fecha15.getDate() + 15);

      if (today >= fecha15) {
        row[3] = "PROPIETARIO"; // Cambio segmento
        row[4] = "Pendiente Renovacion";
        // Asignación de lead si cambió a propietario
        if (!asesorActual || asesorActual.toString().trim() === "") {
          const nuevoAsesor = getNewLeadAssignment();
          row[2] = nuevoAsesor;
        }
        return;
      }
    }

    // Asignación general si no tiene asesor
    if (!asesorActual || asesorActual.toString().trim() === "") {
      const nuevoAsesor = getNewLeadAssignment();
      row[2] = nuevoAsesor;
      if (!estadoActual || estadoActual.toString().trim() === "") {
        row[4] = "Pendiente Renovacion";
      }
    }
  }
}

function getNewLeadAssignment() {
    let asignacion = AssignLead("Renovations");
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

    const finalJsonData = {
      ...datos,
      ...urlsNuevas
    };

    const jsonString = JSON.stringify(finalJsonData).replace(/:\s*NaN\b/g, ': null');

    sheet.getRange(targetRowNum, 5).setValue(observaciones.estadoGestion);
    sheet.getRange(targetRowNum, 6).setValue(jsonString);

    if (observaciones.estadoGestion === "Caso Especial" || observaciones.estadoGestion === "Enviar a Expedicion") {
      var asignado = obtenerSiguienteExpedidor();
      sheet.getRange(targetRowNum, 3).setValue(asignado);
    }
    const cellObs = sheet.getRange(targetRowNum, 7);
    let historial = [];
    try {
      let rawHist = cellObs.getValue().toString().replace(/:\s*NaN\b/g, ': null');
      if (rawHist.startsWith(")]}',")) rawHist = rawHist.substring(5);
      const parsed = JSON.parse(rawHist);
      historial = Array.isArray(parsed) ? parsed : [parsed];
    } catch (e) {
      if (cellObs.getValue()) historial.push({
        fecha: "Previo",
        observacion: cellObs.getValue(),
        usuario: "Sistema"
      });
    }

    const fechaCO = Utilities.formatDate(new Date(), "America/Bogota", "dd/MM/yyyy HH:mm:ss");

    let nuevaEntrada = {
      fecha: fechaCO,
      usuario: Session.getActiveUser().getEmail(),
      observacion: observaciones.observacion,
      estado: observaciones.estadoGestion,
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
    return { status: "error", message: error.message };
  }
}





function buscarCarpetaYGuardarArchivos(archivosBase64, datos, ref) {
  const rootFolder = DriveApp.getFolderById("1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC");
  const urlsNuevas = {};

  const iteradorCandidatos = rootFolder.searchFolders(`title contains '${ref}' and trashed = false`);
  let carpetaCliente = null;

  while (iteradorCandidatos.hasNext()) {
    let candidato = iteradorCandidatos.next();
    let nombreCarpeta = candidato.getName();

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

  console.log("Carpeta encontrada: " + carpetaCliente.getName());

  // Crear o buscar carpeta "Renovacion"
  let carpetaRenovacion;
  const nombreRenovacion = "Renovacion";
  const subCarpetas = carpetaCliente.getFoldersByName(nombreRenovacion);
  carpetaRenovacion = subCarpetas.hasNext() ? subCarpetas.next() : carpetaCliente.createFolder(nombreRenovacion);

  // Guardar archivos principales
  guardarArchivosEnCarpeta(archivosBase64, datos, carpetaRenovacion, urlsNuevas);

  return urlsNuevas;
}


function crearFolderYGuardarArchivos(archivosBase64, datos, referencia) {
  const rootFolder = DriveApp.getFolderById("1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC");
  const urlsNuevas = {};

  console.log("Creando estructura de carpetas para referencia: " + referencia);

  // Crear carpeta principal con el nombre de la referencia
  const carpetaCliente = rootFolder.createFolder(referencia);
  console.log("Carpeta cliente creada: " + carpetaCliente.getName());

  // Crear subcarpeta "Renovacion"
  const carpetaRenovacion = carpetaCliente.createFolder("Renovacion");
  console.log("Subcarpeta Renovacion creada");

  // Guardar archivos en la carpeta de renovación
  guardarArchivosEnCarpeta(archivosBase64, datos, carpetaRenovacion, urlsNuevas);

  return urlsNuevas;
}


function guardarArchivosEnCarpeta(archivosBase64, datos, carpetaRenovacion, urlsNuevas) {
  const guardar = (fileObj, prefix, folder) => {
    try {
      const ext = fileObj.name.split('.').pop();
      const blob = Utilities.newBlob(
        Utilities.base64Decode(fileObj.data),
        fileObj.mimeType,
        `${prefix}_${datos.poliza}.${ext}`
      );
      const file = folder.createFile(blob);
      console.log(`Archivo guardado: ${prefix}_${datos.poliza}.${ext}`);
      return file.getUrl();
    } catch (err) {
      console.error(`Error guardando archivo ${prefix}: ` + err.message);
      return null;
    }
  };

  // Guardar archivos principales
  if (archivosBase64.sarlaft) {
    urlsNuevas.sarlaftArchivoURL = guardar(archivosBase64.sarlaft, 'SARLAFT', carpetaRenovacion);
  }
  if (archivosBase64.propietarioDoc) {
    urlsNuevas.propietarioDocURL = guardar(archivosBase64.propietarioDoc, 'DOC_ID', carpetaRenovacion);
  }

  // Guardar archivos de procesos especiales
  if (datos.procesoEspecial && datos.procesoEspecial !== 'NINGUNO') {
    let especialFolder;
    const subIter = carpetaRenovacion.getFoldersByName("Procesos Especiales");
    especialFolder = subIter.hasNext() ?
      subIter.next() :
      carpetaRenovacion.createFolder("Procesos Especiales");

    console.log("Guardando archivos de proceso especial: " + datos.procesoEspecial);

    if (archivosBase64.otroSi) {
      urlsNuevas.otroSiURL = guardar(archivosBase64.otroSi, 'OTRO_SI', especialFolder);
    }
    if (archivosBase64.cesionCedula) {
      urlsNuevas.cesionCedulaURL = guardar(archivosBase64.cesionCedula, 'CESION_ID', especialFolder);
    }
    if (archivosBase64.cesionDocAdicional) {
      urlsNuevas.cesionDocAdicionalURL = guardar(archivosBase64.cesionDocAdicional, 'CESION_DOC', especialFolder);
    }
  }
}





function enviarOtpWhatsapp(telefono) {
  var authToken = PropertiesService.getScriptProperties().getProperty('infobipAuth');
  var infobipUrl = PropertiesService.getScriptProperties().getProperty('infobipWhatsappUrl') || "https://qgmx9r.api.infobip.com/whatsapp/1/message/template";

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

function RenovaSendWppVencida() {
  const myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");
  myHeaders.append("Authorization", "Basic THVpc2FfU2FudG9zX01rdDpCb2xpMjAyMnZhci4=");

  const raw = JSON.stringify({
    "messages": [
      {
        "from": "573144352014",
        "to": "573222340943",
        "messageId": "a28dd97c-1ffb-4fcf-99f1-0b557ed381da",
        "content": {
          "templateName": "poliza_vencida_propietario_v2",
          "templateData": {
            "body": {
              "placeholders": [
                "Nikol Rodriguez",
                "3123123",
                "Calle 100c sur 7 45"
              ]
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
  });

  const requestOptions = {
    method: "POST",
    headers: myHeaders,
    body: raw,
    redirect: "follow"
  };

  fetch("https://qgmx9r.api.infobip.com/whatsapp/1/message/template", requestOptions)
    .then((response) => response.text())
    .then((result) => console.log(result))
    .catch((error) => console.error(error));
}

function RenovaSendWppProxVen() {
  const myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");
  myHeaders.append("Authorization", "Basic THVpc2FfU2FudG9zX01rdDpCb2xpMjAyMnZhci4=");

  const raw = JSON.stringify({
    "messages": [
      {
        "from": "573144352014",
        "to": "573222340943",
        "messageId": "a28dd97c-1ffb-4fcf-99f1-0b557ed381da",
        "content": {
          "templateName": "poliza_proxima_vencer_propietario",
          "templateData": {
            "body": {
              "placeholders": [
                "Nikol Rodriguez",
                "3123123",
                "Calle 100c sur 7 45"
              ]
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
  });

  const requestOptions = {
    method: "POST",
    headers: myHeaders,
    body: raw,
    redirect: "follow"
  };

  fetch("https://qgmx9r.api.infobip.com/whatsapp/1/message/template", requestOptions)
    .then((response) => response.text())
    .then((result) => console.log(result))
    .catch((error) => console.error(error));
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
