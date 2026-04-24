function doGet(e) {
  let template = HtmlService.createTemplateFromFile("index").evaluate().setTitle("CRM El Libertador Vida, Desempleo").setFaviconUrl("https://www.ellibertador.co/favicon.ico").addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return template;
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

let id_Poliza = PropertiesService.getScriptProperties().getProperty('idPolizas')
let id_Procesamiento = PropertiesService.getScriptProperties().getProperty('idProcesamiento')
let id_DataWereHouse = PropertiesService.getScriptProperties().getProperty('idWareHouse')


const SPREADSHEET_ID = '1uQO1x5npOpVBwV8IJw3s-J2hrQ38RdY4QU6BDB6xfY4';

const idVida = "11F6X0syGZorMvVdLQQMdmPxbk-_CeApu-_pd8qpaAww"

const idProteccionGastos = "1j95i2Nr27bHlq3LK9MvmOvpeW2yV0p09pOtwanxfDww"




const sheetNameVida = 'Base Ventas Vida';
const sheetNameDesempleo = 'Base de ventas protección gastos';

const Polizas = SpreadsheetApp.openById(id_Poliza).getSheetByName("Active_Users");
const Procesamiento = SpreadsheetApp.openById(id_Procesamiento);
const DataWereHouse = SpreadsheetApp.openById(id_DataWereHouse);
const DataCruda = Procesamiento.getSheetByName("DataCruda");
const NewData = Procesamiento.getSheetByName("Filter");
const Leads = DataWereHouse.getSheetByName("Leads");
const DataGestion = DataWereHouse.getSheetByName("Gestion");
const segurosSheet = DataWereHouse.getSheetByName("Seguros");
const formulariosSheet = DataWereHouse.getSheetByName("formularios");

const chunkPolizas = Procesamiento.getSheetByName("Polizas")
const Response1 = Procesamiento.getSheetByName("Response1")
const Response2 = Procesamiento.getSheetByName("Response2")
const Response3 = Procesamiento.getSheetByName("Response3")
const Response4 = Procesamiento.getSheetByName("Response4")
const Response5 = Procesamiento.getSheetByName("Response5")
const filter = Procesamiento.getSheetByName("Filter")
const SheetAplazadas = Procesamiento.getSheetByName("Aplazadas");




function extraerDataCombine1() {
  const columnas = ['A', 'B'];
  const hojasDestino = [Response1, Response2];
  const lastRow = chunkPolizas.getLastRow();
  if (lastRow < 2) {
    console.log("chunkPolizas vacío o sin datos.");
    return;
  }

  const existingSet = new Set();

  hojasDestino.forEach(hoja => {
    const lastRowDestino = hoja.getLastRow();
    if (lastRowDestino >= 2) {
      const existingVals = hoja.getRange(2, 2, lastRowDestino - 1, 1).getDisplayValues().flat();
      existingVals.forEach(v => {
        const s = (v || '').toString().trim();
        if (s) existingSet.add(s);
      });
    }
    console.log("Pólizas existentes en", hoja.getName(), ":", existingSet.size);
  });

  for (let ci = 0; ci < columnas.length; ci++) {
    const col = columnas[ci];
    const hojaDestino = hojasDestino[ci];

    const cells = chunkPolizas.getRange(`${col}2:${col}${lastRow}`).getDisplayValues().flat();
    let allItems = [];
    cells.forEach(cell => {
      const text = (cell || '').toString().trim();
      if (text !== '') {
        const parts = text.split(',').map(p => p.trim()).filter(Boolean);
        if (parts.length) allItems = allItems.concat(parts);
      }
    });

    const uniqueItems = [...new Set(allItems)];
    const toProcess = uniqueItems.filter(p => !existingSet.has(p));
    console.log(`Col ${col}: celdas=${cells.length} items=${allItems.length} unicos=${uniqueItems.length} nuevos=${toProcess.length}`);

    if (toProcess.length > 0) {
      extraerData(toProcess, hojaDestino);
    } else {
      console.log(`Col ${col}: no hay pólizas nuevas - se salta.`);
    }
  }

  console.log("Extraccion de datos para 1250 polizas terminado.");
  validarAsignacionPro()
}

function extraerDataCombine2() {
  const columnas = ['C', 'D', 'E'];
  const hojasDestino = [Response3, Response4, Response5];
  const lastRow = chunkPolizas.getLastRow();
  if (lastRow < 2) {
    console.log("chunkPolizas vacío o sin datos.");
    return;
  }

  const existingSet = new Set();

  hojasDestino.forEach(hoja => {
    const lastRowDestino = hoja.getLastRow();
    if (lastRowDestino >= 2) {
      const existingVals = hoja.getRange(2, 2, lastRowDestino - 1, 1).getDisplayValues().flat();
      existingVals.forEach(v => {
        const s = (v || '').toString().trim();
        if (s) existingSet.add(s);
      });
    }
    console.log("Pólizas existentes en", hoja.getName(), ":", existingSet.size);
  });

  for (let ci = 0; ci < columnas.length; ci++) {
    const col = columnas[ci];
    const hojaDestino = hojasDestino[ci];

    const cells = chunkPolizas.getRange(`${col}2:${col}${lastRow}`).getDisplayValues().flat();
    let allItems = [];
    cells.forEach(cell => {
      const text = (cell || '').toString().trim();
      if (text !== '') {
        const parts = text.split(',').map(p => p.trim()).filter(Boolean);
        if (parts.length) allItems = allItems.concat(parts);
      }
    });

    const uniqueItems = [...new Set(allItems)];
    const toProcess = uniqueItems.filter(p => !existingSet.has(p));
    console.log(`Col ${col}: celdas=${cells.length} items=${allItems.length} unicos=${uniqueItems.length} nuevos=${toProcess.length}`);

    if (toProcess.length > 0) {
      extraerData(toProcess, hojaDestino);
    } else {
      console.log(`Col ${col}: no hay pólizas nuevas - se salta.`);
    }
  }

  console.log("Extraccion de datos terminado: terminado.");
  validarAsignacionPro()
}

function extraerData(polizas, hojaDestino) {
  let arrayPolizas = polizas;
  console.log("Cantidad de polizas a procesar: ", arrayPolizas.length);

  let bloques = chunkArray(arrayPolizas, 100);
  let solicitudes = [];

  const fechaIncionBase = new Date();
  for (let i = 0; i <= 2; i++) {
    const fechaFin = new Date(fechaIncionBase);
    fechaFin.setDate(fechaFin.getDate() - (i * 15));
    const fechaInicio = new Date(fechaFin);
    fechaInicio.setDate(fechaInicio.getDate() - 14);

    const fechaInicioStr = Utilities.formatDate(fechaInicio, "GMT-5", "ddMMyyyy");
    const fechaFinStr = Utilities.formatDate(fechaFin, "GMT-5", "ddMMyyyy");
    console.log("Entre las fechas: ", fechaInicioStr, " y ", fechaFinStr);
    bloques.forEach(function (bloque) {
      let resultados = consultarSolicitudes(bloque, fechaInicioStr, fechaFinStr);
      const data = resultados.flat();
      solicitudes.push(data);
      Utilities.sleep(400);
    });
  }

  let solicitudesUnicas = new Map();
  solicitudes.flat().forEach(row => {
    if (row && row.solicitud) {
      solicitudesUnicas.set(row.solicitud, row);
    }
  });

  const solicitudesSinDuplicados = Array.from(solicitudesUnicas.values());
  console.log("El numero total de solicitudes unicas es", solicitudesSinDuplicados.length);

  let filasParaInsertar = [];
  solicitudesSinDuplicados.forEach(function (row) {
    let NumeroSolicitud = row.solicitud || "";
    let NumeroPoliza = row.poliza || "";
    let NombreInquilino = row.nombreInquilino || "";
    let Identificacion = row.identificacionInquilino || "";
    let Telefono = row.telefonoInquilino || "";
    let Correo = row.correoInquilino || "";
    let Ciudad = row.ciudadInmueble || "";
    let Direccion = row.direccionInmueble || "";
    let TipoInmueble = row.destinoInmueble || "";
    let Canon = row.canon || "";
    let FechaRadicacion = row.fechaRadicacion || "";
    let FechaResultado = row.fechaResultado || "";
    let Estado = row.estadoGeneral || "";
    let TipoDocumento = row.tipoIdentificacion || "";
    let Cuota = row.cuota || "";

    filasParaInsertar.push([
      new Date(), NumeroPoliza, NumeroSolicitud, NombreInquilino, Identificacion, Telefono, Correo,
      Ciudad, Direccion, TipoInmueble, Canon, FechaRadicacion, FechaResultado, Estado, TipoDocumento, Cuota
    ]);
  });

  if (filasParaInsertar.length > 0) {
    const chunkSize = 1000;
    const total = filasParaInsertar.length;
    let insertadas = 0;

    for (let i = 0; i < total; i += chunkSize) {
      const bloque = filasParaInsertar.slice(i, i + chunkSize);
      try {
        let ultimaFila = hojaDestino.getLastRow();
        hojaDestino.getRange(ultimaFila + 1, 1, bloque.length, bloque[0].length)
          .setValues(bloque);
        insertadas += bloque.length;
        console.log(`✅ Se insertaron ${bloque.length} filas en ${hojaDestino.getName()} (total hasta ahora: ${insertadas}/${total}).`);
        Utilities.sleep(500);
      } catch (error) {
        console.error(`❌ Error al insertar bloque ${i / chunkSize + 1} en ${hojaDestino.getName()}:`, error);
      }
    }

    console.log(`✅ Proceso en ${hojaDestino.getName()} completado. Total insertadas: ${insertadas}/${total}.`);
  } else {
    console.log(`No se encontraron nuevas solicitudes para insertar en ${hojaDestino.getName()}.`);
  }
}

function consultarSolicitudes(arrayPolizas, fechaInicio, fechaFin) {
  let urlEndpointSaiProd = PropertiesService.getScriptProperties().getProperty('endpointSaiProd')
  let keyEndpointSaiProd = PropertiesService.getScriptProperties().getProperty('keyEndpointSaiProd')
  const API_BASE_URL = urlEndpointSaiProd;
  const API_KEY = keyEndpointSaiProd;
  const MAX_REINTENTOS = 3;
  const ESTADOS_VALIDOS = ["APROBADA", "APLAZADA", "APLAZADO-NUBE"];


  let resultados = [];
  let polizasPendientes = [...arrayPolizas];
  let intentos = 0;

  while (polizasPendientes.length > 0 && intentos < MAX_REINTENTOS) {
    intentos++;
    Logger.log(`Intento ${intentos}: Procesando ${polizasPendientes.length} pólizas`);

    const requests = polizasPendientes.map(poliza => ({
      url: `${API_BASE_URL}/${poliza}/${fechaInicio}/${fechaFin}`,
      method: "GET",
      headers: { "x-api-key": API_KEY },
      muteHttpExceptions: true
    }));

    try {
      const responses = UrlFetchApp.fetchAll(requests);
      const polizasFallidas = [];

      responses.forEach((response, index) => {
        const poliza = polizasPendientes[index];
        const responseCode = response.getResponseCode();

        if (responseCode === 200) {
          try {
            const data = JSON.parse(response.getContentText());
            if (Array.isArray(data)) {
              resultados.push(...data);
            } else {
              Logger.log(`Póliza ${poliza}: Datos no son array`);
            }
          } catch (error) {
            Logger.log(`Error parsing JSON para póliza ${poliza}: ${error.message}`);
            polizasFallidas.push(poliza);
          }
        } else {
          polizasFallidas.push(poliza);
          Logger.log(`Póliza ${poliza} falló - Código: ${responseCode}`);
        }
      });

      polizasPendientes = polizasFallidas;
      if (polizasPendientes.length === 0) {
        break;
      }
      if (polizasPendientes.length > 0 && intentos < MAX_REINTENTOS) {
        Utilities.sleep(1000);
      }

    } catch (error) {
      Logger.log(`Error en fetchAll - Intento ${intentos}: ${error.message}`);
      break;
    }
  }
  if (polizasPendientes.length > 0) {
    Logger.log(`${polizasPendientes.length} pólizas no se pudieron procesar después de ${MAX_REINTENTOS} intentos:`);
    polizasPendientes.forEach(poliza => Logger.log(`- ${poliza}`));
  }
  const resultadosFiltrados = resultados.filter(row =>
    row && row.estadoGeneral && ESTADOS_VALIDOS.includes(row.estadoGeneral) && row.tipoIdentificacion !== "NT"
  );

  Logger.log(`Procesadas: ${arrayPolizas.length} pólizas, Resultados válidos: ${resultadosFiltrados.length}`);
  return resultadosFiltrados;
}

function chunkArray(array, size) {
  const result = [];
  for (let i = 0; i < array.length; i += size) {
    result.push(array.slice(i, i + size));
  }
  return result;
}

function requestsApiSai() {
  extraerDataCombine1();
  extraerDataCombine2();
  validarAsignacionPro()
}

function validarAsignacionPro() {
  const DataGestionPS = Procesamiento.getSheetByName("DataGestion");
  const hojasResponse = [
    Procesamiento.getSheetByName("Response1"),
    Procesamiento.getSheetByName("Response2"),
    Procesamiento.getSheetByName("Response3"),
    Procesamiento.getSheetByName("Response4"),
    Procesamiento.getSheetByName("Response5")
  ];
  SpreadsheetApp.flush();

  const dataGestion = DataGestionPS.getDataRange().getValues();
  const dataGestionMapeadaCRM = {};
  const dataGestionMapeadaAgencia = {};

  dataGestion.forEach(fila => {
    const solicitudCRM = fila[1];
    if (solicitudCRM) {
      dataGestionMapeadaCRM[solicitudCRM] = {
        estadoGestionCRM: fila[2]
      };
    }

    const solicitudAgencia = fila[6];
    if (solicitudAgencia) {
      dataGestionMapeadaAgencia[solicitudAgencia] = {
        estadoGestionAgencia: fila[11],
        fechaRadicacionAgencia: fila[8]
      };
    }
  });

  hojasResponse.forEach(hoja => {
    if (!hoja) {
      console.log("Error: No se encontró una de las hojas de respuesta.");
      return;
    }

    const lastRow = hoja.getLastRow();
    if (lastRow < 2) {
      console.log(`La hoja ${hoja.getName()} no tiene datos para procesar.`);
      return;
    }

    const valoresC = hoja.getRange(2, 3, lastRow - 1, 1).getValues(); // Columna C
    const valoresM = hoja.getRange(2, 13, lastRow - 1, 1).getValues(); // Columna M
    const valoresN = hoja.getRange(2, 14, lastRow - 1, 1).getValues(); // Columna N

    const resultadosQ = [];
    const resultadosR = [];
    const resultadosS = [];
    const resultadosT = [];

    for (let i = 0; i < valoresC.length; i++) {
      const solicitud = valoresC[i][0];
      const valorM = valoresM[i][0];
      const valorN = valoresN[i][0];

      const datosCruceCRM = dataGestionMapeadaCRM[solicitud] || {};
      const datosCruceAgencia = dataGestionMapeadaAgencia[solicitud] || {};

      const valorQ = datosCruceCRM.estadoGestionCRM || "";
      const valorR = datosCruceAgencia.estadoGestionAgencia || "";
      const valorS = datosCruceAgencia.fechaRadicacionAgencia || "";

      resultadosQ.push([valorQ]);
      resultadosR.push([valorR]);
      resultadosS.push([valorS]);

      let valorEnT = "";

      if (valorQ || valorR) {
        valorEnT = "Gestionado";
      }
      else if ((!valorQ && !valorR) && (valorN === "APLAZADA" || valorN === "APROBADA")) {
        const fechaM = valorM instanceof Date ? valorM : new Date(valorM);
        const fechaS = valorS instanceof Date ? valorS : new Date(valorS);

        if (!isNaN(fechaM.getTime()) && !isNaN(fechaS.getTime())) {
          if (fechaM.getDate() !== fechaS.getDate()) {
            valorEnT = "Gestionado Antes";
          }
        }
      }
      resultadosT.push([valorEnT]);
    }

    hoja.getRange(2, 17, lastRow - 1, 1).setValues(resultadosQ);
    hoja.getRange(2, 18, lastRow - 1, 1).setValues(resultadosR);
    hoja.getRange(2, 19, lastRow - 1, 1).setValues(resultadosS);
    hoja.getRange(2, 20, lastRow - 1, 1).setValues(resultadosT);

    console.log(`Procesamiento de la hoja ${hoja.getName()} completado.`);
  });
}


function filtrarYConsolidarhotHoras() {
  const hojasResponse = [
    Procesamiento.getSheetByName("Response1"),
    Procesamiento.getSheetByName("Response2"),
    Procesamiento.getSheetByName("Response3"),
    Procesamiento.getSheetByName("Response4"),
    Procesamiento.getSheetByName("Response5")
  ];

  if (!NewData || !Polizas) {
    console.error("No se encontró una de las hojas requeridas.");
    return;
  }

  const polizasMapeadas = {};
  const dataPolizas = Polizas.getDataRange().getValues();
  dataPolizas.forEach(fila => {
    const numeroPoliza = fila[0];
    const nombreInmobiliaria = fila[1];
    if (numeroPoliza) {
      polizasMapeadas[numeroPoliza] = nombreInmobiliaria;
    }
  });

  const solicitudesExistentes = new Set();
  const filterLastRow = NewData.getLastRow();
  if (filterLastRow > 1) {
    const solicitudesEnFilter = NewData.getRange(2, 3, filterLastRow - 1, 1).getValues();
    solicitudesEnFilter.forEach(fila => {
      if (fila[0]) {
        solicitudesExistentes.add(fila[0]);
      }
    });
  }
  const registrosAplazados = [];
  const nuevosRegistros = [];
  const hoy = new Date();

  const ayer = new Date(hoy);
  ayer.setDate(hoy.getDate() - 1);

  // Función para corregir el formato de 12 horas a 24 horas
  function corregirHora(fecha) {
    const hora = fecha.getHours();
    if (hora >= 1 && hora <= 6) { // Si la hora es de 1 PM a 6 PM
      fecha.setHours(hora + 12);
    }
    return fecha;
  }

  hojasResponse.forEach(hoja => {
    if (!hoja) return;

    const lastRow = hoja.getLastRow();
    if (lastRow < 2) {
      console.log(`La hoja ${hoja.getName()} no tiene datos para procesar.`);
      return;
    }

    const data = hoja.getRange(2, 1, lastRow - 1, hoja.getLastColumn()).getValues();

    data.forEach(fila => {
      const valorL = fila[11]; // Fecha Radicacion
      const valorM = fila[12]; // Fecha Aprobacion
      const valorN = fila[13]; // Estado
      const valorT = fila[19]; // Columna T (Etiqueta de gestión)
      const solicitud = fila[2];
      const poliza = fila[1]; // Columna B
      const telefono = fila[5];

      const tieneTelefonoValido = !!telefono && telefono.toString().length > 0;
      const solicitudNoEmpiezaCon7 = !solicitud.toString().startsWith('7');
      const estaAprobada = valorN === "APROBADA";
      const estaVaciaT = !valorT;

      const fechaRadicacionDate = valorL instanceof Date ? valorL : new Date(valorL);
      const fechaAprobacionDate = valorM instanceof Date ? valorM : new Date(valorM);

      const esFechaValida = fechaAprobacionDate >= new Date(new Date().setHours(0, 0, 0, 0)) &&
        fechaAprobacionDate < new Date(new Date().setHours(23, 59, 59, 999));

      if (solicitud && esFechaValida && estaAprobada && estaVaciaT && tieneTelefonoValido && solicitudNoEmpiezaCon7 && !solicitudesExistentes.has(solicitud)) {
        const nombreInmobiliaria = polizasMapeadas[poliza] || "";
        fila[16] = nombreInmobiliaria;
        const fechaAprobacionCorregida = corregirHora(new Date(fechaAprobacionDate));
        const fechaRadicacionCorregida = corregirHora(new Date(fechaRadicacionDate));
        const diffRadicacionAprobacionMs = fechaAprobacionCorregida.getTime() - fechaRadicacionCorregida.getTime();
        const diffRadicacionAprobacionHoras = diffRadicacionAprobacionMs / (1000 * 60 * 60);

        const diffAprobacionActualMs = hoy.getTime() - fechaAprobacionCorregida.getTime();
        const diffAprobacionActualHoras = diffAprobacionActualMs / (1000 * 60 * 60);

        let esClienteHot = false;
        if (diffAprobacionActualHoras < 3 && diffRadicacionAprobacionHoras < 12) {
          esClienteHot = true;
        }

        fila.splice(17, 0, esClienteHot ? "HOT LEAD" : "", diffRadicacionAprobacionHoras.toFixed(2), diffAprobacionActualHoras.toFixed(2));

        nuevosRegistros.push(fila);
        solicitudesExistentes.add(solicitud);
      }
      if (valorN === "APLAZADA") {
        registrosAplazados.push(fila)
      }
    });
  });

  if (registrosAplazados.length > 0) {
    const solicitudesExistentes = new Set();
    const ultimaFilaAplazadas = SheetAplazadas.getLastRow();

    if (ultimaFilaAplazadas > 1) {
      const rangoExistente = SheetAplazadas.getRange(2, 3, ultimaFilaAplazadas - 1, 1).getValues();
      rangoExistente.forEach(fila => solicitudesExistentes.add(fila[0]));
    }

    const registrosSinDuplicados = registrosAplazados.filter(registro => {
      const numeroSolicitud = registro[2];
      return !solicitudesExistentes.has(numeroSolicitud);
    });

    if (registrosSinDuplicados.length > 0) {
      const ultimaFilaParaInsertar = SheetAplazadas.getLastRow();
      const numFilas = registrosSinDuplicados.length;
      const numColumnas = registrosSinDuplicados[0].length;
      SheetAplazadas.getRange(ultimaFilaParaInsertar + 1, 1, numFilas, numColumnas).setValues(registrosSinDuplicados);

      console.log(`Se han movido ${numFilas} registros aplazados a la hoja 'Aplazadas'.`);
    } else {
      console.log("No se encontraron nuevos registros aplazados sin duplicados.");
    }
  }


  if (nuevosRegistros.length === 0) {
    console.log("No se encontraron nuevos leads validos.");
    hojasResponse.forEach(hoja => {
      if (hoja) {
        console.log("Se borran todos los registros guardados")
        hoja.getRange(2, 1, hoja.getLastRow(), hoja.getLastColumn()).clearContent();
      }
    });
  } else {
    const maxColumnas = nuevosRegistros.reduce((max, fila) => Math.max(max, fila.length), 0);

    const registrosNormalizados = nuevosRegistros.map(fila => {
      while (fila.length < maxColumnas) {
        fila.push(""); // Agrega celdas vacías al final si faltan columnas
      }
      return fila;
    });

    const startRow = NewData.getLastRow() + 1;
    NewData.getRange(startRow, 1, registrosNormalizados.length, maxColumnas).setValues(registrosNormalizados);

    console.log(`Se han agregado ${registrosNormalizados.length} nuevos registros a la hoja 'Filter'.`);
  }
  LeadhotRDA()
}

function LeadhotRDA() {
  const ultimaFila = SheetAplazadas.getLastRow();
  if (ultimaFila <= 1) {
    console.log("No hay datos en la hoja 'Aplazadas' para procesar.");
    return;
  }
  const registrosAplazados = SheetAplazadas.getRange(2, 1, ultimaFila - 1, 22).getValues();

  const solicitudesEnFilter = NewData.getLastRow() > 1 ? NewData.getRange(2, 3, NewData.getLastRow() - 1, 1).getValues() : [];
  const solicitudesExistentes = new Set(solicitudesEnFilter.flat());

  const Polizas = SpreadsheetApp.openById(id_Poliza).getSheetByName("Active_Users");

  const polizasMapeadas = Polizas.getDataRange().getValues().reduce((mapa, fila) => {
    const numeroPoliza = fila[0];
    const nombreInmobiliaria = fila[1];
    if (numeroPoliza) {
      mapa[numeroPoliza] = nombreInmobiliaria;
    }
    return mapa;
  }, {});

  const hotToAssign = registrosAplazados.filter(fila => {
    const numeroSolicitud = fila[2];
    return (fila[13] === "APROBADA" && !solicitudesExistentes.has(numeroSolicitud));
  });

  if (hotToAssign.length === 0) {
    console.log("No se encontraron leads APROBADOS o nuevos para pasar a la hoja 'Filter'.");
    return;
  }

  const datosParaNewData = hotToAssign.map(fila => {
    const filaDestino = [];

    for (let i = 0; i < 16; i++) {
      filaDestino.push(fila[i]);
    }

    const nombreInmobiliaria = polizasMapeadas[fila[1]] || "";
    filaDestino.push(nombreInmobiliaria); // Poner la inmobiliaria en la posición 16

    // AGREGADO: Agrega la etiqueta "HOT LEAD RDA"
    filaDestino.push("HOT LEAD RDA"); // Poner la etiqueta en la posición 17

    // AGREGADO: Ahora las columnas U y V (índices 20 y 21) se corresponden a los índices 18 y 19
    filaDestino.push(fila[20]);
    filaDestino.push(fila[21]);

    return filaDestino;
  });

  const ultimaFilaDestino = NewData.getLastRow();
  const rangoDestino = NewData.getRange(ultimaFilaDestino + 1, 1, datosParaNewData.length, datosParaNewData[0].length);
  rangoDestino.setValues(datosParaNewData);

  console.log(`Se han movido y marcado ${datosParaNewData.length} registros como "HOT LEAD RDA" en la hoja 'Filter'.`);

  // Bucle para eliminar los registros ya procesados de la hoja de aplazados
  for (let i = registrosAplazados.length - 1; i >= 0; i--) {
    const fila = registrosAplazados[i];
    const numeroSolicitud = fila[2];
    if (fila[13] === "APROBADA" && solicitudesExistentes.has(numeroSolicitud)) {
      // Elimina solo si el registro es APROBADO y ya existía en la hoja de destino.
      // Así se evita un posible bucle infinito al re-procesar registros que no se movieron por duplicados.
      SheetAplazadas.deleteRow(i + 2);
    }
  }
  console.log("Registros APROBADOS y ya existentes en la hoja de destino, eliminados de 'Aplazadas'.");
}

function formatName(fullName) {
  if (!fullName) return "";
  const parts = fullName.split(/\s+/).filter(Boolean);
  if (parts.length < 2) {
    return fullName.toUpperCase();
  }
  const lastPart = parts[parts.length - 1];
  const secondLastPart = parts[parts.length - 2];

  if (lastPart === secondLastPart) {
    parts.pop();
  }
  const firstName = parts[parts.length - 2];
  const secondName = parts[parts.length - 1];
  const firstLastName = parts[0];
  const secondLastName = parts[1];
  const formattedName = `${firstName} ${secondName} ${firstLastName} ${secondLastName}`;
  return formattedName.toUpperCase();
}

function UpdateDataLead() {
  let dataToAssign = NewData.getRange("A2:T" + NewData.getLastRow()).getDisplayValues();
  let rangoBusqueda = Leads.getRange("B2:B" + Leads.getLastRow()).getValues();
  let solicitudesExistentes = new Set(rangoBusqueda.flat().map(s => (s || '').toString().trim()));

  let capacidadTotal = DataGestion.getRange("D2:H" + DataGestion.getLastRow()).getValues();

  let capacidadFiltrada = capacidadTotal.filter(function (row) {
    return (row[0] === "Seguro de Vida" || row[0] === "Seguro de Desempleo") && row[1] === ""
  })

  let capacidad = capacidadFiltrada.reduce(function (acc, row) {
    let capacidadFila = Number(row[2]) || 0;
    let pendienGestion = Number(row[3]) || 0;
    return acc + (capacidadFila - pendienGestion);
  }, 0);

  const lastRow = NewData.getLastRow();
  if (lastRow < 2) {
    console.log("No hay datos nuevos para asignar.");
    filtrarYConsolidar();
    return;
  }

  console.log("La capacidad ", capacidad);
  console.log("Para Asignar", dataToAssign.length);

  dataToAssign.sort((a, b) => {
    const tipoA = a[17];
    const tipoB = b[17];
    const esRdaA = tipoA === "HOT LEAD RDA";
    const esRdaB = tipoB === "HOT LEAD RDA";
    const esHotA = tipoA === "HOT LEAD";
    const esHotB = tipoB === "HOT LEAD";
    if (esRdaA && !esRdaB) return -1;
    if (!esRdaA && esRdaB) return 1;
    if (esHotA && !esHotB) return -1;
    if (!esHotA && esHotB) return 1;
    const tiempoA = parseFloat(a[19]);
    const tiempoB = parseFloat(b[19]);
    return tiempoA - tiempoB;
  });

  const solicitudesAsignadas = [];

  for (let i = 0; i < capacidad && i < dataToAssign.length; i++) {
    let asignacion = AssignLead("Sales");
    let fila = dataToAssign[i];

    let NumeroSolicitud = fila[2];

    if (NumeroSolicitud === "Numero Solicitud" || solicitudesExistentes.has(NumeroSolicitud) || !NumeroSolicitud) {
      continue;
    }
    solicitudesExistentes.add(NumeroSolicitud);
    solicitudesAsignadas.push(NumeroSolicitud);

    let NumeroPoliza = fila[1];
    let NombreInquilino = fila[3];
    let Identificacion = fila[4];
    let Telefono = fila[5];
    let Correo = fila[6];
    let Ciudad = fila[7];
    let Direccion = fila[8];
    let TipoInmueble = fila[9];
    let CanonString = String(fila[10]).replace(/[$.]/g, '');
    let Canon = Number(CanonString) || 0;
    let FechaRadicacion = fila[11];
    let FechaResultado = fila[12];
    let Estado = fila[13];
    let EtapaFunel = "Contacto Inicial y Datos";
    let EstadoGestion = fila[17] === "HOT LEAD" || fila[17] === "HOT LEAD RDA" ? "HOT Lead" : "Pendiente Gestion";

    let AsesorAsignado = asignacion.email;
    let Producto = asignacion.productoAsignado;
    let TipoDocumento = fila[14];
    let Cuota = fila[15];
    let NombreInmobiliaria = fila[16];
    let TiempoProceso = fila[18];
    let TiempoDesdeAprobacion = fila[19];

    let leadData = {
      poliza: NumeroPoliza,
      nombre: NombreInquilino,
      id: Identificacion,
      telefono: Telefono,
      correo: Correo,
      ciudad: Ciudad,
      direccion: Direccion,
      tipoInmueble: TipoInmueble,
      canon: Canon,
      fechaRadicacion: FechaRadicacion,
      fechaAprobacion: FechaResultado,
      estado: Estado,
      tipoDocumento: TipoDocumento,
      cuota: Cuota,
      nombreInmobiliaria: NombreInmobiliaria,
      hot: fila[17],
      tiempoProceso: TiempoProceso,
      tiempoDesdeAprobacion: TiempoDesdeAprobacion
    };
    let infoLead = JSON.stringify(leadData);

    console.log(EstadoGestion)

    Leads.appendRow([
      new Date(), NumeroSolicitud, infoLead, AsesorAsignado, EtapaFunel, EstadoGestion, Producto
    ]);
  }

  if (solicitudesAsignadas.length > 0) {
    solicitudesAsignadas.sort((a, b) => b - a);

    const rangoBusqueda = NewData.getRange(2, 3, NewData.getLastRow() - 1, 1);

    solicitudesAsignadas.forEach(solicitud => {
      const finder = rangoBusqueda.createTextFinder(solicitud).matchEntireCell(true).findNext();
      if (finder) {
        const fila = finder.getRow();
        NewData.deleteRow(fila);
      }
    });
    validarAsignacionPro();
  }
}


const asignacionesEnEjecucion = {};


function AssignLead(deal) {

  let userDataleads = DataGestion.getRange("A1:K" + DataGestion.getLastRow()).getDisplayValues();

  let userData;

  if (deal === "Renovations") {
    userData = userDataleads.filter(function (row) {
      return row[3] === "Renovations";
    })
  } else if (deal === "Sales") {
    userData = userDataleads.filter(function (row) {
      return row[3] === "Seguro de Vida" || row[3] === "Seguro de Desempleo";
    })
  }
  let bestAgent = null;
  let highestEffectiveness = -1;
  let bestSortingKey = Number.POSITIVE_INFINITY;
  let equity = 0.5;

  for (let i = 0; i < userData.length; i++) {
    let id = userData[i][0];
    let agentName = userData[i][1];
    let email = userData[i][2];
    let novelty = userData[i][4];
    let productoAsignado = userData[i][3];
    let totalCapacity = Number(userData[i][5]);
    let totalInProcess = Number(userData[i][6]);
    let effectivenessStr = userData[i][10] ? userData[i][10].toString().replace("%", "").trim() : "0";
    let effectiveness = Number(effectivenessStr);


    let availability = totalCapacity - totalInProcess;

    if ((!novelty || novelty.toString().trim() === "") && availability > 0) {

      console.log(availability)
      let sortingKey = (-(1 - equity) * totalCapacity) + (equity * totalInProcess);

      if (sortingKey < bestSortingKey || (sortingKey === bestSortingKey && effectiveness > highestEffectiveness)) {
        bestSortingKey = sortingKey;
        highestEffectiveness = effectiveness;

        bestAgent = {
          name: agentName,
          email: email,
          productoAsignado: productoAsignado
        };
      }
    }
  }
  if (bestAgent !== null) {
    console.log("The most available agent is: " + bestAgent.name);
    console.log("gestion Asigned: " + bestAgent.productoAsignado)
    return bestAgent;
  } else {
    Logger.log("No agents available.");
    return null;
  }
}

function getDatauserPro() {
  const correoActivo = Session.getActiveUser().getEmail();
  const validarUsuario = DataGestion.getRange("C2:C").createTextFinder(correoActivo).matchEntireCell(true).ignoreDiacritics(true).findNext();

  if (!validarUsuario) {
    return { estado: "No Autenticado" };
  }

  const filaUsuario = validarUsuario.getRow();
  const expertise = DataGestion.getRange(filaUsuario, 4).getDisplayValue();
  const nombreAnalista = DataGestion.getRange(filaUsuario, 2).getDisplayValue();
  const novedad = DataGestion.getRange(filaUsuario, 5).getDisplayValue();
  const totalSolicitudes = DataGestion.getRange(filaUsuario, 9).getDisplayValue();
  const totalGestionadas = DataGestion.getRange(filaUsuario, 12).getDisplayValue();
  const totalVentas = DataGestion.getRange(filaUsuario, 10).getDisplayValue();
  const totalConversion = DataGestion.getRange(filaUsuario, 11).getDisplayValue();
  const contactabilidad = 10;



  let dataFront = [];
  let dataRecuperacion
  let dataGestionada = [];

  if (expertise === "Seguro de Vida" || expertise === "Seguro de Desempleo") {

    let dataSetPlano = Leads.getRange("A1:O" + Leads.getLastRow()).getDisplayValues();
    let leadsDelAsesor = dataSetPlano.filter(row => row[3] && row[3].toString().trim().toLowerCase() === correoActivo.trim().toLowerCase());

    dataFront = leadsDelAsesor.filter(row => row[5] !== "VENTA" && row[5] !== "DESISTIDO");
    dataGestionada = leadsDelAsesor.filter(row => row[5] === "VENTA" || row[5] === "DESISTIDO");

    dataFront = dataFront.map(row => {

      const infoTexto = row[2];
      const historiaGestiones = row[7];
      const gestionSeguroVida = row[9];
      const gestionSeguroDesempleo = row[10];
      const leadData = JSON.parse(infoTexto);

      let gestiones = [];
      if (historiaGestiones && String(historiaGestiones).trim() !== "") {
        let cleanHistoryString = String(historiaGestiones).trim();
        if (cleanHistoryString.startsWith(")]}',")) {
          cleanHistoryString = cleanHistoryString.substring(5);
        }
        try {
          gestiones = JSON.parse(cleanHistoryString);
          gestiones = Array.isArray(gestiones) ? gestiones : [gestiones]; // Asegurar que sea un array
        } catch (e) {
          Logger.log("Error parseando historial de gestiones: " + e.message + " Contenido: " + cleanHistoryString);
          gestiones = [];
        }
      }

      const datosVida = JSON.parse(gestionSeguroVida || "{}");
      const datosDesempleo = JSON.parse(gestionSeguroDesempleo || "{}");

      return {
        fechaIngreso: row[0],
        poliza: leadData.poliza || "",
        numeroSolicitud: row[1],
        nombre: leadData.nombre || "",
        id: leadData.id || "",
        telefono: leadData.telefono || "",
        correo: leadData.correo || "",
        ciudad: leadData.ciudad || "",
        direccion: leadData.direccion || "",
        tipoInmueble: leadData.tipoInmueble || "",
        canon: leadData.canon || "",
        fechaRadicacion: leadData.fechaRadicacion || "",
        fechaAprobacion: leadData.fechaAprobacion || "",
        estado: leadData.estado || "",
        asesorAsignado: row[3],
        etapaFunel: row[4],
        estadoGestion: row[5],
        productoAsignado: row[6],
        historiaGestiones: gestiones,
        datosVida: datosVida,
        datosDesempleo: datosDesempleo,
        tipoDocumento: leadData.tipoDocumento,
        cuota: leadData.cuota,
        nombreInmobiliaria: leadData.nombreInmobiliaria
      };
    });

    dataGestionada = dataGestionada.map(row => {
      const infoTexto = row[2];
      const historiaGestiones = row[7];
      const gestionSeguroVida = row[9];
      const gestionSeguroDesempleo = row[10];

      let leadDataGralString = (infoTexto && String(infoTexto).trim() !== "") ? String(infoTexto).trim() : "{}";
      if (leadDataGralString.startsWith(")]}',")) {
        leadDataGralString = leadDataGralString.substring(5);
      }
      let leadData = JSON.parse(leadDataGralString);

      let gestiones = [];
      if (historiaGestiones && String(historiaGestiones).trim() !== "") {
        let cleanHistoryString = String(historiaGestiones).trim();
        if (cleanHistoryString.startsWith(")]}',")) {
          cleanHistoryString = cleanHistoryString.substring(5);
        }
        try {
          gestiones = JSON.parse(cleanHistoryString);
          gestiones = Array.isArray(gestiones) ? gestiones : [gestiones]; // Asegurar que sea un array
        } catch (e) {
          Logger.log("Error parseando historial de gestiones: " + e.message + " Contenido: " + cleanHistoryString);
          gestiones = [];
        }
      }

      const datosVida = JSON.parse(gestionSeguroVida || "{}");
      const datosDesempleo = JSON.parse(gestionSeguroDesempleo || "{}");

      return {
        fechaIngreso: row[0],
        poliza: leadData.poliza || "",
        numeroSolicitud: row[1],
        nombre: leadData.nombre || "",
        id: leadData.id || "",
        telefono: leadData.telefono || "",
        correo: leadData.correo || "",
        ciudad: leadData.ciudad || "",
        direccion: leadData.direccion || "",
        tipoInmueble: leadData.tipoInmueble || "",
        canon: leadData.canon || "",
        fechaRadicacion: leadData.fechaRadicacion || "",
        fechaAprobacion: leadData.fechaAprobacion || "",
        estado: leadData.estado || "",
        asesorAsignado: row[3],
        etapaFunel: row[4],
        estadoGestion: row[5],
        productoAsignado: row[6],
        historiaGestiones: gestiones,
        datosVida: datosVida,
        datosDesempleo: datosDesempleo,
        tipoDocumento: leadData.tipoDocumento,
        cuota: leadData.cuota,
        nombreInmobiliaria: leadData.nombreInmobiliaria
      };
    });

    console.log(expertise)
  } else if (expertise === "Renovations" || expertise === "CorreccionesBI") {
    const dataSetPlano = DataRenovations.getRange("A1:H" + DataRenovations.getLastRow()).getDisplayValues();
    const estadosTerminales = new Set([
      "enviar a expedicion",
      "caso especial",
      "cliente ya renovo",
      "no interesado",
      "cambio de compañía",
      "cliente fallecido",
      "propietario fallecido",
      "otro",
      "vencido",
      "poliza renovada",
      "expedido",
      "autogestionado",
      "Desistimiento",
      "venta",
      "confirmacion de venta"
    ]);

    const estadosPermitidos = new Set([
      "correccion",
      "pendiente renovacion",
      "volver a llamar",
      "recuperado"
    ]);

    dataFront = dataSetPlano.filter(row => {
      if (!row[2] || row[2].toString().trim().toLowerCase() !== correoActivo.trim().toLowerCase()) {
        return false;
      }
      let segmento = (row[3] || "").toString().trim().toUpperCase();
      let estado = (row[4] || "").toString().trim().toLowerCase();
      let esSegmentoValido = false;

      if (estado.includes("correccion")) {
        esSegmentoValido =
          segmento === "SIN SEGMENTO" ||
          segmento === "PROPIETARIO" ||
          segmento === "BROKER" ||
          segmento === "INMOBILIARIA";
      } else {
        esSegmentoValido =
          segmento === "SIN SEGMENTO" ||
          segmento === "PROPIETARIO" ||
          segmento === "BROKER" ||
          segmento === "INMOBILIARIA";
      }

      if (!esSegmentoValido) {
        return false;
      }

      if (estadosTerminales.has(estado)) {
        return false; // Bloquear estados terminales
      }

      const esEstadoPermitido = Array.from(estadosPermitidos)
        .some(e => estado.includes(e));

      if (esEstadoPermitido) {
        return true;
      }


      // Si no está en ninguna lista, por defecto bloquear
      return false;

    }).map(row => {

      const registro = row[1];
      const historiaGestiones = row[6];
      let datosInquilino = {};
      try {
        let inquilinoRaw = row[7]; // Columna H
        if (inquilinoRaw && inquilinoRaw.trim() !== "") {
          datosInquilino = JSON.parse(inquilinoRaw);
        }
      } catch (e) {
        datosInquilino = { nombre: "Error Datos", identificacion: "" };
      }

      let gestiones = [];
      if (historiaGestiones && String(historiaGestiones).trim() !== "") {
        let cleanHistoryString = String(historiaGestiones).trim();
        if (cleanHistoryString.startsWith(")]}',")) {
          cleanHistoryString = cleanHistoryString.substring(5);
        }
        try {
          gestiones = JSON.parse(cleanHistoryString);
          gestiones = Array.isArray(gestiones) ? gestiones : [gestiones]; // Asegurar que sea un array
        } catch (e) {
          Logger.log("Error parseando historial de gestiones: " + e.message + " Contenido: " + cleanHistoryString);
          gestiones = [];
        }
      }

      let historialGestion = [];
      leadData = parseLeadData(registro)

      if (historiaGestiones !== "") {
        historialGestion = gestiones || [];
      }
      return {
        fechaIngreso: row[0],
        leadData: leadData,
        nombreAgente: row[2],
        etapaFunel: row[3],
        estadoGestion: row[4],
        historialGestiones: historialGestion,
        datosInquilino: datosInquilino
      };
    });

    var estadosTerminalesRecuperacion = new Set([
      "expedido", "desistido", "cliente ya renovo", "no interesado",
      "desocupación", "venta inmueble", "no justifica motivo",
      "cambio de arrendatario", "póliza revocada", "ilocalizado",
      "no renueva", "vencido", "cambio de compañía", "autogestionado",
      "caso revisado"
    ]);

    dataRecuperacion = dataSetPlano.filter(function(row) {
      var estado = (row[4] || "").toString().trim().toLowerCase();
      return estadosTerminalesRecuperacion.has(estado);
    }).map(function(row) {
      var leadData = parseLeadData(row[1]);
      if (!leadData || !leadData.poliza) return null;

      var historialGestiones = [];
      try {
        var rawHist = (row[6] || "").toString().trim().replace(/:\s*NaN\b/g, ': null');
        if (rawHist && rawHist !== "") {
          if (rawHist.startsWith(")]}',")) rawHist = rawHist.substring(5);
          var parsed = JSON.parse(rawHist);
          historialGestiones = Array.isArray(parsed) ? parsed : [parsed];
        }
      } catch (e) { historialGestiones = []; }

      return {
        fechaIngreso: row[0],
        leadData: leadData,
        nombreAgente: row[2],
        etapaFunel: row[3],
        estadoGestion: row[4],
        historialGestiones: historialGestiones
      };
    }).filter(Boolean);
  }


  return {
    estado: "Autenticado",
    dataFront: dataFront,
    dataRecuperacion: dataRecuperacion,
    dataGestionada: dataGestionada,
    correoActivo: correoActivo,
    expertise: expertise,
    nombreAnalista: nombreAnalista,
    novedad: novedad,
    totalGestionadas: totalGestionadas,
    totalVentas: totalVentas,
    totalSolicitudes: totalSolicitudes,
    totalConversion: totalConversion,
    contactabilidad: contactabilidad
  };
}
function parseLeadData(registro) {
  if (!registro) return {};
  let str = String(registro).trim().replace(/\bNaN\b/g, "null");
  try {
    return JSON.parse(str);
  } catch (e) {
    Logger.log("Error parseando registro: " + e.message + " → " + str);
    return {};
  }
}


function recuperarPoliza(poliza, observacion) {
  try {
    var sheet = SpreadsheetApp.openById("1wxqoUCggSYXE0vOUHdgLnwDYfBnlATeyfTZ8CgQQEY4").getSheetByName("JSON");
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: "error", message: "Base de datos vacia." };

    var columnBData = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    var targetRowNum = -1;

    for (var i = 0; i < columnBData.length; i++) {
      var rawJson = columnBData[i][0].toString().replace(/:\s*NaN\b/g, ': null');
      try {
        var json = JSON.parse(rawJson);
        if (String(json.poliza).trim() === String(poliza).trim()) {
          targetRowNum = i + 2;
          break;
        }
      } catch (e) { continue; }
    }

    if (targetRowNum === -1) return { status: "error", message: "Poliza no encontrada." };

    var correoActivo = Session.getActiveUser().getEmail();
    sheet.getRange(targetRowNum, 5).setValue("Recuperado");
    sheet.getRange(targetRowNum, 3).setValue(correoActivo);

    var cellObs = sheet.getRange(targetRowNum, 7);
    var historial = [];
    try {
      var rawHist = cellObs.getValue().toString().replace(/:\s*NaN\b/g, ': null');
      if (rawHist && rawHist.trim() !== "") {
        if (rawHist.startsWith(")]}',")) rawHist = rawHist.substring(5);
        var parsed = JSON.parse(rawHist);
        historial = Array.isArray(parsed) ? parsed : [parsed];
      }
    } catch (e) {
      if (cellObs.getValue()) historial.push({ fecha: "Previo", observacion: cellObs.getValue(), usuario: "Sistema" });
    }

    var fechaCO = Utilities.formatDate(new Date(), "America/Bogota", "dd/MM/yyyy HH:mm:ss");
    historial.push({
      fecha: fechaCO,
      usuario: correoActivo,
      observacion: observacion,
      estado: "Recuperado"
    });
    cellObs.setValue(JSON.stringify(historial));

    return { status: "success" };
  } catch (error) {
    Logger.log("ERROR recuperarPoliza: " + error.stack);
    return { status: "error", message: error.message };
  }
}


function saveData(dataFromFrontend) {
  let finderResult = Leads.getRange("B:B").createTextFinder(dataFromFrontend.solicitud).matchEntireCell(true).ignoreDiacritics(true).findNext();
  Logger.log("Data recibida en saveData: " + JSON.stringify(dataFromFrontend));

  if (finderResult) {
    let rowNum = finderResult.getRow();

    Leads.getRange(rowNum, 5).setValue(dataFromFrontend.leadEtapaActual);
    Leads.getRange(rowNum, 6).setValue(dataFromFrontend.accionRealizada);
    Leads.getRange(rowNum, 7).setValue(dataFromFrontend.productosGestion);
    Leads.getRange(rowNum, 9).setValue(dataFromFrontend.fecha);


    Leads.getRange(rowNum, 8).setValue(dataFromFrontend.observacionesGestion);

    let currentVidaDataInSheet = Leads.getRange(rowNum, 10).getValue();
    let currentDesempleoDataInSheet = Leads.getRange(rowNum, 11).getValue();

    if (dataFromFrontend.datosCompletosVida) {
      Leads.getRange(rowNum, 10).setValue(dataFromFrontend.datosCompletosVida);
    } else if (currentVidaDataInSheet) {
      Leads.getRange(rowNum, 10).setValue(currentVidaDataInSheet);
    } else {
      Leads.getRange(rowNum, 10).setValue(null);
    }

    if (dataFromFrontend.datosCompletosDesempleo) {
      Leads.getRange(rowNum, 11).setValue(dataFromFrontend.datosCompletosDesempleo);
    } else if (currentDesempleoDataInSheet) {
      Leads.getRange(rowNum, 11).setValue(currentDesempleoDataInSheet);
    } else {
      Leads.getRange(rowNum, 11).setValue(null);
    }

    const accion = (String(dataFromFrontend.accionRealizada || ""));

    console.log(dataFromFrontend.datosLeadGeneral)

    Utilities.sleep(400);
    if (accion === "DESISTIDO" || accion === "VENTA" || accion.includes("VENTA") || accion.includes("Confirmacion de venta") || accion === "COMPLETE_LEAD" || accion === "COMPLETE") {

      console.log("entro")
      try {
        appendLeadDataToSheet(dataFromFrontend, dataFromFrontend.productosGestion || "");
        Logger.log("saveData: appendLeadDataToSheet llamado para solicitud " + dataFromFrontend.solicitud + " con accion " + accion);
      } catch (e) {
        Logger.log("saveData: Error llamando appendLeadDataToSheet para solicitud " + dataFromFrontend.solicitud + ": " + e.message);
      }
    } else {
      console.log("es la accion strinsificada")
    }


    Logger.log("Solicitud " + dataFromFrontend.solicitud + " guardada exitosamente en fila " + rowNum);
    return "Gestión guardada exitosamente.";
  } else {
    Logger.log("Error: No se encontró la solicitud: " + dataFromFrontend.solicitud);
    return "Error: No se encontró la solicitud para guardar.";
  }
}

function appendLeadDataToSheet(leadGestionData, productType) {
  const ssVida = SpreadsheetApp.openById(idVida);
  const ssDesempleo = SpreadsheetApp.openById(idProteccionGastos);
  let sheet;

  if (productType === "Seguro de Vida") {
    sheet = ssVida.getSheetByName(sheetNameVida);
  } else if (productType === "Seguro de Desempleo") {
    sheet = ssDesempleo.getSheetByName(sheetNameDesempleo);
  } else {
    Logger.log("Producto no reconocido: " + productType);
    return;
  }
  if (!sheet) {
    Logger.log("Error: La hoja '" + (productType === "Seguro de Vida" ? sheetNameVida : sheetNameDesempleo) + "' no fue encontrada.");
    return;
  }
  const esVenta = leadGestionData.accionRealizada === "VENTA";
  const esDesistimiento = leadGestionData.accionRealizada === "DESISTIDO";

  let rowData = Array(81).fill("");

  if (productType === "Seguro de Vida") {
    if (esDesistimiento) {

      rowData[6] = new Date()
      rowData[7] = leadGestionData.datosLeadGeneral[6] || ""; // Nombre agente
      rowData[8] = "NO"
      rowData[9] = "NO"
      rowData[76] = "Desistir gestión de Venta.";
      rowData[80] = leadGestionData.solicitud || "";
    } else if (esVenta) {

      const vidaData = JSON.parse(leadGestionData.datosCompletosVida || "{}");
      const datosContacto = vidaData["Contacto Inicial y Datos"] || {};
      const datosEmision = vidaData["Emisión de Póliza"] || {};
      const datosCoti = vidaData["Cotización y Pre-Asegurabilidad"] || {};
      const datosDeclaracion = vidaData["Declaración de Asegurabilidad"] || {};
      const datosAutorizacion = vidaData["Autorización de Descuento Débito"] || {};
      const datosBeneficiarios = vidaData["Beneficiarios y Autorizaciones"] || {};
      const leadDataGral = leadGestionData.datosLeadGeneral;

      rowData = [
        "",
        "",
        "",
        "",
        "",
        "",
        new Date(),
        datosContacto.emailCte || leadDataGral[6] || "",
        "Si",
        "Si",
        datosDeclaracion.actividadesLicita ? 'Si' : 'NO',
        datosDeclaracion.procesoPenal ? 'Si' : 'NO',
        datosDeclaracion.frecuenciaDeportesAltoRiesgo || "",
        datosDeclaracion.peso || "",
        datosDeclaracion.estatura || "",
        datosDeclaracion.imcCalculado || "",
        datosDeclaracion.observacionImc || "",
        datosDeclaracion.estadoAsegurabilidad || "",
        datosDeclaracion.hipertensionArterial ? 'Si' : 'NO',
        datosDeclaracion.colesterolTrigliceridosElevados ? 'Si' : 'NO',
        datosDeclaracion.diabetes ? 'Si' : 'NO',
        datosDeclaracion.enfermedadCoronariaInfarto ? 'Si' : 'NO',
        datosDeclaracion.insuficienciaRenal ? 'Si' : 'NO',
        datosDeclaracion.cancerTumores ? 'Si' : 'NO',
        datosDeclaracion.derrameTrombosisCerebral ? 'Si' : 'NO',
        datosDeclaracion.enfermedadesPsicologicasPsiquiatricas ? 'Si' : 'NO',
        datosDeclaracion.sidaVih ? 'Si' : 'NO',
        datosDeclaracion.pruebaVihPositiva ? 'Si' : 'NO',
        datosDeclaracion.secuelasEnfermedadAccidente === true ? 'Si' : (datosDeclaracion.secuelasEnfermedadAccidente === false ? 'NO' : datosDeclaracion.secuelasEnfermedadAccidente || ''),
        datosDeclaracion.diagnosticoEnfermedadFracturas === true ? 'Si' : (datosDeclaracion.diagnosticoEnfermedadFracturas === false ? 'NO' : datosDeclaracion.diagnosticoEnfermedadFracturas || ''),
        datosDeclaracion.tratamientoAlcoholAlucinogenos ? 'Si' : 'No',
        datosDeclaracion.cirugiaProgramada ? 'Si' : 'No',
        (
          datosDeclaracion.secuelasEnfermedadAccidente === true ||
          (typeof datosDeclaracion.secuelasEnfermedadAccidente === 'string' && datosDeclaracion.secuelasEnfermedadAccidente.trim() !== '') ||
          datosDeclaracion.diagnosticoEnfermedadFracturas === true ||
          (typeof datosDeclaracion.diagnosticoEnfermedadFracturas === 'string' && datosDeclaracion.diagnosticoEnfermedadFracturas.trim() !== '')
        ) ? "Consulta Médica" : datosDeclaracion.autorizacionInformacionMedica ? 'Si' : 'No',
        datosAutorizacion.canalDescuento || "",
        datosAutorizacion.canalDescuento === "Cuenta Ahorros" || datosAutorizacion.canalDescuento === "Cuenta Corriente" ? "Débito" : (datosAutorizacion.canalDescuento === "Tarjeta Credito" ? datosAutorizacion.franquiciaDescuento : ""),
        datosAutorizacion.fechaVencimientoTC || "",
        "1",
        datosAutorizacion.entidadBancaria || "",
        datosAutorizacion.numeroCuentaDebito || "",
        datosAutorizacion.valorPrima || "",
        datosAutorizacion.autorizaDescuentoAutomatico ? 'Si' : 'No',
        datosAutorizacion.entiendeMoraPrima ? 'Si' : 'No',
        datosEmision.fechaVigenciaPoliza || "",
        datosEmision.tipoDocumentoCliente || datosContacto.tipoDocumento || leadDataGral[4] || "",
        datosEmision.numeroDocumentoCliente || leadDataGral[4] || "",
        "5010015245601",
        datosCoti.periosidad || "",
        "1010229278",
        "75272",
        "AUTOMÁTICO",
        datosEmision.primerNombreCliente || (leadDataGral[3] || '').split(' ')[2] || '',
        datosEmision.segundoNombreCliente || (leadDataGral[3] || '').split(' ')[3] || '',
        datosEmision.primerApellidoCliente || (leadDataGral[3] || '').split(' ')[0] || '',
        datosEmision.segundoApellidoCliente || (leadDataGral[3] || '').split(' ')[1] || '',
        datosContacto.genero || "", // Sexo (Columna 55)
        datosEmision.fechaNacimientoCliente || datosContacto.fechaNacimiento || "",
        datosEmision.direccionCliente || leadDataGral[8] || "",
        datosEmision.telefonoCliente || leadDataGral[5] || "",
        datosEmision.celularCliente || leadDataGral[5] || "",
        datosEmision.ciudadCorrespondenciaCliente || leadDataGral[7] || "",
        "",
        "VIDA/ITP",
        datosBeneficiarios.apellidoBeneficiario1 || "",
        datosBeneficiarios.nombreBeneficiario1 || "",
        datosBeneficiarios.designacionBeneficiario1 || "",
        datosBeneficiarios.parentescoBeneficiario1 || "",
        datosBeneficiarios.apellidoBeneficiario2 || "",
        datosBeneficiarios.nombreBeneficiario2 || "",
        "CONOCIMIENTO",
        "CONOCIMIENTO",
        datosBeneficiarios.designacionBeneficiario2 || "",
        datosBeneficiarios.parentescoBeneficiario2 || "",
        datosBeneficiarios.apellidoBeneficiario3 || "",
        datosBeneficiarios.nombreBeneficiario3 || "",
        datosBeneficiarios.designacionBeneficiario3 || "",
        datosBeneficiarios.parentescoBeneficiario3 || "",
        "", // Columna 77 (vacía)
        datosEmision.numeroDocumentoCliente || leadDataGral[4] || "",
        esVenta ? "Expedición" : "",
        datosContacto.nomAgt || "",
        leadGestionData.solicitud || ""
      ];
    }
  } else if (productType === "Seguro de Desempleo") {
    if (esDesistimiento) {
      rowData[5] = new Date();
      rowData[6] = leadGestionData.datosLeadGeneral[6] || "";
      rowData[11] = "Desistir gestión de Venta.";
      rowData[62] = "No";
      rowData[67] = "No";
      rowData[70] = "No";
      rowData[71] = leadGestionData.solicitud || "";
    } else if (esVenta || leadGestionData.accionRealizada === "Confirmacion de venta") {
      const desempleoData = JSON.parse(leadGestionData.datosCompletosDesempleo || "{}");
      const datosContacto = desempleoData["Contacto Inicial y Datos"] || {};
      const datosEmision = desempleoData["Emisión de Póliza"] || {};
      const datosAutorizacion = desempleoData["Autorización de Descuento Débito"] || {};
      const datosCotizacion = desempleoData["Cotización y Pre-Asegurabilidad"] || {};
      const leadDataGral = leadGestionData.datosLeadGeneral;

      const continuarVentaIndependiente = (datosContacto.tipoVinculacion === "Independiente" && (datosContacto.continuarVentaIndependiente === true || datosContacto.continuarVentaIndependiente === 'Si')) ? 'Continuar Ind.' : '';

      let continuarVentaEmpleado = 'No';
      let continuarVentaContrato12 = '';
      let continuarVentaContrato3 = '';
      let continuarVentaContrato4 = '';
      let continuarVentaContrato5 = '';
      let continuarVentaContrato6 = '';

      if (datosContacto.tipoVinculacion === "Empleado") {
        switch (datosContacto.tipoContratoLaboral) {
          case 'Termino Indefinido':
          case 'Carrera Administrativa':
            if (datosContacto.continuarVentaContrato12 === true || datosContacto.continuarVentaContrato12 === 'Si') {
              continuarVentaEmpleado = 'Continuar emp.';
              continuarVentaContrato12 = 'Si';
            }
            break;
          case 'Fijo':
            if (datosContacto.continuarVentaContrato3 === true || datosContacto.continuarVentaContrato3 === 'Si') {
              continuarVentaEmpleado = 'Continuar emp.';
              continuarVentaContrato3 = 'Si';
            }
            break;
          case 'Obra o Labor':
            if (datosContacto.continuarVentaContrato4 === true || datosContacto.continuarVentaContrato4 === 'Si') {
              continuarVentaContrato4 = 'Si';
              continuarVentaEmpleado = 'Continuar emp.';
            }
            break;
          case 'Libre Nombramiento':
            if (datosContacto.continuarVentaContrato5 === true || datosContacto.continuarVentaContrato5 === 'Si') {
              continuarVentaEmpleado = 'Continuar emp.';
              continuarVentaContrato5 = 'Si';
            }
            break;
          case 'Provisionalidad':
            if (datosContacto.continuarVentaContrato6 === true || datosContacto.continuarVentaContrato6 === 'Si') {
              continuarVentaContrato6 = 'Si';
              continuarVentaEmpleado = 'Continuar emp.';
            }
            break;
        }
      }

      const nombreParts = (leadDataGral[3] || '').split(/\s+/).filter(Boolean);
      const primerNombre = datosEmision.primerNombreCliente || nombreParts[0] || '';
      const segundoNombre = datosEmision.segundoNombreCliente || nombreParts[1] || '';
      const primerApellido = datosEmision.primerApellidoCliente || nombreParts[2] || '';
      const segundoApellido = datosEmision.segundoApellidoCliente || nombreParts[3] || '';

      rowData = [
        "", "", "", "", "",
        new Date(),
        datosContacto.emailCte || leadDataGral[6] || "",
        datosContacto.emailCte || leadDataGral[6] || "",
        datosContacto.tipoVinculacion || "",
        datosContacto.tipoContratoLaboral || "",
        datosContacto.nombreEmpleador || "",
        "",
        continuarVentaContrato12,
        continuarVentaContrato3,
        continuarVentaContrato4,
        continuarVentaContrato5,
        continuarVentaContrato6,
        continuarVentaIndependiente,
        (datosContacto.autorizaDatosPersonales === true || datosContacto.autorizaDatosPersonales === 'Si') ? 'Si' : 'No',
        "", "",
        (datosCotizacion.ocupacionLicitaInd === true || datosCotizacion.ocupacionLicitaInd === 'Si') ? 'Si' : 'No',
        "", "",
        'Si',
        (datosCotizacion.autorizaDatosInd === true || datosCotizacion.autorizaDatosInd === 'Si') ? 'Si' : 'No',
        "",
        datosCotizacion.otraOcupacionIndText || (datosCotizacion.otraOcupacionInd === 'Si' ? 'Si' : 'No'),
        "", "", "", "",
        datosAutorizacion.canalDescuento || "",
        datosAutorizacion.fechaVencimientoTC || "",
        datosAutorizacion.numeroCuotasTC || "",
        datosAutorizacion.numeroCuentaDebito || "",
        datosAutorizacion.entidadBancaria || "",
        datosAutorizacion.valorPrima || "",
        (datosAutorizacion.autorizaDescuentoAutomatico === true || datosAutorizacion.autorizaDescuentoAutomatico === 'Si') ? 'Si' : 'No',
        (datosAutorizacion.entiendeMoraPrima === true || datosAutorizacion.entiendeMoraPrima === 'Si') ? 'Si' : 'No',
        "33", "506", "5010001000101", "NT", "900957271", "PE", "1", "DB",
        "1/1/0001", "1/1/0001", "75272", "1", "S", "1",
        leadGestionData.solicitud || "",
        "1/1/0001", "12", "1020749992", "506",
        primerNombre, "0",
        (datosContacto.tipoDocumento || leadDataGral[21] || ""),
        (datosContacto.numeroDocumentoCliente || leadDataGral[4] || ""),
        segundoNombre, primerApellido, segundoApellido,
        datosContacto.genero || "",
        datosContacto.fechaNacimiento || "",
        datosContacto.direccionCliente || leadDataGral[8] || "",
        datosContacto.ciudadCorrespondenciaCliente || leadDataGral[7] || "",
        datosContacto.telefonoCliente || leadDataGral[5] || "",
        datosContacto.celularCliente || leadDataGral[5] || "",
        datosContacto.ciudadResidenciaCliente || leadDataGral[7] || "",
        datosContacto.departamentoCliente || "",
        "1",
        datosContacto.tipoContratoLaboral || "",
        "TP 43",
        continuarVentaEmpleado,
        "1", "4",
        leadDataGral[3],
        (datosContacto.tipoDocumento || leadDataGral[21] || ""),
        (datosContacto.numeroDocumentoCliente || leadDataGral[4] || ""),
        "-30",
        "Expedición",
        "OBSERVACION", "OBSERVACION",
        continuarVentaIndependiente,
      ];
    }
  }

  sheet.appendRow(rowData);
  Logger.log("Datos agregados a " + sheet.getName());
}


function sendInformeVP() {
  let Label = GmailApp.getUserLabelByName("Informe CRM Vida Desempleo VP").getThreads();

}

let destinatarios = {
  datosGerente1: {
    nombre: "cristian Mosquera",
    telefono: "3192650209"
  },
}

const message = `¡Hola ${destinatarios.datosGerente1.nombre} ! 👋\n\nAquí tienes el informe del CRM El Libertador.\n\n¡Que tengas un excelente día! ☀️`

function sendWhatsappVPInteractive() {
  var authToken = PropertiesService.getScriptProperties().getProperty('infobipAuth');

  const myHeaders = {
    "Content-Type": "application/json",
    "Authorization": "Basic " + authToken
  }

  const raw = JSON.stringify({
    "from": "573144352014",
    "to": "573192650209",
    "messageId": "a28dd97c-1ffb-4fcf-99f1-0b557ed381da",
    "content": {
      "body": {
        "text": "¡Hola Cristian! 👋\n\nAquí tienes el informe del CRM El Libertador.\n\n¡Que tengas un excelente día! ☀️"
      },
      "action": {
        "displayText": "Ver Tablero",
        "url": "https://lookerstudio.google.com/reporting/06ff7e20-ce33-4917-b021-216742d32bbd"
      }
    },
    "callbackData": "Callback data",
    "notifyUrl": "https://www.example.com/whatsapp"
  });

  const requestOptions = {
    method: "POST",
    headers: myHeaders,
    payload: raw,
    redirect: "follow"
  };

  let response = UrlFetchApp.fetch("https://qgmx9r.api.infobip.com/whatsapp/1/message/interactive/url-button", requestOptions)
  Logger.log(response)
}



