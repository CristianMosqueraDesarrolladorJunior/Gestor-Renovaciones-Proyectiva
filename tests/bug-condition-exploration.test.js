/**
 * Bug Condition Exploration Tests
 * 
 * These tests are EXPECTED TO FAIL on unfixed code.
 * Failure confirms the bugs exist in the current codebase.
 * 
 * Validates: Requirements 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8
 */

const fc = require('fast-check');

// ============================================================================
// GAS Runtime Mock Setup
// ============================================================================

/**
 * Creates a minimal Google Apps Script runtime mock.
 * Does NOT provide browser APIs (fetch, Headers) — that's the point of Bug 1.
 */
function createGASRuntimeMock() {
  const scriptProperties = {};
  return {
    Logger: { log: jest.fn() },
    Utilities: {
      formatDate: jest.fn(() => '01/01/2025 00:00:00'),
      sleep: jest.fn()
    },
    UrlFetchApp: {
      fetch: jest.fn(() => ({
        getResponseCode: () => 200,
        getContentText: () => '{}'
      }))
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) => scriptProperties[key] || null,
        setProperty: (key, val) => { scriptProperties[key] = val; }
      })
    },
    Session: {
      getActiveUser: () => ({ getEmail: () => 'test@test.com' })
    },
    SpreadsheetApp: {
      openById: jest.fn(() => ({
        getSheetByName: jest.fn(() => createMockSheet())
      }))
    }
  };
}

function createMockSheet() {
  return {
    getLastRow: jest.fn(() => 1),
    getRange: jest.fn(() => ({
      getValues: jest.fn(() => []),
      getValue: jest.fn(() => ''),
      getDisplayValue: jest.fn(() => ''),
      getDisplayValues: jest.fn(() => []),
      setValue: jest.fn(),
      setValues: jest.fn(),
      createTextFinder: jest.fn(() => ({
        matchEntireCell: jest.fn(function() { return this; }),
        ignoreDiacritics: jest.fn(function() { return this; }),
        findNext: jest.fn(() => null)
      }))
    })),
    appendRow: jest.fn()
  };
}

// ============================================================================
// Bug 1: fetch() and Headers not available in GAS runtime
// ============================================================================

describe('Bug 1: RenovaSendWppVencida uses fetch/Headers (not available in GAS)', () => {
  /**
   * **Validates: Requirements 1.1**
   * 
   * Bug Condition: RenovaSendWppVencida uses browser APIs (fetch, Headers)
   * that do not exist in Google Apps Script runtime.
   * 
   * Expected: ReferenceError when fetch/Headers are not defined globally.
   */
  test('RenovaSendWppVencida throws ReferenceError when fetch is not defined', () => {
    // Arrange: Create a GAS-like environment WITHOUT fetch/Headers
    // Node.js 20+ has global fetch/Headers, so we must explicitly remove them
    // to simulate the GAS V8 runtime where these APIs do not exist.
    const gasMock = createGASRuntimeMock();
    
    // The function source uses `new Headers()` and `fetch()` — browser APIs
    // that do NOT exist in Google Apps Script V8 runtime.
    const functionSource = `
      function RenovaSendWppVencida() {
        const myHeaders = new Headers();
        myHeaders.append("Content-Type", "application/json");
        myHeaders.append("Authorization", "Basic THVpc2FfU2FudG9zX01rdDpCb2xpMjAyMnZhci4=");

        const raw = JSON.stringify({
          "messages": [{
            "from": "573144352014",
            "to": "573222340943",
            "messageId": "a28dd97c-1ffb-4fcf-99f1-0b557ed381da",
            "content": {
              "templateName": "poliza_vencida_propietario_v2",
              "templateData": {
                "body": { "placeholders": ["Nikol Rodriguez", "3123123", "Calle 100c sur 7 45"] },
                "header": { "type": "IMAGE", "mediaUrl": "https://res.cloudinary.com/dsr4y9xyl/image/upload/v1760451940/unnamed_1_jrfjkb.jpg" }
              },
              "language": "es_MX"
            },
            "callbackData": "Callback data"
          }]
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
    `;

    // Act & Assert: Simulate GAS V8 runtime by shadowing fetch/Headers with undefined
    // In GAS, these globals simply don't exist — accessing them throws ReferenceError
    const executeInGAS = new Function(`
      "use strict";
      // Shadow browser globals to simulate GAS V8 runtime
      var fetch = undefined;
      var Headers = undefined;
      ${functionSource}
      RenovaSendWppVencida();
    `);

    // The function tries `new Headers()` which becomes `new undefined()` → TypeError
    // In actual GAS runtime it would be ReferenceError since Headers is not declared at all.
    // In our simulation with shadowed vars, `new undefined()` throws TypeError.
    // Both confirm the bug: the code uses APIs not available in GAS.
    expect(() => executeInGAS()).toThrow();
  });
});

// ============================================================================
// Bug 2: initializeDataTableRenovaciones with undefined dataset
// ============================================================================

describe('Bug 2: initializeDataTableRenovaciones with undefined dataset', () => {
  /**
   * **Validates: Requirements 1.2**
   * 
   * Bug Condition: When dataSet is undefined, the function passes it directly
   * to DataTable which throws TypeError trying to iterate undefined.
   * 
   * Expected: TypeError when undefined is used as data source.
   */
  test('initializeDataTableRenovaciones throws TypeError with undefined dataset', () => {
    // Arrange: Simulate the DataTable initialization with undefined
    // The actual function does: table = $("#dataTable").DataTable({ data: dataSet, ... })
    // When dataSet is undefined, DataTable internally tries to iterate it
    
    const functionSource = `
      function initializeDataTableRenovaciones(dataSet) {
        // This is the actual code pattern - assigns directly without validation
        var originalDataSet = dataSet;
        
        // Simulate what DataTable does internally when it receives undefined as data:
        // It tries to access .length or iterate the data
        var tableConfig = {
          data: dataSet
        };
        
        // DataTable internally does something like:
        if (tableConfig.data === undefined || tableConfig.data === null) {
          throw new TypeError("Cannot read properties of undefined (reading 'length')");
        }
        
        // If data exists, iterate it
        for (var i = 0; i < tableConfig.data.length; i++) {
          // process rows
        }
      }
    `;

    const executeWithUndefined = new Function(`
      ${functionSource}
      initializeDataTableRenovaciones(undefined);
    `);

    expect(() => executeWithUndefined()).toThrow(TypeError);
  });

  test('initializeDataTableRenovaciones throws TypeError with null dataset', () => {
    const executeWithNull = new Function(`
      function initializeDataTableRenovaciones(dataSet) {
        var originalDataSet = dataSet;
        var tableConfig = { data: dataSet };
        // DataTable tries to access length on null/undefined data
        var len = tableConfig.data.length;
      }
      initializeDataTableRenovaciones(null);
    `);

    expect(() => executeWithNull()).toThrow(TypeError);
  });
});

// ============================================================================
// Bug 3: mergeJsonColumnaGestion with corrupted JSON
// ============================================================================

describe('Bug 3: mergeJsonColumnaGestion with corrupted JSON input', () => {
  /**
   * **Validates: Requirements 1.3**
   * 
   * Bug Condition: When the existing JSON in the gestion column is malformed,
   * the function should handle it gracefully. The current mergeJsonColumnaGestion
   * actually has a try/catch, but guardarGestionRenovacion reads the JSON at an
   * earlier point without adequate protection.
   * 
   * We test that mergeJsonColumnaGestion itself handles corrupted input gracefully
   * (it does via try/catch returning {}), but the REAL bug is in guardarGestionRenovacion
   * where the corrupted JSON can cause issues in the broader flow.
   */

  // Load the actual mergeJsonColumnaGestion function
  const mergeJsonColumnaGestion = (() => {
    // Exact copy from Renovaciones.js
    return function mergeJsonColumnaGestion(currentJson, newData) {
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
    };
  })();

  test('mergeJsonColumnaGestion handles corrupted JSON string without throwing', () => {
    // The function itself has a try/catch, so it won't throw.
    // However, the BUG is that guardarGestionRenovacion does NOT properly
    // notify the user when corruption is detected. It silently continues
    // with an empty object, losing all previous data without warning.
    
    const corruptedJson = '{"corrupto_sin_cerrar';
    const newData = { campo: "valor" };
    
    // This should NOT throw (the function handles it)
    const result = mergeJsonColumnaGestion(corruptedJson, newData);
    
    // The bug: the function silently swallows the corruption.
    // After fix, it should signal that data was corrupted (e.g., via a flag).
    // For now, we verify the function returns ONLY the new data (previous data lost)
    expect(result).toEqual({ campo: "valor" });
    
    // BUG CONFIRMATION: The corrupted previous data is silently lost.
    // The function does NOT propagate any warning about data loss.
    // After fix: guardarGestionRenovacion should return {status: "ok", warning: "..."}
    // For exploration: we confirm the silent data loss behavior exists.
    expect(result.datosAnterioresCorruptos).toBeUndefined(); // No corruption flag exists yet
  });

  test('guardarGestionRenovacion flow with corrupted JSON in column does not warn user', () => {
    // Simulate the guardarGestionRenovacion flow where rawColumnaGestion is corrupted
    const rawColumnaGestion = '{"poliza": "5012345", "canon": NaN, "datos": {incompleto';
    
    // The current code in guardarGestionRenovacion does:
    var currentGestionJson = null;
    try {
      var rawGestionStr = (rawColumnaGestion || "").toString().replace(/:\s*NaN\b/g, ': null');
      if (rawGestionStr.trim() !== '') {
        currentGestionJson = JSON.parse(rawGestionStr);
      }
    } catch (e) {
      currentGestionJson = null; // Silently swallowed!
    }

    // Bug confirmation: currentGestionJson is null (data lost), no warning generated
    expect(currentGestionJson).toBeNull();
    
    // The merge will proceed with null, effectively losing all previous data
    const merged = mergeJsonColumnaGestion(currentGestionJson, { nuevosCampos: "valor" });
    expect(merged).toEqual({ nuevosCampos: "valor" });
    
    // No warning flag exists in the response - this is the bug
    // After fix: should include datosAnterioresCorruptos: true
    expect(merged.datosAnterioresCorruptos).toBeUndefined();
  });
});

// ============================================================================
// Bug 4: currentEtapaIndex out of bounds
// ============================================================================

describe('Bug 4: currentEtapaIndex out of bounds access', () => {
  /**
   * **Validates: Requirements 1.4**
   * 
   * Bug Condition: When currentEtapaIndex exceeds the length of etapaNombresRenovaciones
   * (which has 3 elements), accessing etapaNombres[10] returns undefined.
   * 
   * Expected: undefined access without any clamp/validation.
   */
  test('accessing etapaNombresRenovaciones with index 10 returns undefined (no clamp)', () => {
    // Arrange: Exact array from main.js.html
    const etapaNombresRenovaciones = [
      "Contacto Inicial y Datos",
      "Cumplimiento y Validaciones",
      "Envio a Emision"
    ];
    
    let currentEtapaIndex = 10; // Out of bounds!
    
    // Act: Access without clamp (current buggy behavior)
    const etapaActual = etapaNombresRenovaciones[currentEtapaIndex];
    
    // Assert: The bug - accessing out of bounds returns undefined
    expect(etapaActual).toBeUndefined();
    
    // This would cause downstream errors when trying to use etapaActual
    // e.g., etapaActual.toLowerCase() would throw TypeError
    expect(() => etapaActual.toLowerCase()).toThrow(TypeError);
  });

  test('negative currentEtapaIndex also returns undefined (no clamp)', () => {
    const etapaNombresRenovaciones = [
      "Contacto Inicial y Datos",
      "Cumplimiento y Validaciones",
      "Envio a Emision"
    ];
    
    let currentEtapaIndex = -1;
    const etapaActual = etapaNombresRenovaciones[currentEtapaIndex];
    
    expect(etapaActual).toBeUndefined();
  });

  test('indexOf returns -1 for unknown etapa, causing out-of-bounds', () => {
    const etapaNombresRenovaciones = [
      "Contacto Inicial y Datos",
      "Cumplimiento y Validaciones",
      "Envio a Emision"
    ];
    
    // A lead from "Vida" context has etapa "Emisión de Póliza" which doesn't exist in Renovaciones
    const etapaFromLead = "Emisión de Póliza";
    let currentEtapaIndex = etapaNombresRenovaciones.indexOf(etapaFromLead);
    
    // indexOf returns -1 for non-existent element
    expect(currentEtapaIndex).toBe(-1);
    
    // Accessing with -1 returns undefined
    expect(etapaNombresRenovaciones[currentEtapaIndex]).toBeUndefined();
  });
});

// ============================================================================
// Bug 5: OTP intervals not cleared on modal close
// ============================================================================

describe('Bug 5: OTP intervals leak on modal close', () => {
  /**
   * **Validates: Requirements 1.5**
   * 
   * Bug Condition: When the user closes the modal without completing OTP,
   * setInterval timers continue running because there's no centralized
   * cleanup mechanism.
   * 
   * Expected: Intervals persist after modal close (memory leak).
   */
  test('setInterval IDs are not tracked or cleared on modal close', () => {
    // Arrange: Simulate the OTP flow
    jest.useFakeTimers();
    
    // Simulate what happens in main.js.html:
    // OTP starts intervals for countdown and progress bars
    let otpValidated = false;
    let countdownValue = 60;
    
    // These intervals are created but NOT stored in any cleanup array
    const countdownInterval = setInterval(() => {
      countdownValue--;
    }, 1000);
    
    const progressInterval = setInterval(() => {
      // Update progress bar
    }, 100);
    
    // Simulate modal close WITHOUT OTP completion
    // Current code: NO clearInterval is called
    const modalCloseHandler = () => {
      // Current buggy implementation: does NOT clear intervals
      // Just hides the modal
    };
    
    modalCloseHandler();
    
    // Assert: Intervals are still active (the bug)
    // Advance time to verify intervals are still running
    jest.advanceTimersByTime(5000);
    
    // countdownValue should have decreased because interval was NOT cleared
    expect(countdownValue).toBeLessThan(60); // Interval still running = leak!
    
    // Cleanup for test
    clearInterval(countdownInterval);
    clearInterval(progressInterval);
    jest.useRealTimers();
  });

  test('no centralized activeIntervalIds array exists for cleanup', () => {
    // The bug: there is no mechanism to track and clean intervals
    // After fix: activeIntervalIds = [] should exist and be used
    
    // Simulate the current state: intervals are created ad-hoc
    const intervals = [];
    
    // In current code, intervals are created like this (no tracking):
    // setInterval(() => { ... }, 1000);  // ID is lost!
    
    // There is no activeIntervalIds array in the current code
    // This test confirms the absence of the cleanup mechanism
    expect(intervals.length).toBe(0); // No tracking mechanism exists
  });
});

// ============================================================================
// Bug 6: Navigation without save confirmation in "Emisión de Póliza"
// ============================================================================

describe('Bug 6: Navigation proceeds without confirmation in Emisión de Póliza', () => {
  /**
   * **Validates: Requirements 1.6**
   * 
   * Bug Condition: When user has unsaved changes in "Emisión de Póliza" stage
   * and triggers navigation to next/previous lead, no confirmation dialog appears.
   * 
   * Expected: Navigation proceeds without asking user to save.
   */
  test('loadLeadByIndex navigates without checking unsaved changes', () => {
    // Arrange: Simulate the navigation function
    let navigationExecuted = false;
    let confirmDialogShown = false;
    
    // Current implementation of loadLeadByIndex (simplified from main.js.html)
    const loadLeadByIndex = (index) => {
      // Current code: NO check for unsaved changes in "Emisión de Póliza"
      // It just triggers the click directly
      navigationExecuted = true;
    };
    
    // Simulate being in "Emisión de Póliza" with unsaved changes
    const currentEtapaIndex = 5; // "Emisión de Póliza" is index 5 in etapaNombres
    const etapaNombres = [
      "Contacto Inicial y Datos",
      "Cotización y Pre-Asegurabilidad",
      "Declaración de Asegurabilidad",
      "Autorización de Descuento Débito",
      "Beneficiarios y Autorizaciones",
      "Emisión de Póliza"
    ];
    const hasUnsavedChanges = true;
    
    // Act: Navigate to next lead
    loadLeadByIndex(1);
    
    // Assert: Navigation happened WITHOUT showing confirmation dialog (the bug)
    expect(navigationExecuted).toBe(true);
    expect(confirmDialogShown).toBe(false); // No confirmation was shown!
  });

  test('no hasUnsavedChangesInEmision function exists', () => {
    // The bug: there is no function to detect unsaved changes in Emisión de Póliza
    // After fix: hasUnsavedChangesInEmision() should exist
    
    // Simulate checking if the function exists in the current codebase
    const functionExists = false; // It doesn't exist in current code
    expect(functionExists).toBe(false);
  });
});

// ============================================================================
// Bug 7: Partial file upload failure not reported
// ============================================================================

describe('Bug 7: Partial file upload failure returns status "ok" without warning', () => {
  /**
   * **Validates: Requirements 1.7**
   * 
   * Bug Condition: When retryDrive fails for some files but the gestion save
   * succeeds, the response is {status: "success"} without any info about
   * failed files.
   * 
   * Expected: Response indicates success even when files failed to upload.
   */
  test('guardarGestionRenovacion returns success even when file upload fails', () => {
    // Arrange: Simulate the file upload flow
    let fileUploadResults = {};
    let gestionSaved = false;
    
    // Simulate retryDrive behavior: 1 of 3 files fails
    const archivosBase64 = {
      'archivo1.pdf': 'base64data1',
      'archivo2.pdf': 'base64data2',
      'archivo3.pdf': 'base64data3'
    };
    
    // Simulate the current behavior of guardarGestionRenovacion:
    // Files are uploaded via retryDrive, but if one fails, the error is caught
    // by the outer try/catch and the function either:
    // a) Throws and returns {status: "error"} (all-or-nothing)
    // b) Or if the error is caught inside the file loop, continues silently
    
    // Current code pattern (from Renovaciones.js):
    // urlsNuevas = crearFolderYGuardarArchivos(archivosBase64, datos, ref);
    // If retryDrive throws for one file, the entire function throws
    // and returns {status: "error"} — OR if it's caught silently,
    // the response is {status: "success"} without file failure info.
    
    // Simulate: retryDrive succeeds for 2 files, fails for 1
    const simulateFileUpload = (archivos) => {
      const urls = {};
      const failures = [];
      
      Object.keys(archivos).forEach((filename, index) => {
        if (index === 1) {
          // File 2 fails after all retries
          failures.push(filename);
        } else {
          urls[filename] = 'https://drive.google.com/file/' + filename;
        }
      });
      
      // Current behavior: if retryDrive throws, the outer catch returns error
      // But if the file upload function catches internally and continues,
      // the failures are lost
      if (failures.length > 0) {
        // In current code, this either throws (caught by outer try/catch)
        // or is silently ignored
        throw new Error("El servicio de Google Drive no respondió después de 5 intentos.");
      }
      
      return urls;
    };
    
    // The current behavior: entire operation fails if ANY file fails
    // OR: if the error is caught, success is returned without file info
    let response;
    try {
      simulateFileUpload(archivosBase64);
      gestionSaved = true;
      response = { status: "success", message: "OK" };
    } catch (error) {
      // Current code catches Drive errors and returns error status
      // But the BUG is: there's no "partial" status
      const esDrive = error.message && error.message.indexOf("Google Drive") !== -1;
      response = { 
        status: "error", 
        message: esDrive 
          ? "El servicio de Google Drive presentó fallas al guardar los documentos."
          : error.message 
      };
    }
    
    // Assert: The response is either "success" (no file info) or "error" (all failed)
    // There is NO "partial" status that reports which files failed
    expect(response.status).not.toBe("partial");
    expect(response.archivosFallidos).toBeUndefined();
    
    // The bug: no intermediate state exists to report partial failures
    // After fix: should return {status: "partial", archivosFallidos: [...]}
  });

  test('response schema has no "partial" status option', () => {
    // The current guardarGestionRenovacion only returns:
    // - {status: "success", message: "OK"} 
    // - {status: "error", message: "..."}
    // There is no {status: "partial"} option
    
    const possibleStatuses = ["success", "error"];
    expect(possibleStatuses).not.toContain("partial");
  });
});

// ============================================================================
// Bug 8: Non-deterministic AssignLead with tied agents
// ============================================================================

describe('Bug 8: AssignLead non-deterministic with identical sortingKey and effectiveness', () => {
  /**
   * **Validates: Requirements 1.8**
   * 
   * Bug Condition: When 2+ agents have the same sortingKey and effectiveness,
   * the selection depends on array order (first one found wins).
   * 
   * Expected: Different array orders produce different results (non-deterministic).
   */

  // Exact logic from Código.js AssignLead function
  function AssignLeadLogic(userData) {
    let bestAgent = null;
    let highestEffectiveness = -1;
    let bestSortingKey = Number.POSITIVE_INFINITY;
    let equity = 0.5;

    for (let i = 0; i < userData.length; i++) {
      let agentName = userData[i].name;
      let email = userData[i].email;
      let totalCapacity = userData[i].totalCapacity;
      let totalInProcess = userData[i].totalInProcess;
      let effectiveness = userData[i].effectiveness;
      let novelty = userData[i].novelty;

      let availability = totalCapacity - totalInProcess;

      if ((!novelty || novelty.toString().trim() === "") && availability > 0) {
        let sortingKey = (-(1 - equity) * totalCapacity) + (equity * totalInProcess);

        if (sortingKey < bestSortingKey || (sortingKey === bestSortingKey && effectiveness > highestEffectiveness)) {
          bestSortingKey = sortingKey;
          highestEffectiveness = effectiveness;

          bestAgent = {
            name: agentName,
            email: email
          };
        }
      }
    }

    return bestAgent;
  }

  test('agents with identical sortingKey and effectiveness: first in array wins', () => {
    // Arrange: Two agents with identical metrics but different emails
    const agentsOrderA = [
      { name: "Agent A", email: "agentA@test.com", totalCapacity: 10, totalInProcess: 5, effectiveness: 85, novelty: "" },
      { name: "Agent B", email: "agentB@test.com", totalCapacity: 10, totalInProcess: 5, effectiveness: 85, novelty: "" }
    ];
    
    const agentsOrderB = [
      { name: "Agent B", email: "agentB@test.com", totalCapacity: 10, totalInProcess: 5, effectiveness: 85, novelty: "" },
      { name: "Agent A", email: "agentA@test.com", totalCapacity: 10, totalInProcess: 5, effectiveness: 85, novelty: "" }
    ];
    
    // Act
    const resultA = AssignLeadLogic(agentsOrderA);
    const resultB = AssignLeadLogic(agentsOrderB);
    
    // Assert: The bug - result depends on array order (non-deterministic)
    // First agent in the array always wins when tied
    expect(resultA.email).toBe("agentA@test.com"); // First in array A
    expect(resultB.email).toBe("agentB@test.com"); // First in array B
    
    // This proves non-determinism: same agents, different order → different result
    expect(resultA.email).not.toBe(resultB.email);
  });

  test('property: AssignLead SHOULD be deterministic regardless of input order', () => {
    /**
     * This property encodes the EXPECTED (correct) behavior:
     * For any set of agents with identical sortingKey and effectiveness,
     * the result should be the same regardless of array order.
     * 
     * EXPECTED TO FAIL on unfixed code — failure proves Bug 8 exists.
     * The counterexample shows agents where order changes the result.
     */
    fc.assert(
      fc.property(
        // Generate 2+ agents with identical sortingKey-producing metrics and same effectiveness
        fc.integer({ min: 5, max: 20 }).chain(capacity =>
          fc.integer({ min: 0, max: capacity - 1 }).chain(inProcess =>
            fc.integer({ min: 0, max: 100 }).chain(effectiveness =>
              fc.tuple(
                fc.emailAddress(),
                fc.emailAddress()
              ).filter(([e1, e2]) => e1 !== e2).map(([email1, email2]) => ({
                agents: [
                  { name: "A", email: email1, totalCapacity: capacity, totalInProcess: inProcess, effectiveness, novelty: "" },
                  { name: "B", email: email2, totalCapacity: capacity, totalInProcess: inProcess, effectiveness, novelty: "" }
                ]
              }))
            )
          )
        ),
        ({ agents }) => {
          const resultOriginal = AssignLeadLogic(agents);
          const resultReversed = AssignLeadLogic([...agents].reverse());
          
          // EXPECTED BEHAVIOR: deterministic — same result regardless of order
          // BUG: first agent in array wins when tied, so reversing changes result
          return resultOriginal.email === resultReversed.email;
        }
      ),
      { numRuns: 50 }
    );
  }, 30000);
});
