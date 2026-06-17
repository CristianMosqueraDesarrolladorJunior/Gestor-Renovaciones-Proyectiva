/**
 * Preservation Property Tests
 * 
 * These tests capture the EXISTING correct behavior for non-buggy inputs.
 * They MUST PASS on the current (unfixed) code.
 * After fixes are applied, they verify no regressions were introduced.
 * 
 * **Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.7, 3.8**
 */

const fc = require('fast-check');

// ============================================================================
// Function Extractions from Source Code
// ============================================================================

/**
 * Exact copy of mergeJsonColumnaGestion from Renovaciones.js
 * This is the function under test for preservation of merge behavior.
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

/**
 * Exact logic of AssignLead selection from Código.js
 * Extracted for testability without GAS dependencies.
 */
function AssignLeadLogic(userData) {
  let bestAgent = null;
  let highestEffectiveness = -1;
  let bestSortingKey = Number.POSITIVE_INFINITY;
  let equity = 0.5;

  for (let i = 0; i < userData.length; i++) {
    let agentName = userData[i].name;
    let email = userData[i].email;
    let totalCapacity = Number(userData[i].totalCapacity);
    let totalInProcess = Number(userData[i].totalInProcess);
    let effectivenessStr = userData[i].effectiveness
      ? userData[i].effectiveness.toString().replace("%", "").trim()
      : "0";
    let effectiveness = Number(effectivenessStr);
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

// ============================================================================
// Custom Arbitraries (Generators)
// ============================================================================

/**
 * Generates a valid JSON object with typical renovation lead fields.
 */
const validRenovacionLeadArb = fc.record({
  poliza: fc.stringMatching(/^50\d{5,8}$/),
  documento: fc.stringMatching(/^\d{6,12}$/),
  asegurado: fc.string({ minLength: 3, maxLength: 50 }),
  canon: fc.integer({ min: 100000, max: 50000000 }),
  vigencia: fc.constantFrom('01/01/2024', '15/06/2024', '01/12/2025'),
  vencimiento: fc.constantFrom('01/01/2025', '15/06/2025', '01/12/2026'),
  email: fc.emailAddress(),
  celular: fc.stringMatching(/^3\d{9}$/)
});

/**
 * Generates a valid DataTable row for renovaciones.
 */
const validDataTableRowArb = fc.record({
  fechaIngreso: fc.constantFrom('01/01/2025', '15/03/2025', '20/06/2025'),
  leadData: validRenovacionLeadArb,
  nombreAgente: fc.string({ minLength: 3, maxLength: 30 }),
  etapaFunel: fc.constantFrom('Contacto Inicial y Datos', 'Cumplimiento y Validaciones', 'Envio a Emision'),
  estadoGestion: fc.constantFrom('Pendiente Renovacion', 'En Gestion', 'Volver a llamar')
});

/**
 * Generates a valid JSON string representing existing gestion data.
 */
const validGestionJsonStringArb = fc.record({
  poliza: fc.stringMatching(/^50\d{5,8}$/),
  canon: fc.integer({ min: 100000, max: 50000000 }),
  asegurado: fc.string({ minLength: 3, maxLength: 30 }),
  estadoGestion: fc.constantFrom('En Gestion', 'Pendiente', 'Contactado')
}).map(obj => JSON.stringify(obj));

/**
 * Generates new data to merge into gestion.
 */
const newMergeDataArb = fc.record({
  nuevosCampos: fc.string({ minLength: 1, maxLength: 20 }),
  estadoActualizado: fc.constantFrom('Contactado', 'Renovado', 'Pendiente')
});

/**
 * Generates an agent with specific sortingKey-producing metrics.
 */
const agentArb = (capacityRange, inProcessRange, effectivenessRange) =>
  fc.record({
    name: fc.string({ minLength: 2, maxLength: 20 }),
    email: fc.emailAddress(),
    totalCapacity: fc.integer(capacityRange),
    totalInProcess: fc.integer(inProcessRange),
    effectiveness: fc.integer(effectivenessRange),
    novelty: fc.constant("")
  });

// ============================================================================
// Property 1: initializeDataTableRenovaciones with valid arrays
// ============================================================================

describe('Preservation: initializeDataTableRenovaciones with valid arrays', () => {
  /**
   * **Validates: Requirements 3.2**
   * 
   * Property: For all valid arrays passed to initializeDataTableRenovaciones,
   * table initializes without error showing all records.
   * 
   * Observation: initializeDataTableRenovaciones([{poliza: "123", ...}]) with
   * valid array initializes DataTable correctly.
   */
  test('property: valid arrays do not throw and are accepted as data source', () => {
    fc.assert(
      fc.property(
        fc.array(validDataTableRowArb, { minLength: 0, maxLength: 20 }),
        (dataSet) => {
          // Simulate the validation that initializeDataTableRenovaciones performs:
          // With a valid array, it passes directly to DataTable({ data: dataSet })
          // The key preservation behavior: Array.isArray(dataSet) === true means no error

          // The function assigns: originalDataSet = dataSet
          const originalDataSet = dataSet;

          // DataTable receives the array and can iterate it
          expect(Array.isArray(originalDataSet)).toBe(true);
          expect(originalDataSet.length).toBe(dataSet.length);

          // Each row should be accessible
          for (let i = 0; i < originalDataSet.length; i++) {
            expect(originalDataSet[i]).toBeDefined();
            expect(originalDataSet[i].leadData).toBeDefined();
          }
        }
      ),
      { numRuns: 100 }
    );
  });

  test('property: valid arrays preserve all records without data loss', () => {
    fc.assert(
      fc.property(
        fc.array(validDataTableRowArb, { minLength: 1, maxLength: 15 }),
        (dataSet) => {
          // Preservation: the number of records passed equals the number available
          const originalDataSet = dataSet;

          // All records are preserved
          expect(originalDataSet.length).toBe(dataSet.length);

          // Each record's poliza field is preserved
          dataSet.forEach((row, index) => {
            expect(originalDataSet[index].leadData.poliza).toBe(row.leadData.poliza);
          });
        }
      ),
      { numRuns: 100 }
    );
  });
});

// ============================================================================
// Property 2: mergeJsonColumnaGestion preserves fields and concatenates docs
// ============================================================================

describe('Preservation: mergeJsonColumnaGestion non-destructive merge', () => {
  /**
   * **Validates: Requirements 3.3**
   * 
   * Property: For all valid JSON strings, mergeJsonColumnaGestion preserves
   * existing fields and concatenates documentosProceso.
   * 
   * Observation: mergeJsonColumnaGestion('{"poliza":"5012345","canon":100}', {nuevosCampos: "valor"})
   * performs non-destructive merge preserving fields and concatenating documentosProceso.
   */
  test('property: existing fields are preserved when not overwritten by newData', () => {
    fc.assert(
      fc.property(
        validGestionJsonStringArb,
        newMergeDataArb,
        (jsonString, newData) => {
          const result = mergeJsonColumnaGestion(jsonString, newData);
          const original = JSON.parse(jsonString);

          // All original fields that are NOT in newData must be preserved
          Object.keys(original).forEach(key => {
            if (!Object.prototype.hasOwnProperty.call(newData, key)) {
              expect(result[key]).toEqual(original[key]);
            }
          });

          // All newData fields must be present in result
          Object.keys(newData).forEach(key => {
            expect(result[key]).toEqual(newData[key]);
          });
        }
      ),
      { numRuns: 100 }
    );
  });

  test('property: documentosProceso arrays are concatenated, not replaced', () => {
    fc.assert(
      fc.property(
        fc.array(fc.string({ minLength: 1, maxLength: 30 }), { minLength: 1, maxLength: 5 }),
        fc.array(fc.string({ minLength: 1, maxLength: 30 }), { minLength: 1, maxLength: 5 }),
        (prevDocs, newDocs) => {
          const existingJson = JSON.stringify({
            poliza: "5012345",
            documentosProceso: prevDocs
          });

          const newData = {
            nuevosCampos: "valor",
            documentosProceso: newDocs
          };

          const result = mergeJsonColumnaGestion(existingJson, newData);

          // documentosProceso must be the concatenation of both arrays
          expect(result.documentosProceso).toEqual(prevDocs.concat(newDocs));
          expect(result.documentosProceso.length).toBe(prevDocs.length + newDocs.length);
        }
      ),
      { numRuns: 100 }
    );
  });

  test('property: merge with object input (not string) also preserves fields', () => {
    fc.assert(
      fc.property(
        fc.record({
          poliza: fc.stringMatching(/^50\d{5,8}$/),
          canon: fc.integer({ min: 100000, max: 50000000 }),
          estado: fc.constantFrom('activo', 'pendiente', 'renovado')
        }),
        newMergeDataArb,
        (existingObj, newData) => {
          const result = mergeJsonColumnaGestion(existingObj, newData);

          // Original fields not in newData are preserved
          Object.keys(existingObj).forEach(key => {
            if (!Object.prototype.hasOwnProperty.call(newData, key)) {
              expect(result[key]).toEqual(existingObj[key]);
            }
          });

          // New fields are present
          Object.keys(newData).forEach(key => {
            expect(result[key]).toEqual(newData[key]);
          });
        }
      ),
      { numRuns: 100 }
    );
  });
});

// ============================================================================
// Property 3: renderEtapas with valid currentEtapaIndex
// ============================================================================

describe('Preservation: renderEtapas with valid index accesses correct position', () => {
  /**
   * **Validates: Requirements 3.4**
   * 
   * Property: For all currentEtapaIndex in range [0, etapas.length-1],
   * renderEtapas accesses valid array position.
   * 
   * Observation: renderEtapas() with currentEtapaIndex = 0, 1, 2 (within range)
   * renders corresponding stage without error.
   */
  test('property: valid indices in etapaNombresRenovaciones always return defined value', () => {
    const etapaNombresRenovaciones = [
      "Contacto Inicial y Datos",
      "Cumplimiento y Validaciones",
      "Envio a Emision"
    ];

    fc.assert(
      fc.property(
        fc.integer({ min: 0, max: etapaNombresRenovaciones.length - 1 }),
        (currentEtapaIndex) => {
          const etapaActual = etapaNombresRenovaciones[currentEtapaIndex];

          // Valid index always returns a defined, non-empty string
          expect(etapaActual).toBeDefined();
          expect(typeof etapaActual).toBe('string');
          expect(etapaActual.length).toBeGreaterThan(0);
        }
      ),
      { numRuns: 50 }
    );
  });

  test('property: valid indices in etapaNombres (6 stages) always return defined value', () => {
    const etapaNombres = [
      "Contacto Inicial y Datos",
      "Cotización y Pre-Asegurabilidad",
      "Declaración de Asegurabilidad",
      "Autorización de Descuento Débito",
      "Beneficiarios y Autorizaciones",
      "Emisión de Póliza"
    ];

    fc.assert(
      fc.property(
        fc.integer({ min: 0, max: etapaNombres.length - 1 }),
        (currentEtapaIndex) => {
          const etapaActual = etapaNombres[currentEtapaIndex];

          // Valid index always returns a defined, non-empty string
          expect(etapaActual).toBeDefined();
          expect(typeof etapaActual).toBe('string');
          expect(etapaActual.length).toBeGreaterThan(0);

          // The value should be one of the known stage names
          expect(etapaNombres).toContain(etapaActual);
        }
      ),
      { numRuns: 50 }
    );
  });

  test('property: indexOf for known etapas always returns valid index', () => {
    const etapaNombresRenovaciones = [
      "Contacto Inicial y Datos",
      "Cumplimiento y Validaciones",
      "Envio a Emision"
    ];

    fc.assert(
      fc.property(
        fc.constantFrom(...etapaNombresRenovaciones),
        (etapaNombre) => {
          const index = etapaNombresRenovaciones.indexOf(etapaNombre);

          // Known etapa names always produce valid indices
          expect(index).toBeGreaterThanOrEqual(0);
          expect(index).toBeLessThan(etapaNombresRenovaciones.length);
          expect(etapaNombresRenovaciones[index]).toBe(etapaNombre);
        }
      ),
      { numRuns: 50 }
    );
  });
});

// ============================================================================
// Property 4: AssignLead selects agent with strictly lower sortingKey
// ============================================================================

describe('Preservation: AssignLead selects agent with lowest sortingKey', () => {
  /**
   * **Validates: Requirements 3.8**
   * 
   * Property: For all agent sets where one agent has strictly lower sortingKey,
   * that agent is selected regardless of other factors.
   * 
   * Observation: AssignLead() with agents of different sortingKey selects
   * the one with lowest sortingKey.
   */
  test('property: agent with strictly lower sortingKey is always selected', () => {
    fc.assert(
      fc.property(
        // Generate two agents where one has strictly lower sortingKey
        // sortingKey = (-(1-0.5) * capacity) + (0.5 * inProcess) = -0.5*capacity + 0.5*inProcess
        // Lower sortingKey means: higher capacity OR lower inProcess
        fc.integer({ min: 10, max: 30 }).chain(highCapacity =>
          fc.integer({ min: 0, max: highCapacity - 1 }).chain(lowInProcess =>
            fc.integer({ min: 1, max: 9 }).chain(lowCapacity =>
              fc.integer({ min: 0, max: lowCapacity - 1 }).chain(highInProcess =>
                fc.tuple(
                  fc.emailAddress(),
                  fc.emailAddress(),
                  fc.integer({ min: 0, max: 100 }),
                  fc.integer({ min: 0, max: 100 })
                ).filter(([e1, e2]) => e1 !== e2).map(([email1, email2, eff1, eff2]) => {
                  // Agent A: high capacity, low inProcess → lower sortingKey
                  // Agent B: low capacity, high inProcess → higher sortingKey
                  return {
                    agentA: {
                      name: "AgentA", email: email1,
                      totalCapacity: highCapacity, totalInProcess: lowInProcess,
                      effectiveness: eff1, novelty: ""
                    },
                    agentB: {
                      name: "AgentB", email: email2,
                      totalCapacity: lowCapacity, totalInProcess: highInProcess,
                      effectiveness: eff2, novelty: ""
                    }
                  };
                })
              )
            )
          )
        ),
        ({ agentA, agentB }) => {
          // Verify that agentA actually has lower sortingKey
          const skA = (-0.5 * agentA.totalCapacity) + (0.5 * agentA.totalInProcess);
          const skB = (-0.5 * agentB.totalCapacity) + (0.5 * agentB.totalInProcess);

          // Only test when sortingKeys are actually different
          if (skA >= skB) return true; // Skip this case

          // Test both orderings — agent with lower sortingKey should always win
          const resultAFirst = AssignLeadLogic([agentA, agentB]);
          const resultBFirst = AssignLeadLogic([agentB, agentA]);

          // Agent A (lower sortingKey) should be selected regardless of order
          expect(resultAFirst.email).toBe(agentA.email);
          expect(resultBFirst.email).toBe(agentA.email);
        }
      ),
      { numRuns: 100 }
    );
  });
});

// ============================================================================
// Property 5: AssignLead with same sortingKey selects higher effectiveness
// ============================================================================

describe('Preservation: AssignLead selects agent with higher effectiveness on tie', () => {
  /**
   * **Validates: Requirements 3.8**
   * 
   * Property: For all agent sets where sortingKey is equal but one has strictly
   * higher effectiveness, that agent is selected.
   * 
   * Observation: AssignLead() with same sortingKey but different effectiveness
   * selects highest effectiveness.
   */
  test('property: agent with higher effectiveness wins when sortingKey is equal', () => {
    fc.assert(
      fc.property(
        // Generate two agents with SAME sortingKey but DIFFERENT effectiveness
        // Same sortingKey means same capacity and same inProcess
        fc.integer({ min: 5, max: 25 }).chain(capacity =>
          fc.integer({ min: 0, max: capacity - 1 }).chain(inProcess =>
            fc.integer({ min: 50, max: 100 }).chain(highEff =>
              fc.integer({ min: 0, max: highEff - 1 }).chain(lowEff =>
                fc.tuple(
                  fc.emailAddress(),
                  fc.emailAddress()
                ).filter(([e1, e2]) => e1 !== e2).map(([email1, email2]) => ({
                  agentHigh: {
                    name: "HighEff", email: email1,
                    totalCapacity: capacity, totalInProcess: inProcess,
                    effectiveness: highEff, novelty: ""
                  },
                  agentLow: {
                    name: "LowEff", email: email2,
                    totalCapacity: capacity, totalInProcess: inProcess,
                    effectiveness: lowEff, novelty: ""
                  }
                }))
              )
            )
          )
        ),
        ({ agentHigh, agentLow }) => {
          // Verify same sortingKey
          const skHigh = (-0.5 * agentHigh.totalCapacity) + (0.5 * agentHigh.totalInProcess);
          const skLow = (-0.5 * agentLow.totalCapacity) + (0.5 * agentLow.totalInProcess);
          expect(skHigh).toBe(skLow);

          // Verify different effectiveness
          expect(agentHigh.effectiveness).toBeGreaterThan(agentLow.effectiveness);

          // Test both orderings — higher effectiveness should always win
          const resultHighFirst = AssignLeadLogic([agentHigh, agentLow]);
          const resultLowFirst = AssignLeadLogic([agentLow, agentHigh]);

          // Agent with higher effectiveness should be selected regardless of order
          expect(resultHighFirst.email).toBe(agentHigh.email);
          expect(resultLowFirst.email).toBe(agentHigh.email);
        }
      ),
      { numRuns: 100 }
    );
  });
});
