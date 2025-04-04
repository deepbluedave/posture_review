/**
 * Posture Summary Script (v17 - Enhanced Row Generation Debugging)
 *
 * Reads application posture data from various sheets defined in a 'Config' sheet,
 * aggregates the data based on specified methods (List, Count, Sum, Average, Min, Max, UniqueList),
 * pulls specified fields from the master application list,
 * and writes a summary report to a 'Posture Summary' sheet.
 *
 * ==============================================================================
 * == Maintainer Notes for Office Scripts Compatibility ==
 * ==============================================================================
 * Please adhere to the following guidelines when modifying this script to ensure
 * compatibility with the Office Scripts runtime environment:
 *
 * 1.  **Arrow Functions:** Avoid complex or deeply nested arrow functions, especially
 *     within array methods like `.map()`, `.filter()`, `.forEach()`, `.reduce()`.
 *     - Simple, top-level arrow functions are often okay.
 *     - Inside loops or other callbacks, prefer standard `for...of` or `for` loops
 *       over array methods with arrow function callbacks if issues arise.
 *     - See: https://learn.microsoft.com/en-us/office/dev/scripts/develop/typescript-restrictions#arrow-functions-and-this-keyword
 *
 * 2.  **Type Inference:** Office Scripts requires explicit type declarations more often
 *     than standard TypeScript.
 *     - Always declare types for variables, function parameters, and return types
 *       where inference might be ambiguous (e.g., callback parameters, complex objects).
 *     - Do not rely heavily on type inference, especially within loops or callbacks.
 *
 * 3.  **Console Logging:** Use *only* `console.log()`.
 *     - `console.warn()`, `console.error()`, `console.table()`, etc., are NOT supported
 *       and may cause runtime errors.
 *     - Use a `DEBUG_MODE` flag (as implemented below) to control log verbosity
 *       instead of different console methods.
 *
 * 4.  **`this` Keyword:** Be cautious with the `this` keyword, especially inside
 *     callbacks or nested functions. Arrow functions preserve `this` from the
 *     enclosing scope, while standard functions may not behave as expected in
 *     some contexts within Office Scripts. Sticking to explicit parameters is safer.
 *
 * 5.  **External Libraries & DOM:** You cannot use external JavaScript libraries or
 *     access the HTML DOM. All logic must use built-in JavaScript features and
 *     the Office Scripts ExcelScript APIs.
 *
 * 6.  **Asynchronous Operations:** Use `async`/`await` correctly for all API calls
 *     that interact with the workbook content (e.g., `getValues()`, `setValues()`,
 *     `getRange()`, formatting calls like `autofitColumns()`). Synchronous API calls
 *     (e.g., `getWorksheet()`, `getTable()`, math operations) do not need `await`.
 *
 * By following these guidelines, we can minimize compatibility issues when editing
 * or extending this script.
 * ==============================================================================
 *
 * Key changes in v17:
 * - Added detailed logging INSIDE the posture data column generation loop to
 *   track row length before/after each column push for specific App IDs. Helps
 *   diagnose column count mismatches.
 * - Previous v16 fixes (Office Scripts Compatibility, Maintainer Notes) maintained.
 */
async function main(workbook: ExcelScript.Workbook) {
  // --- Control Script Behavior ---
  const DEBUG_MODE: boolean = true; // Set to false to reduce console output

  if (DEBUG_MODE) console.log("Starting posture summary script (v17: Enhanced Row Debugging)...");
  const startTime = Date.now();

  // --- Overall Constants ---
  const MASTER_APP_SHEET_NAME: string = "Applications";
  const MASTER_APP_ID_HEADER: string = "Application ID";
  const SUMMARY_SHEET_NAME: string = "Posture Summary";
  const CONFIG_SHEET_NAME: string = "Config";
  const CONFIG_TABLE_NAME: string = "ConfigTable";
  const DEFAULT_VALUE_MISSING: string = "";
  const CONCATENATE_SEPARATOR: string = " - ";
  const COUNT_SEPARATOR: string = ": ";

  // --- Type Definitions ---
  type AggregationMethod = "List" | "Count" | "Sum" | "Average" | "Min" | "Max" | "UniqueList";

  type PostureSheetConfig = {
      isEnabled: boolean; sheetName: string; appIdHeaders: string[];
      dataHeadersToPull: string[]; aggregationType: AggregationMethod;
      valueHeaderForAggregation?: string; masterFieldsForRow?: string[];
  };

  type PostureDataObject = { [appId: string]: { [header: string]: (string | number | boolean)[] } };
  type MasterAppData = { [fieldName: string]: string | number | boolean; }
  type MasterAppDataMap = { [appId: string]: MasterAppData; }

  // --- Helper Functions (No changes needed in helpers for this issue) ---
  function findColumnIndex(headerRowValues: (string | number | boolean)[], possibleHeaders: string[]): number {
      for (const header of possibleHeaders) { if (!header) continue; const lowerHeader = header.toString().toLowerCase().trim(); const index: number = headerRowValues.findIndex(h => h?.toString().toLowerCase().trim() === lowerHeader); if (index !== -1) { return index; } } return -1;
  }
  function parseNumber(value: string | number | boolean | null | undefined): number | null { if (value === null || typeof value === 'undefined' || value === "") { return null; } const cleanedValue: string | number | boolean = typeof value === 'string' ? value.replace(/[^0-9.-]+/g, "") : value; const num: number = Number(cleanedValue); return isNaN(num) ? null : num; }
  function getValuesFromMap(dataMap: PostureDataObject, appId: string, headerName: string): (string | number | boolean)[] { const appData = dataMap[appId]; return appData?.[headerName] ?? []; }
  function getAlignedValuesForRowConcatenation(dataMap: PostureDataObject, appId: string, headers: string[]): { valueLists: (string | number | boolean)[][]; maxRows: number } { const valueLists: (string | number | boolean)[][] = []; let maxRows: number = 0; for (const header of headers) { const values: (string | number | boolean)[] = getValuesFromMap(dataMap, appId, header); valueLists.push(values); if (values.length > maxRows) { maxRows = values.length; } } return { valueLists, maxRows }; }

  // --- 1. Read Configuration ---
  // [No changes needed in Config reading logic for this specific error]
  if (DEBUG_MODE) console.log(`Reading configuration from sheet: ${CONFIG_SHEET_NAME}`);
  const configSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet(CONFIG_SHEET_NAME);
  if (!configSheet) { console.log(`Error: Config sheet "${CONFIG_SHEET_NAME}" not found.`); return; }
  let configValues: (string | number | boolean)[][] = [];
  let configHeaderRow: (string | number | boolean)[];
  const configTable: ExcelScript.Table | undefined = configSheet.getTable(CONFIG_TABLE_NAME);
  try { /* ... existing try/catch block for reading config ... */
      if (configTable) {
          if (DEBUG_MODE) console.log(`Using table "${CONFIG_TABLE_NAME}"...`);
          const tableRange: ExcelScript.Range = configTable.getRange(); if (!tableRange || tableRange.getRowCount() === 0) { console.log(`Warning: Config table "${CONFIG_TABLE_NAME}" is empty.`); return; } const configRangeWithHeader: ExcelScript.Range = configTable.getHeaderRowRange().getResizedRange(configTable.getRowCount(), 0); configValues = await configRangeWithHeader.getValues(); if (configValues.length <= 1) { console.log(`Info: Config table "${CONFIG_TABLE_NAME}" has only headers or is empty.`); return; } configHeaderRow = configValues[0];
      } else {
          if (DEBUG_MODE) console.log(`Using used range on "${CONFIG_SHEET_NAME}" (no table named "${CONFIG_TABLE_NAME}")...`);
          const configRange: ExcelScript.Range | undefined = configSheet.getUsedRange(); if (!configRange || configRange.getRowCount() <= 1) { console.log(`Info: Config sheet "${CONFIG_SHEET_NAME}" is empty or has only a header row.`); return; } configValues = await configRange.getValues(); if (configValues.length <= 1) { console.log(`Info: Config sheet "${CONFIG_SHEET_NAME}" has only a header row or is empty.`); return; } configHeaderRow = configValues[0];
      }
  } catch (error) { console.log(`Error reading config data: ${error instanceof Error ? error.message : String(error)}`); return; }
  const colIdxIsEnabled: number = findColumnIndex(configHeaderRow, ["IsEnabled", "Enabled"]); const colIdxSheetName: number = findColumnIndex(configHeaderRow, ["SheetName", "Sheet Name"]); const colIdxAppIdHeaders: number = findColumnIndex(configHeaderRow, ["AppIdHeaders", "App ID Headers", "Application ID Headers"]); const colIdxDataHeaders: number = findColumnIndex(configHeaderRow, ["DataHeadersToPull", "Data Headers"]); const colIdxAggType: number = findColumnIndex(configHeaderRow, ["AggregationType", "Aggregation Type"]); const colIdxValueHeader: number = findColumnIndex(configHeaderRow, ["ValueHeaderForAggregation", "Value Header"]); const colIdxMasterFields: number = findColumnIndex(configHeaderRow, ["MasterAppFieldsToPull", "Master Fields"]);
  const essentialCols: { [key: string]: number } = { "IsEnabled": colIdxIsEnabled, "SheetName": colIdxSheetName, "AppIdHeaders": colIdxAppIdHeaders, "AggregationType": colIdxAggType }; const missingEssential: string[] = Object.entries(essentialCols).filter(([_, index]) => index === -1).map(([name, _]) => name); if (missingEssential.length > 0) { console.log(`Error: Missing essential config columns: ${missingEssential.join(', ')}.`); return; }
  if (DEBUG_MODE) { /* ... debug logs for optional columns ... */ }
  const POSTURE_SHEETS_CONFIG: PostureSheetConfig[] = []; const uniqueMasterFields = new Set<string>(); let configIsValid: boolean = true;
  for (let i = 1; i < configValues.length; i++) { /* ... existing config parsing and validation loop ... */
      const row: (string | number | boolean)[] = configValues[i]; const maxIndex = Math.max(colIdxIsEnabled, colIdxSheetName, colIdxAppIdHeaders, colIdxAggType, colIdxDataHeaders, colIdxValueHeader, colIdxMasterFields); if (row.length <= maxIndex) { if (DEBUG_MODE) console.log(`Debug: Config (Row ${i + 1}): Skipping row due to insufficient columns (needs ${maxIndex + 1}, has ${row.length}).`); continue; }
      const cleanRow: (string | number | boolean)[] = row.map(val => typeof val === 'string' ? val.trim() : val); const isEnabled: boolean = cleanRow[colIdxIsEnabled]?.toString().toUpperCase() === "TRUE"; if (!isEnabled) continue;
      const sheetName: string = cleanRow[colIdxSheetName]?.toString() ?? ""; const appIdHeadersRaw: string = cleanRow[colIdxAppIdHeaders]?.toString() ?? ""; const aggTypeRaw: string = cleanRow[colIdxAggType]?.toString() || "List"; const dataHeadersRaw: string = (colIdxDataHeaders !== -1 && cleanRow[colIdxDataHeaders] != null) ? cleanRow[colIdxDataHeaders].toString() : ""; const valueHeader: string | undefined = (colIdxValueHeader !== -1 && cleanRow[colIdxValueHeader] != null) ? cleanRow[colIdxValueHeader].toString().trim() : undefined; const masterFieldsRaw: string = (colIdxMasterFields !== -1 && cleanRow[colIdxMasterFields] != null) ? cleanRow[colIdxMasterFields].toString() : "";
      if (!sheetName || !appIdHeadersRaw) { console.log(`Warning: Config (Row ${i + 1}): Missing SheetName or AppIdHeaders. Skipping row.`); continue; }
      const appIdHeaders: string[] = appIdHeadersRaw.split(',').map(h => h.trim()).filter(h => h); const dataHeadersToPull: string[] = dataHeadersRaw.split(',').map(h => h.trim()).filter(h => h); const masterFieldsForRow: string[] = masterFieldsRaw.split(',').map(h => h.trim()).filter(h => h); masterFieldsForRow.forEach(field => uniqueMasterFields.add(field)); if (appIdHeaders.length === 0) { console.log(`Warning: Config (Row ${i + 1}, Sheet: ${sheetName}): AppIdHeaders empty after parsing. Skipping.`); continue; }
      let aggregationType: AggregationMethod = "List"; const normalizedAggType = aggTypeRaw.charAt(0).toUpperCase() + aggTypeRaw.slice(1).toLowerCase(); const allowedAggTypes: AggregationMethod[] = ["List", "Count", "Sum", "Average", "Min", "Max", "UniqueList"]; if (allowedAggTypes.includes(normalizedAggType as AggregationMethod)) { aggregationType = normalizedAggType as AggregationMethod; } else { console.log(`Warning: Config (Row ${i + 1}, Sheet: ${sheetName}): Invalid AggregationType "${aggTypeRaw}". Defaulting to "List".`); }
      let rowIsValid: boolean = true; const needsDataHeaders: AggregationMethod[] = ["List", "Count", "UniqueList"]; const needsValueHeader: AggregationMethod[] = ["Sum", "Average", "Min", "Max"];
      if (needsDataHeaders.includes(aggregationType) && dataHeadersToPull.length === 0) { console.log(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): Type '${aggregationType}' requires 'DataHeadersToPull'.`); rowIsValid = false; }
      else if (needsValueHeader.includes(aggregationType)) { if (!valueHeader) { console.log(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): Type '${aggregationType}' requires 'ValueHeaderForAggregation'.`); rowIsValid = false; } else { let found = false; for (const header of dataHeadersToPull) { if (header === valueHeader) { found = true; break; } } if (!found) { console.log(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): 'ValueHeaderForAggregation' ("${valueHeader}") must be in 'DataHeadersToPull'.`); rowIsValid = false; } } }
      else if (aggregationType === "UniqueList" && dataHeadersToPull.length > 1 && DEBUG_MODE) { console.log(`Debug: Config (Row ${i + 1}, Sheet "${sheetName}"): 'UniqueList' uses only first header ("${dataHeadersToPull[0]}").`); }
      if (rowIsValid) { const configEntry: PostureSheetConfig = { isEnabled: true, sheetName: sheetName, appIdHeaders: appIdHeaders, dataHeadersToPull: dataHeadersToPull, aggregationType: aggregationType, valueHeaderForAggregation: valueHeader, masterFieldsForRow: masterFieldsForRow }; POSTURE_SHEETS_CONFIG.push(configEntry); } else { configIsValid = false; }
  }
  if (!configIsValid) { console.log("Error: Configuration contains errors (see logs above). Please fix and rerun."); return; }
  if (POSTURE_SHEETS_CONFIG.length === 0) { console.log("Info: No enabled and valid configurations found in the Config sheet."); return; }
  const masterFieldsToPull: string[] = Array.from(uniqueMasterFields);
  if (DEBUG_MODE) console.log(`Loaded ${POSTURE_SHEETS_CONFIG.length} posture sheet configurations.`);
  if (masterFieldsToPull.length > 0 && DEBUG_MODE) console.log(`Debug: Will attempt to pull master fields: ${masterFieldsToPull.join(', ')}`);

  // --- 2. Read Master App Data ---
  // [No changes needed in Master App Data reading logic for this specific error]
  if (DEBUG_MODE) console.log(`Reading master App data from sheet: ${MASTER_APP_SHEET_NAME}...`);
  const masterSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet(MASTER_APP_SHEET_NAME); if (!masterSheet) { console.log(`Error: Master application sheet "${MASTER_APP_SHEET_NAME}" not found.`); return; } const masterRange: ExcelScript.Range | undefined = masterSheet.getUsedRange(); if (!masterRange) { console.log(`Info: Master sheet "${MASTER_APP_SHEET_NAME}" appears empty.`); return; } let masterValues: (string | number | boolean)[][] = []; try { masterValues = await masterRange.getValues(); } catch (e) { console.log(`Error reading master sheet data: ${e instanceof Error ? e.message : String(e)}`); return; } if (masterValues.length <= 1) { console.log(`Info: Master sheet "${MASTER_APP_SHEET_NAME}" has only a header row or is empty.`); return; }
  const masterHeaderRow: (string | number | boolean)[] = masterValues[0]; const masterAppIdColIndex: number = findColumnIndex(masterHeaderRow, [MASTER_APP_ID_HEADER]); if (masterAppIdColIndex === -1) { console.log(`Error: Master App ID header "${MASTER_APP_ID_HEADER}" not found in sheet "${MASTER_APP_SHEET_NAME}".`); return; } const masterFieldColIndices = new Map<string, number>(); for (const field of masterFieldsToPull) { const index: number = findColumnIndex(masterHeaderRow, [field]); if (index !== -1) { masterFieldColIndices.set(field, index); } else { console.log(`Warning: Requested master field "${field}" not found in sheet "${MASTER_APP_SHEET_NAME}". It will be skipped.`); } }
  const masterAppIds = new Set<string>(); const masterAppDataMap: MasterAppDataMap = {}; for (let i = 1; i < masterValues.length; i++) { const row: (string | number | boolean)[] = masterValues[i]; if (row.length <= masterAppIdColIndex) continue; const appId: string | undefined = row[masterAppIdColIndex]?.toString().trim(); if (appId && appId !== "") { if (!masterAppIds.has(appId)) { masterAppIds.add(appId); const appData: MasterAppData = {}; masterFieldColIndices.forEach((colIndex: number, fieldName: string) => { if (row.length > colIndex) { appData[fieldName] = row[colIndex]; } }); masterAppDataMap[appId] = appData; } else { if (DEBUG_MODE) console.log(`Debug: Duplicate master App ID "${appId}" found on row ${i + 1}. Using data from first occurrence.`); } } }
  if (DEBUG_MODE) console.log(`Found ${masterAppIds.size} unique App IDs in the master list.`); if (masterAppIds.size === 0) console.log("Warning: No valid App IDs found in the master list.");

  // --- 3. Process Posture Sheets ---
  // [No changes needed in Posture Sheet processing logic for this specific error]
   if (DEBUG_MODE) console.log("Processing posture sheets..."); const postureDataMap: PostureDataObject = {}; for (const config of POSTURE_SHEETS_CONFIG) { if (DEBUG_MODE) console.log(`Processing sheet: ${config.sheetName}...`); const postureSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet(config.sheetName); if (!postureSheet) { console.log(`Warning: Sheet "${config.sheetName}" not found. Skipping.`); continue; } const postureRange: ExcelScript.Range | undefined = postureSheet.getUsedRange(); if (!postureRange || postureRange.getRowCount() <= 1) { console.log(`Info: Sheet "${config.sheetName}" is empty or has only headers. Skipping.`); continue; } let postureValues: (string | number | boolean)[][] = []; try { postureValues = await postureRange.getValues(); } catch (e) { console.log(`Error reading data from posture sheet "${config.sheetName}": ${e instanceof Error ? e.message : String(e)}. Skipping sheet.`); continue; } const postureHeaderRow: (string | number | boolean)[] = postureValues[0]; const appIdColIndex: number = findColumnIndex(postureHeaderRow, config.appIdHeaders); if (appIdColIndex === -1) { console.log(`Warning: App ID header (tried: ${config.appIdHeaders.join(', ')}) not found in sheet "${config.sheetName}". Skipping sheet.`); continue; } const dataColIndicesMap = new Map<string, number>(); let requiredHeadersAvailable: boolean = true; const headersToCheckSet = new Set<string>([...config.dataHeadersToPull]); if (config.valueHeaderForAggregation) { headersToCheckSet.add(config.valueHeaderForAggregation); } const headersRequiredForThisConfig: string[] = Array.from(headersToCheckSet); for (const header of headersRequiredForThisConfig) { if (!header) continue; const index: number = findColumnIndex(postureHeaderRow, [header]); if (index !== -1) { dataColIndicesMap.set(header, index); } else { let isCritical: boolean = false; const isDataHeader = config.dataHeadersToPull.includes(header); const isValueHeader = header === config.valueHeaderForAggregation; if (config.aggregationType === 'List' || config.aggregationType === 'Count') { if (isDataHeader) isCritical = true; } else if (config.aggregationType === 'UniqueList') { if (header === config.dataHeadersToPull[0]) isCritical = true; } else if (['Sum', 'Average', 'Min', 'Max'].includes(config.aggregationType)) { if (isValueHeader) isCritical = true; } if (isCritical) { console.log(`Error: Critical header "${header}" for type "${config.aggregationType}" in sheet "${config.sheetName}" not found. Skipping config.`); requiredHeadersAvailable = false; break; } else if (isDataHeader) { console.log(`Warning: Non-critical header "${header}" requested in DataHeadersToPull not found in ${config.sheetName}.`); } } } if (!requiredHeadersAvailable) continue; let columnsAvailableForProcessing: boolean = true; if (config.aggregationType === 'List' || config.aggregationType === 'Count') { for (const h of config.dataHeadersToPull) { if (!dataColIndicesMap.has(h)) { console.log(`Warning: Not all headers in 'DataHeadersToPull' found for Sheet "${config.sheetName}" (Type: ${config.aggregationType}). Header "${h}" missing. Skipping config.`); columnsAvailableForProcessing = false; break; } } } else if (config.aggregationType === 'UniqueList') { if (!config.dataHeadersToPull[0] || !dataColIndicesMap.has(config.dataHeadersToPull[0])) { console.log(`Warning: First header ("${config.dataHeadersToPull[0] ?? 'N/A'}") for Sheet "${config.sheetName}" (Type: UniqueList) not found. Skipping config.`); columnsAvailableForProcessing = false; } } else if (['Sum', 'Average', 'Min', 'Max'].includes(config.aggregationType)) { if (!config.valueHeaderForAggregation || !dataColIndicesMap.has(config.valueHeaderForAggregation)) { console.log(`Warning: 'ValueHeaderForAggregation' ("${config.valueHeaderForAggregation ?? 'N/A'}") for Sheet "${config.sheetName}" (Type: ${config.aggregationType}) not found. Skipping config.`); columnsAvailableForProcessing = false; } } if (!columnsAvailableForProcessing) continue; let rowsProcessed: number = 0; for (let i = 1; i < postureValues.length; i++) { const row: (string | number | boolean)[] = postureValues[i]; if (row.length <= appIdColIndex) continue; const appId: string | undefined = row[appIdColIndex]?.toString().trim(); if (appId && masterAppIds.has(appId)) { if (!postureDataMap[appId]) postureDataMap[appId] = {}; const appData = postureDataMap[appId]; dataColIndicesMap.forEach((colIndex: number, headerName: string) => { if (row.length > colIndex) { const value = row[colIndex]; if (value !== null && typeof value !== 'undefined' && value !== "") { if (!appData[headerName]) appData[headerName] = []; appData[headerName].push(value); } } }); rowsProcessed++; } } if (DEBUG_MODE) console.log(`Processed ${rowsProcessed} relevant rows for sheet "${config.sheetName}".`); } if (DEBUG_MODE) console.log("Finished processing posture sheets.");

  // --- 4. Prepare and Write Summary Sheet ---
  if (DEBUG_MODE) console.log(`Preparing summary sheet: ${SUMMARY_SHEET_NAME}`);
  workbook.getWorksheet(SUMMARY_SHEET_NAME)?.delete();
  const summarySheet: ExcelScript.Worksheet = workbook.addWorksheet(SUMMARY_SHEET_NAME);
  summarySheet.activate(); // Sync

  const summaryHeaders: string[] = [MASTER_APP_ID_HEADER, ...masterFieldsToPull];
  const addedPostureHeaders = new Set<string>();
  const postureColumnConfigMap = new Map<string, PostureSheetConfig>();
  for (const config of POSTURE_SHEETS_CONFIG) { const header: string = config.sheetName; if (!addedPostureHeaders.has(header)) { summaryHeaders.push(header); postureColumnConfigMap.set(header, config); addedPostureHeaders.add(header); } else { console.log(`Warning: Duplicate config found for SheetName "${header}". Only the first encountered configuration will be used.`); } }
  if (DEBUG_MODE) console.log(`Generated ${summaryHeaders.length} summary headers: ${summaryHeaders.join(', ')}`);
  if (summaryHeaders.length > 0) { const headerRange: ExcelScript.Range = summarySheet.getRangeByIndexes(0, 0, 1, summaryHeaders.length); if (headerRange) { await headerRange.setValues([summaryHeaders]); const headerFormat: ExcelScript.RangeFormat = headerRange.getFormat(); headerFont = headerFormat.getFont(); headerFill = headerFormat.getFill(); headerFont.setBold(true); headerFill.setColor("#4472C4"); headerFont.setColor("white"); } else { console.log("Error: Could not get header range for writing."); } } else { console.log("Warning: No headers generated for the summary sheet."); }

  // --- 4b. Generate Summary Data Rows ---
  const outputData: (string | number | boolean)[][] = [];
  const masterAppIdArray: string[] = Array.from(masterAppIds).sort();
  if (DEBUG_MODE) console.log(`Processing ${masterAppIdArray.length} unique master App IDs for summary rows.`);

  for (const appId of masterAppIdArray) { // Outer loop per App ID
      const masterData: MasterAppData = masterAppDataMap[appId] ?? {};
      const row: (string | number | boolean)[] = [appId];
      for (const field of masterFieldsToPull) { row.push(masterData[field] ?? DEFAULT_VALUE_MISSING); }

      // === START OF POSTURE DATA LOOP ===
      if (DEBUG_MODE) {
           console.log(` --> Processing App ID: ${appId}. Initial row length (AppID + Master Fields): ${row.length}`);
           console.log(` --> Expected # Posture Columns: ${postureColumnConfigMap.size}`);
      }

      // Inner loop per configured posture sheet (unique sheet names)
      postureColumnConfigMap.forEach((config: PostureSheetConfig, headerName: string) => {
          const aggType: AggregationMethod = config.aggregationType;
          let outputValue: string | number | boolean = DEFAULT_VALUE_MISSING;

           if (DEBUG_MODE) console.log(`      -> Processing Header/Sheet: ${headerName} (Type: ${aggType})`);

          try {
              // --- Aggregation Logic Switch ---
              switch (aggType) {
                  case "Count": {
                      const headersToGroup: string[] = config.dataHeadersToPull;
                      const { valueLists, maxRows } = getAlignedValuesForRowConcatenation(postureDataMap, appId, headersToGroup);
                      if (maxRows > 0) { const groupCounts = new Map<string, number>(); const internalSep: string = "|||"; for (let i = 0; i < maxRows; i++) { const keyParts: string[] = []; for (let j = 0; j < headersToGroup.length; j++) { keyParts.push((valueLists[j]?.[i] ?? "").toString()); } const groupKey: string = keyParts.join(internalSep); groupCounts.set(groupKey, (groupCounts.get(groupKey) || 0) + 1); } const sortedEntries: [string, number][] = Array.from(groupCounts.entries()).sort((a: [string, number], b: [string, number]) => a[0].localeCompare(b[0])); const formattedLines: string[] = []; for (const [key, count] of sortedEntries) { formattedLines.push(`${key.split(internalSep).join(CONCATENATE_SEPARATOR)}${COUNT_SEPARATOR}${count}`); } outputValue = formattedLines.join('\n'); } else { outputValue = 0; } break;
                  }
                  case "List": {
                      const headersToConcat: string[] = config.dataHeadersToPull;
                      const { valueLists, maxRows } = getAlignedValuesForRowConcatenation(postureDataMap, appId, headersToConcat);
                      if (maxRows > 0) { const concatenatedLines: string[] = []; for (let i = 0; i < maxRows; i++) { const lineParts: string[] = []; for (let j = 0; j < headersToConcat.length; j++) { lineParts.push((valueLists[j]?.[i] ?? "").toString()); } concatenatedLines.push(lineParts.join(CONCATENATE_SEPARATOR)); } outputValue = concatenatedLines.join('\n'); } break;
                  }
                  case "Sum": case "Average": case "Min": case "Max": {
                      const valueHeader = config.valueHeaderForAggregation!; const values: (string | number | boolean)[] = getValuesFromMap(postureDataMap, appId, valueHeader); const numericValues: number[] = []; for (const v of values) { const num: number | null = parseNumber(v); if (num !== null) { numericValues.push(num); } } if (numericValues.length > 0) { if (aggType === "Sum") outputValue = numericValues.reduce((s, c) => s + c, 0); else if (aggType === "Average") { let sum = numericValues.reduce((s, c) => s + c, 0); outputValue = parseFloat((sum / numericValues.length).toFixed(2)); } else if (aggType === "Min") outputValue = Math.min(...numericValues); else if (aggType === "Max") outputValue = Math.max(...numericValues); } break;
                  }
                  case "UniqueList": {
                      const header: string = config.dataHeadersToPull[0]; const values: (string | number | boolean)[] = getValuesFromMap(postureDataMap, appId, header); if (values.length > 0) { const stringValues: string[] = []; for (const v of values) { stringValues.push(v?.toString() ?? ""); } const uniqueValuesSet = new Set<string>(stringValues); const uniqueValuesArray: string[] = Array.from(uniqueValuesSet); uniqueValuesArray.sort(); outputValue = uniqueValuesArray.join('\n'); } break;
                  }
              } // End switch
          } catch (e: unknown) {
              const errorMsg = e instanceof Error ? e.message : String(e);
              console.log(`Error during aggregation: Type "${aggType}", App "${appId}", Sheet "${config.sheetName}". Details: ${errorMsg}`);
              outputValue = 'ERROR';
          }

          // --- DEBUG LOGS AROUND PUSH ---
          if (DEBUG_MODE) {
              console.log(`         [AppID: ${appId}, Header: ${headerName}] Calculated value: "${outputValue}" (Type: ${typeof outputValue})`);
              console.log(`         [AppID: ${appId}, Header: ${headerName}] Row length BEFORE push: ${row.length}`);
          }

          row.push(outputValue); // Push the single value for this posture sheet column

          if (DEBUG_MODE) {
              console.log(`         [AppID: ${appId}, Header: ${headerName}] Row length AFTER push: ${row.length}`);
          }
          // --- END DEBUG LOGS ---

      }); // === END OF POSTURE DATA LOOP ===

      // Final check for the completed row for this App ID
      if (row.length !== summaryHeaders.length) {
           console.log(`Error: Final row for App ID "${appId}" has ${row.length} columns, but expected ${summaryHeaders.length} (Headers: ${summaryHeaders.join(', ')}). Skipping row.`);
           // Do not push the row if column count mismatch
      } else {
           if (DEBUG_MODE && appId.substring(0, 4) === 'XXXX') { // Example: Log content for specific IDs if needed
               console.log(`Debug: Final row content for App ID "${appId}": [${row.join(' | ')}]`);
           }
          outputData.push(row);
      }
  } // End masterAppIdArray loop


  // --- 4c. Write Data ---
  if (outputData.length > 0) {
      if (DEBUG_MODE) { /* ... existing pre-write diagnostic logs ... */
           console.log(`Attempting to write ${outputData.length} data rows.`); console.log(`Expected columns based on headers: ${summaryHeaders.length}`); if (outputData[0]) { console.log(`Actual columns in first data row: ${outputData[0].length}`); if (outputData[0].length !== summaryHeaders.length) { console.log(`COLUMN COUNT MISMATCH! Header cols (${summaryHeaders.length}): ${summaryHeaders.join(' | ')}`); console.log(`First data row cols (${outputData[0].length}): ${outputData[0].join(' | ')}`); } } else { console.log("First data row is undefined."); }
      }
      if (outputData[0] && outputData[0].length === summaryHeaders.length) {
          try { const dataRange: ExcelScript.Range = summarySheet.getRangeByIndexes(1, 0, outputData.length, summaryHeaders.length); await dataRange.setValues(outputData); if (DEBUG_MODE) console.log(`Successfully wrote ${outputData.length} rows of data.`); } catch (e) { console.log(`Error during final setValues: ${e instanceof Error ? e.message : String(e)}`); console.log(`Data dimensions tried: ${outputData.length} rows, ${summaryHeaders.length} cols.`); }
      } else if (outputData.length > 0) { console.log(`Error: Halting before setValues due to column count mismatch detected earlier. Check preceding logs.`); } else { console.log("Info: No valid data rows to write (outputData is empty)."); }
  } else { if (DEBUG_MODE) console.log(`No data rows generated for the summary.`); }

  // --- 5. Apply Basic Formatting ---
  const finalUsedRange: ExcelScript.Range | undefined = summarySheet.getUsedRange(); if (finalUsedRange && outputData.length > 0) { if (DEBUG_MODE) console.log("Applying formatting..."); try { const usedRangeFormat: ExcelScript.RangeFormat = finalUsedRange.getFormat(); usedRangeFormat.setWrapText(true); usedRangeFormat.setVerticalAlignment(ExcelScript.VerticalAlignment.top); await usedRangeFormat.autofitColumns(); if (DEBUG_MODE) console.log("Applied formatting and autofit columns."); } catch (e) { console.log(`Warning: Error applying formatting: ${e instanceof Error ? e.message : String(e)}`); } } else if (DEBUG_MODE) { console.log("Skipping formatting as no data rows were written or sheet is empty."); }

  // --- Finish ---
  try { const targetSheet = workbook.getWorksheet(SUMMARY_SHEET_NAME); if (targetSheet) { await targetSheet.getCell(0, 0).select(); } } catch (e) { if (DEBUG_MODE) console.log(`Debug: Minor error selecting cell A1: ${e instanceof Error ? e.message : String(e)}`); }
  const endTime = Date.now(); const duration = (endTime - startTime) / 1000; console.log(`Script finished in ${duration.toFixed(2)} seconds.`);

} // End main function

// Declare helper variables potentially used outside initial declaration scope
let headerFont: ExcelScript.RangeFont;
let headerFill: ExcelScript.RangeFill;