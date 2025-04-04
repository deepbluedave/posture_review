/**
 * Posture Summary Script (v16 - Office Scripts Compatibility & Maintainer Notes)
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
 * Key changes in v16:
 * - Fixed type inference error on Map.forEach callback parameters (Line ~446).
 * - Replaced .map() with arrow function inside a loop with a standard for loop
 *   for compatibility (Line ~460).
 * - Added detailed maintainer notes above regarding Office Scripts restrictions.
 * - Previous v15 fixes (Debug Logging, Error Fix Attempt, Default Header) maintained.
 */
async function main(workbook: ExcelScript.Workbook) {
  // --- Control Script Behavior ---
  const DEBUG_MODE: boolean = true; // Set to false to reduce console output

  if (DEBUG_MODE) console.log("Starting posture summary script (v16: Office Scripts Compatibility)...");
  const startTime = Date.now();

  // --- Overall Constants ---
  const MASTER_APP_SHEET_NAME: string = "Applications";
  const MASTER_APP_ID_HEADER: string = "Application ID"; // Changed Default
  const SUMMARY_SHEET_NAME: string = "Posture Summary";
  const CONFIG_SHEET_NAME: string = "Config";
  const CONFIG_TABLE_NAME: string = "ConfigTable"; // Optional table on Config sheet
  const DEFAULT_VALUE_MISSING: string = "";
  const CONCATENATE_SEPARATOR: string = " - ";
  const COUNT_SEPARATOR: string = ": ";

  // --- Type Definitions ---
  type AggregationMethod = "List" | "Count" | "Sum" | "Average" | "Min" | "Max" | "UniqueList";

  type PostureSheetConfig = {
      isEnabled: boolean;
      sheetName: string;
      appIdHeaders: string[];
      dataHeadersToPull: string[];
      aggregationType: AggregationMethod;
      valueHeaderForAggregation?: string;
      masterFieldsForRow?: string[];
  };

  type PostureDataObject = { [appId: string]: { [header: string]: (string | number | boolean)[] } };
  type MasterAppData = { [fieldName: string]: string | number | boolean; }
  type MasterAppDataMap = { [appId: string]: MasterAppData; }

  // --- Helper Functions ---

  function findColumnIndex(headerRowValues: (string | number | boolean)[], possibleHeaders: string[]): number {
      for (const header of possibleHeaders) {
          if (!header) continue;
          const lowerHeader = header.toString().toLowerCase().trim();
          const index: number = headerRowValues.findIndex(h => h?.toString().toLowerCase().trim() === lowerHeader); // Simple arrow func usually OK here
          if (index !== -1) { return index; }
      }
      return -1;
  }

  function parseNumber(value: string | number | boolean | null | undefined): number | null {
      if (value === null || typeof value === 'undefined' || value === "") { return null; }
      const cleanedValue: string | number | boolean = typeof value === 'string' ? value.replace(/[^0-9.-]+/g, "") : value;
      const num: number = Number(cleanedValue);
      return isNaN(num) ? null : num;
  }

  function getValuesFromMap(dataMap: PostureDataObject, appId: string, headerName: string): (string | number | boolean)[] {
      // Added explicit check for appId existence in dataMap for robustness
      const appData = dataMap[appId];
      return appData?.[headerName] ?? [];
  }

  function getAlignedValuesForRowConcatenation(dataMap: PostureDataObject, appId: string, headers: string[]): { valueLists: (string | number | boolean)[][]; maxRows: number } {
      // Using standard loop for map operation for maximum compatibility
      const valueLists: (string | number | boolean)[][] = [];
      let maxRows: number = 0;
      for (const header of headers) {
           const values: (string | number | boolean)[] = getValuesFromMap(dataMap, appId, header);
           valueLists.push(values);
           if (values.length > maxRows) {
               maxRows = values.length;
           }
      }
      // Replaced this .map(): const valueLists = headers.map(header => getValuesFromMap(dataMap, appId, header));
      // Replaced this .map(): const maxRows = Math.max(0, ...valueLists.map(list => list.length));
      // MaxRows calculation moved inside the loop above.

      return { valueLists, maxRows };
  }

  // --- 1. Read Configuration ---
  if (DEBUG_MODE) console.log(`Reading configuration from sheet: ${CONFIG_SHEET_NAME}`);
  const configSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet(CONFIG_SHEET_NAME);
  if (!configSheet) {
      console.log(`Error: Config sheet "${CONFIG_SHEET_NAME}" not found.`);
      return;
  }

  let configValues: (string | number | boolean)[][] = [];
  let configHeaderRow: (string | number | boolean)[];
  const configTable: ExcelScript.Table | undefined = configSheet.getTable(CONFIG_TABLE_NAME);

  try {
      if (configTable) {
          if (DEBUG_MODE) console.log(`Using table "${CONFIG_TABLE_NAME}"...`);
          const tableRange: ExcelScript.Range = configTable.getRange();
           if (!tableRange || tableRange.getRowCount() === 0) {
                console.log(`Warning: Config table "${CONFIG_TABLE_NAME}" is empty.`);
                return;
           }
          const configRangeWithHeader: ExcelScript.Range = configTable.getHeaderRowRange().getResizedRange(configTable.getRowCount(), 0);
          configValues = await configRangeWithHeader.getValues();

          if (configValues.length <= 1) { console.log(`Info: Config table "${CONFIG_TABLE_NAME}" has only headers or is empty.`); return; }
          configHeaderRow = configValues[0];
      } else {
          if (DEBUG_MODE) console.log(`Using used range on "${CONFIG_SHEET_NAME}" (no table named "${CONFIG_TABLE_NAME}")...`);
          const configRange: ExcelScript.Range | undefined = configSheet.getUsedRange();
          // Added check for configRange existence
          if (!configRange || configRange.getRowCount() <= 1) { console.log(`Info: Config sheet "${CONFIG_SHEET_NAME}" is empty or has only a header row.`); return; }
          configValues = await configRange.getValues();
          // Check again after getting values, usedRange might return header even if no data
          if (configValues.length <= 1) { console.log(`Info: Config sheet "${CONFIG_SHEET_NAME}" has only a header row or is empty.`); return; }
          configHeaderRow = configValues[0];
      }
  } catch (error) {
      console.log(`Error reading config data: ${error instanceof Error ? error.message : String(error)}`);
      return;
  }

  // Find config indices
  const colIdxIsEnabled: number = findColumnIndex(configHeaderRow, ["IsEnabled", "Enabled"]);
  const colIdxSheetName: number = findColumnIndex(configHeaderRow, ["SheetName", "Sheet Name"]);
  const colIdxAppIdHeaders: number = findColumnIndex(configHeaderRow, ["AppIdHeaders", "App ID Headers", "Application ID Headers"]); // Added alias
  const colIdxDataHeaders: number = findColumnIndex(configHeaderRow, ["DataHeadersToPull", "Data Headers"]);
  const colIdxAggType: number = findColumnIndex(configHeaderRow, ["AggregationType", "Aggregation Type"]);
  const colIdxValueHeader: number = findColumnIndex(configHeaderRow, ["ValueHeaderForAggregation", "Value Header"]);
  const colIdxMasterFields: number = findColumnIndex(configHeaderRow, ["MasterAppFieldsToPull", "Master Fields"]);

  // Check essential columns
  const essentialCols: { [key: string]: number } = { "IsEnabled": colIdxIsEnabled, "SheetName": colIdxSheetName, "AppIdHeaders": colIdxAppIdHeaders, "AggregationType": colIdxAggType };
  const missingEssential: string[] = Object.entries(essentialCols).filter(([_, index]) => index === -1).map(([name, _]) => name);
  if (missingEssential.length > 0) {
      console.log(`Error: Missing essential config columns: ${missingEssential.join(', ')}.`);
      return;
  }
  if (DEBUG_MODE) {
      if (colIdxDataHeaders === -1) console.log("Debug: Config column 'DataHeadersToPull' not found.");
      if (colIdxValueHeader === -1) console.log("Debug: Config column 'ValueHeaderForAggregation' not found.");
      if (colIdxMasterFields === -1) console.log("Debug: Config column 'MasterAppFieldsToPull' not found.");
  }

  // Parse config & Collect Master Fields
  const POSTURE_SHEETS_CONFIG: PostureSheetConfig[] = [];
  const uniqueMasterFields = new Set<string>();
  let configIsValid: boolean = true;

  // Use standard for loop for iterating config rows
  for (let i = 1; i < configValues.length; i++) {
      const row: (string | number | boolean)[] = configValues[i];
      // Check length against highest possible index used
      const maxIndex = Math.max(colIdxIsEnabled, colIdxSheetName, colIdxAppIdHeaders, colIdxAggType, colIdxDataHeaders, colIdxValueHeader, colIdxMasterFields);
      if (row.length <= maxIndex) {
          if (DEBUG_MODE) console.log(`Debug: Config (Row ${i + 1}): Skipping row due to insufficient columns (needs ${maxIndex + 1}, has ${row.length}).`);
          continue;
      }

      // Using map with simple arrow func here is generally safe
      const cleanRow: (string | number | boolean)[] = row.map(val => typeof val === 'string' ? val.trim() : val);

      const isEnabled: boolean = cleanRow[colIdxIsEnabled]?.toString().toUpperCase() === "TRUE";
      if (!isEnabled) continue;

      const sheetName: string = cleanRow[colIdxSheetName]?.toString() ?? "";
      const appIdHeadersRaw: string = cleanRow[colIdxAppIdHeaders]?.toString() ?? "";
      const aggTypeRaw: string = cleanRow[colIdxAggType]?.toString() || "List";
      const dataHeadersRaw: string = (colIdxDataHeaders !== -1 && cleanRow[colIdxDataHeaders] != null) ? cleanRow[colIdxDataHeaders].toString() : "";
      const valueHeader: string | undefined = (colIdxValueHeader !== -1 && cleanRow[colIdxValueHeader] != null) ? cleanRow[colIdxValueHeader].toString().trim() : undefined;
      const masterFieldsRaw: string = (colIdxMasterFields !== -1 && cleanRow[colIdxMasterFields] != null) ? cleanRow[colIdxMasterFields].toString() : "";

      if (!sheetName || !appIdHeadersRaw) {
          console.log(`Warning: Config (Row ${i + 1}): Missing SheetName or AppIdHeaders. Skipping row.`);
          continue;
      }

      // Using map/filter with simple arrow funcs here is generally safe
      const appIdHeaders: string[] = appIdHeadersRaw.split(',').map(h => h.trim()).filter(h => h);
      const dataHeadersToPull: string[] = dataHeadersRaw.split(',').map(h => h.trim()).filter(h => h);
      const masterFieldsForRow: string[] = masterFieldsRaw.split(',').map(h => h.trim()).filter(h => h);

      // Using forEach with simple arrow func here is generally safe
      masterFieldsForRow.forEach(field => uniqueMasterFields.add(field));

      if (appIdHeaders.length === 0) {
          console.log(`Warning: Config (Row ${i + 1}, Sheet: ${sheetName}): AppIdHeaders empty after parsing. Skipping.`);
          continue;
      }

      let aggregationType: AggregationMethod = "List"; // Explicit default type
      const normalizedAggType = aggTypeRaw.charAt(0).toUpperCase() + aggTypeRaw.slice(1).toLowerCase();
      // Explicit check against allowed types
      const allowedAggTypes: AggregationMethod[] = ["List", "Count", "Sum", "Average", "Min", "Max", "UniqueList"];
      if (allowedAggTypes.includes(normalizedAggType as AggregationMethod)) {
          aggregationType = normalizedAggType as AggregationMethod;
      } else {
          console.log(`Warning: Config (Row ${i + 1}, Sheet: ${sheetName}): Invalid AggregationType "${aggTypeRaw}". Defaulting to "List".`);
      }

      let rowIsValid: boolean = true;
      const needsDataHeaders: AggregationMethod[] = ["List", "Count", "UniqueList"];
      const needsValueHeader: AggregationMethod[] = ["Sum", "Average", "Min", "Max"];

      // Validation logic... (no changes needed here related to compatibility fixes)
      if (needsDataHeaders.includes(aggregationType) && dataHeadersToPull.length === 0) {
          console.log(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): Type '${aggregationType}' requires 'DataHeadersToPull'.`); rowIsValid = false;
      } else if (needsValueHeader.includes(aggregationType)) {
          if (!valueHeader) {
              console.log(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): Type '${aggregationType}' requires 'ValueHeaderForAggregation'.`); rowIsValid = false;
          } else if (!dataHeadersToPull.includes(valueHeader)) {
              // Check if dataHeadersToPull actually contains valueHeader
              let found = false;
              for (const header of dataHeadersToPull) {
                  if (header === valueHeader) {
                      found = true;
                      break;
                  }
              }
              if (!found) {
                  console.log(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): 'ValueHeaderForAggregation' ("${valueHeader}") must be in 'DataHeadersToPull'.`); rowIsValid = false;
              }
          }
      } else if (aggregationType === "UniqueList" && dataHeadersToPull.length > 1 && DEBUG_MODE) {
          console.log(`Debug: Config (Row ${i + 1}, Sheet "${sheetName}"): 'UniqueList' uses only first header ("${dataHeadersToPull[0]}").`);
      }


      if (rowIsValid) {
          // Create config entry with explicit types where possible
          const configEntry: PostureSheetConfig = {
              isEnabled: true,
              sheetName: sheetName,
              appIdHeaders: appIdHeaders,
              dataHeadersToPull: dataHeadersToPull,
              aggregationType: aggregationType,
              valueHeaderForAggregation: valueHeader, // undefined if not applicable/found
              masterFieldsForRow: masterFieldsForRow
          };
          POSTURE_SHEETS_CONFIG.push(configEntry);
      } else {
          configIsValid = false; // Mark overall config invalid
      }
  }

  if (!configIsValid) {
      console.log("Error: Configuration contains errors (see logs above). Please fix and rerun.");
      return;
  }
  if (POSTURE_SHEETS_CONFIG.length === 0) {
      console.log("Info: No enabled and valid configurations found in the Config sheet.");
      return;
  }
  // Use Array.from for Set conversion
  const masterFieldsToPull: string[] = Array.from(uniqueMasterFields);
  if (DEBUG_MODE) console.log(`Loaded ${POSTURE_SHEETS_CONFIG.length} posture sheet configurations.`);
  if (masterFieldsToPull.length > 0 && DEBUG_MODE) console.log(`Debug: Will attempt to pull master fields: ${masterFieldsToPull.join(', ')}`);


  // --- 2. Read Master App Data ---
  if (DEBUG_MODE) console.log(`Reading master App data from sheet: ${MASTER_APP_SHEET_NAME}...`);
  const masterSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet(MASTER_APP_SHEET_NAME);
  if (!masterSheet) { console.log(`Error: Master application sheet "${MASTER_APP_SHEET_NAME}" not found.`); return; }
  const masterRange: ExcelScript.Range | undefined = masterSheet.getUsedRange();
  if (!masterRange) { console.log(`Info: Master sheet "${MASTER_APP_SHEET_NAME}" appears empty.`); return; }
  let masterValues: (string | number | boolean)[][] = [];
  try {
      masterValues = await masterRange.getValues();
  } catch (e) {
      console.log(`Error reading master sheet data: ${e instanceof Error ? e.message : String(e)}`);
      return;
  }
  if (masterValues.length <= 1) { console.log(`Info: Master sheet "${MASTER_APP_SHEET_NAME}" has only a header row or is empty.`); return; }

  const masterHeaderRow: (string | number | boolean)[] = masterValues[0];
  const masterAppIdColIndex: number = findColumnIndex(masterHeaderRow, [MASTER_APP_ID_HEADER]);
  if (masterAppIdColIndex === -1) {
      console.log(`Error: Master App ID header "${MASTER_APP_ID_HEADER}" not found in sheet "${MASTER_APP_SHEET_NAME}".`); return;
  }

  const masterFieldColIndices = new Map<string, number>();
  // Use standard for loop for iterating master fields to find indices
  for (const field of masterFieldsToPull) {
      const index: number = findColumnIndex(masterHeaderRow, [field]);
      if (index !== -1) {
          masterFieldColIndices.set(field, index);
      } else {
          console.log(`Warning: Requested master field "${field}" not found in sheet "${MASTER_APP_SHEET_NAME}". It will be skipped.`);
      }
  }


  const masterAppIds = new Set<string>();
  const masterAppDataMap: MasterAppDataMap = {};
  // Use standard for loop for iterating master data rows
  for (let i = 1; i < masterValues.length; i++) {
      const row: (string | number | boolean)[] = masterValues[i];
      if (row.length <= masterAppIdColIndex) continue;
      const appId: string | undefined = row[masterAppIdColIndex]?.toString().trim(); // App ID could be number, convert safely
      if (appId && appId !== "") {
          if (!masterAppIds.has(appId)) {
               masterAppIds.add(appId);
               const appData: MasterAppData = {};
               // Use Map.forEach here - parameters need type hints but usually okay
               masterFieldColIndices.forEach((colIndex: number, fieldName: string) => {
                   if (row.length > colIndex) {
                      appData[fieldName] = row[colIndex];
                   }
               });
               masterAppDataMap[appId] = appData;
          } else {
              if (DEBUG_MODE) console.log(`Debug: Duplicate master App ID "${appId}" found on row ${i+1}. Using data from first occurrence.`);
          }
      }
  }
  if (DEBUG_MODE) console.log(`Found ${masterAppIds.size} unique App IDs in the master list.`);
  if (masterAppIds.size === 0) console.log("Warning: No valid App IDs found in the master list.");


  // --- 3. Process Posture Sheets ---
  if (DEBUG_MODE) console.log("Processing posture sheets...");
  const postureDataMap: PostureDataObject = {};

  // Use standard for...of loop for iterating configurations
  for (const config of POSTURE_SHEETS_CONFIG) {
      if (DEBUG_MODE) console.log(`Processing sheet: ${config.sheetName}...`);
      const postureSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet(config.sheetName);
      if (!postureSheet) { console.log(`Warning: Sheet "${config.sheetName}" not found. Skipping.`); continue; }
      const postureRange: ExcelScript.Range | undefined = postureSheet.getUsedRange();
      if (!postureRange || postureRange.getRowCount() <= 1) { console.log(`Info: Sheet "${config.sheetName}" is empty or has only headers. Skipping.`); continue; }

      let postureValues: (string | number | boolean)[][] = [];
      try {
          postureValues = await postureRange.getValues();
      } catch (e) {
           console.log(`Error reading data from posture sheet "${config.sheetName}": ${e instanceof Error ? e.message : String(e)}. Skipping sheet.`);
           continue;
      }
      const postureHeaderRow: (string | number | boolean)[] = postureValues[0];
      const appIdColIndex: number = findColumnIndex(postureHeaderRow, config.appIdHeaders);
      if (appIdColIndex === -1) { console.log(`Warning: App ID header (tried: ${config.appIdHeaders.join(', ')}) not found in sheet "${config.sheetName}". Skipping sheet.`); continue; }

      const dataColIndicesMap = new Map<string, number>();
      let requiredHeadersAvailable: boolean = true;
      // Use Set and spread for efficiency if needed, standard loops are safer
      const headersToCheckSet = new Set<string>([...config.dataHeadersToPull]);
      if (config.valueHeaderForAggregation) {
           headersToCheckSet.add(config.valueHeaderForAggregation);
      }
      const headersRequiredForThisConfig: string[] = Array.from(headersToCheckSet); // Back to array for iteration

      // Use standard for loop for iterating headers to check
      for (const header of headersRequiredForThisConfig) {
          if (!header) continue;
          const index: number = findColumnIndex(postureHeaderRow, [header]);
          if (index !== -1) {
              dataColIndicesMap.set(header, index);
          } else {
              let isCritical: boolean = false;
              // Simplified critical check logic
              const isDataHeader = config.dataHeadersToPull.includes(header);
              const isValueHeader = header === config.valueHeaderForAggregation;

              if (config.aggregationType === 'List' || config.aggregationType === 'Count') {
                  if (isDataHeader) isCritical = true;
              } else if (config.aggregationType === 'UniqueList') {
                  if (header === config.dataHeadersToPull[0]) isCritical = true; // Only first is critical
              } else if (['Sum', 'Average', 'Min', 'Max'].includes(config.aggregationType)) {
                  if (isValueHeader) isCritical = true; // Only value header is critical
              }


              if (isCritical) {
                  console.log(`Error: Critical header "${header}" for type "${config.aggregationType}" in sheet "${config.sheetName}" not found. Skipping config.`);
                  requiredHeadersAvailable = false;
                  break; // Exit header check loop early if critical is missing
              } else if (isDataHeader) {
                   // Warn only if it was requested in DataHeadersToPull but wasn't critical for the current agg type
                  console.log(`Warning: Non-critical header "${header}" requested in DataHeadersToPull not found in ${config.sheetName}.`);
              }
          }
      }


      if (!requiredHeadersAvailable) continue; // Skip config if critical headers missing

      // Final check for necessary columns availability
      let columnsAvailableForProcessing: boolean = true;
      // Use standard loops for checking column availability
      if (config.aggregationType === 'List' || config.aggregationType === 'Count') {
          for (const h of config.dataHeadersToPull) {
              if (!dataColIndicesMap.has(h)) {
                   console.log(`Warning: Not all headers in 'DataHeadersToPull' found for Sheet "${config.sheetName}" (Type: ${config.aggregationType}). Header "${h}" missing. Skipping config.`);
                   columnsAvailableForProcessing = false;
                   break;
              }
          }
      } else if (config.aggregationType === 'UniqueList') {
           if (!config.dataHeadersToPull[0] || !dataColIndicesMap.has(config.dataHeadersToPull[0])) {
               console.log(`Warning: First header ("${config.dataHeadersToPull[0] ?? 'N/A'}") for Sheet "${config.sheetName}" (Type: UniqueList) not found. Skipping config.`);
               columnsAvailableForProcessing = false;
           }
      } else if (['Sum', 'Average', 'Min', 'Max'].includes(config.aggregationType)) {
           if (!config.valueHeaderForAggregation || !dataColIndicesMap.has(config.valueHeaderForAggregation)) {
               console.log(`Warning: 'ValueHeaderForAggregation' ("${config.valueHeaderForAggregation ?? 'N/A'}") for Sheet "${config.sheetName}" (Type: ${config.aggregationType}) not found. Skipping config.`);
               columnsAvailableForProcessing = false;
           }
      }


      if (!columnsAvailableForProcessing) continue;

      let rowsProcessed: number = 0;
      // Use standard for loop for iterating posture data rows
      for (let i = 1; i < postureValues.length; i++) {
          const row: (string | number | boolean)[] = postureValues[i];
          if (row.length <= appIdColIndex) continue;

          const appId: string | undefined = row[appIdColIndex]?.toString().trim();
          if (appId && masterAppIds.has(appId)) {
              if (!postureDataMap[appId]) postureDataMap[appId] = {};
              const appData = postureDataMap[appId];

              // Use Map.forEach (parameters need type hints but usually okay)
              dataColIndicesMap.forEach((colIndex: number, headerName: string) => {
                  if (row.length > colIndex) {
                      const value = row[colIndex];
                      if (value !== null && typeof value !== 'undefined' && value !== "") {
                          if (!appData[headerName]) appData[headerName] = [];
                          appData[headerName].push(value);
                      }
                  }
              });
              rowsProcessed++;
          }
      }
      if (DEBUG_MODE) console.log(`Processed ${rowsProcessed} relevant rows for sheet "${config.sheetName}".`);
  }
  if (DEBUG_MODE) console.log("Finished processing posture sheets.");


  // --- 4. Prepare and Write Summary Sheet ---
  if (DEBUG_MODE) console.log(`Preparing summary sheet: ${SUMMARY_SHEET_NAME}`);
  // Use optional chaining for safety
  workbook.getWorksheet(SUMMARY_SHEET_NAME)?.delete();
  const summarySheet: ExcelScript.Worksheet = workbook.addWorksheet(SUMMARY_SHEET_NAME);
  summarySheet.activate(); // Sync

  // Generate headers
  const summaryHeaders: string[] = [MASTER_APP_ID_HEADER, ...masterFieldsToPull];
  const addedPostureHeaders = new Set<string>();
  const postureColumnConfigMap = new Map<string, PostureSheetConfig>();

  // Use standard for...of loop for iterating configurations
  for (const config of POSTURE_SHEETS_CONFIG) {
      const header: string = config.sheetName;
      if (!addedPostureHeaders.has(header)) {
          summaryHeaders.push(header);
          postureColumnConfigMap.set(header, config);
          addedPostureHeaders.add(header);
      } else {
          console.log(`Warning: Duplicate config found for SheetName "${header}". Only the first encountered configuration will be used.`);
      }
  }

  if (DEBUG_MODE) console.log(`Generated ${summaryHeaders.length} summary headers: ${summaryHeaders.join(', ')}`);

  if (summaryHeaders.length > 0) {
      const headerRange: ExcelScript.Range = summarySheet.getRangeByIndexes(0, 0, 1, summaryHeaders.length);
      // Check range exists before setting values
      if(headerRange){
           await headerRange.setValues([summaryHeaders]); // Async
           const headerFormat: ExcelScript.RangeFormat = headerRange.getFormat(); // Sync
           headerFont = headerFormat.getFont(); // Sync
           headerFill = headerFormat.getFill(); // Sync
           headerFont.setBold(true); // Sync
           headerFill.setColor("#4472C4"); // Sync
           headerFont.setColor("white"); // Sync
      } else {
           console.log("Error: Could not get header range for writing.");
      }
  } else {
      console.log("Warning: No headers generated for the summary sheet.");
  }

  // --- 4b. Generate Summary Data Rows ---
  const outputData: (string | number | boolean)[][] = [];
  // Use Array.from for Set conversion
  const masterAppIdArray: string[] = Array.from(masterAppIds).sort();

  if (DEBUG_MODE) console.log(`Processing ${masterAppIdArray.length} unique master App IDs for summary rows.`);

  // Use standard for...of loop for iterating App IDs
  for (const appId of masterAppIdArray) {
      const masterData: MasterAppData = masterAppDataMap[appId] ?? {};
      // Start row: AppID + Master Fields
      const row: (string | number | boolean)[] = [
          appId,
          // Use standard loop instead of map for master fields for max compatibility
          // ...masterFieldsToPull.map(field => masterData[field] ?? DEFAULT_VALUE_MISSING)
      ];
      for(const field of masterFieldsToPull){
          row.push(masterData[field] ?? DEFAULT_VALUE_MISSING);
      }


      // Add posture data columns
      // Use Map.forEach - REQUIRES EXPLICIT TYPES FOR PARAMS (FIXED HERE)
      postureColumnConfigMap.forEach((config: PostureSheetConfig, headerName: string) => {
          const aggType: AggregationMethod = config.aggregationType;
          let outputValue: string | number | boolean = DEFAULT_VALUE_MISSING;

          try {
              switch (aggType) {
                  case "Count": {
                      const headersToGroup: string[] = config.dataHeadersToPull;
                      const { valueLists, maxRows } = getAlignedValuesForRowConcatenation(postureDataMap, appId, headersToGroup);
                      if (maxRows > 0) {
                          const groupCounts = new Map<string, number>();
                          const internalSep: string = "|||";
                          for (let i = 0; i < maxRows; i++) {
                              // FIX LINE 451: Replace map with standard for loop
                              // const keyParts = headersToGroup.map((_, j) => (valueLists[j]?.[i] ?? "").toString());
                              const keyParts: string[] = [];
                              for (let j = 0; j < headersToGroup.length; j++) {
                                   keyParts.push((valueLists[j]?.[i] ?? "").toString());
                              }
                              const groupKey: string = keyParts.join(internalSep);
                              groupCounts.set(groupKey, (groupCounts.get(groupKey) || 0) + 1);
                          }

                          // Use Array.from / map / sort here - generally safe but monitor if issues arise
                          const sortedEntries: [string, number][] = Array.from(groupCounts.entries())
                              .sort((a: [string, number], b: [string, number]) => a[0].localeCompare(b[0])); // Sort by internal key

                          const formattedLines: string[] = [];
                          // Use standard for...of loop for iterating sorted entries
                          for(const [key, count] of sortedEntries){
                              formattedLines.push(`${key.split(internalSep).join(CONCATENATE_SEPARATOR)}${COUNT_SEPARATOR}${count}`);
                          }
                          outputValue = formattedLines.join('\n');

                      } else { outputValue = 0; }
                      break;
                  }
                  case "List": {
                      const headersToConcat: string[] = config.dataHeadersToPull;
                      const { valueLists, maxRows } = getAlignedValuesForRowConcatenation(postureDataMap, appId, headersToConcat);
                      if (maxRows > 0) {
                          const concatenatedLines: string[] = [];
                          // Use standard for loop
                          for (let i = 0; i < maxRows; i++) {
                              const lineParts: string[] = [];
                              // Use standard for loop
                              for (let j = 0; j < headersToConcat.length; j++) {
                                   lineParts.push((valueLists[j]?.[i] ?? "").toString());
                              }
                              concatenatedLines.push(lineParts.join(CONCATENATE_SEPARATOR));
                          }
                          outputValue = concatenatedLines.join('\n');
                      } // else default missing
                      break;
                  }
                  case "Sum": case "Average": case "Min": case "Max": {
                      const valueHeader = config.valueHeaderForAggregation!;
                      const values: (string | number | boolean)[] = getValuesFromMap(postureDataMap, appId, valueHeader);
                      // Use standard loop + push instead of map/filter for numeric parsing
                      const numericValues: number[] = [];
                      for (const v of values) {
                           const num: number | null = parseNumber(v);
                           if (num !== null) {
                               numericValues.push(num);
                           }
                      }
                      // const numericValues = values.map(parseNumber).filter(n => n !== null) as number[]; // Replaced line
                      if (numericValues.length > 0) {
                           // .reduce is often OK, but could be replaced with loop if needed
                          if (aggType === "Sum") outputValue = numericValues.reduce((s, c) => s + c, 0);
                          else if (aggType === "Average") { let sum = numericValues.reduce((s, c) => s + c, 0); outputValue = parseFloat((sum / numericValues.length).toFixed(2)); }
                          // Math.min/max with spread (...) is OK.
                          else if (aggType === "Min") outputValue = Math.min(...numericValues);
                          else if (aggType === "Max") outputValue = Math.max(...numericValues);
                      } // else default missing
                      break;
                  }
                  case "UniqueList": {
                      const header: string = config.dataHeadersToPull[0];
                      const values: (string | number | boolean)[] = getValuesFromMap(postureDataMap, appId, header);
                      if (values.length > 0) {
                          // Set creation, Array.from, map, sort - often OK but complex chain
                          const stringValues: string[] = [];
                          for(const v of values){ // Use loop instead of map
                              stringValues.push(v?.toString() ?? "");
                          }
                          const uniqueValuesSet = new Set<string>(stringValues);
                          const uniqueValuesArray: string[] = Array.from(uniqueValuesSet);
                          uniqueValuesArray.sort(); // Standard sort is fine
                          outputValue = uniqueValuesArray.join('\n');
                      } // else default missing
                      break;
                  }
              } // End switch
          } catch (e: unknown) {
              const errorMsg = e instanceof Error ? e.message : String(e);
              console.log(`Error during aggregation: Type "${aggType}", App "${appId}", Sheet "${config.sheetName}". Details: ${errorMsg}`);
              outputValue = 'ERROR'; // Set cell to ERROR
          }
          row.push(outputValue); // Push value (or default/ERROR) for this posture column
      }); // End postureColumnConfigMap.forEach

      // Final check: Ensure the generated row has the correct number of columns
      if (row.length !== summaryHeaders.length) {
           console.log(`Error: Row for App ID "${appId}" has ${row.length} columns, but expected ${summaryHeaders.length} (headers: ${summaryHeaders.join(', ')}). Skipping row.`);
           // Do not push the row if column count mismatch
      } else {
          outputData.push(row);
      }
  } // End masterAppIdArray loop


  // --- 4c. Write Data ---
  if (outputData.length > 0) {
      if (DEBUG_MODE) {
          console.log(`Attempting to write ${outputData.length} data rows.`);
          console.log(`Expected columns based on headers: ${summaryHeaders.length}`);
          if (outputData[0]) {
              console.log(`Actual columns in first data row: ${outputData[0].length}`);
               if(outputData[0].length !== summaryHeaders.length){
                   console.log(`COLUMN COUNT MISMATCH! Header columns (${summaryHeaders.length}): ${summaryHeaders.join(' | ')}`);
                   console.log(`First data row cols (${outputData[0].length}): ${outputData[0].join(' | ')}`);
               }
          } else {
              console.log("First data row is undefined.");
          }
      }

      // Final check before writing (redundant if inner check works, but safe)
      if (outputData[0] && outputData[0].length === summaryHeaders.length) {
           try {
               const dataRange: ExcelScript.Range = summarySheet.getRangeByIndexes(1, 0, outputData.length, summaryHeaders.length);
               await dataRange.setValues(outputData); // Async
               if (DEBUG_MODE) console.log(`Successfully wrote ${outputData.length} rows of data.`);
           } catch (e) {
               console.log(`Error during final setValues: ${e instanceof Error ? e.message : String(e)}`);
               console.log(`Data dimensions tried: ${outputData.length} rows, ${summaryHeaders.length} cols.`);
           }
      } else if (outputData.length > 0) {
           // This condition implies the check inside the loop caught the mismatch
           console.log(`Error: Halting before setValues due to column count mismatch detected earlier. Check preceding logs.`);
      } else {
           console.log("Info: No valid data rows to write (outputData is empty).");
      }

  } else {
      if (DEBUG_MODE) console.log(`No data rows generated for the summary.`);
  }

  // --- 5. Apply Basic Formatting ---
  // Check if sheet and range exist before formatting
  const finalUsedRange: ExcelScript.Range | undefined = summarySheet.getUsedRange();
  if (finalUsedRange && outputData.length > 0) {
      if (DEBUG_MODE) console.log("Applying formatting...");
      try {
          const usedRangeFormat: ExcelScript.RangeFormat = finalUsedRange.getFormat();
          usedRangeFormat.setWrapText(true); // Sync
          usedRangeFormat.setVerticalAlignment(ExcelScript.VerticalAlignment.top); // Sync
          await usedRangeFormat.autofitColumns(); // Async
          if (DEBUG_MODE) console.log("Applied formatting and autofit columns.");
      } catch (e) {
          console.log(`Warning: Error applying formatting: ${e instanceof Error ? e.message : String(e)}`);
      }
  } else if (DEBUG_MODE) {
      console.log("Skipping formatting as no data rows were written or sheet is empty.");
  }


  // --- Finish ---
  try {
      // Select cell A1 only if the sheet still exists
      const targetSheet = workbook.getWorksheet(SUMMARY_SHEET_NAME);
      if(targetSheet){
           await targetSheet.getCell(0, 0).select(); // Async
      }
  } catch (e) {
      if (DEBUG_MODE) console.log(`Debug: Minor error selecting cell A1: ${e instanceof Error ? e.message : String(e)}`);
  }

  const endTime = Date.now();
  const duration = (endTime - startTime) / 1000;
  console.log(`Script finished in ${duration.toFixed(2)} seconds.`);

} // End main function

// Declare helper variables potentially used outside initial declaration scope
// Explicitly type them
let headerFont: ExcelScript.RangeFont;
let headerFill: ExcelScript.RangeFill;