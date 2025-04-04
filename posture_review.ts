/**
 * Posture Summary Script (v15 - Debug Logging, Error Fix Attempt, Default Header)
 *
 * Reads application posture data from various sheets defined in a 'Config' sheet,
 * aggregates the data based on specified methods (List, Count, Sum, Average, Min, Max, UniqueList),
 * pulls specified fields from the master application list,
 * and writes a summary report to a 'Posture Summary' sheet.
 *
 * Key changes:
 * - Changed default MASTER_APP_ID_HEADER to "Application ID".
 * - Replaced console.warn/error with console.log.
 * - Added DEBUG_MODE flag to control verbose logging output.
 * - Added diagnostic logging before the final setValues call to help debug dimension mismatch errors.
 * - Standardized headers and aggregation display logic from v14 remains.
 */
async function main(workbook: ExcelScript.Workbook) {
  // --- Control Script Behavior ---
  const DEBUG_MODE: boolean = true; // Set to false to reduce console output

  if (DEBUG_MODE) console.log("Starting posture summary script (v15: Debug Logging, Error Fix)...");
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
          const index = headerRowValues.findIndex(h => h?.toString().toLowerCase().trim() === lowerHeader);
          if (index !== -1) { return index; }
      }
      return -1;
  }

  function parseNumber(value: string | number | boolean | null | undefined): number | null {
      if (value === null || typeof value === 'undefined' || value === "") { return null; }
      const cleanedValue = typeof value === 'string' ? value.replace(/[^0-9.-]+/g, "") : value;
      const num: number = Number(cleanedValue);
      return isNaN(num) ? null : num;
  }

  function getValuesFromMap(dataMap: PostureDataObject, appId: string, headerName: string): (string | number | boolean)[] {
      return dataMap[appId]?.[headerName] ?? [];
  }

  function getAlignedValuesForRowConcatenation(dataMap: PostureDataObject, appId: string, headers: string[]): { valueLists: (string | number | boolean)[][]; maxRows: number } {
      const valueLists = headers.map(header => getValuesFromMap(dataMap, appId, header));
      const maxRows = Math.max(0, ...valueLists.map(list => list.length));
      return { valueLists, maxRows };
  }

  // --- 1. Read Configuration ---
  if (DEBUG_MODE) console.log(`Reading configuration from sheet: ${CONFIG_SHEET_NAME}`);
  const configSheet = workbook.getWorksheet(CONFIG_SHEET_NAME);
  if (!configSheet) {
      console.log(`Error: Config sheet "${CONFIG_SHEET_NAME}" not found.`); // Use console.log
      return;
  }

  let configValues: (string | number | boolean)[][] = [];
  let configHeaderRow: (string | number | boolean)[];
  const configTable = configSheet.getTable(CONFIG_TABLE_NAME);

  try {
      if (configTable) {
          if (DEBUG_MODE) console.log(`Using table "${CONFIG_TABLE_NAME}"...`);
          const tableRange = configTable.getRange();
          // Ensure we get header and body if table exists
           if (!tableRange || tableRange.getRowCount() === 0) {
                console.log(`Warning: Config table "${CONFIG_TABLE_NAME}" is empty.`);
                return;
           }
           // Fetch header and data together
          const configRangeWithHeader = configTable.getHeaderRowRange().getResizedRange(configTable.getRowCount(), 0);
          configValues = await configRangeWithHeader.getValues();

          if (configValues.length <= 1) { console.log(`Info: Config table "${CONFIG_TABLE_NAME}" has only headers or is empty.`); return; }
          configHeaderRow = configValues[0];
      } else {
          if (DEBUG_MODE) console.log(`Using used range on "${CONFIG_SHEET_NAME}" (no table named "${CONFIG_TABLE_NAME}")...`);
          const configRange = configSheet.getUsedRange();
          if (!configRange || configRange.getRowCount() <= 1) { console.log(`Info: Config sheet "${CONFIG_SHEET_NAME}" is empty or has only a header row.`); return; }
          configValues = await configRange.getValues();
          configHeaderRow = configValues[0];
      }
  } catch (error) {
      console.log(`Error reading config data: ${error instanceof Error ? error.message : String(error)}`);
      return;
  }

  // Find config indices
  const colIdxIsEnabled = findColumnIndex(configHeaderRow, ["IsEnabled", "Enabled"]);
  const colIdxSheetName = findColumnIndex(configHeaderRow, ["SheetName", "Sheet Name"]);
  const colIdxAppIdHeaders = findColumnIndex(configHeaderRow, ["AppIdHeaders", "App ID Headers", "Application ID Headers"]); // Added alias
  const colIdxDataHeaders = findColumnIndex(configHeaderRow, ["DataHeadersToPull", "Data Headers"]);
  const colIdxAggType = findColumnIndex(configHeaderRow, ["AggregationType", "Aggregation Type"]);
  const colIdxValueHeader = findColumnIndex(configHeaderRow, ["ValueHeaderForAggregation", "Value Header"]);
  const colIdxMasterFields = findColumnIndex(configHeaderRow, ["MasterAppFieldsToPull", "Master Fields"]);

  const essentialCols = { "IsEnabled": colIdxIsEnabled, "SheetName": colIdxSheetName, "AppIdHeaders": colIdxAppIdHeaders, "AggregationType": colIdxAggType };
  const missingEssential = Object.entries(essentialCols).filter(([_, index]) => index === -1).map(([name, _]) => name);
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
  let configIsValid = true;

  for (let i = 1; i < configValues.length; i++) {
      const row = configValues[i];
      if (row.length <= Math.max(colIdxIsEnabled, colIdxSheetName, colIdxAppIdHeaders, colIdxAggType)) {
          if (DEBUG_MODE) console.log(`Debug: Config (Row ${i + 1}): Skipping row due to insufficient columns.`);
          continue;
      }

      const cleanRow = row.map(val => typeof val === 'string' ? val.trim() : val);
      const isEnabled = cleanRow[colIdxIsEnabled]?.toString().toUpperCase() === "TRUE";
      if (!isEnabled) continue;

      const sheetName = cleanRow[colIdxSheetName]?.toString() ?? "";
      const appIdHeadersRaw = cleanRow[colIdxAppIdHeaders]?.toString() ?? "";
      const aggTypeRaw = cleanRow[colIdxAggType]?.toString() || "List";
      const dataHeadersRaw = (colIdxDataHeaders !== -1 && cleanRow[colIdxDataHeaders] != null) ? cleanRow[colIdxDataHeaders].toString() : "";
      const valueHeader = (colIdxValueHeader !== -1 && cleanRow[colIdxValueHeader] != null) ? cleanRow[colIdxValueHeader].toString().trim() : undefined;
      const masterFieldsRaw = (colIdxMasterFields !== -1 && cleanRow[colIdxMasterFields] != null) ? cleanRow[colIdxMasterFields].toString() : "";

      if (!sheetName || !appIdHeadersRaw) {
          console.log(`Warning: Config (Row ${i + 1}): Missing SheetName or AppIdHeaders. Skipping row.`);
          continue;
      }

      const appIdHeaders = appIdHeadersRaw.split(',').map(h => h.trim()).filter(h => h);
      const dataHeadersToPull = dataHeadersRaw.split(',').map(h => h.trim()).filter(h => h);
      const masterFieldsForRow = masterFieldsRaw.split(',').map(h => h.trim()).filter(h => h);

      masterFieldsForRow.forEach(field => uniqueMasterFields.add(field));

      if (appIdHeaders.length === 0) {
          console.log(`Warning: Config (Row ${i + 1}, Sheet: ${sheetName}): AppIdHeaders empty after parsing. Skipping.`);
          continue;
      }

      let aggregationType = "List" as AggregationMethod;
      const normalizedAggType = aggTypeRaw.charAt(0).toUpperCase() + aggTypeRaw.slice(1).toLowerCase();
      if (["List", "Count", "Sum", "Average", "Min", "Max", "UniqueList"].includes(normalizedAggType)) {
          aggregationType = normalizedAggType as AggregationMethod;
      } else {
          console.log(`Warning: Config (Row ${i + 1}, Sheet: ${sheetName}): Invalid AggregationType "${aggTypeRaw}". Defaulting to "List".`);
      }

      let rowIsValid = true;
      const needsDataHeaders = ["List", "Count", "UniqueList"];
      const needsValueHeader = ["Sum", "Average", "Min", "Max"];

      if (needsDataHeaders.includes(aggregationType) && dataHeadersToPull.length === 0) {
          console.log(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): Type '${aggregationType}' requires 'DataHeadersToPull'.`); rowIsValid = false;
      } else if (needsValueHeader.includes(aggregationType)) {
          if (!valueHeader) {
              console.log(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): Type '${aggregationType}' requires 'ValueHeaderForAggregation'.`); rowIsValid = false;
          } else if (!dataHeadersToPull.includes(valueHeader)) {
              console.log(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): 'ValueHeaderForAggregation' ("${valueHeader}") must be in 'DataHeadersToPull'.`); rowIsValid = false;
          }
      } else if (aggregationType === "UniqueList" && dataHeadersToPull.length > 1 && DEBUG_MODE) {
          console.log(`Debug: Config (Row ${i + 1}, Sheet "${sheetName}"): 'UniqueList' uses only first header ("${dataHeadersToPull[0]}").`);
      }

      if (rowIsValid) {
          POSTURE_SHEETS_CONFIG.push({ isEnabled: true, sheetName, appIdHeaders, dataHeadersToPull, aggregationType, valueHeaderForAggregation: valueHeader, masterFieldsForRow });
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
  const masterFieldsToPull: string[] = Array.from(uniqueMasterFields);
  if (DEBUG_MODE) console.log(`Loaded ${POSTURE_SHEETS_CONFIG.length} posture sheet configurations.`);
  if (masterFieldsToPull.length > 0 && DEBUG_MODE) console.log(`Debug: Will attempt to pull master fields: ${masterFieldsToPull.join(', ')}`);


  // --- 2. Read Master App Data ---
  if (DEBUG_MODE) console.log(`Reading master App data from sheet: ${MASTER_APP_SHEET_NAME}...`);
  const masterSheet = workbook.getWorksheet(MASTER_APP_SHEET_NAME);
  if (!masterSheet) { console.log(`Error: Master application sheet "${MASTER_APP_SHEET_NAME}" not found.`); return; }
  const masterRange = masterSheet.getUsedRange();
  if (!masterRange) { console.log(`Info: Master sheet "${MASTER_APP_SHEET_NAME}" appears empty.`); return; }
  let masterValues: (string | number | boolean)[][] = [];
  try {
      masterValues = await masterRange.getValues();
  } catch (e) {
      console.log(`Error reading master sheet data: ${e instanceof Error ? e.message : String(e)}`);
      return;
  }
  if (masterValues.length <= 1) { console.log(`Info: Master sheet "${MASTER_APP_SHEET_NAME}" has only a header row or is empty.`); return; }

  const masterHeaderRow = masterValues[0];
  const masterAppIdColIndex = findColumnIndex(masterHeaderRow, [MASTER_APP_ID_HEADER]);
  if (masterAppIdColIndex === -1) {
      console.log(`Error: Master App ID header "${MASTER_APP_ID_HEADER}" not found in sheet "${MASTER_APP_SHEET_NAME}".`); return;
  }

  const masterFieldColIndices = new Map<string, number>();
  masterFieldsToPull.forEach(field => {
      const index = findColumnIndex(masterHeaderRow, [field]);
      if (index !== -1) { masterFieldColIndices.set(field, index); }
      else { console.log(`Warning: Requested master field "${field}" not found in sheet "${MASTER_APP_SHEET_NAME}". It will be skipped.`); }
  });

  const masterAppIds = new Set<string>();
  const masterAppDataMap: MasterAppDataMap = {};
  for (let i = 1; i < masterValues.length; i++) {
      const row = masterValues[i];
      if (row.length <= masterAppIdColIndex) continue;
      const appId = row[masterAppIdColIndex]?.toString().trim();
      if (appId && appId !== "") {
          if (!masterAppIds.has(appId)) { // Process only first occurrence of App ID if duplicates exist
               masterAppIds.add(appId);
               const appData: MasterAppData = {};
               masterFieldColIndices.forEach((colIndex, fieldName) => {
                   if (row.length > colIndex) appData[fieldName] = row[colIndex];
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

  for (const config of POSTURE_SHEETS_CONFIG) {
      if (DEBUG_MODE) console.log(`Processing sheet: ${config.sheetName}...`);
      const postureSheet = workbook.getWorksheet(config.sheetName);
      if (!postureSheet) { console.log(`Warning: Sheet "${config.sheetName}" not found. Skipping.`); continue; }
      const postureRange = postureSheet.getUsedRange();
      if (!postureRange || postureRange.getRowCount() <= 1) { console.log(`Info: Sheet "${config.sheetName}" is empty or has only headers. Skipping.`); continue; }

      let postureValues: (string | number | boolean)[][] = [];
      try {
          postureValues = await postureRange.getValues();
      } catch (e) {
           console.log(`Error reading data from posture sheet "${config.sheetName}": ${e instanceof Error ? e.message : String(e)}. Skipping sheet.`);
           continue;
      }
      const postureHeaderRow = postureValues[0];
      const appIdColIndex = findColumnIndex(postureHeaderRow, config.appIdHeaders);
      if (appIdColIndex === -1) { console.log(`Warning: App ID header (tried: ${config.appIdHeaders.join(', ')}) not found in sheet "${config.sheetName}". Skipping sheet.`); continue; }

      const dataColIndicesMap = new Map<string, number>();
      let requiredHeadersAvailable = true;
      const headersRequiredForThisConfig = new Set<string>([...config.dataHeadersToPull, ...(config.valueHeaderForAggregation ? [config.valueHeaderForAggregation] : [])]);

      headersRequiredForThisConfig.forEach(header => {
          if (!header) return;
          const index = findColumnIndex(postureHeaderRow, [header]);
          if (index !== -1) {
              dataColIndicesMap.set(header, index);
          } else {
              let isCritical = false;
              if ((config.aggregationType === 'List' || config.aggregationType === 'Count' || config.aggregationType === 'UniqueList') && config.dataHeadersToPull.includes(header)) isCritical = true;
              else if (['Sum', 'Average', 'Min', 'Max'].includes(config.aggregationType) && header === config.valueHeaderForAggregation) isCritical = true;

              if (isCritical) {
                  console.log(`Error: Critical header "${header}" for type "${config.aggregationType}" in sheet "${config.sheetName}" not found. Skipping config.`);
                  requiredHeadersAvailable = false;
              } else if (config.dataHeadersToPull.includes(header)) {
                  // Warn only if it was requested but wasn't critical for the current agg type
                  console.log(`Warning: Non-critical header "${header}" requested in DataHeadersToPull not found in ${config.sheetName}.`);
              }
          }
      });

      if (!requiredHeadersAvailable) continue;

      let columnsAvailableForProcessing = true;
      if ((config.aggregationType === 'List' || config.aggregationType === 'Count') && !config.dataHeadersToPull.every(h => dataColIndicesMap.has(h))) {
          console.log(`Warning: Not all headers in 'DataHeadersToPull' found for Sheet "${config.sheetName}" (Type: ${config.aggregationType}). Skipping config.`); columnsAvailableForProcessing = false;
      } else if (config.aggregationType === 'UniqueList' && !dataColIndicesMap.has(config.dataHeadersToPull[0])) {
          console.log(`Warning: First header ("${config.dataHeadersToPull[0]}") for Sheet "${config.sheetName}" (Type: UniqueList) not found. Skipping config.`); columnsAvailableForProcessing = false;
      } else if (['Sum', 'Average', 'Min', 'Max'].includes(config.aggregationType) && !dataColIndicesMap.has(config.valueHeaderForAggregation!)) {
          console.log(`Warning: 'ValueHeaderForAggregation' ("${config.valueHeaderForAggregation}") for Sheet "${config.sheetName}" (Type: ${config.aggregationType}) not found. Skipping config.`); columnsAvailableForProcessing = false;
      }

      if (!columnsAvailableForProcessing) continue;

      let rowsProcessed = 0;
      for (let i = 1; i < postureValues.length; i++) {
          const row = postureValues[i];
          if (row.length <= appIdColIndex) continue;

          const appId = row[appIdColIndex]?.toString().trim();
          if (appId && masterAppIds.has(appId)) { // Check if App ID is in our master list
              if (!postureDataMap[appId]) postureDataMap[appId] = {};
              const appData = postureDataMap[appId];

              dataColIndicesMap.forEach((colIndex, headerName) => {
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
  workbook.getWorksheet(SUMMARY_SHEET_NAME)?.delete();
  const summarySheet = workbook.addWorksheet(SUMMARY_SHEET_NAME);
  summarySheet.activate();

  const summaryHeaders: string[] = [MASTER_APP_ID_HEADER, ...masterFieldsToPull];
  const addedPostureHeaders = new Set<string>();
  const postureColumnConfigMap = new Map<string, PostureSheetConfig>();

  POSTURE_SHEETS_CONFIG.forEach(config => {
      const header = config.sheetName;
      if (!addedPostureHeaders.has(header)) {
          summaryHeaders.push(header);
          postureColumnConfigMap.set(header, config);
          addedPostureHeaders.add(header);
      } else {
          console.log(`Warning: Duplicate config found for SheetName "${header}". Only the first encountered configuration will be used.`);
      }
  });

  if (DEBUG_MODE) console.log(`Generated ${summaryHeaders.length} summary headers: ${summaryHeaders.join(', ')}`);

  if (summaryHeaders.length > 0) {
      const headerRange = summarySheet.getRangeByIndexes(0, 0, 1, summaryHeaders.length);
      await headerRange.setValues([summaryHeaders]);
      const headerFormat = headerRange.getFormat();
      headerFont = headerFormat.getFont();
      headerFill = headerFormat.getFill();
      headerFont.setBold(true);
      headerFill.setColor("#4472C4");
      headerFont.setColor("white");
  } else {
      console.log("Warning: No headers generated for the summary sheet.");
      // Decide if we should stop if no headers? For now, continue, might write empty data.
  }

  // --- 4b. Generate Summary Data Rows ---
  const outputData: (string | number | boolean)[][] = [];
  const masterAppIdArray = Array.from(masterAppIds).sort();

  if (DEBUG_MODE) console.log(`Processing ${masterAppIdArray.length} unique master App IDs for summary rows.`);

  masterAppIdArray.forEach(appId => {
      const masterData = masterAppDataMap[appId] ?? {};
      // Start row: AppID + Master Fields
      const row: (string | number | boolean)[] = [
          appId,
          ...masterFieldsToPull.map(field => masterData[field] ?? DEFAULT_VALUE_MISSING)
      ];

      // Add posture data columns
      postureColumnConfigMap.forEach((config, headerName) => {
          const aggType = config.aggregationType;
          let outputValue: string | number | boolean = DEFAULT_VALUE_MISSING;

          try {
              switch (aggType) {
                  case "Count": {
                      const headersToGroup = config.dataHeadersToPull;
                      const { valueLists, maxRows } = getAlignedValuesForRowConcatenation(postureDataMap, appId, headersToGroup);
                      if (maxRows > 0) {
                          const groupCounts = new Map<string, number>();
                          const internalSep = "|||"; // Simpler internal separator
                          for (let i = 0; i < maxRows; i++) {
                              const keyParts = headersToGroup.map((_, j) => (valueLists[j]?.[i] ?? "").toString());
                              const groupKey = keyParts.join(internalSep);
                              groupCounts.set(groupKey, (groupCounts.get(groupKey) || 0) + 1);
                          }
                          const formattedEntries = Array.from(groupCounts.entries())
                              .sort((a, b) => a[0].localeCompare(b[0])) // Sort by internal key
                              .map(([key, count]) => `${key.split(internalSep).join(CONCATENATE_SEPARATOR)}${COUNT_SEPARATOR}${count}`);
                          outputValue = formattedEntries.join('\n');
                      } else { outputValue = 0; }
                      break;
                  }
                  case "List": {
                      const headersToConcat = config.dataHeadersToPull;
                      const { valueLists, maxRows } = getAlignedValuesForRowConcatenation(postureDataMap, appId, headersToConcat);
                      if (maxRows > 0) {
                          const concatenatedLines = [];
                          for (let i = 0; i < maxRows; i++) {
                              const lineParts = headersToConcat.map((_, j) => (valueLists[j]?.[i] ?? "").toString());
                              concatenatedLines.push(lineParts.join(CONCATENATE_SEPARATOR));
                          }
                          outputValue = concatenatedLines.join('\n');
                      } // else default missing
                      break;
                  }
                  case "Sum": case "Average": case "Min": case "Max": {
                      const valueHeader = config.valueHeaderForAggregation!;
                      const values = getValuesFromMap(postureDataMap, appId, valueHeader);
                      const numericValues = values.map(parseNumber).filter(n => n !== null) as number[];
                      if (numericValues.length > 0) {
                          if (aggType === "Sum") outputValue = numericValues.reduce((s, c) => s + c, 0);
                          else if (aggType === "Average") { let avg = numericValues.reduce((s, c) => s + c, 0) / numericValues.length; outputValue = parseFloat(avg.toFixed(2)); }
                          else if (aggType === "Min") outputValue = Math.min(...numericValues);
                          else if (aggType === "Max") outputValue = Math.max(...numericValues);
                      } // else default missing
                      break;
                  }
                  case "UniqueList": {
                      const header = config.dataHeadersToPull[0];
                      const values = getValuesFromMap(postureDataMap, appId, header);
                      if (values.length > 0) {
                          const uniqueValues = Array.from(new Set(values.map(v => v?.toString() ?? ""))).sort();
                          outputValue = uniqueValues.join('\n');
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
           // Optionally, you could push a row of ERRORs, but skipping is safer for setValues
      } else {
          outputData.push(row);
      }
  }); // End masterAppIdArray.forEach


  // --- 4c. Write Data ---
  if (outputData.length > 0) {
      if (DEBUG_MODE) {
          console.log(`Attempting to write ${outputData.length} data rows.`);
          // --- DIAGNOSTIC LOGS ---
          console.log(`Expected columns based on headers: ${summaryHeaders.length}`);
          if (outputData[0]) {
              console.log(`Actual columns in first data row: ${outputData[0].length}`);
               if(outputData[0].length !== summaryHeaders.length){
                   console.log(`COLUMN COUNT MISMATCH! Header columns: ${summaryHeaders.join(' | ')}`);
                   console.log(`First data row columns: ${outputData[0].join(' | ')}`);
               }
          } else {
              console.log("First data row is undefined (outputData might be empty despite check).");
          }
           // Check last row as well?
           const lastRow = outputData[outputData.length -1];
           if (lastRow && lastRow.length !== summaryHeaders.length) {
              console.log(`COLUMN COUNT MISMATCH! Last data row (${outputData.length}) columns: ${lastRow.length}. Data: ${lastRow.join(' | ')}`);
           }
          // --- END DIAGNOSTIC ---
      }

      // Double-check dimensions before writing
      if (outputData[0] && outputData[0].length === summaryHeaders.length) {
           try {
               const dataRange = summarySheet.getRangeByIndexes(1, 0, outputData.length, summaryHeaders.length);
               await dataRange.setValues(outputData);
               if (DEBUG_MODE) console.log(`Successfully wrote ${outputData.length} rows of data.`);
           } catch (e) {
               console.log(`Error during final setValues: ${e instanceof Error ? e.message : String(e)}`);
               console.log(`Data dimensions tried: ${outputData.length} rows, ${summaryHeaders.length} cols.`);
               // Log more details if in debug mode
               if (DEBUG_MODE && outputData.length > 0) {
                  console.log("First row data sample: " + JSON.stringify(outputData[0].slice(0, 10))); // Log first 10 cols sample
               }
           }
      } else if (outputData.length > 0) {
           console.log(`Error: Halting before setValues due to column count mismatch detected between headers (${summaryHeaders.length}) and data rows (${outputData[0]?.length ?? 'N/A'}). Check preceding logs.`);
      } else {
           // This case should technically not be reached if outputData.length > 0 check passed, but included for safety.
           console.log("Info: No valid data rows to write (outputData is empty or first row invalid).");
      }

  } else {
      if (DEBUG_MODE) console.log(`No data rows generated for the summary.`);
  }

  // --- 5. Apply Basic Formatting ---
  const usedRange = summarySheet.getUsedRange();
  if (usedRange && outputData.length > 0) { // Only format if data was potentially written
      if (DEBUG_MODE) console.log("Applying formatting...");
      try {
          const usedRangeFormat = usedRange.getFormat();
          usedRangeFormat.setWrapText(true);
          usedRangeFormat.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
          await usedRangeFormat.autofitColumns();
          if (DEBUG_MODE) console.log("Applied formatting and autofit columns.");
      } catch (e) {
          console.log(`Warning: Error applying formatting: ${e instanceof Error ? e.message : String(e)}`);
      }
  } else if (DEBUG_MODE && outputData.length === 0) {
      console.log("Skipping formatting as no data rows were written.");
  }


  // --- Finish ---
  try {
      await summarySheet.getCell(0, 0).select();
  } catch (e) {
      if (DEBUG_MODE) console.log(`Debug: Minor error selecting cell A1: ${e instanceof Error ? e.message : String(e)}`);
  }

  const endTime = Date.now();
  const duration = (endTime - startTime) / 1000;
  console.log(`Script finished in ${duration.toFixed(2)} seconds.`);

} // End main function

// Declare helper variables potentially used outside initial declaration scope
let headerFont: ExcelScript.RangeFont;
let headerFill: ExcelScript.RangeFill;