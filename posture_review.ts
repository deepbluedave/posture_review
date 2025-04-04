/**
 * Posture Summary Script (v13 - Concatenate Aggregation & Master Fields)
 *
 * Reads application posture data from various sheets defined in a 'Config' sheet,
 * aggregates the data based on specified methods (List, Count, Sum, Average, Min, Max, UniqueList, Concatenate),
 * pulls specified fields from the master application list,
 * and writes a summary report to a 'Posture Summary' sheet.
 *
 * Key changes:
 * - Added "Concatenate" aggregation type: Creates one column per sheet, concatenating
 *   values from 'DataHeadersToPull' for each related row in the source sheet.
 * - Added ability to pull additional fields from the master "Applications" sheet
 *   into the summary, configured via a new "MasterAppFieldsToPull" column in Config.
 * - Modified header and data row generation accordingly.
 * - Previous v12 fixes maintained.
 */
async function main(workbook: ExcelScript.Workbook) {
  console.log("Starting posture summary script (v13: Concatenate Aggregation & Master Fields)...");
  const startTime = Date.now();

  // --- Overall Constants ---
  const MASTER_APP_SHEET_NAME: string = "Applications";
  const MASTER_APP_ID_HEADER: string = "UniqueID"; // This remains the key identifier
  const SUMMARY_SHEET_NAME: string = "Posture Summary";
  const CONFIG_SHEET_NAME: string = "Config";
  const CONFIG_TABLE_NAME: string = "ConfigTable";
  const DEFAULT_VALUE_MISSING: string = "";
  const CONCATENATE_SEPARATOR: string = " - "; // Separator for concatenated fields

  // --- Type Definitions ---
  // Added "Concatenate"
  type AggregationMethod = "List" | "Count" | "Sum" | "Average" | "Min" | "Max" | "UniqueList" | "Concatenate";

  type PostureSheetConfig = {
    isEnabled: boolean;
    sheetName: string;
    appIdHeaders: string[];
    dataHeadersToPull: string[];
    aggregationType: AggregationMethod;
    countByHeader?: string;
    valueHeaderForAggregation?: string;
    // Master fields requested by this specific config row (used to build the full set)
    masterFieldsForRow?: string[];
  };

  // Stores raw data fetched from posture sheets
  type PostureDataObject = {
    [appId: string]: {
      // Header -> Array of values found for that appID/header combo
      [header: string]: (string | number | boolean)[]
    }
  };

  // Stores data fetched from the Master Application sheet
  type MasterAppData = {
      [fieldName: string]: string | number | boolean;
  }
  type MasterAppDataMap = {
      [appId: string]: MasterAppData;
  }


  // --- Helper Functions ---

  /**
   * Finds the 0-based index of the first matching header in a row. Case-insensitive.
   */
  function findColumnIndex(headerRowValues: (string | number | boolean)[], possibleHeaders: string[]): number {
    for (const header of possibleHeaders) {
      if (!header) continue;
      const lowerHeader = header.toString().toLowerCase().trim();
      const index = headerRowValues.findIndex(h => h?.toString().toLowerCase().trim() === lowerHeader);
      if (index !== -1) { return index; }
    }
    return -1;
  }

  /**
   * Safely parses a value into a number. Returns null if not a valid number.
   */
  function parseNumber(value: string | number | boolean | null | undefined): number | null {
    if (value === null || typeof value === 'undefined' || value === "") { return null; }
    // Attempt to remove common currency/grouping symbols, keep decimal points and negative signs
    const cleanedValue = typeof value === 'string' ? value.replace(/[^0-9.-]+/g, "") : value;
    const num: number = Number(cleanedValue);
    return isNaN(num) ? null : num;
  }

  /**
   * Helper function to get values for a specific app and header from the main data object.
   * Returns an empty array if data is not found, simplifying downstream checks.
   */
  function getValuesFromMap(
    dataMap: PostureDataObject,
    currentAppId: string,
    headerName: string): (string | number | boolean)[] {
      const appData = dataMap[currentAppId];
      return appData?.[headerName] ?? []; // Return empty array if not found
  }


  // --- 1. Read Configuration ---
  console.log(`Reading configuration from sheet: ${CONFIG_SHEET_NAME}`);
  const configSheet = workbook.getWorksheet(CONFIG_SHEET_NAME);
  if (!configSheet) {
    console.error(`Error: Config sheet "${CONFIG_SHEET_NAME}" not found.`);
    return;
  }

  let configValues: (string | number | boolean)[][] = [];
  let configHeaderRow: (string | number | boolean)[];
  const configTable = configSheet.getTable(CONFIG_TABLE_NAME);

  try {
    if (configTable) {
      console.log(`Using table "${CONFIG_TABLE_NAME}"...`);
      const configRangeWithHeader = configTable.getHeaderRowRange().getResizedRange(configTable.getRowCount(), 0);
      configValues = await configRangeWithHeader.getValues();
      if (configValues.length <= 1) { console.log("Config table is empty or has only headers."); return; }
      configHeaderRow = configValues[0];
    } else {
      console.log(`Using used range on "${CONFIG_SHEET_NAME}"...`);
      const configRange = configSheet.getUsedRange();
      if (!configRange || configRange.getRowCount() <= 1) { console.log("Config sheet is empty or has only a header row."); return; }
      configValues = await configRange.getValues();
      configHeaderRow = configValues[0];
    }
  } catch (error) {
      console.error(`Error reading config data: ${error instanceof Error ? error.message : String(error)}`);
      return;
  }


  // Find config indices
  const colIdxIsEnabled = findColumnIndex(configHeaderRow, ["IsEnabled", "Enabled"]);
  const colIdxSheetName = findColumnIndex(configHeaderRow, ["SheetName", "Sheet Name"]);
  const colIdxAppIdHeaders = findColumnIndex(configHeaderRow, ["AppIdHeaders", "App ID Headers"]);
  const colIdxDataHeaders = findColumnIndex(configHeaderRow, ["DataHeadersToPull", "Data Headers"]);
  const colIdxAggType = findColumnIndex(configHeaderRow, ["AggregationType", "Aggregation Type"]);
  const colIdxCountBy = findColumnIndex(configHeaderRow, ["CountByHeader", "Count By"]);
  const colIdxValueHeader = findColumnIndex(configHeaderRow, ["ValueHeaderForAggregation", "Value Header"]);
  // New column index for Master App Fields
  const colIdxMasterFields = findColumnIndex(configHeaderRow, ["MasterAppFieldsToPull", "Master Fields"]);

  if ([colIdxIsEnabled, colIdxSheetName, colIdxAppIdHeaders, colIdxAggType].some(idx => idx === -1)) {
     console.error("Error: Missing one or more essential config columns: IsEnabled, SheetName, AppIdHeaders, AggregationType.");
     return;
  }
  if (colIdxDataHeaders === -1) {
    console.warn("Warning: Config column 'DataHeadersToPull' not found. 'List' and 'Concatenate' aggregations might not work as expected.");
  }
   if (colIdxMasterFields === -1) {
    console.warn("Warning: Config column 'MasterAppFieldsToPull' not found. No additional master fields will be pulled.");
  }

  // Parse config & Collect Master Fields
  const POSTURE_SHEETS_CONFIG: PostureSheetConfig[] = [];
  const uniqueMasterFields = new Set<string>(); // Collect all requested master fields
  let configIsValid = true;

  for (let i = 1; i < configValues.length; i++) { // Start from 1 to skip header
       const row = configValues[i];
       // Ensure row has enough columns before accessing indices
       if (row.length <= Math.max(colIdxIsEnabled, colIdxSheetName, colIdxAppIdHeaders, colIdxAggType, colIdxDataHeaders, colIdxCountBy, colIdxValueHeader, colIdxMasterFields)) {
          console.warn(`Warning: Config (Row ${i + 1}): Skipping row due to insufficient columns.`);
          continue;
       }

       const cleanRow = row.map(val => typeof val === 'string' ? val.trim() : val);
       const isEnabled = cleanRow[colIdxIsEnabled]?.toString().toUpperCase() === "TRUE";
       if (!isEnabled) continue;

       const sheetName = cleanRow[colIdxSheetName]?.toString() ?? "";
       const appIdHeadersRaw = cleanRow[colIdxAppIdHeaders]?.toString() ?? "";
       const dataHeadersRaw = (colIdxDataHeaders !== -1 && cleanRow[colIdxDataHeaders] !== null && typeof cleanRow[colIdxDataHeaders] !== 'undefined') ? cleanRow[colIdxDataHeaders].toString() : "";
       const aggTypeRaw = cleanRow[colIdxAggType]?.toString() ?? "List"; // Default to List
       const countByHeader = colIdxCountBy !== -1 ? cleanRow[colIdxCountBy]?.toString() ?? undefined : undefined;
       const valueHeader = colIdxValueHeader !== -1 ? cleanRow[colIdxValueHeader]?.toString() ?? undefined : undefined;
       const masterFieldsRaw = (colIdxMasterFields !== -1 && cleanRow[colIdxMasterFields] !== null && typeof cleanRow[colIdxMasterFields] !== 'undefined') ? cleanRow[colIdxMasterFields].toString() : "";

       // Basic validation for core fields
       if (!sheetName || !appIdHeadersRaw) {
          console.warn(`Warning: Config (Row ${i + 1}): Missing SheetName or AppIdHeaders. Skipping row.`);
          continue;
       }

       const appIdHeaders = appIdHeadersRaw.split(',').map(h => h.trim()).filter(h => h);
       const dataHeadersToPull = dataHeadersRaw.split(',').map(h => h.trim()).filter(h => h);
       const masterFieldsForRow = masterFieldsRaw.split(',').map(h => h.trim()).filter(h => h);

       // Add master fields from this row to the overall set
       masterFieldsForRow.forEach(field => uniqueMasterFields.add(field));

       if (appIdHeaders.length === 0) {
          console.warn(`Warning: Config (Row ${i + 1}, Sheet: ${sheetName}): AppIdHeaders empty after parsing. Skipping.`);
          continue;
       }

       // Normalize and validate aggregation type
       let aggregationType = "List" as AggregationMethod;
       const normalizedAggType = aggTypeRaw.charAt(0).toUpperCase() + aggTypeRaw.slice(1).toLowerCase();
       if (["List", "Count", "Sum", "Average", "Min", "Max", "UniqueList", "Concatenate"].includes(normalizedAggType)) {
           aggregationType = normalizedAggType as AggregationMethod;
       } else {
           console.warn(`Warning: Config (Row ${i + 1}, Sheet: ${sheetName}): Invalid AggregationType "${aggTypeRaw}". Defaulting to "List".`);
           aggregationType = "List";
       }

       // --- Aggregation Type Specific Validation ---
       let rowIsValid = true;
       const needsDataHeaders = ["List", "UniqueList", "Concatenate"];
       const needsValueHeader = ["Sum", "Average", "Min", "Max"];

       if (needsDataHeaders.includes(aggregationType) && dataHeadersToPull.length === 0) {
           console.error(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): Aggregation type '${aggregationType}' requires 'DataHeadersToPull'.`);
           rowIsValid = false;
       }
       if (aggregationType === "Count") {
           if (!countByHeader) {
               console.error(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): Aggregation type 'Count' requires 'CountByHeader'.`);
               rowIsValid = false;
           }
       } else if (needsValueHeader.includes(aggregationType)) {
           if (!valueHeader) {
               console.error(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): Aggregation type '${aggregationType}' requires 'ValueHeaderForAggregation'.`);
               rowIsValid = false;
           } else if (!dataHeadersToPull.includes(valueHeader)) {
               // Ensure the value header is actually requested to be pulled
               console.error(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): 'ValueHeaderForAggregation' ("${valueHeader}") must also be listed in 'DataHeadersToPull'.`);
               rowIsValid = false;
           }
       } else if (aggregationType === "UniqueList") {
           if (dataHeadersToPull.length > 1) {
               console.warn(`Warning: Config (Row ${i + 1}, Sheet "${sheetName}"): 'UniqueList' uses only the first header in 'DataHeadersToPull' ("${dataHeadersToPull[0]}").`);
           }
       }

       if (rowIsValid) {
            const configEntry: PostureSheetConfig = {
                isEnabled: true,
                sheetName,
                appIdHeaders,
                dataHeadersToPull,
                aggregationType,
                countByHeader,
                valueHeaderForAggregation: valueHeader,
                masterFieldsForRow // Keep track for potential future per-row logic, though we use the unique set now
            };
           POSTURE_SHEETS_CONFIG.push(configEntry);
       } else {
           configIsValid = false; // Mark overall config as invalid if any row fails critical validation
       }
  }

  if (!configIsValid) {
      console.error("Error: Configuration contains errors. Please fix the issues listed above in the Config sheet and rerun.");
      return;
  }
  if (POSTURE_SHEETS_CONFIG.length === 0) {
      console.log("No enabled and valid configurations found in the Config sheet.");
      return;
  }
  const masterFieldsToPull: string[] = Array.from(uniqueMasterFields); // Final list of master fields
  console.log(`Loaded ${POSTURE_SHEETS_CONFIG.length} posture sheet configurations.`);
  if (masterFieldsToPull.length > 0) {
      console.log(`Will attempt to pull master fields: ${masterFieldsToPull.join(', ')}`);
  }


  // --- 2. Read Master App Data (including additional fields) ---
  console.log(`Reading master App data from sheet: ${MASTER_APP_SHEET_NAME}...`);
  const masterSheet = workbook.getWorksheet(MASTER_APP_SHEET_NAME);
  if (!masterSheet) {
    console.error(`Error: Master application sheet "${MASTER_APP_SHEET_NAME}" not found.`);
    return;
  }

  const masterRange = masterSheet.getUsedRange();
  if (!masterRange) { console.log("Master sheet appears empty."); return; }
  const masterValues = await masterRange.getValues();
  if (masterValues.length <= 1) { console.log("Master sheet has only a header row or is empty."); return; }

  const masterHeaderRow = masterValues[0];
  const masterAppIdColIndex = findColumnIndex(masterHeaderRow, [MASTER_APP_ID_HEADER]);
  if (masterAppIdColIndex === -1) {
    console.error(`Error: Master App ID header "${MASTER_APP_ID_HEADER}" not found in sheet "${MASTER_APP_SHEET_NAME}".`);
    return;
  }

  // Find indices for requested master fields
  const masterFieldColIndices = new Map<string, number>();
  let allMasterFieldsFound = true;
  for (const field of masterFieldsToPull) {
      const index = findColumnIndex(masterHeaderRow, [field]);
      if (index !== -1) {
          masterFieldColIndices.set(field, index);
      } else {
          console.warn(`Warning: Requested master field "${field}" not found in sheet "${MASTER_APP_SHEET_NAME}". It will be skipped.`);
          allMasterFieldsFound = false; // Keep track, but don't stop execution
      }
  }

  // Store Master App IDs and their associated data
  const masterAppIds = new Set<string>();
  const masterAppDataMap: MasterAppDataMap = {};

  for (let i = 1; i < masterValues.length; i++) {
      const row = masterValues[i];
      // Ensure row has enough columns before accessing indices
       if (row.length <= masterAppIdColIndex || masterFieldsToPull.some(f => masterFieldColIndices.has(f) && row.length <= masterFieldColIndices.get(f)!)) {
           console.warn(`Warning: Master sheet (Row ${i + 1}): Skipping row due to insufficient columns for ID or requested fields.`);
           continue;
       }

      const appId = row[masterAppIdColIndex]?.toString().trim();
      if (appId && appId !== "") {
          masterAppIds.add(appId);

          // Store data for this App ID
          const appData: MasterAppData = {};
          masterFieldColIndices.forEach((colIndex, fieldName) => {
              appData[fieldName] = row[colIndex];
          });
          masterAppDataMap[appId] = appData;
      }
  }

  console.log(`Found ${masterAppIds.size} unique App IDs in the master list.`);
  if (masterAppIds.size === 0) {
      console.warn("Warning: No App IDs found in the master list. The summary sheet will likely be empty.");
      // Consider returning here if App IDs are essential
  }


  // --- 3. Process Posture Sheets ---
  console.log("Processing posture sheets...");
  const postureDataMap: PostureDataObject = {}; // Stores { appId: { header: [value1, value2,...] } }

  for (const config of POSTURE_SHEETS_CONFIG) {
    console.log(`Processing sheet: ${config.sheetName}...`);
    const postureSheet = workbook.getWorksheet(config.sheetName);
    if (!postureSheet) {
        console.warn(`Warning: Sheet "${config.sheetName}" not found. Skipping.`);
        continue;
    }

    const postureRange = postureSheet.getUsedRange();
    if (!postureRange || postureRange.getRowCount() <= 1) {
        console.warn(`Warning: Sheet "${config.sheetName}" is empty or has only headers. Skipping.`);
        continue;
    }

    const postureValues = await postureRange.getValues();
    const postureHeaderRow = postureValues[0];

    // Find App ID column in this posture sheet
    const appIdColIndex = findColumnIndex(postureHeaderRow, config.appIdHeaders);
    if (appIdColIndex === -1) {
        console.warn(`Warning: App ID header (tried: ${config.appIdHeaders.join(', ')}) not found in sheet "${config.sheetName}". Skipping sheet.`);
        continue;
    }

    // Find indices for all required data columns for this sheet's config
    const dataColIndicesMap = new Map<string, number>();
    let requiredHeadersAvailable = true;
    const headersToCheck = new Set<string>([
        ...config.dataHeadersToPull,
        ...(config.countByHeader ? [config.countByHeader] : []), // Add countBy if needed
        ...(config.valueHeaderForAggregation ? [config.valueHeaderForAggregation] : []) // Add valueHeader if needed
    ]);

    headersToCheck.forEach(header => {
        if (!header) return; // Skip empty header names if any
        const index = findColumnIndex(postureHeaderRow, [header]);
        if (index !== -1) {
            dataColIndicesMap.set(header, index);
        } else {
            // Is this missing header critical for the chosen aggregation?
             const isCritical =
                  (config.aggregationType === "Count" && header === config.countByHeader) ||
                  (["Sum", "Average", "Min", "Max"].includes(config.aggregationType) && header === config.valueHeaderForAggregation) ||
                  (config.aggregationType === "Concatenate" && config.dataHeadersToPull.includes(header)) || // All headers needed for concat
                  (config.aggregationType === "UniqueList" && header === config.dataHeadersToPull[0]) || // First header needed for UniqueList
                  (config.aggregationType === "List" && config.dataHeadersToPull.includes(header)); // All headers needed for List type

             if (isCritical) {
                 console.error(`Error: Critical header "${header}" required for aggregation type "${config.aggregationType}" not found in sheet "${config.sheetName}". Skipping this sheet configuration.`);
                 requiredHeadersAvailable = false;
             } else {
                  // Only warn if it's a non-critical header (e.g., listed in DataHeadersToPull but not the primary one for Count/Sum etc.)
                  if (config.dataHeadersToPull.includes(header)) {
                      console.warn(`Warning: Non-critical data column "${header}" not found in ${config.sheetName}. It won't be included for this sheet.`);
                  }
                  // If it was a countBy or valueHeader but NOT the one needed for the *current* aggregation, it's also non-critical here.
             }
        }
    });

    if (!requiredHeadersAvailable) {
      continue; // Skip this config if critical headers are missing
    }

    // Check if *any* data columns were actually found for relevant aggregation types
    if (["List", "Concatenate", "UniqueList", "Sum", "Average", "Min", "Max"].includes(config.aggregationType) && ![...dataColIndicesMap.keys()].some(key => config.dataHeadersToPull.includes(key))) {
         console.warn(`Warning: No columns specified in 'DataHeadersToPull' were found for sheet "${config.sheetName}". Skipping aggregation.`);
         continue;
    }
     if (config.aggregationType === "Count" && !dataColIndicesMap.has(config.countByHeader!)) {
          console.warn(`Warning: Column specified in 'CountByHeader' ("${config.countByHeader}") was not found for sheet "${config.sheetName}". Skipping count aggregation.`);
          continue;
     }


    // Populate the postureDataMap
    let rowsProcessed = 0;
    for (let i = 1; i < postureValues.length; i++) { // Start row 1 (data)
      const row = postureValues[i];
       // Ensure row has enough columns before accessing indices
       if (row.length <= appIdColIndex || ![...dataColIndicesMap.values()].every(idx => row.length > idx)) {
           // console.warn(`Warning: Posture sheet "${config.sheetName}" (Row ${i + 1}): Skipping row due to insufficient columns.`);
           continue; // Silently skip rows that are too short
       }

      const appId = row[appIdColIndex]?.toString().trim();

      // Only process if the App ID is valid and exists in the master list
      if (appId && masterAppIds.has(appId)) {
        if (!postureDataMap[appId]) {
          postureDataMap[appId] = {}; // Initialize object for this App ID if first time seen
        }
        const appData = postureDataMap[appId];

        // Iterate through the headers relevant to this config that were found
        dataColIndicesMap.forEach((colIndex, headerName) => {
          const value = row[colIndex];
          // Store the value if it's not null/empty/undefined
          if (value !== null && typeof value !== 'undefined' && value !== "") {
            if (!appData[headerName]) {
              appData[headerName] = []; // Initialize array for this header if first time seen for this app
            }
            appData[headerName].push(value);
          }
        });
        rowsProcessed++;
      }
    }
    console.log(`Processed ${rowsProcessed} relevant rows for sheet "${config.sheetName}".`);
  }
  console.log("Finished processing posture sheets.");


  // --- 4. Prepare and Write Summary Sheet ---
  console.log(`Preparing summary sheet: ${SUMMARY_SHEET_NAME}`);
  workbook.getWorksheet(SUMMARY_SHEET_NAME)?.delete(); // Delete existing sheet if it exists
  const summarySheet = workbook.addWorksheet(SUMMARY_SHEET_NAME);
  summarySheet.activate();

  // Generate headers: Master Fields first, then App ID, then Posture Data
  const summaryHeaders: string[] = [
      MASTER_APP_ID_HEADER, // App ID first seems more standard
      ...masterFieldsToPull // Add the requested master fields
  ];

  // Store mapping for data generation phase
  const postureColumnConfigs: { header: string, config: PostureSheetConfig, sourceValueHeader?: string }[] = [];

  POSTURE_SHEETS_CONFIG.forEach(config => {
      const aggType = config.aggregationType;
      let header = "";
      let sourceHeader: string | undefined = undefined;

      switch (aggType) {
            case "Count":
                header = `${config.countByHeader} Count (${config.sheetName})`; // Added sheet name for context
                sourceHeader = config.countByHeader;
                break;
            case "Sum":
            case "Average":
            case "Min":
            case "Max":
                header = `${config.sheetName} ${aggType} (${config.valueHeaderForAggregation})`;
                sourceHeader = config.valueHeaderForAggregation;
                break;
            case "UniqueList":
                const uniqueHeaderSource = config.dataHeadersToPull[0]; // Use first header
                header = `${uniqueHeaderSource} Unique List (${config.sheetName})`; // Added sheet name
                sourceHeader = uniqueHeaderSource;
                break;
           case "Concatenate": // NEW CASE
                header = config.sheetName; // Use sheet name as the header
                // No single source header, uses all in dataHeadersToPull
                break;
            case "List": // Original List behavior (join single column values)
            default:
                // This still creates one column per header in DataHeadersToPull for "List" type
                 config.dataHeadersToPull.forEach(originalHeader => {
                     const listHeader = `${originalHeader} (${config.sheetName})`; // Add sheet name context
                     summaryHeaders.push(listHeader);
                     // Add mapping for each generated header in "List" case
                     postureColumnConfigs.push({ header: listHeader, config: config, sourceValueHeader: originalHeader });
                 });
                return; // Skip adding a single entry below for List type
      }
      // Add the generated header (for non-List types)
      summaryHeaders.push(header);
      postureColumnConfigs.push({ header, config, sourceValueHeader: sourceHeader });
  });


  // Write header row
  if (summaryHeaders.length > 0) {
      const headerRange = summarySheet.getRangeByIndexes(0, 0, 1, summaryHeaders.length);
      await headerRange.setValues([summaryHeaders]);

      // Apply header formatting
      const headerFormat = headerRange.getFormat();
      const headerFont = headerFormat.getFont();
      const headerFill = headerFormat.getFill();
      headerFont.setBold(true);
      headerFill.setColor("#4472C4"); // Blue background
      headerFont.setColor("white"); // White text
  } else {
      console.warn("No headers generated for the summary sheet.");
      // Optional: return if no headers
  }


  // --- 4b. Generate Summary Data Rows ---
  const outputData: (string | number | boolean)[][] = [];
  const masterAppIdArray = Array.from(masterAppIds).sort(); // Sort App IDs for consistent output

  masterAppIdArray.forEach(appId => {
    const masterData = masterAppDataMap[appId] ?? {}; // Get master data, default to empty object
    // Start row with AppID and then the Master App Fields
    const row: (string | number | boolean)[] = [
        appId,
        ...masterFieldsToPull.map(field => masterData[field] ?? DEFAULT_VALUE_MISSING) // Map fields to values
    ];

    // Now process the posture sheet aggregations using the postureColumnConfigs map
    postureColumnConfigs.forEach(({ config, sourceValueHeader }) => {
        const aggType = config.aggregationType;
        let outputValue: string | number | boolean = DEFAULT_VALUE_MISSING; // Default value

        try {
            switch (aggType) {
                case "Count": {
                    const valuesToCount = getValuesFromMap(postureDataMap, appId, sourceValueHeader!);
                    if (valuesToCount.length > 0) {
                        const counts = new Map<string | number | boolean, number>();
                        for (const value of valuesToCount) {
                            counts.set(value, (counts.get(value) || 0) + 1);
                        }
                        const sortedEntries = Array.from(counts.entries())
                            .sort((a, b) => a[0].toString().localeCompare(b[0].toString()));
                        outputValue = sortedEntries.map(([value, count]) => `${value}: ${count}`).join('\n');
                    } else {
                        outputValue = 0; // Show 0 if no relevant entries found for the app
                    }
                    break;
                }
                case "Sum":
                case "Average":
                case "Min":
                case "Max": {
                    const valuesToAggregate = getValuesFromMap(postureDataMap, appId, sourceValueHeader!);
                    const numericValues: number[] = [];
                    if (valuesToAggregate) {
                        for (const value of valuesToAggregate) {
                            const num = parseNumber(value);
                            if (num !== null) {
                                numericValues.push(num);
                            }
                        }
                    }

                    if (numericValues.length > 0) {
                        if (aggType === "Sum") { outputValue = numericValues.reduce((s, c) => s + c, 0); }
                        else if (aggType === "Average") { let avg = numericValues.reduce((s, c) => s + c, 0) / numericValues.length; outputValue = parseFloat(avg.toFixed(2)); } // Keep 2 decimal places
                        else if (aggType === "Min") { outputValue = Math.min(...numericValues); }
                        else if (aggType === "Max") { outputValue = Math.max(...numericValues); }
                    }
                    // else: outputValue remains DEFAULT_VALUE_MISSING
                    break;
                }
                case "UniqueList": {
                    const valuesToList = getValuesFromMap(postureDataMap, appId, sourceValueHeader!);
                    if (valuesToList.length > 0) {
                        const uniqueValues = Array.from(new Set(valuesToList.map(v => v?.toString() ?? "")));
                        uniqueValues.sort((a, b) => a.localeCompare(b));
                        outputValue = uniqueValues.join('\n');
                    }
                    // else: outputValue remains DEFAULT_VALUE_MISSING
                    break;
                }
                // --- NEW CONCATENATE LOGIC ---
                case "Concatenate": {
                    const headersToConcat = config.dataHeadersToPull;
                    const valueLists: (string | number | boolean)[][] = headersToConcat.map(header =>
                        getValuesFromMap(postureDataMap, appId, header)
                    );

                    // Find the maximum number of entries across all pulled headers for this app/sheet
                    const maxRows = Math.max(0, ...valueLists.map(list => list.length));

                    if (maxRows > 0) {
                        const concatenatedLines: string[] = [];
                        for (let i = 0; i < maxRows; i++) {
                            const lineParts: string[] = [];
                            for (let j = 0; j < headersToConcat.length; j++) {
                                // Get the value for the j-th header at the i-th row index
                                const value = valueLists[j]?.[i] ?? ""; // Use empty string if value missing for this row/header
                                lineParts.push(value.toString());
                            }
                            // Join parts for this line, e.g., "ValueA - ValueB - ValueC"
                            concatenatedLines.push(lineParts.join(CONCATENATE_SEPARATOR));
                        }
                        // Join all concatenated lines with newline for the final cell value
                        outputValue = concatenatedLines.join('\n');
                    }
                    // else: outputValue remains DEFAULT_VALUE_MISSING
                    break;
                }
                 // --- Original List Logic (handles cases where it expanded headers) ---
                case "List": {
                     // Need to find the *specific* header this column represents
                     // We look up the `sourceValueHeader` associated with the current summary `header`
                     // Note: The header in postureColumnConfigs for "List" includes the sheet name,
                     // but sourceValueHeader is the original header name.
                     const valuesToList = getValuesFromMap(postureDataMap, appId, sourceValueHeader!);
                     if (valuesToList.length > 0) {
                         const sortedValues = valuesToList.map(v => v?.toString() ?? "").sort((a, b) => a.localeCompare(b));
                         outputValue = sortedValues.join('\n');
                     }
                     // else: outputValue remains DEFAULT_VALUE_MISSING
                     break;
                 }

            } // End switch
        } catch (e: unknown) {
            const errorMessage = e instanceof Error ? e.message : String(e);
            console.error(`Error during aggregation type "${aggType}" for App "${appId}", Sheet "${config.sheetName}", Header "${sourceValueHeader ?? config.sheetName}": ${errorMessage}`);
            outputValue = 'ERROR'; // Indicate error in the cell
        }
        row.push(outputValue); // Add the calculated value to the row
    }); // End inner postureColumnConfigs.forEach

    outputData.push(row); // Add the completed row for this App ID
  }); // End outer masterAppIdArray.forEach


  // --- 4c. Write Data ---
  if (outputData.length > 0) {
    console.log(`Writing ${outputData.length} rows of data...`);
    const dataRange = summarySheet.getRangeByIndexes(1, 0, outputData.length, summaryHeaders.length);
    await dataRange.setValues(outputData);
  } else {
     console.log(`No data rows generated for the summary.`);
  }

  // --- 5. Apply Basic Formatting ---
  const usedRange = summarySheet.getUsedRange();
  if (usedRange) {
    console.log("Applying formatting...");
    const usedRangeFormat = usedRange.getFormat();
    usedRangeFormat.setWrapText(true);
    usedRangeFormat.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
    // Autofit columns after data is written and formatting applied
    await usedRangeFormat.autofitColumns();
    console.log("Applied basic formatting and autofit columns.");
  }

  // --- Finish ---
  await summarySheet.getCell(0,0).select(); // Select A1
  const endTime = Date.now();
  const duration = (endTime - startTime) / 1000;
  console.log(`Script finished successfully in ${duration.toFixed(2)} seconds.`);

} // End main function