/**
 * Posture Summary Script (v14 - Standardized Headers & Revised Aggregation Display)
 *
 * Reads application posture data from various sheets defined in a 'Config' sheet,
 * aggregates the data based on specified methods (List, Count, Sum, Average, Min, Max, UniqueList),
 * pulls specified fields from the master application list,
 * and writes a summary report to a 'Posture Summary' sheet.
 *
 * Key changes:
 * - Posture sheet columns in the summary are now always named after the source SheetName.
 * - Removed "Concatenate" AggregationType.
 * - "List" now performs row-by-row concatenation (like former "Concatenate").
 * - "Count" now concatenates the 'DataHeadersToPull' values for each group, followed by the count (e.g., "Val1 - Val2: 5").
 * - Other aggregation types (Sum, Avg, Min, Max, UniqueList) display their result under the SheetName column.
 * - Master field pulling remains the same.
 * - Previous v12/v13 fixes maintained.
 */
async function main(workbook: ExcelScript.Workbook) {
  console.log("Starting posture summary script (v14: Standardized Headers)...");
  const startTime = Date.now();

  // --- Overall Constants ---
  const MASTER_APP_SHEET_NAME: string = "Applications";
  const MASTER_APP_ID_HEADER: string = "UniqueID";
  const SUMMARY_SHEET_NAME: string = "Posture Summary";
  const CONFIG_SHEET_NAME: string = "Config";
  const CONFIG_TABLE_NAME: string = "ConfigTable";
  const DEFAULT_VALUE_MISSING: string = "";
  const CONCATENATE_SEPARATOR: string = " - "; // Separator for concatenated fields within a row/group
  const COUNT_SEPARATOR: string = ": ";       // Separator between concatenated group and count value

  // --- Type Definitions ---
  // Removed "Concatenate"
  type AggregationMethod = "List" | "Count" | "Sum" | "Average" | "Min" | "Max" | "UniqueList";

  type PostureSheetConfig = {
    isEnabled: boolean;
    sheetName: string;
    appIdHeaders: string[];
    dataHeadersToPull: string[];
    aggregationType: AggregationMethod;
    // Optional fields based on AggregationType:
    countByHeader?: string; // Kept for potential backward compatibility check, but DataHeadersToPull is primary now for Count
    valueHeaderForAggregation?: string; // Still needed for Sum, Avg, Min, Max
    masterFieldsForRow?: string[];
  };

  // Stores raw data fetched from posture sheets: { appId: { header: [value1, value2,...] } }
  type PostureDataObject = {
    [appId: string]: {
      [header: string]: (string | number | boolean)[]
    }
  };

  // Stores data fetched from the Master Application sheet: { appId: { fieldName: value } }
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

  /**
   * Gets all values for a specific app and header from the posture data map.
   * Returns an empty array if not found.
   */
  function getValuesFromMap(
    dataMap: PostureDataObject,
    currentAppId: string,
    headerName: string): (string | number | boolean)[] {
      const appData = dataMap[currentAppId];
      return appData?.[headerName] ?? [];
  }

  /**
   * Gets values for multiple headers for a specific app, aligned row by row.
   * Returns an object containing the value lists and the maximum row count.
   */
  function getAlignedValuesForRowConcatenation(
      dataMap: PostureDataObject,
      appId: string,
      headersToGet: string[]
  ): { valueLists: (string | number | boolean)[][]; maxRows: number } {
      const valueLists: (string | number | boolean)[][] = headersToGet.map(header =>
          getValuesFromMap(dataMap, appId, header)
      );
      const maxRows = Math.max(0, ...valueLists.map(list => list.length));
      return { valueLists, maxRows };
  }


  // --- 1. Read Configuration ---
  console.log(`Reading configuration from sheet: ${CONFIG_SHEET_NAME}`);
  const configSheet = workbook.getWorksheet(CONFIG_SHEET_NAME);
  if (!configSheet) {
    console.log(`Error: Config sheet "${CONFIG_SHEET_NAME}" not found.`);
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
      console.log(`Error reading config data: ${error instanceof Error ? error.message : String(error)}`);
      return;
  }

  // Find config indices
  const colIdxIsEnabled = findColumnIndex(configHeaderRow, ["IsEnabled", "Enabled"]);
  const colIdxSheetName = findColumnIndex(configHeaderRow, ["SheetName", "Sheet Name"]);
  const colIdxAppIdHeaders = findColumnIndex(configHeaderRow, ["AppIdHeaders", "App ID Headers"]);
  const colIdxDataHeaders = findColumnIndex(configHeaderRow, ["DataHeadersToPull", "Data Headers"]);
  const colIdxAggType = findColumnIndex(configHeaderRow, ["AggregationType", "Aggregation Type"]);
  // const colIdxCountBy = findColumnIndex(configHeaderRow, ["CountByHeader", "Count By"]); // Less relevant now
  const colIdxValueHeader = findColumnIndex(configHeaderRow, ["ValueHeaderForAggregation", "Value Header"]);
  const colIdxMasterFields = findColumnIndex(configHeaderRow, ["MasterAppFieldsToPull", "Master Fields"]);

  // --- Essential Column Check ---
  const essentialCols = {
      "IsEnabled": colIdxIsEnabled,
      "SheetName": colIdxSheetName,
      "AppIdHeaders": colIdxAppIdHeaders,
      "AggregationType": colIdxAggType // Still needed to know how to process
  };
  const missingEssential = Object.entries(essentialCols)
                                 .filter(([_, index]) => index === -1)
                                 .map(([name, _]) => name);
  if (missingEssential.length > 0) {
     console.log(`Error: Missing essential config columns: ${missingEssential.join(', ')}.`);
     return;
  }
   // --- Warn for potentially missing columns needed by specific types ---
  if (colIdxDataHeaders === -1) console.log("Warning: Config column 'DataHeadersToPull' not found. 'List' and 'Count' aggregations require this.");
  if (colIdxValueHeader === -1) console.log("Warning: Config column 'ValueHeaderForAggregation' not found. 'Sum', 'Average', 'Min', 'Max' aggregations require this.");
  if (colIdxMasterFields === -1) console.log("Warning: Config column 'MasterAppFieldsToPull' not found. No additional master fields will be pulled.");

  // Parse config & Collect Master Fields
  const POSTURE_SHEETS_CONFIG: PostureSheetConfig[] = [];
  const uniqueMasterFields = new Set<string>();
  let configIsValid = true;

  for (let i = 1; i < configValues.length; i++) { // Start from 1 (data row)
       const row = configValues[i];
       // Basic check for enough columns based on essential indices
       if (row.length <= Math.max(colIdxIsEnabled, colIdxSheetName, colIdxAppIdHeaders, colIdxAggType)) {
          console.log(`Warning: Config (Row ${i + 1}): Skipping row due to insufficient columns for essential fields.`);
          continue;
       }

       const cleanRow = row.map(val => typeof val === 'string' ? val.trim() : val);
       const isEnabled = cleanRow[colIdxIsEnabled]?.toString().toUpperCase() === "TRUE";
       if (!isEnabled) continue;

       const sheetName = cleanRow[colIdxSheetName]?.toString() ?? "";
       const appIdHeadersRaw = cleanRow[colIdxAppIdHeaders]?.toString() ?? "";
       // Default AggregationType to "List" if empty or invalid
       const aggTypeRaw = cleanRow[colIdxAggType]?.toString() || "List";
       // Optional fields
       const dataHeadersRaw = (colIdxDataHeaders !== -1 && cleanRow[colIdxDataHeaders] != null) ? cleanRow[colIdxDataHeaders].toString() : "";
       const valueHeader = (colIdxValueHeader !== -1 && cleanRow[colIdxValueHeader] != null) ? cleanRow[colIdxValueHeader].toString().trim() : undefined;
       const masterFieldsRaw = (colIdxMasterFields !== -1 && cleanRow[colIdxMasterFields] != null) ? cleanRow[colIdxMasterFields].toString() : "";

       // Basic validation
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

       // Normalize and validate aggregation type
       let aggregationType = "List" as AggregationMethod; // Default to List
       const normalizedAggType = aggTypeRaw.charAt(0).toUpperCase() + aggTypeRaw.slice(1).toLowerCase();
       // Use the allowed types (Concatenate removed)
       if (["List", "Count", "Sum", "Average", "Min", "Max", "UniqueList"].includes(normalizedAggType)) {
           aggregationType = normalizedAggType as AggregationMethod;
       } else {
           console.log(`Warning: Config (Row ${i + 1}, Sheet: ${sheetName}): Invalid AggregationType "${aggTypeRaw}". Defaulting to "List".`);
           // aggregationType remains "List"
       }

       // --- Aggregation Type Specific Validation ---
       let rowIsValid = true;
       const needsDataHeaders = ["List", "Count", "UniqueList"]; // Sum/Avg etc. also need it, but specifically ValueHeader
       const needsValueHeader = ["Sum", "Average", "Min", "Max"];

       if (needsDataHeaders.includes(aggregationType) && dataHeadersToPull.length === 0) {
           console.log(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): Aggregation type '${aggregationType}' requires 'DataHeadersToPull'.`);
           rowIsValid = false;
       } else if (needsValueHeader.includes(aggregationType)) {
           if (!valueHeader) {
               console.log(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): Aggregation type '${aggregationType}' requires 'ValueHeaderForAggregation'.`);
               rowIsValid = false;
           } else if (!dataHeadersToPull.includes(valueHeader)) {
               // Value Header *must* be in the list of headers to pull
               console.log(`Error: Config (Row ${i + 1}, Sheet "${sheetName}"): 'ValueHeaderForAggregation' ("${valueHeader}") must also be listed in 'DataHeadersToPull'.`);
               rowIsValid = false;
           }
       } else if (aggregationType === "UniqueList" && dataHeadersToPull.length > 1) {
           console.log(`Warning: Config (Row ${i + 1}, Sheet "${sheetName}"): 'UniqueList' uses only the first header in 'DataHeadersToPull' ("${dataHeadersToPull[0]}").`);
       }
       // No specific validation needed for 'List' beyond requiring DataHeadersToPull (checked above)
       // 'Count' now primarily relies on DataHeadersToPull (checked above)

       if (rowIsValid) {
            const configEntry: PostureSheetConfig = {
                isEnabled: true,
                sheetName,
                appIdHeaders,
                dataHeadersToPull,
                aggregationType,
                // countByHeader: countByHeader, // Not strictly needed for new logic
                valueHeaderForAggregation: valueHeader,
                masterFieldsForRow
            };
           POSTURE_SHEETS_CONFIG.push(configEntry);
       } else {
           configIsValid = false;
       }
  }

  if (!configIsValid) {
      console.log("Error: Configuration contains errors. Please fix the issues listed above in the Config sheet and rerun.");
      return;
  }
  if (POSTURE_SHEETS_CONFIG.length === 0) {
      console.log("No enabled and valid configurations found in the Config sheet.");
      return;
  }
  const masterFieldsToPull: string[] = Array.from(uniqueMasterFields);
  console.log(`Loaded ${POSTURE_SHEETS_CONFIG.length} posture sheet configurations.`);
  if (masterFieldsToPull.length > 0) console.log(`Will attempt to pull master fields: ${masterFieldsToPull.join(', ')}`);


  // --- 2. Read Master App Data ---
  console.log(`Reading master App data from sheet: ${MASTER_APP_SHEET_NAME}...`);
  const masterSheet = workbook.getWorksheet(MASTER_APP_SHEET_NAME);
  if (!masterSheet) {
    console.log(`Error: Master application sheet "${MASTER_APP_SHEET_NAME}" not found.`); return;
  }
  const masterRange = masterSheet.getUsedRange();
  if (!masterRange) { console.log("Master sheet appears empty."); return; }
  const masterValues = await masterRange.getValues();
  if (masterValues.length <= 1) { console.log("Master sheet has only a header row or is empty."); return; }

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
      if (row.length <= masterAppIdColIndex) continue; // Skip short rows
      const appId = row[masterAppIdColIndex]?.toString().trim();
      if (appId && appId !== "") {
          masterAppIds.add(appId);
          const appData: MasterAppData = {};
          masterFieldColIndices.forEach((colIndex, fieldName) => {
              if (row.length > colIndex) appData[fieldName] = row[colIndex];
          });
          masterAppDataMap[appId] = appData;
      }
  }
  console.log(`Found ${masterAppIds.size} unique App IDs in the master list.`);
  if (masterAppIds.size === 0) console.log("Warning: No App IDs found in the master list.");


  // --- 3. Process Posture Sheets ---
  console.log("Processing posture sheets...");
  const postureDataMap: PostureDataObject = {};

  for (const config of POSTURE_SHEETS_CONFIG) {
    console.log(`Processing sheet: ${config.sheetName}...`);
    const postureSheet = workbook.getWorksheet(config.sheetName);
    if (!postureSheet) { console.log(`Warning: Sheet "${config.sheetName}" not found. Skipping.`); continue; }
    const postureRange = postureSheet.getUsedRange();
    if (!postureRange || postureRange.getRowCount() <= 1) { console.log(`Warning: Sheet "${config.sheetName}" is empty or has only headers. Skipping.`); continue; }

    const postureValues = await postureRange.getValues();
    const postureHeaderRow = postureValues[0];
    const appIdColIndex = findColumnIndex(postureHeaderRow, config.appIdHeaders);
    if (appIdColIndex === -1) { console.log(`Warning: App ID header (tried: ${config.appIdHeaders.join(', ')}) not found in sheet "${config.sheetName}". Skipping sheet.`); continue; }

    // Find indices for *all* headers mentioned in the config for this sheet
    const dataColIndicesMap = new Map<string, number>();
    let requiredHeadersAvailable = true;
    const headersRequiredForThisConfig = new Set<string>([
        ...config.dataHeadersToPull, // Always need these if specified
        ...(config.valueHeaderForAggregation ? [config.valueHeaderForAggregation] : []) // Add value header if specified
    ]);

    headersRequiredForThisConfig.forEach(header => {
        if (!header) return;
        const index = findColumnIndex(postureHeaderRow, [header]);
        if (index !== -1) {
            dataColIndicesMap.set(header, index);
        } else {
            // Check if this *missing* header is absolutely critical for the *current* config's aggregation type
             let isCritical = false;
             if ((config.aggregationType === 'List' || config.aggregationType === 'Count' || config.aggregationType === 'UniqueList') && config.dataHeadersToPull.includes(header)) {
                // For these types, *all* listed DataHeadersToPull are considered essential for the intended output
                isCritical = true;
             } else if (['Sum', 'Average', 'Min', 'Max'].includes(config.aggregationType) && header === config.valueHeaderForAggregation) {
                 // For these types, the specific ValueHeaderForAggregation is critical
                isCritical = true;
             }
             // Note: A header might be listed in dataHeadersToPull but *not* be the valueHeaderForAggregation.
             // In the Sum/Avg/Min/Max case, missing such a header is not critical *for the calculation*,
             // although it means it wasn't fetched (which might be unexpected but not fatal).

             if (isCritical) {
                 console.log(`Error: Critical header "${header}" required for aggregation type "${config.aggregationType}" in sheet "${config.sheetName}" not found. Skipping this sheet configuration.`);
                 requiredHeadersAvailable = false;
             } else {
                  // Warn if it was requested in DataHeadersToPull but not found (and wasn't critical)
                  if (config.dataHeadersToPull.includes(header)) {
                       console.log(`Warning: Non-critical data column "${header}" requested in DataHeadersToPull not found in ${config.sheetName}. It won't be included in concatenation/grouping.`);
                  }
             }
        }
    });

    if (!requiredHeadersAvailable) continue; // Skip this config if critical headers missing

    // Final check: Do we have the necessary columns based on the indices we *did* find?
    let columnsAvailableForProcessing = true;
    if ((config.aggregationType === 'List' || config.aggregationType === 'Count') && !config.dataHeadersToPull.every(h => dataColIndicesMap.has(h))) {
        console.log(`Warning: Not all headers in 'DataHeadersToPull' for Sheet "${config.sheetName}" (Type: ${config.aggregationType}) were found. Results may be incomplete. Skipping this configuration.`);
        columnsAvailableForProcessing = false;
    } else if (config.aggregationType === 'UniqueList' && !dataColIndicesMap.has(config.dataHeadersToPull[0])) {
        console.log(`Warning: The first header ("${config.dataHeadersToPull[0]}") in 'DataHeadersToPull' for Sheet "${config.sheetName}" (Type: UniqueList) was not found. Skipping.`);
        columnsAvailableForProcessing = false;
    } else if (['Sum', 'Average', 'Min', 'Max'].includes(config.aggregationType) && !dataColIndicesMap.has(config.valueHeaderForAggregation!)) {
        console.log(`Warning: The 'ValueHeaderForAggregation' ("${config.valueHeaderForAggregation}") for Sheet "${config.sheetName}" (Type: ${config.aggregationType}) was not found. Skipping.`);
        columnsAvailableForProcessing = false;
    }

    if (!columnsAvailableForProcessing) continue;


    // Populate the postureDataMap
    let rowsProcessed = 0;
    for (let i = 1; i < postureValues.length; i++) {
      const row = postureValues[i];
      if (row.length <= appIdColIndex) continue; // Skip short rows

      const appId = row[appIdColIndex]?.toString().trim();
      if (appId && masterAppIds.has(appId)) {
        if (!postureDataMap[appId]) postureDataMap[appId] = {};
        const appData = postureDataMap[appId];

        // Store values for all headers found for this config
        dataColIndicesMap.forEach((colIndex, headerName) => {
           if (row.length > colIndex) { // Check row length again for this specific column
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
    console.log(`Processed ${rowsProcessed} relevant rows for sheet "${config.sheetName}".`);
  }
  console.log("Finished processing posture sheets.");


  // --- 4. Prepare and Write Summary Sheet ---
  console.log(`Preparing summary sheet: ${SUMMARY_SHEET_NAME}`);
  workbook.getWorksheet(SUMMARY_SHEET_NAME)?.delete();
  const summarySheet = workbook.addWorksheet(SUMMARY_SHEET_NAME);
  summarySheet.activate();

  // Generate headers: Master Fields first, then App ID, then one column per Posture Sheet config
  const summaryHeaders: string[] = [
      MASTER_APP_ID_HEADER,
      ...masterFieldsToPull
  ];
  // Use a simple map to track headers added to prevent duplicates from multiple configs for the *same* sheet
  const addedPostureHeaders = new Set<string>();
  // Store mapping for data generation phase: Header Name -> Config Object
  const postureColumnConfigMap = new Map<string, PostureSheetConfig>();

  POSTURE_SHEETS_CONFIG.forEach(config => {
      const header = config.sheetName; // Use SheetName directly as header
      // Only add the header and config mapping if we haven't already processed this sheet name
      if (!addedPostureHeaders.has(header)) {
          summaryHeaders.push(header);
          postureColumnConfigMap.set(header, config);
          addedPostureHeaders.add(header);
      } else {
          // If the same sheet name appears again (e.g., different aggregation on the same sheet),
          // we currently ignore subsequent entries because we can only have one column per name.
          // Log a warning. A more advanced approach might append "(1)", "(2)" etc.
           console.log(`Warning: Configuration for SheetName "${header}" appears multiple times in Config. Only the first configuration row encountered for this sheet name will be used in the summary.`);
      }
  });

  // Write header row
  if (summaryHeaders.length > 0) {
      const headerRange = summarySheet.getRangeByIndexes(0, 0, 1, summaryHeaders.length);
      await headerRange.setValues([summaryHeaders]);
      // Apply header formatting
      const headerFormat = headerRange.getFormat();
      headerFont = headerFormat.getFont(); // Re-get font object
      headerFill = headerFormat.getFill(); // Re-get fill object
      headerFont.setBold(true);
      headerFill.setColor("#4472C4");
      headerFont.setColor("white");
  } else {
      console.log("No headers generated for the summary sheet.");
  }


  // --- 4b. Generate Summary Data Rows ---
  const outputData: (string | number | boolean)[][] = [];
  const masterAppIdArray = Array.from(masterAppIds).sort();

  masterAppIdArray.forEach(appId => {
    const masterData = masterAppDataMap[appId] ?? {};
    const row: (string | number | boolean)[] = [
        appId,
        ...masterFieldsToPull.map(field => masterData[field] ?? DEFAULT_VALUE_MISSING)
    ];

    // Process posture data using the simplified headers
    postureColumnConfigMap.forEach((config, headerName) => { // Iterate Map(Header -> Config)
        const aggType = config.aggregationType;
        let outputValue: string | number | boolean = DEFAULT_VALUE_MISSING;

        try {
            switch (aggType) {
                // --- NEW Count Logic ---
                case "Count": {
                    const headersToGroup = config.dataHeadersToPull;
                    const { valueLists, maxRows } = getAlignedValuesForRowConcatenation(postureDataMap, appId, headersToGroup);

                    if (maxRows > 0) {
                        const groupCounts = new Map<string, number>();
                        // Internal separator unlikely to appear in real data
                        const internalSep = "|||---|||";

                        for (let i = 0; i < maxRows; i++) {
                            // Create a unique key for this row's combination of values
                            const keyParts: string[] = [];
                            for (let j = 0; j < headersToGroup.length; j++) {
                                const value = valueLists[j]?.[i] ?? "";
                                keyParts.push(value.toString());
                            }
                            const groupKey = keyParts.join(internalSep); // Use internal separator for map key
                            groupCounts.set(groupKey, (groupCounts.get(groupKey) || 0) + 1);
                        }

                        // Format output: "Val1 - Val2: Count"
                         const formattedEntries: string[] = [];
                         // Sort by the concatenated key representation for consistent output
                         const sortedKeys = Array.from(groupCounts.keys()).sort();

                         sortedKeys.forEach(groupKey => {
                             const count = groupCounts.get(groupKey)!;
                             // Reconstruct the display string using the user-facing separator
                             const displayKey = groupKey.split(internalSep).join(CONCATENATE_SEPARATOR);
                             formattedEntries.push(`${displayKey}${COUNT_SEPARATOR}${count}`);
                         });

                        outputValue = formattedEntries.join('\n');
                    } else {
                        outputValue = 0; // Or perhaps "" or "No entries found"? Let's stick with 0 count.
                    }
                    break;
                }
                // --- NEW List Logic (Row-wise Concatenation) ---
                case "List": {
                    const headersToConcat = config.dataHeadersToPull;
                    const { valueLists, maxRows } = getAlignedValuesForRowConcatenation(postureDataMap, appId, headersToConcat);

                    if (maxRows > 0) {
                        const concatenatedLines: string[] = [];
                        for (let i = 0; i < maxRows; i++) {
                            const lineParts: string[] = [];
                            for (let j = 0; j < headersToConcat.length; j++) {
                                const value = valueLists[j]?.[i] ?? "";
                                lineParts.push(value.toString());
                            }
                            concatenatedLines.push(lineParts.join(CONCATENATE_SEPARATOR));
                        }
                        outputValue = concatenatedLines.join('\n');
                    }
                    // else: outputValue remains DEFAULT_VALUE_MISSING
                    break;
                }
                // --- Sum/Avg/Min/Max --- (Logic unchanged, just target column is SheetName)
                case "Sum":
                case "Average":
                case "Min":
                case "Max": {
                    const valueHeader = config.valueHeaderForAggregation!;
                    const valuesToAggregate = getValuesFromMap(postureDataMap, appId, valueHeader);
                    const numericValues: number[] = [];
                    if (valuesToAggregate) {
                        for (const value of valuesToAggregate) {
                            const num = parseNumber(value);
                            if (num !== null) numericValues.push(num);
                        }
                    }
                    if (numericValues.length > 0) {
                        if (aggType === "Sum") outputValue = numericValues.reduce((s, c) => s + c, 0);
                        else if (aggType === "Average") { let avg = numericValues.reduce((s, c) => s + c, 0) / numericValues.length; outputValue = parseFloat(avg.toFixed(2)); }
                        else if (aggType === "Min") outputValue = Math.min(...numericValues);
                        else if (aggType === "Max") outputValue = Math.max(...numericValues);
                    }
                    // else: outputValue remains DEFAULT_VALUE_MISSING
                    break;
                }
                // --- UniqueList --- (Logic unchanged, just target column is SheetName)
                case "UniqueList": {
                    const headerForUniqueList = config.dataHeadersToPull[0]; // Use first header
                    const valuesToList = getValuesFromMap(postureDataMap, appId, headerForUniqueList);
                    if (valuesToList.length > 0) {
                        const uniqueValues = Array.from(new Set(valuesToList.map(v => v?.toString() ?? "")));
                        uniqueValues.sort((a, b) => a.localeCompare(b));
                        outputValue = uniqueValues.join('\n');
                    }
                    // else: outputValue remains DEFAULT_VALUE_MISSING
                    break;
                }
            } // End switch
        } catch (e: unknown) {
            const errorMessage = e instanceof Error ? e.message : String(e);
            console.log(`Error during aggregation type "${aggType}" for App "${appId}", Sheet "${config.sheetName}": ${errorMessage}`);
            outputValue = 'ERROR';
        }
        row.push(outputValue); // Add the calculated value for this SheetName column
    }); // End inner postureColumnConfigMap.forEach

    outputData.push(row);
  }); // End outer masterAppIdArray.forEach


  // --- 4c. Write Data ---
  if (outputData.length > 0) {
    console.log(`Writing ${outputData.length} rows of data...`);
    // Ensure data range matches the actual number of headers generated
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
    await usedRangeFormat.autofitColumns(); // Autofit columns
    console.log("Applied basic formatting and autofit columns.");
  }

  // --- Finish ---
  await summarySheet.getCell(0,0).select();
  const endTime = Date.now();
  const duration = (endTime - startTime) / 1000;
  console.log(`Script finished successfully in ${duration.toFixed(2)} seconds.`);

} // End main function


// Declare helper variables used in formatting that might be flagged otherwise
let headerFont: ExcelScript.RangeFont;
let headerFill: ExcelScript.RangeFill;