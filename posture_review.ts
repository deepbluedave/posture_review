/**
 * Posture Summary Script (v10 - Moved Nested Function Definition)
 *
 * Reads application posture data from various sheets defined in a 'Config' sheet,
 * aggregates the data based on specified methods (List, Count, Sum, Average, Min, Max, UniqueList),
 * and writes a summary report to a 'Posture Summary' sheet.
 * Conditional formatting has been removed.
 *
 * Key changes:
 * - Moved the nested 'getValues' helper function outside the main data processing loop
 *   to comply with Office Scripts restrictions on nested function declarations.
 * - Confirmed necessary 'await' usage.
 * - Verified all array method callbacks use arrow functions.
 * - Addressed other previous compatibility fixes.
 */
async function main(workbook: ExcelScript.Workbook) {
    console.log("Starting posture summary script (v10: Moved Nested Function Definition)...");
    const startTime = Date.now();
  
    // --- Overall Constants ---
    const MASTER_APP_SHEET_NAME: string = "Applications";
    const MASTER_APP_ID_HEADER: string = "UniqueID";
    const SUMMARY_SHEET_NAME: string = "Posture Summary";
    const CONFIG_SHEET_NAME: string = "Config";
    const CONFIG_TABLE_NAME: string = "ConfigTable";
    const DEFAULT_VALUE_MISSING: string = "";
  
    // --- Type Definitions ---
    type AggregationMethod = "List" | "Count" | "Sum" | "Average" | "Min" | "Max" | "UniqueList";
  
    type PostureSheetConfig = {
      isEnabled: boolean;
      sheetName: string;
      appIdHeaders: string[];
      dataHeadersToPull: string[];
      aggregationType: AggregationMethod;
      countByHeader?: string;
      valueHeaderForAggregation?: string;
    };
  
    type PostureDataObject = {
      [appId: string]: {
        [header: string]: (string | number | boolean)[]
      }
    };
  
    // --- Helper Functions ---
  
    /**
     * Finds the 0-based index of the first matching header in a row. Case-insensitive search.
     */
    function findColumnIndex(headerRowValues: (string | number | boolean)[], possibleHeaders: string[]): number {
      // Synchronous
      for (const header of possibleHeaders) {
        if (!header) continue;
        const lowerHeader = header.toString().toLowerCase();
        const index = headerRowValues.findIndex(h => h?.toString().toLowerCase() === lowerHeader); // Arrow func OK
        if (index !== -1) {
          return index;
        }
      }
      return -1;
    }
  
    /**
     * Safely parses a value into a number.
     */
    function parseNumber(value: string | number | boolean | null | undefined): number | null {
      // Synchronous
      if (value === null || typeof value === 'undefined' || value === "") {
        return null;
      }
      const cleanedValue = typeof value === 'string' ? value.replace(/[^0-9.-]+/g, "") : value;
      const num: number = Number(cleanedValue);
      return isNaN(num) ? null : num;
    }
  
    /**
     * Helper function to get values for a specific app and header from the main data object.
     * Moved outside the loop to comply with Office Scripts restrictions.
     */
    function getValuesFromMap(
      dataMap: PostureDataObject,
      currentAppId: string,
      headerName: string): (string | number | boolean)[] | undefined {
        // Synchronous
        const appData = dataMap[currentAppId];
        return appData?.[headerName];
    }
  
  
    // --- 1. Read Configuration ---
    console.log(`Reading configuration from sheet: ${CONFIG_SHEET_NAME}`);
    const configSheet = workbook.getWorksheet(CONFIG_SHEET_NAME); // Sync
    if (!configSheet) { /* ... error handling ... */ console.log("Config sheet not found"); return; }
  
    let configValues: (string | number | boolean)[][] = [];
    let configHeaderRow: (string | number | boolean)[];
    const configTable = configSheet.getTable(CONFIG_TABLE_NAME); // Sync
  
    if (configTable) {
      console.log(`Using table "${CONFIG_TABLE_NAME}"...`);
      const configRangeWithHeader = configTable.getHeaderRowRange().getResizedRange(configTable.getRowCount(), 0); // Sync
      configValues = await configRangeWithHeader.getValues(); // Async - Await REQUIRED
      if (configValues.length <= 1) { /* ... error handling ... */ console.log("Config table empty"); return; }
      configHeaderRow = configValues[0];
    } else {
      console.log(`Using used range on "${CONFIG_SHEET_NAME}"...`);
      const configRange = configSheet.getUsedRange(); // Sync
      if (!configRange || configRange.getRowCount() <= 1) { /* ... error handling ... */ console.log("Config sheet empty"); return; }
      configValues = await configRange.getValues(); // Async - Await REQUIRED
      configHeaderRow = configValues[0];
    }
  
    // Find config indices (Sync)
    const colIdxIsEnabled = findColumnIndex(configHeaderRow, ["IsEnabled"]);
    // ... other findColumnIndex calls ...
    const colIdxSheetName = findColumnIndex(configHeaderRow, ["SheetName"]);
    const colIdxAppIdHeaders = findColumnIndex(configHeaderRow, ["AppIdHeaders", "App ID Headers"]);
    const colIdxDataHeaders = findColumnIndex(configHeaderRow, ["DataHeadersToPull", "Data Headers"]);
    const colIdxAggType = findColumnIndex(configHeaderRow, ["AggregationType", "Aggregation Type"]);
    const colIdxCountBy = findColumnIndex(configHeaderRow, ["CountByHeader", "Count By"]);
    const colIdxValueHeader = findColumnIndex(configHeaderRow, ["ValueHeaderForAggregation", "Value Header"]);
  
  
    // Validate essential columns (Sync, uses arrow func callback)
    if ([colIdxIsEnabled, colIdxSheetName, colIdxAppIdHeaders, colIdxDataHeaders, colIdxAggType].some(idx => idx === -1)) {
       /* ... error handling ... */ console.log("Missing essential config columns"); return;
    }
  
    // Parse config (Sync loop, uses arrow funcs in map/filter)
    const POSTURE_SHEETS_CONFIG: PostureSheetConfig[] = [];
    let configIsValid = true;
    for (let i = 1; i < configValues.length; i++) {
        // ... parsing logic using map/filter with arrow functions ...
        // ... validation logic ...
         const row = configValues[i];
         const cleanRow = row.map(val => typeof val === 'string' ? val.trim() : val); // Arrow func
         const isEnabled = cleanRow[colIdxIsEnabled]?.toString().toUpperCase() === "TRUE";
         if (!isEnabled) continue;
         // ... (rest of parsing as before) ...
         const sheetName = cleanRow[colIdxSheetName]?.toString() ?? "";
         const appIdHeadersRaw = cleanRow[colIdxAppIdHeaders]?.toString() ?? "";
         const dataHeadersRaw = cleanRow[colIdxDataHeaders]?.toString() ?? "";
         const aggTypeRaw = cleanRow[colIdxAggType]?.toString() ?? "List";
         const countByHeader = colIdxCountBy !== -1 ? cleanRow[colIdxCountBy]?.toString() ?? undefined : undefined;
         const valueHeader = colIdxValueHeader !== -1 ? cleanRow[colIdxValueHeader]?.toString() ?? undefined : undefined;
  
          if (!sheetName || !appIdHeadersRaw || (!dataHeadersRaw && aggTypeRaw !== "Count")) {
              console.log(`Warning: Config (Row ${i + 1}): Skipping row due to missing SheetName, AppIdHeaders, or DataHeadersToPull (required unless AggregationType is Count).`);
              continue;
          }
         const appIdHeaders = appIdHeadersRaw.split(',').map(h => h.trim()).filter(h => h); // Arrow funcs
         const dataHeadersToPull = dataHeadersRaw.split(',').map(h => h.trim()).filter(h => h); // Arrow funcs
  
          if (appIdHeaders.length === 0) {
              console.log(`Warning: Config (Row ${i + 1}, Sheet: ${sheetName}): AppIdHeaders field contains no valid header names after parsing. Skipping.`);
              continue;
          }
  
         let aggregationType = "List" as AggregationMethod;
         const normalizedAggType = aggTypeRaw.charAt(0).toUpperCase() + aggTypeRaw.slice(1).toLowerCase();
         if (["List", "Count", "Sum", "Average", "Min", "Max", "UniqueList"].includes(normalizedAggType)) {
           aggregationType = normalizedAggType as AggregationMethod;
         } else {
           console.log(`Warning: Config (Row ${i + 1}, Sheet: ${sheetName}): Invalid AggregationType "${aggTypeRaw}". Defaulting to "List".`);
           aggregationType = "List";
         }
         const configEntry: PostureSheetConfig = { isEnabled: true, sheetName, appIdHeaders, dataHeadersToPull, aggregationType, countByHeader, valueHeaderForAggregation: valueHeader };
  
         let rowIsValid = true;
          // ... (validation logic as before) ...
          if (aggregationType === "Count") {
              if (!countByHeader) { console.log(`Error: Config (Row ${i+1}, Sheet "${sheetName}"): AggregationType "Count" requires 'CountByHeader'.`); rowIsValid = false; }
          } else if (["Sum", "Average", "Min", "Max"].includes(aggregationType)) {
              if (!valueHeader) { console.log(`Error: Config (Row ${i+1}, Sheet "${sheetName}"): AggregationType "${aggregationType}" requires 'ValueHeaderForAggregation'.`); rowIsValid = false; }
              else if (!dataHeadersToPull.includes(valueHeader)) { console.log(`Error: Config (Row ${i+1}, Sheet "${sheetName}"): 'ValueHeaderForAggregation' ("${valueHeader}") must also be listed in 'DataHeadersToPull'.`); rowIsValid = false; }
          } else if (aggregationType === "UniqueList") {
               if (dataHeadersToPull.length === 0) { console.log(`Error: Config (Row ${i+1}, Sheet "${sheetName}"): AggregationType "UniqueList" requires at least one header in 'DataHeadersToPull'.`); rowIsValid = false; }
               if (dataHeadersToPull.length > 1) { console.log(`Warning: Config (Row ${i+1}, Sheet "${sheetName}"): AggregationType "UniqueList" uses only the first header in 'DataHeadersToPull' ("${dataHeadersToPull[0]}"). Others ignored.`);}
          } else if (aggregationType === "List") {
              if (dataHeadersToPull.length === 0) { console.log(`Error: Config (Row ${i+1}, Sheet "${sheetName}"): AggregationType "List" requires at least one header in 'DataHeadersToPull'.`); rowIsValid = false; }
          }
  
  
         if (rowIsValid) { POSTURE_SHEETS_CONFIG.push(configEntry); }
         else { configIsValid = false; }
    }
    if (!configIsValid) { /* ... error handling ... */ console.log("Config invalid"); return; }
    if (POSTURE_SHEETS_CONFIG.length === 0) { /* ... warning ... */ console.log("No valid configs"); }
    console.log(`Loaded ${POSTURE_SHEETS_CONFIG.length} configurations.`);
  
  
    // --- 2. Read Master App IDs ---
    console.log(`Reading master App IDs from sheet: ${MASTER_APP_SHEET_NAME}...`);
    const masterSheet = workbook.getWorksheet(MASTER_APP_SHEET_NAME); // Sync
    if (!masterSheet) { /* ... error handling ... */ console.log("Master sheet not found"); return; }
    const masterRange = masterSheet.getUsedRange(); // Sync
    if (!masterRange) { /* ... error handling ... */ console.log("Master sheet empty"); return; }
    const masterValues = await masterRange.getValues(); // Async - Await REQUIRED
    if (masterValues.length <= 1) { /* ... error handling ... */ console.log("Master sheet header only"); return; }
    const masterHeaderRow = masterValues[0];
    const masterAppIdColIndex = findColumnIndex(masterHeaderRow, [MASTER_APP_ID_HEADER]); // Sync
    if (masterAppIdColIndex === -1) { /* ... error handling ... */ console.log("Master App ID header not found"); return; }
    const masterAppIds = new Set<string>(); // Sync
    for (let i = 1; i < masterValues.length; i++) { /* ... sync loop ... */ masterAppIds.add(masterValues[i][masterAppIdColIndex]?.toString().trim()); }
    console.log(`Found ${masterAppIds.size} unique App IDs.`);
    if (masterAppIds.size === 0) { /* ... warning ... */ }
  
  
    // --- 3. Process Posture Sheets ---
    console.log("Processing posture sheets...");
    const postureDataMap: PostureDataObject = {}; // Sync init
  
    for (const config of POSTURE_SHEETS_CONFIG) { // Sync loop, allows await inside
      console.log(`Processing sheet: ${config.sheetName}...`);
      const postureSheet = workbook.getWorksheet(config.sheetName); // Sync
      if (!postureSheet) { /* ... warning ... */ continue; }
      const postureRange = postureSheet.getUsedRange(); // Sync
      if (!postureRange || postureRange.getRowCount() <= 1) { /* ... warning ... */ continue; }
      const postureValues = await postureRange.getValues(); // Async - Await REQUIRED
      const postureHeaderRow = postureValues[0];
      const appIdColIndex = findColumnIndex(postureHeaderRow, config.appIdHeaders); // Sync
      if (appIdColIndex === -1) { /* ... warning ... */ continue; }
  
      // Find indices (Sync, uses arrow funcs in forEach)
      const dataColIndicesMap = new Map<string, number>();
      let requiredHeaderMissing = false;
      config.dataHeadersToPull.forEach(header => { /* ... find index logic ... */
          const index = findColumnIndex(postureHeaderRow, [header]);
          if (index !== -1) { dataColIndicesMap.set(header, index); }
          else {
              console.log(`Warning: Data column "${header}" not found in sheet "${config.sheetName}".`);
              // ... check if critical header ...
              if ((["Sum", "Average", "Min", "Max"].includes(config.aggregationType) && header === config.valueHeaderForAggregation) ||
                  (config.aggregationType === "UniqueList" && header === config.dataHeadersToPull[0] && config.dataHeadersToPull.length === 1))
              {
                  console.log(`Error: Critical header "${header}" missing for ${config.aggregationType}. Skipping sheet.`);
                  requiredHeaderMissing = true;
              }
          }
      });
      if (requiredHeaderMissing) continue;
  
      if (config.aggregationType === "Count" && config.countByHeader) { /* ... find CountBy index logic ... */
          const countByColIndex = findColumnIndex(postureHeaderRow, [config.countByHeader]);
          if (countByColIndex === -1) { requiredHeaderMissing = true; console.log(`Error: CountByHeader "${config.countByHeader}" missing.`);}
          else if (!dataColIndicesMap.has(config.countByHeader)) { dataColIndicesMap.set(config.countByHeader, countByColIndex); }
      }
       if (requiredHeaderMissing) continue;
  
      if (["Sum", "Average", "Min", "Max"].includes(config.aggregationType) && config.valueHeaderForAggregation) { /* ... check ValueHeader exists ... */
          if (!dataColIndicesMap.has(config.valueHeaderForAggregation)) { requiredHeaderMissing = true; console.log(`Error: ValueHeader "${config.valueHeaderForAggregation}" missing from map.`);}
      }
      if (requiredHeaderMissing) continue;
  
  
      if (dataColIndicesMap.size === 0) { /* ... warning ... */ continue; }
  
      // Populate map (Sync loop, uses arrow func in forEach)
      let rowsProcessed = 0;
      for (let i = 1; i < postureValues.length; i++) {
        const row = postureValues[i];
        const appId = row[appIdColIndex]?.toString().trim();
        if (appId && masterAppIds.has(appId)) {
          if (!postureDataMap[appId]) { postureDataMap[appId] = {}; }
          const appData = postureDataMap[appId];
          dataColIndicesMap.forEach((colIndex, headerName) => {
            const value = row[colIndex];
            if (value !== null && typeof value !== 'undefined' && value !== "") {
              if (!appData[headerName]) { appData[headerName] = []; }
              appData[headerName].push(value);
            }
          });
          rowsProcessed++;
        }
      }
      console.log(`Processed ${rowsProcessed} rows for ${config.sheetName}.`);
    }
    console.log("Finished processing posture sheets.");
  
  
    // --- 4. Prepare and Write Summary Sheet ---
    console.log(`Preparing summary sheet: ${SUMMARY_SHEET_NAME}`);
    workbook.getWorksheet(SUMMARY_SHEET_NAME)?.delete(); // Sync
    const summarySheet = workbook.addWorksheet(SUMMARY_SHEET_NAME); // Sync
    summarySheet.activate(); // Sync
  
    // Generate headers (Sync, uses arrow funcs in forEach)
    const summaryHeaders: string[] = [MASTER_APP_ID_HEADER];
    const headerConfigMapping: { header: string, config: PostureSheetConfig, sourceValueHeader?: string }[] = [];
    POSTURE_SHEETS_CONFIG.forEach(config => { /* ... switch case logic ... */
        const aggType = config.aggregationType;
        switch (aggType) {
              case "Count": { const header = `${config.countByHeader} Count Summary`; summaryHeaders.push(header); headerConfigMapping.push({ header, config, sourceValueHeader: config.countByHeader }); break; }
              case "Sum": case "Average": case "Min": case "Max": { const header = `${config.sheetName} ${aggType} (${config.valueHeaderForAggregation})`; summaryHeaders.push(header); headerConfigMapping.push({ header, config, sourceValueHeader: config.valueHeaderForAggregation }); break; }
              case "UniqueList": { const uniqueHeaderSource = config.dataHeadersToPull[0]; const header = `${config.sheetName} Unique ${uniqueHeaderSource}`; summaryHeaders.push(header); headerConfigMapping.push({ header, config, sourceValueHeader: uniqueHeaderSource }); break; }
              case "List": default: { config.dataHeadersToPull.forEach(originalHeader => { summaryHeaders.push(originalHeader); headerConfigMapping.push({ header: originalHeader, config }); }); break; }
        }
    });
  
    // Get range/format objects (Sync)
    const headerRange = summarySheet.getRangeByIndexes(0, 0, 1, summaryHeaders.length);
    const headerFormat = headerRange.getFormat();
    const headerFont = headerFormat.getFont();
    const headerFill = headerFormat.getFill();
  
    // Write header values (Async - Await REQUIRED)
    await headerRange.setValues([summaryHeaders]);
  
    // Set header format (Sync - No await)
    headerFont.setBold(true);
    headerFill.setColor("#4472C4");
    headerFont.setColor("white");
  
    // --- 4b. Generate Summary Data Rows ---
    const outputData: (string | number | boolean)[][] = []; // Sync init
    const masterAppIdArray = Array.from(masterAppIds).sort(); // Sync
  
    // Process data (Sync loop, uses arrow funcs)
    // The loop where the error was previously reported (around L454)
    masterAppIdArray.forEach(appId => { // Outer forEach (arrow func OK)
      const row: (string | number | boolean)[] = [appId];
  
      // Call the helper function moved outside the loop
      // Inner forEach (arrow func OK)
      POSTURE_SHEETS_CONFIG.forEach(config => {
        const aggType = config.aggregationType;
        try {
          // All aggregation logic IS synchronous
          switch (aggType) {
            case "Count": {
              let outputValue: string | number = DEFAULT_VALUE_MISSING;
              // Use the moved helper function
              const valuesToCount = getValuesFromMap(postureDataMap, appId, config.countByHeader!);
              if (valuesToCount && valuesToCount.length > 0) {
                const counts = new Map<string | number | boolean, number>();
                valuesToCount.forEach(value => { counts.set(value, (counts.get(value) || 0) + 1); }); // Arrow func
                const countEntries: string[] = [];
                Array.from(counts.entries()) // Arrow funcs
                    .sort((a, b) => a[0].toString().localeCompare(b[0].toString()))
                    .forEach(([value, count]) => { countEntries.push(`${value}: ${count}`); });
                outputValue = countEntries.join('\n');
              } else { outputValue = 0; }
              row.push(outputValue);
              break;
            }
            case "Sum": case "Average": case "Min": case "Max": {
              let outputValue: string | number | boolean = DEFAULT_VALUE_MISSING;
              // Use the moved helper function
              const valuesToAggregate = getValuesFromMap(postureDataMap, appId, config.valueHeaderForAggregation!);
              // Arrow funcs used in map/filter
              const numericValues = valuesToAggregate?.map(parseNumber).filter(n => n !== null) as number[] | undefined;
              if (numericValues && numericValues.length > 0) {
                // Arrow funcs used in reduce
                if (aggType === "Sum") { outputValue = numericValues.reduce((s, c) => s + c, 0); }
                else if (aggType === "Average") { let avg = numericValues.reduce((s, c) => s + c, 0) / numericValues.length; outputValue = parseFloat(avg.toFixed(2)); }
                else if (aggType === "Min") { outputValue = Math.min(...numericValues); }
                else if (aggType === "Max") { outputValue = Math.max(...numericValues); }
              }
              row.push(outputValue);
              break;
            }
            case "UniqueList": {
              let outputValue: string | number | boolean = DEFAULT_VALUE_MISSING;
              const headerForUniqueList = config.dataHeadersToPull[0];
              if (headerForUniqueList) {
                // Use the moved helper function
                const valuesToList = getValuesFromMap(postureDataMap, appId, headerForUniqueList);
                if (valuesToList && valuesToList.length > 0) {
                  // Arrow funcs used in map/sort
                  const uniqueValues = Array.from(new Set(valuesToList.map(v => v?.toString() ?? "")));
                  uniqueValues.sort((a, b) => a.localeCompare(b));
                  outputValue = uniqueValues.join('\n');
                }
              }
              row.push(outputValue);
              break;
            }
            case "List": default: {
              // Arrow func used in forEach
              config.dataHeadersToPull.forEach(header => {
                let listOutput: string | number | boolean = DEFAULT_VALUE_MISSING;
                // Use the moved helper function
                const valuesToList = getValuesFromMap(postureDataMap, appId, header);
                if (valuesToList && valuesToList.length > 0) {
                   // Arrow funcs used in map/sort
                   const sortedValues = valuesToList.map(v => v?.toString() ?? "").sort((a,b) => a.localeCompare(b));
                   listOutput = sortedValues.join('\n');
                }
                row.push(listOutput);
              });
              break;
            }
          } // end switch
        } catch (e: unknown) { /* ... error handling ... */
              let errorMessage = "Unknown error";
              if (e instanceof Error) { errorMessage = e.message; }
              else { errorMessage = String(e); }
              console.log(`Error during aggregation type "${aggType}" for App "${appId}", Sheet "${config.sheetName}": ${errorMessage}`);
              if (aggType === 'List') { config.dataHeadersToPull.forEach(() => row.push('ERROR')); } // Arrow func
              else { row.push('ERROR'); }
        }
      }); // End inner POSTURE_SHEETS_CONFIG.forEach
      outputData.push(row);
    }); // End outer masterAppIdArray.forEach
  
  
    // --- 4c. Write Data ---
    let dataRange: ExcelScript.Range | undefined = undefined; // Sync init
    if (outputData.length > 0) { // Sync check
      dataRange = summarySheet.getRangeByIndexes(1, 0, outputData.length, summaryHeaders.length); // Sync
      // Writing data IS asynchronous
      await dataRange.setValues(outputData); // Await REQUIRED
      console.log(`Wrote ${outputData.length} rows of data.`);
    } else {
       console.log(`No data rows to write.`);
    }
  
    // --- 5. Apply Basic Formatting ---
    const usedRange = summarySheet.getUsedRange(); // Sync
    if (usedRange) {
      const usedRangeFormat = usedRange.getFormat(); // Sync
      // Setting format IS sync (no await)
      usedRangeFormat.setWrapText(true);
      usedRangeFormat.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
      // Autofitting IS async
      await usedRangeFormat.autofitColumns(); // Await REQUIRED
      // await usedRangeFormat.autofitRows(); // Optional - Await REQUIRED if used
      console.log("Applied basic formatting.");
    }
  
    // --- Finish ---
    // Select IS asynchronous
    await summarySheet.getCell(0,0).select(); // Await REQUIRED
    const endTime = Date.now(); // Sync
    const duration = (endTime - startTime) / 1000; // Sync
    console.log(`Script finished in ${duration.toFixed(2)} seconds.`);
  
  } // End main function