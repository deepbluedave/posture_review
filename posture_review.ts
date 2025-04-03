/**
 * Posture Summary Script (v9 - Final Await/Arrow Func Cleanup)
 *
 * Reads application posture data from various sheets defined in a 'Config' sheet,
 * aggregates the data based on specified methods (List, Count, Sum, Average, Min, Max, UniqueList),
 * and writes a summary report to a 'Posture Summary' sheet.
 * Conditional formatting has been removed.
 *
 * Key changes:
 * - Replaced console.warn/error with console.log.
 * - Removed status bar interactions.
 * - Switched nested Map structure to use plain objects.
 * - Added explicit typing for catch block errors (e: unknown).
 * - Corrected 'any' type usage in parseNumber.
 * - Removed unnecessary 'await' from synchronous calls (e.g., getFormat, getFont, setColor).
 * - Confirmed necessary 'await' remains on asynchronous calls (getValues, setValues, select, autofitColumns) despite potential editor warnings.
 * - Corrected autofitColumns method usage.
 * - Verified all array method callbacks use arrow functions.
 */
async function main(workbook: ExcelScript.Workbook) {
    console.log("Starting posture summary script (v9: Final Await/Arrow Func Cleanup)...");
    const startTime = Date.now(); // Synchronous
  
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
     * Uses arrow function for findIndex callback.
     */
    function findColumnIndex(headerRowValues: (string | number | boolean)[], possibleHeaders: string[]): number {
      // This function is synchronous.
      for (const header of possibleHeaders) {
        if (!header) continue;
        const lowerHeader = header.toString().toLowerCase();
        // findIndex is synchronous, callback uses arrow function
        const index = headerRowValues.findIndex(h => h?.toString().toLowerCase() === lowerHeader);
        if (index !== -1) {
          return index;
        }
      }
      return -1;
    }
  
    /**
     * Safely parses a value into a number. Handles various non-numeric inputs.
     */
    function parseNumber(value: string | number | boolean | null | undefined): number | null {
      // This function is synchronous.
      if (value === null || typeof value === 'undefined' || value === "") {
        return null;
      }
      const cleanedValue = typeof value === 'string' ? value.replace(/[^0-9.-]+/g, "") : value;
      const num: number = Number(cleanedValue);
      return isNaN(num) ? null : num;
    }
  
    // --- 1. Read Configuration from the "Config" Sheet ---
    console.log(`Reading configuration from sheet: ${CONFIG_SHEET_NAME}`);
    // Getting workbook objects is synchronous
    const configSheet = workbook.getWorksheet(CONFIG_SHEET_NAME);
    if (!configSheet) {
      console.log(`Error: Configuration sheet "${CONFIG_SHEET_NAME}" not found. Script cannot proceed.`);
      return;
    }
  
    let configValues: (string | number | boolean)[][] = [];
    let configHeaderRow: (string | number | boolean)[];
    // Getting table object is synchronous
    const configTable = configSheet.getTable(CONFIG_TABLE_NAME);
  
    if (configTable) {
      console.log(`Using table "${CONFIG_TABLE_NAME}" for configuration.`);
      // Getting ranges/sizes is synchronous
      const configRangeWithHeader = configTable.getHeaderRowRange().getResizedRange(configTable.getRowCount(), 0);
      // Reading data from Excel IS asynchronous
      configValues = await configRangeWithHeader.getValues(); // Await REQUIRED (L97 area)
      if (configValues.length <= 1) {
        console.log(`Error: Configuration table "${CONFIG_TABLE_NAME}" has no data rows.`);
        return;
      }
      configHeaderRow = configValues[0];
    } else {
      console.log(`Table "${CONFIG_TABLE_NAME}" not found. Using used range on "${CONFIG_SHEET_NAME}".`);
      // Getting range/size is synchronous
      const configRange = configSheet.getUsedRange();
      if (!configRange || configRange.getRowCount() <= 1) {
        console.log(`Error: Configuration sheet "${CONFIG_SHEET_NAME}" is empty or has only headers.`);
        return;
      }
      // Reading data from Excel IS asynchronous
      configValues = await configRange.getValues(); // Await REQUIRED (L110 area)
      configHeaderRow = configValues[0];
    }
  
    // Finding indices is synchronous
    const colIdxIsEnabled = findColumnIndex(configHeaderRow, ["IsEnabled"]);
    const colIdxSheetName = findColumnIndex(configHeaderRow, ["SheetName"]);
    const colIdxAppIdHeaders = findColumnIndex(configHeaderRow, ["AppIdHeaders", "App ID Headers"]);
    const colIdxDataHeaders = findColumnIndex(configHeaderRow, ["DataHeadersToPull", "Data Headers"]);
    const colIdxAggType = findColumnIndex(configHeaderRow, ["AggregationType", "Aggregation Type"]);
    const colIdxCountBy = findColumnIndex(configHeaderRow, ["CountByHeader", "Count By"]);
    const colIdxValueHeader = findColumnIndex(configHeaderRow, ["ValueHeaderForAggregation", "Value Header"]);
  
    // .some uses arrow function callback, synchronous check
    if ([colIdxIsEnabled, colIdxSheetName, colIdxAppIdHeaders, colIdxDataHeaders, colIdxAggType].some(idx => idx === -1)) {
      console.log("Error: One or more essential columns (IsEnabled, SheetName, AppIdHeaders, DataHeadersToPull, AggregationType) are missing in the Config sheet/table headers.");
      return;
    }
  
    const POSTURE_SHEETS_CONFIG: PostureSheetConfig[] = []; // Synchronous init
    let configIsValid = true;
    // Parsing config loop is synchronous
    for (let i = 1; i < configValues.length; i++) {
      const row = configValues[i];
      // .map uses arrow function callback, synchronous
      const cleanRow = row.map(val => typeof val === 'string' ? val.trim() : val);
  
      const isEnabled = cleanRow[colIdxIsEnabled]?.toString().toUpperCase() === "TRUE";
      if (!isEnabled) continue;
  
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
  
      // .map/.filter use arrow function callbacks, synchronous
      const appIdHeaders = appIdHeadersRaw.split(',').map(h => h.trim()).filter(h => h);
      const dataHeadersToPull = dataHeadersRaw.split(',').map(h => h.trim()).filter(h => h);
  
      if (appIdHeaders.length === 0) {
          console.log(`Warning: Config (Row ${i + 1}, Sheet: ${sheetName}): AppIdHeaders field contains no valid header names after parsing. Skipping.`);
          continue;
      }
  
      let aggregationType = "List" as AggregationMethod;
      const normalizedAggType = aggTypeRaw.charAt(0).toUpperCase() + aggTypeRaw.slice(1).toLowerCase();
      // .includes is synchronous
      if (["List", "Count", "Sum", "Average", "Min", "Max", "UniqueList"].includes(normalizedAggType)) {
        aggregationType = normalizedAggType as AggregationMethod;
      } else {
        console.log(`Warning: Config (Row ${i + 1}, Sheet: ${sheetName}): Invalid AggregationType "${aggTypeRaw}". Defaulting to "List".`);
        aggregationType = "List";
      }
  
      const configEntry: PostureSheetConfig = {
        isEnabled: true, sheetName, appIdHeaders, dataHeadersToPull, aggregationType, countByHeader, valueHeaderForAggregation: valueHeader
      };
  
      let rowIsValid = true;
      // Validation logic is synchronous
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
  
      if (rowIsValid) {
        POSTURE_SHEETS_CONFIG.push(configEntry);
      } else {
        configIsValid = false;
      }
    } // End config parsing loop
  
    if (!configIsValid) {
      console.log("Error: Configuration errors detected. Please review the Config sheet and messages above. Script halted.");
      return;
    }
    if (POSTURE_SHEETS_CONFIG.length === 0) {
      console.log("Warning: No enabled and valid configurations found in the Config sheet. Summary will be empty or just headers.");
    }
    console.log(`Successfully loaded ${POSTURE_SHEETS_CONFIG.length} valid configurations.`);
  
  
    // --- 2. Read Master App IDs ---
    console.log(`Reading master App IDs from sheet: ${MASTER_APP_SHEET_NAME}, Header: ${MASTER_APP_ID_HEADER}`);
    const masterSheet = workbook.getWorksheet(MASTER_APP_SHEET_NAME); // Synchronous
    if (!masterSheet) {
        console.log(`Error: Master sheet "${MASTER_APP_SHEET_NAME}" not found.`);
        return;
    }
  
    const masterRange = masterSheet.getUsedRange(); // Synchronous
    if (!masterRange) {
        console.log(`Error: Master sheet "${MASTER_APP_SHEET_NAME}" appears to be empty.`);
        return;
    }
    // Reading data IS asynchronous
    const masterValues = await masterRange.getValues(); // Await REQUIRED (L219 area)
    if (masterValues.length <= 1) {
        console.log(`Error: Master sheet "${MASTER_APP_SHEET_NAME}" contains only headers or is empty.`);
         return;
    }
  
    const masterHeaderRow = masterValues[0];
    const masterAppIdColIndex = findColumnIndex(masterHeaderRow, [MASTER_APP_ID_HEADER]); // Synchronous
    if (masterAppIdColIndex === -1) {
        console.log(`Error: Master App ID column "${MASTER_APP_ID_HEADER}" not found in sheet "${MASTER_APP_SHEET_NAME}".`);
        return;
    }
  
    const masterAppIds = new Set<string>(); // Synchronous init
    // Populating set IS synchronous
    for (let i = 1; i < masterValues.length; i++) {
      const appId = masterValues[i][masterAppIdColIndex]?.toString().trim();
      if (appId && appId !== "") {
          masterAppIds.add(appId);
      }
    }
    console.log(`Found ${masterAppIds.size} unique App IDs in master list.`);
    if (masterAppIds.size === 0) {
        console.log(`Warning: No valid App IDs found in the master sheet "${MASTER_APP_SHEET_NAME}". Summary will be empty.`);
    }
  
  
    // --- 3. Process Posture Sheets and Aggregate Data ---
    console.log("Processing posture sheets based on configuration...");
    const postureDataMap: PostureDataObject = {}; // Synchronous init
  
    // Using 'for...of' allows 'await' for getValues inside
    for (const config of POSTURE_SHEETS_CONFIG) {
      console.log(`Processing sheet: ${config.sheetName} (Type: ${config.aggregationType})`);
      const postureSheet = workbook.getWorksheet(config.sheetName); // Synchronous
      if (!postureSheet) {
          console.log(`Warning: Sheet "${config.sheetName}" specified in config not found. Skipping.`);
          continue;
      }
  
      const postureRange = postureSheet.getUsedRange(); // Synchronous
      if (!postureRange || postureRange.getRowCount() <= 1) { // getRowCount is synchronous
          console.log(`Warning: Sheet "${config.sheetName}" is empty or contains only headers. Skipping.`);
          continue;
      }
  
      // Reading data IS asynchronous
      const postureValues = await postureRange.getValues(); // Await REQUIRED (L264 area)
      const postureHeaderRow = postureValues[0];
  
      // Finding indices is synchronous
      const appIdColIndex = findColumnIndex(postureHeaderRow, config.appIdHeaders);
      if (appIdColIndex === -1) {
          console.log(`Warning: Could not find any specified App ID header (${config.appIdHeaders.join(', ')}) in sheet "${config.sheetName}". Skipping.`);
          continue;
      }
  
      const dataColIndicesMap = new Map<string, number>(); // Synchronous init
      let requiredHeaderMissing = false;
  
      // Finding indices and validating config IS synchronous
      // .forEach uses arrow function callback
      config.dataHeadersToPull.forEach(header => {
        const index = findColumnIndex(postureHeaderRow, [header]);
        if (index !== -1) {
          dataColIndicesMap.set(header, index);
        } else {
          console.log(`Warning: Data column "${header}" listed in DataHeadersToPull not found in sheet "${config.sheetName}". It will be skipped.`);
          if (["Sum", "Average", "Min", "Max"].includes(config.aggregationType) && header === config.valueHeaderForAggregation) {
              console.log(`Error: Critical header "${header}" (ValueHeaderForAggregation) is missing in sheet "${config.sheetName}". Skipping sheet.`);
              requiredHeaderMissing = true;
          }
          if (config.aggregationType === "UniqueList" && header === config.dataHeadersToPull[0] && config.dataHeadersToPull.length === 1) {
               console.log(`Error: The only header specified for UniqueList ("${header}") is missing in sheet "${config.sheetName}". Skipping sheet.`);
               requiredHeaderMissing = true;
          }
        }
      });
      if (requiredHeaderMissing) continue;
  
      if (config.aggregationType === "Count" && config.countByHeader) { // Synchronous check
          const countByColIndex = findColumnIndex(postureHeaderRow, [config.countByHeader]);
          if (countByColIndex === -1) {
              console.log(`Error: Critical header "${config.countByHeader}" (CountByHeader) is missing in sheet "${config.sheetName}". Skipping sheet.`);
              requiredHeaderMissing = true;
          } else if (!dataColIndicesMap.has(config.countByHeader)) {
               dataColIndicesMap.set(config.countByHeader, countByColIndex);
               console.log(`Info: CountByHeader "${config.countByHeader}" found and will be used for counting.`);
          }
      }
      if (requiredHeaderMissing) continue;
  
      if (["Sum", "Average", "Min", "Max"].includes(config.aggregationType) && config.valueHeaderForAggregation) { // Synchronous check
          if (!dataColIndicesMap.has(config.valueHeaderForAggregation)) {
               console.log(`Error: Internal Error/Config Issue: ValueHeaderForAggregation "${config.valueHeaderForAggregation}" not found in collected indices map. Skipping sheet "${config.sheetName}".`);
               requiredHeaderMissing = true;
          }
      }
      if (requiredHeaderMissing) continue;
  
      if (dataColIndicesMap.size === 0) { // Synchronous check
        console.log(`Warning: No relevant data columns found in "${config.sheetName}". Skipping.`);
        continue;
      }
  
      // Populating the postureDataMap object from postureValues IS synchronous
      let rowsProcessed = 0;
      for (let i = 1; i < postureValues.length; i++) { // Synchronous loop
        const row = postureValues[i];
        const appId = row[appIdColIndex]?.toString().trim();
  
        if (appId && masterAppIds.has(appId)) { // Synchronous check
          if (!postureDataMap[appId]) {
            postureDataMap[appId] = {};
          }
          const appData = postureDataMap[appId];
  
          // .forEach uses arrow function callback
          dataColIndicesMap.forEach((colIndex, headerName) => {
            const value = row[colIndex];
            if (value !== null && typeof value !== 'undefined' && value !== "") {
              if (!appData[headerName]) {
                appData[headerName] = [];
              }
              appData[headerName].push(value);
            }
          });
          rowsProcessed++;
        }
      }
      console.log(`Finished processing sheet ${config.sheetName}. Found data for ${rowsProcessed} relevant rows.`);
    } // End loop through posture sheet configs
    console.log("Finished processing all configured posture sheets.");
  
  
    // --- 4. Prepare and Write Summary Sheet ---
    console.log(`Preparing summary sheet: ${SUMMARY_SHEET_NAME}`);
    // Sheet operations: delete/add/activate are synchronous
    workbook.getWorksheet(SUMMARY_SHEET_NAME)?.delete();
    const summarySheet = workbook.addWorksheet(SUMMARY_SHEET_NAME);
    summarySheet.activate();
  
    // Generating headers IS synchronous
    const summaryHeaders: string[] = [MASTER_APP_ID_HEADER];
    const headerConfigMapping: { header: string, config: PostureSheetConfig, sourceValueHeader?: string }[] = [];
    // .forEach uses arrow function callback
    POSTURE_SHEETS_CONFIG.forEach(config => {
      const aggType = config.aggregationType;
      switch (aggType) {
          case "Count": {
              const header = `${config.countByHeader} Count Summary`;
              summaryHeaders.push(header);
              headerConfigMapping.push({ header: header, config: config, sourceValueHeader: config.countByHeader });
              break;
          }
          case "Sum": case "Average": case "Min": case "Max": {
              const header = `${config.sheetName} ${aggType} (${config.valueHeaderForAggregation})`;
              summaryHeaders.push(header);
              headerConfigMapping.push({ header: header, config: config, sourceValueHeader: config.valueHeaderForAggregation });
              break;
          }
          case "UniqueList": {
              const uniqueHeaderSource = config.dataHeadersToPull[0];
              const header = `${config.sheetName} Unique ${uniqueHeaderSource}`;
              summaryHeaders.push(header);
              headerConfigMapping.push({ header: header, config: config, sourceValueHeader: uniqueHeaderSource });
              break;
          }
          case "List": default: {
              // .forEach uses arrow function callback
              config.dataHeadersToPull.forEach(originalHeader => {
                  summaryHeaders.push(originalHeader);
                  headerConfigMapping.push({ header: originalHeader, config: config });
              });
              break;
          }
      }
    });
  
    // Getting range/format objects IS synchronous
    const headerRange = summarySheet.getRangeByIndexes(0, 0, 1, summaryHeaders.length);
    const headerFormat = headerRange.getFormat();
    const headerFont = headerFormat.getFont();
    const headerFill = headerFormat.getFill();
  
    // Writing values IS asynchronous
    await headerRange.setValues([summaryHeaders]); // Await REQUIRED (L402 area)
  
    // Setting format properties IS synchronous (no await)
    headerFont.setBold(true);
    headerFill.setColor("#4472C4");
    headerFont.setColor("white");
  
    // --- 4b. Generate Summary Data Rows ---
    const outputData: (string | number | boolean)[][] = []; // Synchronous init
    // Array.from/sort are synchronous
    const masterAppIdArray = Array.from(masterAppIds).sort();
  
    // Processing data in memory IS synchronous
    // .forEach uses arrow function callback
    masterAppIdArray.forEach(appId => { // START of block potentially related to L454 error report
      const row: (string | number | boolean)[] = [appId];
      const appMapData = postureDataMap[appId];
  
      // Helper function uses arrow function syntax, synchronous
      const getValues = (headerName: string): (string | number | boolean)[] | undefined => {
          return appMapData?.[headerName];
      }; // End of getValues function definition
  
      // .forEach uses arrow function callback - THIS is the loop starting near L454
      POSTURE_SHEETS_CONFIG.forEach(config => {
        const aggType = config.aggregationType;
        try {
          // All aggregation logic IS synchronous
          switch (aggType) {
            case "Count": {
              let outputValue: string | number = DEFAULT_VALUE_MISSING;
              const valuesToCount = getValues(config.countByHeader!);
              if (valuesToCount && valuesToCount.length > 0) {
                const counts = new Map<string | number | boolean, number>();
                // .forEach uses arrow function callback
                valuesToCount.forEach(value => { counts.set(value, (counts.get(value) || 0) + 1); });
                const countEntries: string[] = [];
                // Array.from, .sort (arrow func), .forEach (arrow func)
                Array.from(counts.entries())
                    .sort((a, b) => a[0].toString().localeCompare(b[0].toString()))
                    .forEach(([value, count]) => { countEntries.push(`${value}: ${count}`); });
                outputValue = countEntries.join('\n');
              } else {
                  outputValue = 0;
              }
              row.push(outputValue);
              break;
            }
            case "Sum": case "Average": case "Min": case "Max": {
              let outputValue: string | number | boolean = DEFAULT_VALUE_MISSING;
              const valuesToAggregate = getValues(config.valueHeaderForAggregation!);
              // .map (implicit arrow func for parseNumber), .filter (arrow func)
              const numericValues = valuesToAggregate?.map(parseNumber).filter(n => n !== null) as number[] | undefined;
  
              if (numericValues && numericValues.length > 0) {
                // .reduce uses arrow function callback
                if (aggType === "Sum") { outputValue = numericValues.reduce((s, c) => s + c, 0); }
                // .reduce uses arrow function callback
                else if (aggType === "Average") {
                    let avg = numericValues.reduce((s, c) => s + c, 0) / numericValues.length;
                    outputValue = parseFloat(avg.toFixed(2));
                }
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
                const valuesToList = getValues(headerForUniqueList);
                if (valuesToList && valuesToList.length > 0) {
                  // Array.from/Set, .map (arrow func), .sort (arrow func)
                  const uniqueValues = Array.from(new Set(valuesToList.map(v => v?.toString() ?? "")));
                  uniqueValues.sort((a, b) => a.localeCompare(b));
                  outputValue = uniqueValues.join('\n');
                }
              }
              row.push(outputValue);
              break;
            }
            case "List": default: {
              // .forEach uses arrow function callback
              config.dataHeadersToPull.forEach(header => {
                let listOutput: string | number | boolean = DEFAULT_VALUE_MISSING;
                const valuesToList = getValues(header);
                if (valuesToList && valuesToList.length > 0) {
                   // .map (arrow func), .sort (arrow func)
                   const sortedValues = valuesToList.map(v => v?.toString() ?? "").sort((a,b) => a.localeCompare(b));
                   listOutput = sortedValues.join('\n');
                }
                row.push(listOutput);
              });
              break;
            }
          } // end switch
        } catch (e: unknown) {
          let errorMessage = "Unknown error";
          if (e instanceof Error) { errorMessage = e.message; }
          else { errorMessage = String(e); }
          console.log(`Error during aggregation type "${aggType}" for App "${appId}", Sheet "${config.sheetName}": ${errorMessage}`);
  
          if (aggType === 'List') {
            // .forEach uses arrow function callback
            config.dataHeadersToPull.forEach(() => row.push('ERROR'));
          } else {
            row.push('ERROR');
          }
        }
      }); // End POSTURE_SHEETS_CONFIG.forEach
      outputData.push(row);
    }); // End masterAppIdArray.forEach
  
  
    // --- 4c. Write Data to Sheet ---
    let dataRange: ExcelScript.Range | undefined = undefined; // Synchronous init
    if (outputData.length > 0) { // Synchronous check
      // Getting range IS synchronous
      dataRange = summarySheet.getRangeByIndexes(1, 0, outputData.length, summaryHeaders.length);
      // Writing data IS asynchronous
      await dataRange.setValues(outputData); // Await REQUIRED (L536 area)
      console.log(`Wrote ${outputData.length} rows of data to ${SUMMARY_SHEET_NAME}.`);
    } else {
       console.log(`No data rows to write to ${SUMMARY_SHEET_NAME}.`);
    }
  
    // --- 5. Apply Basic Formatting (No Conditional Formatting) ---
    const usedRange = summarySheet.getUsedRange(); // Synchronous
    if (usedRange) {
      // Getting format object IS synchronous
      const usedRangeFormat = usedRange.getFormat();
      // Setting format properties IS synchronous (no await)
      usedRangeFormat.setWrapText(true);
      usedRangeFormat.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
      // Autofitting IS asynchronous
      await usedRangeFormat.autofitColumns(); // Await REQUIRED (L543 area), called on Format object
      // await usedRangeFormat.autofitRows(); // Optional - Await REQUIRED if used
      console.log("Applied basic formatting (Wrap Text, Top Align, Autofit Columns).");
    }
  
    // --- Finish ---
    // Select IS asynchronous
    await summarySheet.getCell(0,0).select(); // Await REQUIRED
    const endTime = Date.now(); // Synchronous
    const duration = (endTime - startTime) / 1000; // Synchronous
    console.log(`Script finished successfully in ${duration.toFixed(2)} seconds.`);
  
  } // End main function