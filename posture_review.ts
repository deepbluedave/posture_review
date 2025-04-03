/**
 * Posture Summary Script (v7 - Office Scripts Compatibility Fixes)
 *
 * Reads application posture data from various sheets defined in a 'Config' sheet,
 * aggregates the data based on specified methods (List, Count, Sum, Average, Min, Max, UniqueList),
 * and writes a summary report to a 'Posture Summary' sheet.
 * Conditional formatting has been removed.
 *
 * Key changes for Office Scripts compatibility:
 * - Replaced console.warn/error with console.log.
 * - Removed status bar interactions.
 * - Switched nested Map structure to use plain objects.
 * - Added explicit typing for catch block errors (e: unknown).
 * - Corrected 'any' type usage in parseNumber.
 * - Removed unnecessary 'await' from synchronous formatting/object retrieval calls.
 * - Ensured necessary 'await' keywords are present for asynchronous operations.
 */
async function main(workbook: ExcelScript.Workbook) {
    console.log("Starting posture summary script (v7: Office Scripts Compatibility Fixes)...");
    const startTime = Date.now(); // For performance timing
  
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
      for (const header of possibleHeaders) {
        if (!header) continue;
        const lowerHeader = header.toString().toLowerCase();
        const index = headerRowValues.findIndex(h => h?.toString().toLowerCase() === lowerHeader);
        if (index !== -1) {
          return index;
        }
      }
      return -1;
    }
  
    /**
     * Safely parses a value into a number. Handles various non-numeric inputs.
     * Uses specific types instead of 'any' for better compatibility.
     * @param value The value to parse (typically string, number, boolean, null, or undefined from getValues()).
     * @returns The parsed number, or null if parsing fails or input is invalid.
     */
    function parseNumber(value: string | number | boolean | null | undefined): number | null { // Changed 'any' to specific types
      if (value === null || typeof value === 'undefined' || value === "") {
        return null;
      }
      // Clean string values before converting
      const cleanedValue = typeof value === 'string' ? value.replace(/[^0-9.-]+/g, "") : value;
      // Explicitly type 'num'
      const num: number = Number(cleanedValue);
      return isNaN(num) ? null : num;
    }
  
    // --- 1. Read Configuration from the "Config" Sheet ---
    console.log(`Reading configuration from sheet: ${CONFIG_SHEET_NAME}`);
    const configSheet = workbook.getWorksheet(CONFIG_SHEET_NAME);
    if (!configSheet) {
      console.log(`Error: Configuration sheet "${CONFIG_SHEET_NAME}" not found. Script cannot proceed.`);
      return;
    }
  
    let configValues: (string | number | boolean)[][] = [];
    let configHeaderRow: (string | number | boolean)[];
    const configTable = configSheet.getTable(CONFIG_TABLE_NAME);
  
    if (configTable) {
      console.log(`Using table "${CONFIG_TABLE_NAME}" for configuration.`);
      const configRangeWithHeader = configTable.getHeaderRowRange().getResizedRange(configTable.getRowCount(), 0);
      // Reading values from Excel IS asynchronous
      configValues = await configRangeWithHeader.getValues(); // Await NEEDED here
      if (configValues.length <= 1) {
        console.log(`Error: Configuration table "${CONFIG_TABLE_NAME}" has no data rows.`);
        return;
      }
      configHeaderRow = configValues[0];
    } else {
      console.log(`Table "${CONFIG_TABLE_NAME}" not found. Using used range on "${CONFIG_SHEET_NAME}".`);
      const configRange = configSheet.getUsedRange();
      if (!configRange || configRange.getRowCount() <= 1) {
        console.log(`Error: Configuration sheet "${CONFIG_SHEET_NAME}" is empty or has only headers.`);
        return;
      }
      // Reading values from Excel IS asynchronous
      configValues = await configRange.getValues(); // Await NEEDED here
      configHeaderRow = configValues[0];
    }
  
    // Find indices (synchronous operations)
    const colIdxIsEnabled = findColumnIndex(configHeaderRow, ["IsEnabled"]);
    const colIdxSheetName = findColumnIndex(configHeaderRow, ["SheetName"]);
    const colIdxAppIdHeaders = findColumnIndex(configHeaderRow, ["AppIdHeaders", "App ID Headers"]);
    const colIdxDataHeaders = findColumnIndex(configHeaderRow, ["DataHeadersToPull", "Data Headers"]);
    const colIdxAggType = findColumnIndex(configHeaderRow, ["AggregationType", "Aggregation Type"]);
    const colIdxCountBy = findColumnIndex(configHeaderRow, ["CountByHeader", "Count By"]);
    const colIdxValueHeader = findColumnIndex(configHeaderRow, ["ValueHeaderForAggregation", "Value Header"]);
  
    if ([colIdxIsEnabled, colIdxSheetName, colIdxAppIdHeaders, colIdxDataHeaders, colIdxAggType].some(idx => idx === -1)) {
      console.log("Error: One or more essential columns (IsEnabled, SheetName, AppIdHeaders, DataHeadersToPull, AggregationType) are missing in the Config sheet/table headers.");
      return;
    }
  
    const POSTURE_SHEETS_CONFIG: PostureSheetConfig[] = [];
    let configIsValid = true;
    // Parsing config values is synchronous
    for (let i = 1; i < configValues.length; i++) {
      const row = configValues[i];
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
  
      const appIdHeaders = appIdHeadersRaw.split(',').map(h => h.trim()).filter(h => h);
      const dataHeadersToPull = dataHeadersRaw.split(',').map(h => h.trim()).filter(h => h);
  
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
    }
  
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
    const masterSheet = workbook.getWorksheet(MASTER_APP_SHEET_NAME);
    if (!masterSheet) {
        console.log(`Error: Master sheet "${MASTER_APP_SHEET_NAME}" not found.`);
        return;
    }
  
    const masterRange = masterSheet.getUsedRange();
    if (!masterRange) {
        console.log(`Error: Master sheet "${MASTER_APP_SHEET_NAME}" appears to be empty.`);
        return;
    }
    // Reading values IS asynchronous
    const masterValues = await masterRange.getValues(); // Await NEEDED
    if (masterValues.length <= 1) {
        console.log(`Error: Master sheet "${MASTER_APP_SHEET_NAME}" contains only headers or is empty.`);
         return;
    }
  
    const masterHeaderRow = masterValues[0];
    // Finding index is synchronous
    const masterAppIdColIndex = findColumnIndex(masterHeaderRow, [MASTER_APP_ID_HEADER]);
    if (masterAppIdColIndex === -1) {
        console.log(`Error: Master App ID column "${MASTER_APP_ID_HEADER}" not found in sheet "${MASTER_APP_SHEET_NAME}".`);
        return;
    }
  
    // Populating the Set is synchronous
    const masterAppIds = new Set<string>();
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
    const postureDataMap: PostureDataObject = {};
  
    // Using 'for...of' allows 'await' for getValues inside the loop
    for (const config of POSTURE_SHEETS_CONFIG) {
      console.log(`Processing sheet: ${config.sheetName} (Type: ${config.aggregationType})`);
      const postureSheet = workbook.getWorksheet(config.sheetName);
      if (!postureSheet) {
          console.log(`Warning: Sheet "${config.sheetName}" specified in config not found. Skipping.`);
          continue;
      }
  
      const postureRange = postureSheet.getUsedRange();
      if (!postureRange || postureRange.getRowCount() <= 1) {
          console.log(`Warning: Sheet "${config.sheetName}" is empty or contains only headers. Skipping.`);
          continue;
      }
  
      // Reading values IS asynchronous
      const postureValues = await postureRange.getValues(); // Await NEEDED
      const postureHeaderRow = postureValues[0];
  
      // Finding indices is synchronous
      const appIdColIndex = findColumnIndex(postureHeaderRow, config.appIdHeaders);
      if (appIdColIndex === -1) {
          console.log(`Warning: Could not find any specified App ID header (${config.appIdHeaders.join(', ')}) in sheet "${config.sheetName}". Skipping.`);
          continue;
      }
  
      const dataColIndicesMap = new Map<string, number>();
      let requiredHeaderMissing = false;
  
      // Finding indices and validating config is synchronous
      config.dataHeadersToPull.forEach(header => { /* ... synchronous ... */
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
  
      if (config.aggregationType === "Count" && config.countByHeader) { /* ... synchronous ... */
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
  
      if (["Sum", "Average", "Min", "Max"].includes(config.aggregationType) && config.valueHeaderForAggregation) { /* ... synchronous ... */
          if (!dataColIndicesMap.has(config.valueHeaderForAggregation)) {
               console.log(`Error: Internal Error/Config Issue: ValueHeaderForAggregation "${config.valueHeaderForAggregation}" not found in collected indices map. Skipping sheet "${config.sheetName}".`);
               requiredHeaderMissing = true;
          }
      }
      if (requiredHeaderMissing) continue;
  
      if (dataColIndicesMap.size === 0) {
        console.log(`Warning: No relevant data columns found in "${config.sheetName}". Skipping.`);
        continue;
      }
  
      // Populating the postureDataMap object from the retrieved postureValues IS synchronous
      let rowsProcessed = 0;
      for (let i = 1; i < postureValues.length; i++) {
        const row = postureValues[i];
        const appId = row[appIdColIndex]?.toString().trim();
  
        if (appId && masterAppIds.has(appId)) {
          if (!postureDataMap[appId]) {
            postureDataMap[appId] = {};
          }
          const appData = postureDataMap[appId];
  
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
    }
    console.log("Finished processing all configured posture sheets.");
  
  
    // --- 4. Prepare and Write Summary Sheet ---
    console.log(`Preparing summary sheet: ${SUMMARY_SHEET_NAME}`);
    // Deleting/Adding sheets IS asynchronous
    workbook.getWorksheet(SUMMARY_SHEET_NAME)?.delete(); // delete() doesn't need await
    const summarySheet = workbook.addWorksheet(SUMMARY_SHEET_NAME); // addWorksheet doesn't need await
    summarySheet.activate(); // activate() doesn't need await
  
    // Generating headers IS synchronous
    const summaryHeaders: string[] = [MASTER_APP_ID_HEADER];
    const headerConfigMapping: { header: string, config: PostureSheetConfig, sourceValueHeader?: string }[] = [];
    POSTURE_SHEETS_CONFIG.forEach(config => { /* ... synchronous ... */
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
              config.dataHeadersToPull.forEach(originalHeader => {
                  summaryHeaders.push(originalHeader);
                  headerConfigMapping.push({ header: originalHeader, config: config });
              });
              break;
          }
      }
    });
  
    // Getting range and format objects is synchronous
    const headerRange = summarySheet.getRangeByIndexes(0, 0, 1, summaryHeaders.length);
    const headerFormat = headerRange.getFormat();
    const headerFont = headerFormat.getFont();
    const headerFill = headerFormat.getFill();
  
    // Writing values IS asynchronous
    await headerRange.setValues([summaryHeaders]); // Await NEEDED
  
    // Setting format properties is synchronous, no await needed
    headerFont.setBold(true);
    headerFill.setColor("#4472C4");
    headerFont.setColor("white");
  
    // --- 4b. Generate Summary Data Rows ---
    const outputData: (string | number | boolean)[][] = [];
    const masterAppIdArray = Array.from(masterAppIds).sort(); // Synchronous
  
    // Processing data in memory IS synchronous
    masterAppIdArray.forEach(appId => {
      const row: (string | number | boolean)[] = [appId];
      const appMapData = postureDataMap[appId];
  
      const getValues = (headerName: string): (string | number | boolean)[] | undefined => {
          return appMapData?.[headerName];
      }
  
      POSTURE_SHEETS_CONFIG.forEach(config => {
        const aggType = config.aggregationType;
        try {
          // All aggregation logic here uses standard JS and is synchronous
          switch (aggType) {
            case "Count": { /* ... synchronous ... */
              let outputValue: string | number = DEFAULT_VALUE_MISSING;
              const valuesToCount = getValues(config.countByHeader!);
              if (valuesToCount && valuesToCount.length > 0) {
                const counts = new Map<string | number | boolean, number>();
                valuesToCount.forEach(value => { counts.set(value, (counts.get(value) || 0) + 1); });
                const countEntries: string[] = [];
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
            case "Sum": case "Average": case "Min": case "Max": { /* ... synchronous ... */
              let outputValue: string | number | boolean = DEFAULT_VALUE_MISSING;
              const valuesToAggregate = getValues(config.valueHeaderForAggregation!);
              const numericValues = valuesToAggregate?.map(parseNumber).filter(n => n !== null) as number[] | undefined;
  
              if (numericValues && numericValues.length > 0) {
                if (aggType === "Sum") { outputValue = numericValues.reduce((s, c) => s + c, 0); }
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
            case "UniqueList": { /* ... synchronous ... */
              let outputValue: string | number | boolean = DEFAULT_VALUE_MISSING;
              const headerForUniqueList = config.dataHeadersToPull[0];
              if (headerForUniqueList) {
                const valuesToList = getValues(headerForUniqueList);
                if (valuesToList && valuesToList.length > 0) {
                  const uniqueValues = Array.from(new Set(valuesToList.map(v => v?.toString() ?? "")));
                  uniqueValues.sort((a, b) => a.localeCompare(b));
                  outputValue = uniqueValues.join('\n');
                }
              }
              row.push(outputValue);
              break;
            }
            case "List": default: { /* ... synchronous ... */
              config.dataHeadersToPull.forEach(header => {
                let listOutput: string | number | boolean = DEFAULT_VALUE_MISSING;
                const valuesToList = getValues(header);
                if (valuesToList && valuesToList.length > 0) {
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
            config.dataHeadersToPull.forEach(() => row.push('ERROR'));
          } else {
            row.push('ERROR');
          }
        }
      });
      outputData.push(row);
    });
  
  
    // --- 4c. Write Data to Sheet ---
    let dataRange: ExcelScript.Range | undefined = undefined;
    if (outputData.length > 0) {
      dataRange = summarySheet.getRangeByIndexes(1, 0, outputData.length, summaryHeaders.length);
      // Writing values IS asynchronous
      await dataRange.setValues(outputData); // Await NEEDED
      console.log(`Wrote ${outputData.length} rows of data to ${SUMMARY_SHEET_NAME}.`);
    } else {
       console.log(`No data rows to write to ${SUMMARY_SHEET_NAME}.`);
    }
  
    // --- 5. Apply Basic Formatting (No Conditional Formatting) ---
    const usedRange = summarySheet.getUsedRange();
    if (usedRange) {
      // Getting format object is synchronous
      const usedRangeFormat = usedRange.getFormat();
      // Setting format properties is synchronous
      usedRangeFormat.setWrapText(true);
      usedRangeFormat.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
      // Autofitting IS asynchronous
      await usedRange.autofitColumns(); // Await NEEDED
      // await usedRange.autofitRows(); // Optional - Await NEEDED if used
      console.log("Applied basic formatting (Wrap Text, Top Align, Autofit Columns).");
    }
  
    // --- Finish ---
    // Select IS asynchronous
    await summarySheet.getCell(0,0).select(); // Await NEEDED
    const endTime = Date.now(); // Synchronous
    const duration = (endTime - startTime) / 1000; // Synchronous
    console.log(`Script finished successfully in ${duration.toFixed(2)} seconds.`);
  
  } // End main function