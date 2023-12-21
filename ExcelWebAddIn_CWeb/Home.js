﻿

(function () {
    "use strict";

    // Declaration of global variables for later use
    let messageBanner;
    let dialog
    let accessToken;  // used to store user's access token
    let runningEnvir
    let tableListeners = { "sensei_lessonslearned": null }
    let myTables = { "sensei_lessonslearned": null }


    // Constants for client ID, redirect URL, and resource domain for authentication
    const clientId = "be63874f-f40e-433a-9f35-46afa1aef385"
    const redirectUrl = "https://seamus77x.github.io/index.html"
    const resourceDomain = "https://gsis-pmo-australia-sensei-dev.crm6.dynamics.com/"

    // Initialization function that runs each time a new page is loaded.
    Office.initialize = function (reason) {
        $(function () {
            try {

                //Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                //Office.context.document.settings.saveAsync();

                Office.addin.setStartupBehavior(Office.StartupBehavior.load);
                Office.addin.showAsTaskpane();
                //Office.addin.hide();
                //Office.addin.setStartupBehavior(Office.StartupBehavior.none);

                switch (Office.context.platform) {
                    case Office.PlatformType.PC:
                        runningEnvir = Office.PlatformType.PC
                        console.log('Excel Platform: Desktop Excel on Windows');
                        break;
                    case Office.PlatformType.Mac:
                        runningEnvir = Office.PlatformType.Mac
                        console.log('Excel Platform: Desktop Excel on Mac');
                        break;
                    case Office.PlatformType.OfficeOnline:
                        runningEnvir = Office.PlatformType.OfficeOnline
                        console.log('Excel Platform: Web Excel');
                        break;
                    case Office.PlatformType.iOS:
                        runningEnvir = Office.PlatformType.iOS
                        console.log('Excel Platform: Excel on iOS');
                        break;
                    case Office.PlatformType.Android:
                        runningEnvir = Office.PlatformType.Android
                        console.log('Excel Platform: Excel on Android');
                        break;
                    // You can add more cases here as needed
                    default:
                        runningEnvir = PlatformNotFound
                        console.log('Excel Platform: Not Identified');
                        break;
                }

                // Notification mechanism initialization and hiding it initially
                let element = document.querySelector('.MessageBanner');
                messageBanner = new components.MessageBanner(element);
                messageBanner.hideBanner();

                // Fallback logic for versions of Excel older than 2016
                if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                    throw new Error("Sorry, this add-in only works with newer versions of Excel.")
                }

                // add external js
                //$('#myScriptX').attr('src', 'Test.js')
                //$.getScript('Test.js', function () {
                //    externalFun()
                //})

                // UI text setting for buttons and descriptions
                $('#button1-text').text("Download");
                $("#button1").attr("title", "Load Data to Excel")
                $('#button1').on("click", loadSampleData);

                $('#button2-text').text("Button 2");

                // Authentication and access token retrieval logic
                if (typeof accessToken === 'undefined') {
                    // Constructing authentication URL
                    let authUrl = "https://login.microsoftonline.com/common/oauth2/authorize" +
                        "?client_id=" + clientId +
                        "&response_type=token" +
                        "&redirect_uri=" + redirectUrl +
                        "&response_mode=fragment" +
                        "&resource=" + resourceDomain;

                    // Displaying authentication dialog
                    Office.context.ui.displayDialogAsync(authUrl, { height: 30, width: 30, requireHTTPS: true },
                        function (result) {
                            if (result.status === Office.AsyncResultStatus.Failed) {
                                // If the dialog fails to open, throw an error
                                throw new Error("Failed to open dialog: " + result.error.message);
                            }
                            dialog = result.value;
                            dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                        }
                    );
                }
            } catch (error) {
                errorHandler(error.message)
            }
        });
    }

    Office.actions.associate("buttonFunction", function (event) {
        console.log('Hey, you just pressed a ribbon button.')

        console.log("Undo List")
        console.log(rowChangeHistory_Undo)
        console.log("Redo List")
        console.log(rowChangeHistory_Redo)
        event.completed();
    })

    // Process message (access token) received from the dialog
    function processMessage(arg) {
        try {
            // Check if the message is present
            if (!arg.message) {
                throw new Error("No message received from the dialog.");
            }

            // Parse the JSON message received from the dialog
            const response = JSON.parse(arg.message);

            // Check the status of the response
            if (response.Status === "Success") {
                // store the token in memory for later use
                accessToken = response.AccessToken
                console.log("Authentication Result: Passed")
            } else if (response.Status === "Error") {
                // Handle the error scenario
                errorHandler(response.Message || "An error occurred.");
            } else {
                // Handle unexpected status
                errorHandler("Unexpected response status.");
            }

        } catch (error) {
            // Handle any errors that occur during processing
            errorHandler(error.message);
        } finally {
            // Close the dialog, regardless of whether an error occurred
            if (dialog) {
                dialog.close();
            }
        }
    }

    // Function to load sample data
    async function loadSampleData() {
        const tableName = 'sensei_lessonslearned'
        const excludedColsNames = ['@odata.etag', 'sensei_lessonlearnedid']
        const odataCondition = '?$select=sensei_lessonlearnedid,sensei_name,sensei_lessonlearned,sensei_observation,sensei_actiontaken'

        await loadData(`${resourceDomain}api/data/v9.2/${tableName}${odataCondition}`
            , tableName, 1, 'Sheet1', 'A1', excludedColsNames)
    }
    //sc_integrationrecentgranulartransactions
    //sensei_financialtransaction
    //sensei_financialtransactions?$select=sc_kbrkey,sc_vendorname,sensei_value,sc_docdate,sensei_financialtransactionid&$top=50000


    // Function to retrieve data from Dynamics 365
    async function loadData(resourceUrl, tableName, Col_To_Paste_In_Table = 1, defaultSheet = 'Sheet1', defaultTpLeftRng = 'A1', excludedColsNames = ['@odata.etag']) {
        try {
            // turn off the listener to the table when refreshing
            await toggleEventListener(false)

            let DataArr = await Read_D365(resourceUrl);

            // act as the corresponding table in memory, which records the change in Excel table
            myTables[tableName] = JSON.parse(JSON.stringify(DataArr))

            // delete unwanted cols from the array which is going to be pasted into Excel
            let colIndices = excludedColsNames.map(colName => DataArr[0].indexOf(colName)).filter(index => index !== -1);
            // Sort the indices in descending order to avoid index shifting issues during removal
            colIndices.sort((a, b) => b - a);
            // Remove the columns with the found indices
            DataArr.map(row => {
                colIndices.forEach(colIndex => row.splice(colIndex, 1));
            });
            // report an error and interupt if failed to read data from Dataverse
            if (!DataArr || DataArr.length === 0) {
                throw new Error("No data retrieved or data array is empty");
            }
            // paste data into Excel worksheet 
            await Excel.run(async (ctx) => {
                const ThisWorkbook = ctx.workbook;
                const Worksheets = ThisWorkbook.worksheets;
                ctx.application.calculationMode = Excel.CalculationMode.manual;
                Worksheets.load("items/tables/items/name");

                await ctx.sync();

                let tableFound = false;
                let table;
                let oldRangeAddress;
                let oldFirstRow_formula
                let sheet

                if (tableName !== 'not using a table') {

                    // Attempt to find the existing table.
                    for (sheet of Worksheets.items) {
                        const tables = sheet.tables;

                        // Check if the table exists in the current sheet
                        table = tables.items.find(t => t.name === tableName);

                        // if the table found, delete the existing data
                        if (table) {
                            tableFound = true;
                            // Clear the data body range.
                            const dataBodyRange = table.getDataBodyRange();
                            dataBodyRange.load("address");
                            let firstRow = dataBodyRange.getRow(0);
                            firstRow.load('formulas');

                            dataBodyRange.clear();
                            await ctx.sync();
                            // Load the address of the range for new data insertion.
                            oldRangeAddress = dataBodyRange.address.split('!')[1];
                            oldFirstRow_formula = firstRow.formulas;
                            break;
                        }
                    }

                    if (tableFound) {
                        // delete header row of DataArr
                        DataArr.shift()

                        // add LHS and RHS formula cols to expand dataArr
                        let excelTableRightColNo = columnNameToNumber(oldRangeAddress.split(":")[1].replace(/\d+$/, ''))
                        let ppTableRightColNo = columnNameToNumber(oldRangeAddress.split(":")[0].replace(/\d+$/, '')) + Col_To_Paste_In_Table - 1 + DataArr[0].length - 1
                        DataArr.forEach(row => {
                            if (Col_To_Paste_In_Table > 1) {
                                let tempRowFormula = oldFirstRow_formula
                                row.unshift(...tempRowFormula[0].slice(0, Col_To_Paste_In_Table - 1))
                            }

                            if (excelTableRightColNo > ppTableRightColNo) {
                                let tempRowFormula = oldFirstRow_formula
                                row.push(...tempRowFormula[0].slice(ppTableRightColNo - excelTableRightColNo))
                            }
                        })

                        let newRangeAdress = oldRangeAddress.replace(/\d+$/, parseInt(oldRangeAddress.match(/\d+/)[0], 10) + DataArr.length - 1)
                        let range = sheet.getRange(newRangeAdress);

                        if (runningEnvir !== Office.PlatformType.OfficeOnline) {
                            range.values = DataArr;
                        } else {
                            pasteChunksToExcel(splitArrayIntoSmallPieces(DataArr), newRangeAdress, sheet, ctx)
                        }

                        // include header row when resize
                        let newRangeAdressWithHeader = newRangeAdress.replace(/\d+/, oldRangeAddress.match(/\d+/)[0] - 1)
                        let WholeTableRange = sheet.getRange(newRangeAdressWithHeader)
                        table.resize(WholeTableRange)

                        range.format.autofitColumns();
                        range.format.autofitRows();
                    } else {
                        // Situation 2: If the table doesn't exist, create a new one.
                        let tgtSheet = Worksheets.getItem(defaultSheet);
                        let endCellCol = columnNumberToName(columnNameToNumber(defaultTpLeftRng.replace(/\d+$/, "")) - 1 + DataArr[0].length)
                        let endCellRow = parseInt(defaultTpLeftRng.match(/\d+$/)[0], 10) + DataArr.length - 1
                        let rangeAddress = defaultTpLeftRng + ":" + endCellCol + endCellRow;
                        let range = tgtSheet.getRange(rangeAddress);

                        if (runningEnvir !== Office.PlatformType.OfficeOnline) {
                            range.values = DataArr;
                        } else {
                            pasteChunksToExcel(splitArrayIntoSmallPieces(DataArr), rangeAddress, tgtSheet, ctx)
                        }

                        let newTable = tgtSheet.tables.add(rangeAddress, true /* hasHeaders */);
                        newTable.name = tableName;

                        newTable.getRange().format.autofitColumns();
                        newTable.getRange().format.autofitRows();
                    }

                } else {
                    // Situation 3: paste the data in sheet directly, no table format
                    let tgtSheet = Worksheets.getItem(defaultSheet);
                    let endCellCol = columnNumberToName(columnNameToNumber(defaultTpLeftRng.replace(/\d+$/, "")) - 1 + DataArr[0].length)
                    let endCellRow = parseInt(defaultTpLeftRng.match(/\d+$/)[0], 10) + DataArr.length - 1
                    let rangeAddress = defaultTpLeftRng + ":" + endCellCol + endCellRow;
                    let range = tgtSheet.getRange(rangeAddress);

                    if (runningEnvir !== Office.PlatformType.OfficeOnline) {
                        range.values = DataArr;
                    } else {
                        pasteChunksToExcel(splitArrayIntoSmallPieces(DataArr), rangeAddress, tgtSheet, ctx)
                    }

                    range.format.autofitColumns();
                    range.format.autofitRows();
                }

                await ctx.sync();
            })  // end of pasting data
        } catch (error) {
            errorHandler(error.message)
        } finally {
            // turn on the auto calculation in Excel
            Excel.run(async (ctx) => {
                ctx.application.calculationMode = Excel.CalculationMode.automatic;
                await ctx.sync()
            })
            // turn on the listener to the table when refreshing
            toggleEventListener(true)
            // add listener to the table if no listener
            registerTableChangeEvent(tableName)

            // clearn redo and undo lists when refreshing
            rowChangeHistory_Undo = []
            rowChangeHistory_Redo = []
        }
    }

    function splitArrayIntoSmallPieces(data, maxChunkSizeInMB = 3.3) {

        const jsonString = JSON.stringify(data);
        const sizeInBytes = new TextEncoder().encode(jsonString).length;
        const sizeInMB = sizeInBytes / (1024 * 1024);

        console.log(`Total size: ${sizeInMB} MB`);

        if (sizeInMB <= maxChunkSizeInMB) {
            return [data]; // No need to chunk
        }

        let chunks = [];
        const totalRows = data.length;
        const rowsPerChunk = Math.ceil(totalRows * maxChunkSizeInMB / sizeInMB);

        for (let i = 0; i < totalRows; i += rowsPerChunk) {
            const chunk = data.slice(i, i + rowsPerChunk);
            chunks.push(chunk);
        }

        return chunks;
    }
    async function pasteChunksToExcel(chunks, rangeAddressToPaste, sheet, ctx) {
        const startCol = rangeAddressToPaste.match(/[A-Za-z]+/)[0];
        let startRow = parseInt(rangeAddressToPaste.match(/\d+/)[0], 10);

        const numberOfCols = chunks[0][0].length;
        const endCol = columnNumberToName(columnNameToNumber(startCol) + numberOfCols - 1);

        for (const chunk of chunks) {
            const chunkRowCount = chunk.length;
            const endRow = startRow + chunkRowCount - 1; // Calculate end row for the current chunk
            const rangeAddress = `${startCol}${startRow}:${endCol}${endRow}`;
            const range = sheet.getRange(rangeAddress);
            range.values = chunk;
            await ctx.sync();

            startRow += chunkRowCount; // Update startRow for the next chunk
        }
    }
    async function toggleEventListener(eventBoolean) {
        await Excel.run(async (context) => {
            //context.runtime.load("enableEvents");
            //await context.sync();

            //let eventBoolean = !context.runtime.enableEvents;
            context.runtime.enableEvents = eventBoolean;
            if (eventBoolean) {
                console.log("Events Status: On");
            } else {
                console.log("Events Status: Off");
            }

            await context.sync();
        });
    }

    async function updateData() {
        Update_D365('sensei_lessonslearned', '0f0db491-3421-ee11-9966-000d3a798402', { 'sc_additionalcommentsnotes': 'Update Test' })
        //Create_D365('sensei_lessonslearned', { 'sensei_name': 'Add Test', 'sc_additionalcommentsnotes': 'ADD Redo_Undo_Event_Done from Web Add-In' })
        //Delete_D365('sensei_lessonslearned','f38edda5-8d8d-ee11-be35-6045bd3db52a')
    }


    // Function to create data in Dynamics 365
    async function Create_D365(entityLogicalName, addedData, guidColName) {
        const url = `${resourceDomain}api/data/v9.2/${entityLogicalName}`;

        try {
            const response = await fetch(url, {
                method: 'POST',
                headers: {
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Accept': 'application/json',
                    'Content-Type': 'application/json; charset=utf-8',
                    'Authorization': `Bearer ${accessToken}`,
                    'Prefer': 'return=representation'
                },
                body: JSON.stringify(addedData)
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(`Server responded with status ${response.status}: ${errorData.error?.message}`);
            }

            const responseData = await response.json();
            console.log(`Synchronisation Result: record [${responseData[guidColName]}] is successfully created. - HTTP ${response.status}`);

            return responseData[guidColName]
        } catch (error) {
            if (error.name === 'TypeError') {
                // Handle network errors (e.g., no internet connection)
                errorHandler("Network error: " + error.message);
            } else {
                // Handle other types of errors (e.g., server responded with error code)
                errorHandler("Error encountered when adding new records in Dataverse:" + error.message);
            }
        }
    }
    // Function to read data in Dynamics 365
    async function Read_D365(url) {
        let totalRecords = 0;
        let finalArr = [];
        let startTime = new Date().getTime();

        try {
            do {
                let response = await fetch(url, {
                    method: 'GET',
                    headers: {
                        'OData-MaxVersion': '4.0',
                        'OData-Version': '4.0',
                        'Accept': 'application/json',
                        'Content-Type': 'application/json; charset=utf-8',
                        'Authorization': `Bearer ${accessToken}`,
                    }
                });

                if (!response.ok) {
                    const errorData = await response.json();
                    throw new Error(`Server responded with status ${response.status}: ${errorData.error?.message}`);
                }

                let jsonObj = await response.json();
                let headers = [];
                let tempArr_5k = [];

                if (jsonObj["value"] && jsonObj["value"].length > 0) {
                    for (let fieldName in jsonObj["value"][0]) {
                        if (typeof jsonObj["value"][0][fieldName] === "object" && jsonObj["value"][0][fieldName] != null) {
                            for (let relatedField in jsonObj["value"][0][fieldName]) {
                                let expandedFieldName = `${fieldName} / ${relatedField}`;
                                headers.push(expandedFieldName);
                            }
                        } else {
                            headers.push(fieldName);
                        }
                    }

                    tempArr_5k = [headers];

                    jsonObj["value"].forEach((row) => {
                        let itemWithRelatedFields = {};

                        for (let cell in row) {
                            if (typeof row[cell] === "object" && row[cell] !== null) {
                                for (let field in row[cell]) {
                                    let relatedFieldName = `${cell} / ${field}`;
                                    itemWithRelatedFields[relatedFieldName] = row[cell][field];
                                }
                            } else {
                                itemWithRelatedFields[cell] = row[cell];
                            }
                        }

                        let tempValRow = headers.map((header) => {
                            return itemWithRelatedFields[header] || null;
                        });

                        tempArr_5k.push(tempValRow);
                    });

                    if (totalRecords >= 1) {

                        let tempArr = [];
                        let headerRow = tempArr_5k[0];

                        for (let row of tempArr_5k) {
                            let tempValRow = [];
                            for (let fieldName of finalArr[0]) {
                                let trueColNo = headerRow.indexOf(fieldName);
                                tempValRow.push(row[trueColNo] || null);
                            }
                            tempArr.push(tempValRow);
                        }

                        tempArr.splice(0, 1);
                        finalArr = finalArr.concat(tempArr);
                    } else {
                        finalArr = finalArr.concat(tempArr_5k);
                    }
                }

                if (jsonObj["@odata.nextLink"]) {
                    url = jsonObj["@odata.nextLink"];
                } else {
                    url = null; // No more pages to retrieve
                }

                totalRecords += 1;
                console.log(`Retrieving Data: Page ${totalRecords} - HTTP ${response.status}`)

            } while (url != null);

            // Update Excel with the collected data
            if (finalArr.length > 0) {
                let finishTime = new Date().getTime();
                console.log(`Time Used: ${(finishTime - startTime) / 1000}s (${finalArr.length} rows, ${finalArr[0].length} columns)`);

                return finalArr
            } else {
                throw new Error("No data downloaded");
            }


        } catch (error) {
            if (error.name === 'TypeError') {
                // Handle network errors (e.g., no internet connection)
                errorHandler("Network error: " + error.message);
            } else {
                // Handle other types of errors (e.g., server responded with error code)
                errorHandler("Error encountered when retrieving records from Dataverse:" + error.message);
            }
        }
    }
    // Function to update data in Dynamics 365
    async function Update_D365(entityLogicalName, recordId, updatedData) {
        const url = `${resourceDomain}api/data/v9.2/${entityLogicalName}(${recordId})`;

        try {
            const response = await fetch(url, {
                method: 'PATCH',
                headers: {
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Accept': 'application/json',
                    'Content-Type': 'application/json; charset=utf-8',
                    'Authorization': `Bearer ${accessToken}`,
                    //'Prefer': 'return=representation'
                },
                body: JSON.stringify(updatedData)
            });

            if (!response.ok) {
                // If the server responded with a non-OK status, handle the error
                const errorData = await response.json();
                throw new Error(`Server responded with status ${response.status}: ${errorData.error?.message}`);
            }

            console.log(`Synchronisation Result: record [${recordId}] is successfully updated. - HTTP ${response.status}`);
        } catch (error) {
            if (error.name === 'TypeError') {
                // Handle network errors (e.g., no internet connection)
                errorHandler("Network error: " + error.message);
            } else {
                // Handle other types of errors (e.g., server responded with error code)
                errorHandler("Error encountered when updating records in Dataverse" + error.message);
            }
        }
    }
    // Function to delete data in Dynamics 365
    async function Delete_D365(entityLogicalName, recordId) {
        const url = `${resourceDomain}api/data/v9.2/${entityLogicalName}(${recordId})`;

        try {
            const response = await fetch(url, {
                method: 'DELETE',
                headers: {
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json; charset=utf-8'
                }
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(`Server responded with status ${response.status}: ${errorData.error?.message}`);
            }

            console.log(`Record with ID [${recordId}] deleted successfully.`);
        } catch (error) {
            if (error.name === 'TypeError') {
                // Handle network errors (e.g., no internet connection)
                errorHandler("Network error: " + error.message);
            } else {
                // Handle other types of errors (e.g., server responded with error code)
                errorHandler("Error encountered when deleting new records in Dataverse:" + error.message);
            }
        }
    }

    // Progress bar update function
    function updateProgressBar(progress) {
        let elem = document.getElementById("myProgressBar");
        elem.style.width = progress + '%';
        //elem.innerHTML = progress + '%';
    }
    //// Example: Update the progress bar every second
    //let progress = 0;
    //let interval = setInterval(function () {
    //    progress += 10; // Increment progress
    //    updateProgressBar(progress);

    //    if (progress >= 100) clearInterval(interval); // Clear interval at 100%
    //}, 1000);


    // Utility function to convert column number to name
    function columnNumberToName(columnNumber) {
        let columnName = "";
        while (columnNumber > 0) {
            let remainder = (columnNumber - 1) % 26;
            columnName = String.fromCharCode(65 + remainder) + columnName;
            columnNumber = Math.floor((columnNumber - 1) / 26);
        }
        return columnName;
    }
    // Utility function to convert column name to number
    function columnNameToNumber(columnName) {
        let columnNumber = 0;
        for (let i = 0; i < columnName.length; i++) {
            columnNumber *= 26;
            columnNumber += columnName.charCodeAt(i) - 64;
        }
        return columnNumber;
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.error("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }


    async function registerTableChangeEvent(tableName) {
        try {
            if (tableListeners[tableName]) {
                return
            }

            let ThisWorkbook;
            let Worksheets;

            Excel.run(function (ctx) {

                ThisWorkbook = ctx.workbook;
                Worksheets = ThisWorkbook.worksheets;
                Worksheets.load("items/tables/items/name, items/tables/items/id");
                return ctx.sync().then(() => {
                    for (let sheet of Worksheets.items) {
                        const tables = sheet.tables;
                        // Check if the 'Test' table exists in the current sheet
                        let table = tables.items.find(t => t.name === tableName);

                        if (table) {
                            // if the table found, then listen to the change in the table
                            tableListeners[tableName] = table.onChanged.add(handleTableChange)
                            table.worksheet.onSelectionChanged.add((eventArgs) => {
                                handleSelectionChange(eventArgs, table.id);
                            });
                            console.log(`Events Listener Added: ${tableName}`)
                            break;
                        }
                    }

                    if (!tableListeners[tableName]) {
                        // if the table not found, then raise an error
                        throw new Error(`[${tableName}] table is not found in Excel`);
                    }
                }).then(ctx.sync);

            })
        } catch (error) {
            // Error handling for issues within the Excel.run block
            errorHandler("Error in registerTableChangeEvent: " + error.message);
        }
    }


    let guidPromise
    let undo_redo
    let changesTracker = [[], []]
    let rowChangeHistory_Undo = []
    let rowChangeHistory_Redo = []
    let rowChange = []
    // row change events handlers
    function rowInsertedHandler(startRow, endRow, startCol, endCol, tableName, tableData, rangeAddress) {
        for (let r = startRow; r <= endRow; r++) {
            let jsonPayLoad = {}
            for (let c = startCol; c <= endCol; c++) {
                let displayColName = tableData[0][c]
                // need a fieldNameConverter - mapping table required.
                let logicalColNum = myTables[tableName][0].indexOf(displayColName)
                if (logicalColNum > -1) {
                    let logicalColName = myTables[tableName][0][logicalColNum]
                    jsonPayLoad[logicalColName] = tableData[r][c]
                }
            }
            if (Object.keys(jsonPayLoad).length > 0) {
                // add the row in memory table as well
                if (r + 1 > myTables[tableName].length) {
                    myTables[tableName].push([Date.now(), "waiting for guid"])
                } else {
                    myTables[tableName].splice(r, 0, [Date.now(), "waiting for guid"])
                }
                // Add in P+ table
                (async function () {
                    guidPromise = Create_D365(tableName, jsonPayLoad, "sensei_lessonlearnedid");
                    myTables[tableName][r][1] = await guidPromise
                })()
                // keep a record in rowHistory array
                rowChange.push([
                    "RowInserted",  // change type
                    r,              // relative row num in the table
                    rangeAddress,   // range address in the sheet
                    changesTracker[0],   // change type
                    ""
                ])
            }
        }
    }
    async function rowDeletedHandler(startRow, endRow, tableName, rangeAddress) {
        // collect the row num of the rows deleted
        for (let r = endRow; r >= startRow; r--) {
            //let guidColNum = myTables[table.name][0].indexOf(guidColName)
            let guidColNum = 1

            await guidPromise
            // delete the rows in P+ table
            Delete_D365(tableName, myTables[tableName][r][guidColNum])
            // delete the row in memeory table as well
            myTables[tableName].splice(r, 1)
            // keep a record in rowHistory array
            rowChange.push([
                "RowDeleted",  // change type
                r,              // relative row num in the table
                rangeAddress,   // range address in the sheet
                changesTracker[0],   // change type
                ""
            ])
        }
    }
    // hanle table change.    tip: get after value from Excel if multiple range changes
    function handleTableChange(eventArgs) {
        try {
            if (undo_redo = true) { undo_redo = undefined }
            if (changesTracker[1][0] === "Fulfilled") { changesTracker = [[], []] }
            changesTracker[0] += "B"

            let thisTableChangeType = eventArgs.changeType
            Excel.run(function (ctx) {
                // get the Range changed and the table changed
                let range = ctx.workbook.worksheets.getActiveWorksheet().getRange(eventArgs.address)
                range.load("values, address, rowIndex, columnIndex, cellCount, rowCount, columnCount")
                let table = ctx.workbook.tables.getItem(eventArgs.tableId);
                table.load("name")
                let tableRange = table.getRange()
                tableRange.load("rowIndex, columnIndex, rowCount,columnCount, values")

                return ctx.sync().then(async function () {
                    changesTracker[1][0] = "Fulfilled"

                    if (undo_redo === true && changesTracker[0] === "AB") {
                        // stop the AB case for redo row deletion
                        return
                    } else if (undo_redo === true && changesTracker[0] === "ABA") {
                        // allow the ABA case for undoing and redoing row addition
                    }

                    let tableData = tableRange.values
                    let tableStartRow = tableRange.rowIndex;
                    let tableStartCol = tableRange.columnIndex;

                    let startRangeRowRelative = range.rowIndex - tableStartRow;
                    let startRangeColRelative = range.columnIndex - tableStartCol;

                    let endRangeRowRelative = startRangeRowRelative + range.rowCount - 1
                    let endRangeColRelative = startRangeColRelative + range.columnCount - 1

                    // range content change event handler
                    async function rangeChangeHandler(startRow, endRow, startCol, endCol) {
                        // construct the JSON Payload
                        let jsonPayLoadColl = []
                        let guidColl = []

                        for (let r = startRow; r <= endRow; r++) {
                            // if this is triggered by adding new rows from bottom
                            if (Date.now() - myTables[table.name][r][0] <= 10) {
                                return "ignore this"
                            }

                            let jsonPayLoad = {}
                            for (let c = startCol; c <= endCol; c++) {
                                let displayColName = tableData[0][c]
                                // need a fieldNameConverter - mapping table required.
                                let logicalColNum = myTables[table.name][0].indexOf(displayColName)
                                if (logicalColNum > -1) {
                                    let logicalColName = myTables[table.name][0][logicalColNum]
                                    jsonPayLoad[logicalColName] = tableData[r][c]
                                }
                            }
                            if (Object.keys(jsonPayLoad).length > 0) {
                                jsonPayLoadColl.push(jsonPayLoad)
                                //let guidColNum = myTables[table.name][0].indexOf(guidColName)
                                let guidColNum = 1
                                await guidPromise
                                guidColl.push(myTables[table.name][r][guidColNum])
                            }
                        }

                        // start syncing by sending http request to D365 API
                        if (guidColl.length > 0) {
                            if (range.cellCount === 1) {
                                if (eventArgs.details !== undefined) {
                                    if (JSON.stringify(eventArgs.details.valueAsJsonAfter) !== JSON.stringify(eventArgs.details.valueAsJsonBefore)) {
                                        // single cell is changed
                                        Update_D365(table.name, guidColl[0], jsonPayLoadColl[0])
                                    } else {
                                        return "ignore this"
                                    }
                                } else {
                                    Update_D365(table.name, guidColl[0], jsonPayLoadColl[0])
                                }
                            } else {
                                // multiple range are changed
                                guidColl.forEach(function (rowGUID, index) {
                                    Update_D365(table.name, rowGUID, jsonPayLoadColl[index])
                                })
                            }
                        }
                    }

                    if (myTables[table.name].length < tableRange.rowCount && thisTableChangeType === "RangeEdited") {
                        if (changesTracker[0] === "BAB") {
                            // allow BAB case for normal row addition
                        } else if (changesTracker[0] === "BA") {
                            // stop the BA case for undoing row deletion
                            return
                        }
                    }

                    switch (thisTableChangeType) {
                        case 'RangeEdited':
                            let result = await rangeChangeHandler(startRangeRowRelative, endRangeRowRelative, startRangeColRelative, endRangeColRelative)

                            if (result === "ignore this") {
                                break;
                            } else {
                                // clean redo if new change
                                if (changesTracker[0] !== "ABA") {
                                    rowChangeHistory_Redo = []
                                }

                                console.log(`Range Updated: [${eventArgs.address}] in '${table.name}' table.`);
                                break;
                            }
                        case "RowInserted":
                            rowChange = []
                            await rowInsertedHandler(startRangeRowRelative, endRangeRowRelative, startRangeColRelative, endRangeColRelative, table.name, tableData, range.address.split("!")[1])
                            rowChangeHistory_Undo.push(rowChange)
                            rowChange = []
                            // clean redo if new change
                            rowChangeHistory_Redo = []

                            console.log(`Row Inserted: [${eventArgs.address}] in '${table.name}' table.`)
                            break;
                        case "RowDeleted":
                            rowChange = []
                            await rowDeletedHandler(startRangeRowRelative, endRangeRowRelative, table.name, range.address.split("!")[1])
                            rowChangeHistory_Undo.push(rowChange)
                            rowChange = []
                            // clean redo if new change
                            rowChangeHistory_Redo = []

                            console.log(`Row Deleted: [${eventArgs.address}] in '${table.name}' table.`)
                            break;
                        case "ColumnInserted":

                            console.log(`Column Inserted: [${eventArgs.address}] in '${table.name}' table.`)
                            break;
                        case "ColumnDeleted":

                            console.log(`Column Deleted [${eventArgs.address}] in '${table.name}' table.`)
                            break;
                        case "CellInserted":
                            console.log(`Cell [${eventArgs.address}] in '${table.name}' table.`)
                            break;
                        case "CellDeleted":
                            console.log(`Cell [${eventArgs.address}] in '${table.name}' table.`)
                            break;
                        default:
                            console.log(`Unknown action.`)
                            break;
                    }
                })
            })
        } catch (error) {
            errorHandler(error.message)
        }
    }

    function handleSelectionChange(eventArgs, tableID) {
        try {
            if (undo_redo = true) { undo_redo = undefined }
            if (changesTracker[1][0] === "Fulfilled") { changesTracker = [[], []] }
            changesTracker[0] += "A"

            Excel.run(function (ctx) {
                let table = ctx.workbook.tables.getItem(tableID);
                table.load("name")
                let tableRange = table.getRange()
                tableRange.load("rowCount,columnCount, values")
                let triggerCell = tableRange.getCell(1, 0)
                triggerCell.load("values")

                return ctx.sync().then(async () => {
                    changesTracker[1][0] = "Fulfilled"

                    if (myTables[table.name].length === tableRange.rowCount) {
                        // stop the BBA case for normal row addition
                        return
                    } else if (myTables[table.name].length < tableRange.rowCount &&
                            (changesTracker[0] === "BAB" || changesTracker[0] === "AB")) {
                        // stop the BAB and AB case for normal row addition;
                        return
                    }


                    if (rowChangeHistory_Undo.length > 0 || rowChangeHistory_Redo.length > 0) {
                        let startRangeRowRelative_history
                        let startRangeColRelative_history
                        let endRangeRowRelative_history
                        let endRangeColRelative_history

                        let previousRowChangeType
                        let previousRowAddress
                        let previousRowsCount
                        let previousChangeTracker_normal
                        if (rowChangeHistory_Undo.length > 0) {
                            previousRowChangeType = rowChangeHistory_Undo.at(-1)[0][0]
                            previousRowAddress = rowChangeHistory_Undo.at(-1)[0][2]
                            previousRowsCount = previousRowChangeType === "RowInserted" ? - rowChangeHistory_Undo.at(-1).length : + rowChangeHistory_Undo.at(-1).length
                            previousChangeTracker_normal = rowChangeHistory_Undo.at(-1)[0][3]
                        }

                        let nextRowChangeType
                        let nextRowAddress
                        let nextRowsCount
                        let nextChangeTracker_undo
                        if (rowChangeHistory_Redo.length > 0) {
                            nextRowChangeType = rowChangeHistory_Redo.at(-1)[0][0]
                            nextRowAddress = rowChangeHistory_Redo.at(-1)[0][2]
                            nextRowsCount = nextRowChangeType === "RowInserted" ? + rowChangeHistory_Redo.at(-1).length : - rowChangeHistory_Redo.at(-1).length
                            nextChangeTracker_undo = rowChangeHistory_Redo.at(-1)[0][4]
                        }

                        // stop undo and redo similarity issue
                        let rowsGap = tableRange.rowCount - myTables[table.name].length
                        let advancedEvidence
                        if (changesTracker[0] === "A") {
                            if (previousRowsCount === nextRowsCount) {
                                if (nextChangeTracker_undo !== "A") {
                                    advancedEvidence = "Redo"
                                } else if (!(previousChangeTracker_normal === "B" || previousChangeTracker_normal === "BBA")) {
                                    advancedEvidence = "Undo"
                                } else {
                                    errorHandler("Cannot distinguish between redo and undo")
                                }
                                
                            } else if (previousRowsCount === rowsGap) {
                                advancedEvidence = "Undo"
                            } else if (nextRowsCount === rowsGap) {
                                advancedEvidence = "Redo"
                            }
                        }

                        if (myTables[table.name].length < tableRange.rowCount) {
                            undo_redo = true
                            if (previousRowChangeType === "RowDeleted"
                                && changesTracker[0] === "BA"
                                || advancedEvidence === "Undo") {
                                // undo row deletion
                                startRangeRowRelative_history = rowChangeHistory_Undo.at(-1).at(-1)[1]
                                startRangeColRelative_history = 0
                                endRangeRowRelative_history = rowChangeHistory_Undo.at(-1)[0][1]
                                endRangeColRelative_history = tableRange.columnCount

                                rowInsertedHandler(startRangeRowRelative_history, endRangeRowRelative_history, startRangeColRelative_history, endRangeColRelative_history, table.name, tableRange.values)
                                rowChangeHistory_Redo.push(rowChangeHistory_Undo.pop()) // add this rowAdded history // delete previous rowDeleted history
                                rowChangeHistory_Redo.at(-1)[0][4] = changesTracker[0]
                                rowChange = []

                                console.log(`Undo Row Deletion: [${previousRowAddress}] in '${table.name}' table.`)
                            } else if (nextRowChangeType === "RowInserted"
                                && (changesTracker[0] === "AA" || changesTracker[0] === "ABA")
                                || advancedEvidence === "Redo") {
                                // redo row addition
                                startRangeRowRelative_history = rowChangeHistory_Redo.at(-1)[0][1]
                                startRangeColRelative_history = 0
                                endRangeRowRelative_history = rowChangeHistory_Redo.at(-1).at(-1)[1]
                                endRangeColRelative_history = tableRange.columnCount

                                rowInsertedHandler(startRangeRowRelative_history, endRangeRowRelative_history, startRangeColRelative_history, endRangeColRelative_history, table.name, tableRange.values)
                                rowChangeHistory_Undo.push(rowChangeHistory_Redo.pop()) // add this rowAdded history // delete previous rowDeleted history
                                rowChange = []

                                console.log(`Redo Row Addition: [${nextRowAddress}] in '${table.name}' table.`)
                            }
                        } else if (myTables[table.name].length > tableRange.rowCount) {
                            undo_redo = true
                            if (previousRowChangeType === "RowInserted"
                                && (changesTracker[0] === "AA" || changesTracker[0] === "ABA")
                                || advancedEvidence === "Undo") {
                                // undo row addition
                                startRangeRowRelative_history = rowChangeHistory_Undo.at(-1)[0][1]
                                endRangeRowRelative_history = rowChangeHistory_Undo.at(-1).at(-1)[1]

                                rowDeletedHandler(startRangeRowRelative_history, endRangeRowRelative_history, table.name)
                                rowChangeHistory_Redo.push(rowChangeHistory_Undo.pop()) // delete previous rowAdded history // add this rowDeleted history
                                rowChangeHistory_Redo.at(-1)[0][4] = changesTracker[0]
                                rowChange = []

                                console.log(`Undo Row Addition: [${previousRowAddress}] in '${table.name}' table.`)
                            } else if (nextRowChangeType === "RowDeleted"
                                && changesTracker[0] === "AB"
                                || advancedEvidence === "Redo") {
                                // redo row deletion
                                startRangeRowRelative_history = rowChangeHistory_Redo.at(-1).at(-1)[1]
                                endRangeRowRelative_history = rowChangeHistory_Redo.at(-1)[0][1]

                                rowDeletedHandler(startRangeRowRelative_history, endRangeRowRelative_history, table.name)
                                rowChangeHistory_Undo.push(rowChangeHistory_Redo.pop()) // delete previous rowAdded history // add this rowDeleted history
                                rowChange = []

                                console.log(`Redo Row Deletion: [${nextRowAddress}] in '${table.name}' table.`)
                            }
                        }

                        // stop the redo and undo similarity issue
                        if (rowChangeHistory_Undo.length > 0) {
                            previousRowChangeType = rowChangeHistory_Undo.at(-1)[0][0]
                            previousChangeTracker_normal = rowChangeHistory_Undo.at(-1)[0][3]
                            previousRowsCount = previousRowChangeType === "RowInserted" ? - rowChangeHistory_Undo.at(-1).length : + rowChangeHistory_Undo.at(-1).length
                        }
                        if (rowChangeHistory_Redo.length > 0) {
                            nextRowChangeType = rowChangeHistory_Redo.at(-1)[0][0]
                            nextChangeTracker_undo = rowChangeHistory_Redo.at(-1)[0][4]
                            nextRowsCount = nextRowChangeType === "RowInserted" ? + rowChangeHistory_Redo.at(-1).length : - rowChangeHistory_Redo.at(-1).length
                        }
                        
                        if (previousRowsCount === nextRowsCount) {
                            if (nextChangeTracker_undo === "A"
                                && (previousChangeTracker_normal === "B" || previousChangeTracker_normal === "BBA")
                            ) {
                                rowChangeHistory_Redo = []
                                rowChangeHistory_Undo
                                ctx.workbook.worksheets.getActiveWorksheet().getRange("XFD1048576").values = null
                                ctx.sync()
                            }
                        }
                    }
                })
            })
        } catch (error) {
            errorHandler(error.message)
        }

    }















    function batRequestTest(entityLogicalName) {
        //Batch requests can contain up to 1000 individual requests and can't contain other batch requests.

        const boundary = "batch_" + new Date().getTime();
        const batchUrl = `${resourceDomain}api/data/v9.2/$batch`

        const batchBody =
            `--${boundary}
OData-MaxVersion: 4.0,
OData-Version: 4.0,
Accept: application/json,
Content-Type: application/http
Content-Transfer-Encoding: binary

POST /api/data/v9.2/${entityLogicalName} HTTP/1.1
Content-Type: application/json;type=entry

{sensei_name: "Batch Request Testing", sensei_lessonlearned: ""}

--${boundary}--
\n\n`;

        //`--${boundary}
        //Content-Type: application/http
        //Content-Transfer-Encoding: binary

        //GET /api/data/v9.2/${entityLogicalName} HTTP/1.1

        //--${boundary}--
        //\n\n`;

        console.log(batchBody)

        fetch(batchUrl, {
            method: "POST",
            headers: {
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'Accept': 'application/json',
                'Content-Type': `multipart/mixed;boundary=${boundary}`,
                'Authorization': `Bearer ${accessToken}`,
            },
            body: batchBody
        })

    }



})();

