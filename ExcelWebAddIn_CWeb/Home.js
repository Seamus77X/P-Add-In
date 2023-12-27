

(function () {
    "use strict";

    // Declaration of global variables for later use
    let messageBanner;
    let accessToken;  // used to store user's access token
    let workbookGUID
    let runningEnvir
    let pp_eacb_rowIdMapping = { }
    let pp_eacb_fieldNameMapping = {
        'sensei_lessonslearned': [
            ["sensei_lessonlearnedid", "Row ID"],
            ["modifiedon", "Entry Date"],
            ["sensei_name", "Title"],
            ["sensei_category", "Category"],
            ["sensei_lessonlearned", "Detailed Description"],
            ["sc_projectimpact", "Project Impact"],
            ["sensei_observation", "When is it likely to occur/when did it occur"],
            ["sensei_recommendation", "Recommendations"],
            ["sc_additionalcommentsnotes", "Additional Comments/Notes"]
        ],
        'sensei_risks': [],
        'sc_variations': []
    }
    let EntityAttributes = {}


    // Constants for client ID, redirect URL, and resource domain for authentication
    const clientId = "be63874f-f40e-433a-9f35-46afa1aef385"
    const redirectUrl = "https://seamus77x.github.io/index.html"
    const resourceDomain = "https://gsis-pmo-australia-sensei-demo.crm6.dynamics.com/"

    // Initialization function that runs each time a new page is loaded.
    Office.onReady().then(() => {

        $(function () {
            try {

                //////////////////////////////////////////////////
                // Settting: auto open add-in and show taskpane once the add-in is manually opened bu a user
                //////////////////////////////////////////////////
                //Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                //Office.context.document.settings.saveAsync();

                Office.addin.setStartupBehavior(Office.StartupBehavior.load);
                Office.addin.showAsTaskpane();
                //Office.addin.hide();
                //Office.addin.setStartupBehavior(Office.StartupBehavior.none);

                //////////////////////////////////////////////////
                // Initialise Add-In taskpane page
                //////////////////////////////////////////////////
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

                //////////////////////////////////////////////////
                // Print the Excel Platform
                //////////////////////////////////////////////////
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

                //////////////////////////////////////////////////
                // Authentication and access token retrieval logic
                //////////////////////////////////////////////////
                setTimeout(authFunction, 500)
                function authFunction() {
                    if (accessToken === undefined) {
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
                                let dialog = result.value;
                                dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent)
                                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);

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
                                        dialog.close();
                                    }
                                }
                                function processDialogEvent(arg) {
                                    switch (arg.error) {
                                        case 12002:
                                            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
                                            break;
                                        case 12003:
                                            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");
                                            break;
                                        case 12006:
                                            showNotification("Dialog closed.");
                                            break;
                                        default:
                                            showNotification("Unknown error in dialog box.");
                                            break;
                                    }
                                }
                            }
                        )
                    }
                }


                //////////////////////////////////////////////////
                // Register the workbook
                //////////////////////////////////////////////////
                (async () => {
                    Excel.run(async (context) => {
                        const workbook = context.workbook;
                        const properties = workbook.properties;
                        const customProperties = properties.custom;

                        properties.load('creationDate');
                        customProperties.load("items/key, items/value");

                        await context.sync();

                        // Check if the property already exists
                        const existingProperty = customProperties.items.find(prop => prop.key === "Workbook ID");
                        if (existingProperty) {
                            workbookGUID = existingProperty.value
                            console.log("Workbook ID retrieved.");
                        } else {
                            workbookGUID = `[${uuid.v4()}] - ${properties.creationDate}`
                            customProperties.add("Workbook ID", workbookGUID); // Add new property

                            await context.sync();
                            console.log("Workbook ID created.");
                        }
                    });
                })()

                //////////////////////////////////////////////////
                // Add function to Excel ribbon buttons
                //////////////////////////////////////////////////
                Office.actions.associate("buttonFunction", function (event) {
                    console.log('Hey, you just pressed a ribbon button.')
                    //Create_D365('sensei_lessonslearned', { 'sensei_name': 'Add Test', 'sc_additionalcommentsnotes': 'ADD Redo_Undo_Event_Done from Web Add-In' })

                    console.log(pp_eacb_rowIdMapping)
                    console.log(EntityAttributes)
                    event.completed();
                })

            } catch (error) {
                errorHandler(error.message)
            }
        });
    })

    async function loadSampleData() {

        //sc_integrationrecentgranulartransactions
        //sensei_financialtransactions?$select=sc_kbrkey,sc_vendorname,sensei_value,sc_docdate,sensei_financialtransactionid&$top=50000

        await Excel.run(async (ctx) => {
            ctx.application.calculationMode = Excel.CalculationMode.manual;
            ctx.runtime.enableEvents = false;
            await ctx.sync();

            await loadData(['sensei_lessonslearned']);

            ctx.application.calculationMode = Excel.CalculationMode.automatic;
            ctx.runtime.enableEvents = true;

            await ctx.sync();
        });
    }


    // Function to retrieve data from Dynamics 365
    async function loadDaaata(resourceUrl, tableName, FirstDataColumnIndex = 1, defaultSheet = 'Sheet1', defaultTpLeftRng = 'A1', excludedColsNames = ['@odata.etag']) {
        try {
            let DataArr = await Read_D365(resourceUrl);

            // act as the corresponding table in memory, which records the change in Excel table
            pp_eacb_rowIdMapping[tableName] = _.cloneDeep(DataArr)

            //if (DataArr.length === 0) {return}

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
                        // Situation 1: If the table exists, update existing one
                        // delete header row of DataArr
                        DataArr.shift()

                        // add LHS and RHS formula cols to expand dataArr
                        let excelTableRightColNo = columnNameToNumber(oldRangeAddress.split(":")[1].replace(/\d+$/, ''))
                        let ppTableRightColNo = columnNameToNumber(oldRangeAddress.split(":")[0].replace(/\d+$/, '')) + FirstDataColumnIndex - 1 + DataArr[0].length - 1
                        DataArr.forEach(row => {
                            if (FirstDataColumnIndex > 1) {
                                let tempRowFormula = oldFirstRow_formula
                                row.unshift(...tempRowFormula[0].slice(0, FirstDataColumnIndex - 1))
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
            // add listener to the table if no listener
            registerTableChangeEvent(tableName)
        }
    }


    async function loadData(validTables) {
        try {
            await Excel.run(async (context) => {
                const workbook = context.workbook;
                const worksheets = workbook.worksheets;
                worksheets.load('items/tables/items/name, items/tables/items/id, items/tables/items/columns/items/name, items/tables/items/columns/items/index');
                await context.sync();

                const tablePromises = [];
                for (const sheet of worksheets.items) {
                    for (const table of sheet.tables.items) {
                        if (validTables.includes(table.name)) {
                            tablePromises.push(processTable(table, sheet, context));
                            // add listener to the table if no listener
                            registerTableChangeEvent(table)
                        }
                    }
                }

                // Execute all table updates concurrently
                await Promise.all(tablePromises);
            });
        } catch (error) {
            errorHandler(error);
        }

        async function processTable(table, sheet, context) {
            let selectArr
            let filterArr
            let selectCondition
            let filterCondition
            let expandCondition
            let entityPath
            let url
            let mappingArray = pp_eacb_fieldNameMapping[table.name];
            let tablePromises = []

            // get Logical Name of the table
            entityPath = 'EntityDefinitions'
            selectCondition = '?$select=LogicalName'
            filterCondition = `&$filter=EntitySetName eq '${table.name}'`
            url = `${resourceDomain}api/data/v9.2/${entityPath}${selectCondition}${filterCondition}&LabelLanguages=1033`
            let EntityLogicalName = await Read_D365(url).then((result) => {
                return result[1][result[0].indexOf("LogicalName")]
            })

            // get the Field Proerties of the table
            EntityAttributes[table.name] = {}

            entityPath = `EntityDefinitions(LogicalName='${EntityLogicalName}')/Attributes`
            selectArr = ['MetadataId', 'LogicalName', 'AttributeType', 'IsPrimaryId', 'IsPrimaryName']
            filterArr = mappingArray.map(entry => `LogicalName eq '${entry[0]}'`);
            selectCondition = `?$select=${selectArr.join(',')}`
            filterCondition = `&$filter=${filterArr.join(' or ')}`
            
            url = `${resourceDomain}api/data/v9.2/${entityPath}${selectCondition}${filterCondition}&LabelLanguages=1033`;
            tablePromises.push(Read_D365(url).then((result) => {
                let IsPrimaryId_colNo = result[0].indexOf('IsPrimaryId')
                let LogicalName_colNo = result[0].indexOf('LogicalName')
                let primaryIdRow = result.find(row => row[IsPrimaryId_colNo] === true)
                let primaryColName = primaryIdRow ? primaryIdRow[LogicalName_colNo] : undefined;

                let excludedCols_index = ['MetadataId', '@odata.type', 'IsPrimaryId', 'IsPrimaryName'].map(fieldName => result[0].indexOf(fieldName))
                let properties = result.map(row => row.filter((item, index) => !excludedCols_index.includes(index)))

                EntityAttributes[table.name]["PrimaryName"] = primaryColName
                EntityAttributes[table.name]["EntityDefinitions"] = properties
            }))

            // get the PickLists Field Properties of the table
            selectArr = ['MetadataId', 'LogicalName']
            selectCondition = `?$select=${selectArr.join(',')}`
            entityPath = `EntityDefinitions(LogicalName='${EntityLogicalName}')/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata`
            expandCondition = `&$expand=GlobalOptionSet($select=Options)` 

            url = `${resourceDomain}api/data/v9.2/${entityPath}${selectCondition}${expandCondition}&LabelLanguages=1033`;
            tablePromises.push(Read_D365(url).then((result) => {
                let thisPickList = {}

                let headerRow = result.shift()
                let picklistName_colNum = headerRow.indexOf('LogicalName')
                let picklist_colNum = headerRow.indexOf('GlobalOptionSet / Options')

                result.forEach(row => {
                    let picklist_name = row[picklistName_colNum]
                    let picklist_options = row[picklist_colNum].map(option => {
                        let value = option['Value']
                        let key = option['Label']['UserLocalizedLabel']['Label']
                        return [key, value]
                    })
                    thisPickList[picklist_name] = picklist_options
                }) 

                EntityAttributes[table.name]["PickLists"] = thisPickList
            }))

            // get P+ table
            entityPath = table.name
            selectCondition = `?$select=${mappingArray.map(entry => entry[0]).join(',')}`;
            url = `${resourceDomain}api/data/v9.2/${entityPath}${selectCondition}`;
            tablePromises.push(Read_D365(url))

            let results = await Promise.all(tablePromises);
            let DataArr = results.at(-1).map(row => 
                row.map(item => item === null ? '' : item)
            )

            // Update existing table
            let logicalHeaderNames = DataArr.shift(); // Remove header row
            let primaryColNum = logicalHeaderNames.indexOf(EntityAttributes[table.name]["PrimaryName"])
            // Create an array of promises
            let mappingPromises = DataArr.map(row => {
                return hashString(JSON.stringify(row))
                    .then(hash => [row[primaryColNum], hash]);
            });

            // Resolve all promises and assign the results
            Promise.all(mappingPromises).then(mappingTab => {
                pp_eacb_rowIdMapping[workbookGUID + table.name] = mappingTab;
            })

            const firstRowRange = table.getDataBodyRange().getRow(0).load('formulas');
            const dataBodyRange = table.getDataBodyRange().load('address');
            await context.sync();

            const firstRowFormulas = firstRowRange.formulas[0];
            const updatedData = DataArr.map(row => {
                return table.columns.items.map(column => {
                    const mappingEntry = mappingArray.find(entry => entry[1] === column.name);
                    if (mappingEntry) {
                        const columnIndexInUpdatedData = logicalHeaderNames.findIndex(col => col === mappingEntry[0]);
                        return columnIndexInUpdatedData >= 0 ? row[columnIndexInUpdatedData] : null;
                    } else {
                        return firstRowFormulas[column.index];
                    }
                });
            });

            // Clear existing data in the table
            table.getDataBodyRange().clear()

            // Paste the updated data into Excel
            let startRow = parseInt(dataBodyRange.address.split("!")[1].match(/\d+/), 10);
            let newEndRow = startRow + DataArr.length - 1;
            let newRangeAddress = dataBodyRange.address.split("!")[1].replace(/\d+$/, newEndRow);

            const updatedRange = sheet.getRange(newRangeAddress);
            updatedRange.values = updatedData;

            // Resize the table including the header row
            startRow = parseInt(newRangeAddress.match(/\d+/), 10);
            let headerRowNumber = startRow - 1;
            let resizedRangeAddress = newRangeAddress.replace(/\d+/, headerRowNumber);

            let resizedRange = sheet.getRange(resizedRangeAddress);
            table.resize(resizedRange);

            resizedRange.format.autofitColumns();
            resizedRange.format.autofitRows();

            await context.sync();
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
    function pasteChunksToExcel(chunks, rangeAddressToPaste, sheet, ctx) {
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
            ctx.sync();

            startRow += chunkRowCount; // Update startRow for the next chunk
        }
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
                console.log("No data found");
                return finalArr
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

    let sheetEventListeners = {};
    function registerTableChangeEvent(table) {
        try {
            // Check if listener has already been added
            if (!sheetEventListeners[table.id]) {
                Excel.run(function (ctx) {
                    // if the table found, then listen to the change in the table
                    table.onChanged.add(handleTableChange)
                    table.worksheet.onSelectionChanged.add((eventArgs) => {
                        handleSelectionChange(eventArgs, table.id);
                    });
                    console.log(`Events Listener Added: ${table.name}`)
                    // Update the map
                    sheetEventListeners[table.id] = true;
                    return ctx.sync()
                })
            }
        } catch (error) {
            // Error handling for issues within the Excel.run block
            errorHandler(error.message);
        }
    }



    let undo_redo
    let tableEventAddress
    let sheetEventAddress
    let previousTableData
    let eventsTracker = ['', '']
    let multi_undo_redo = ''
    // row change events handlers
    function rowInsertedHandler(startRow, endRow, startCol, endCol, tableName, tableData) {
        let rowId_Mapping = pp_eacb_rowIdMapping[workbookGUID + tableName]
        let promiseArray = [];

        for (let r = startRow; r <= endRow; r++) {
            let jsonPayLoad = {};
            for (let c = startCol; c <= endCol; c++) {
                let displayColName = tableData[0][c];
                let mappingEntry = pp_eacb_fieldNameMapping[tableName].find(entry => entry[1] === displayColName);
                if (mappingEntry) {
                    let logicalColName = mappingEntry[0];
                    jsonPayLoad[logicalColName] = tableData[r][c];
                }
            }

            if (Object.keys(jsonPayLoad).length > 0) {
                let createD365Promise = Create_D365(tableName, jsonPayLoad, EntityAttributes[tableName]["PrimaryName"])
                // update row mapping table
                let tempID = uuid.v4()
                if (r - 1 + 1 > rowId_Mapping.length) {
                    rowId_Mapping.push([createD365Promise, "Hash", tempID]);
                } else {
                    rowId_Mapping.splice(r - 1, 0, [createD365Promise, "Hash", tempID]);
                }
                // replace promise with the guid once promise is resolved
                (async () => {
                    let guid = createD365Promise instanceof Promise ? await createD365Promise : createD365Promise;
                    let existedRow = rowId_Mapping.find(rowInfo => rowInfo[2] === tempID)
                    if (existedRow) {
                        rowId_Mapping[r - 1] = [guid, "Hash"]
                    }
                })()
            }
        }
    }

    function rowDeletedHandler(startRow, endRow, tableName) {
        let rowId_Mapping = pp_eacb_rowIdMapping[workbookGUID + tableName]

        // collect the row num of the rows deleted
        for (let r = endRow; r >= startRow; r--) {
            // Ensure the GUID is resolved
            let guidOrPromise = rowId_Mapping[r - 1][0];

            (async () => {
                let guid = guidOrPromise instanceof Promise ? await guidOrPromise : guidOrPromise;
                // delete the rows in P+ table
                Delete_D365(tableName, guid)
            })()
            
            // delete the row in memeory table as well
            rowId_Mapping.splice(r-1, 1)
        }
    }
    // range content change event handler
    function rangeChangeHandler(startRow, endRow, startCol, endCol, thisTableData, thisTableName, thisEventArgs) {
        // construct the JSON PayloadRowNo_RowGUID_MappingTable
        let rowId_Mapping = pp_eacb_rowIdMapping[workbookGUID + thisTableName]

        for (let r = startRow; r <= endRow; r++) {
            let jsonPayLoad = {};
            for (let c = startCol; c <= endCol; c++) {
                let displayColName = thisTableData[0][c];
                // Use pp_eacb_fieldNameMapping for converting display names to logical names.
                let mappingEntry = pp_eacb_fieldNameMapping[thisTableName].find(entry => entry[1] === displayColName);
                if (mappingEntry) {
                    let logicalColName = mappingEntry[0];
                    jsonPayLoad[logicalColName] = thisTableData[r][c];
                }
            }

            if (Object.keys(jsonPayLoad).length > 0) {
                // Ensure the GUID is resolved
                let guidOrPromise = rowId_Mapping[r - 1][0];

                (async () => {
                    let guid = guidOrPromise instanceof Promise ? await guidOrPromise : guidOrPromise;
                    if (thisEventArgs !== undefined && thisEventArgs.details !== undefined
                        && JSON.stringify(thisEventArgs.details.valueAsJsonAfter) === JSON.stringify(thisEventArgs.details.valueAsJsonBefore)) {
                        // if range content is unchanged, then do not sync
                    } else {
                        Update_D365(thisTableName, guid, jsonPayLoad)
                    }
                })()
            }
        }
    }

    // hanle table change.
    function handleTableChange(eventArgs) {
        try {
            if (undo_redo = true) { undo_redo = undefined }
            if (eventsTracker[1] === "Fulfilled") { eventsTracker = ['', ''] }
            eventsTracker[0] += "B"
            tableEventAddress = eventArgs.address

            let thisTableChangeType = eventArgs.changeType
            Excel.run( (ctx) => {
                // load necessary info from Excel
                let range
                let isDiscontinuousRanges
                if (eventArgs.address.includes(",")) {
                    range = ctx.workbook.worksheets.getActiveWorksheet().getRanges(eventArgs.address).areas
                    range.load("items/values, items/address, items/rowIndex, items/columnIndex, items/cellCount, items/rowCount, items/columnCount")
                    isDiscontinuousRanges= true
                } else {
                    range = ctx.workbook.worksheets.getActiveWorksheet().getRange(eventArgs.address)
                    range.load("values, address, rowIndex, columnIndex, cellCount, rowCount, columnCount")
                    isDiscontinuousRanges = false
                }
                
                let table = ctx.workbook.tables.getItem(eventArgs.tableId);
                table.load("name")
                let tableRange = table.getRange()
                tableRange.load("rowIndex, columnIndex, rowCount, columnCount, values")

                return ctx.sync().then(() => {
                    let rowId_Mapping = pp_eacb_rowIdMapping[workbookGUID + table.name]
                    eventsTracker[1] = "Fulfilled"
                    previousTableData = tableRange.values // update the previous data

                    ////////////////////////////////////////////
                    // Special Case Check
                    ////////////////////////////////////////////

                    let discontinuousRangeChange_undo_redo = false
                    // continue if undo or redo discontinuous range content change
                    if (sheetEventAddress !== undefined && sheetEventAddress.includes(",")) {
                        let numberOfRanges = sheetEventAddress.split(",").length
                        if (eventsTracker[0] === "B".repeat(numberOfRanges) + "A"
                            || "A" + eventsTracker[0] === "B".repeat(numberOfRanges)) {
                            discontinuousRangeChange_undo_redo = true
                        }
                    // quit and let sheet event listener handle this if mutiple continuous undo or redo row changes
                    } else if (multi_undo_redo === true || eventsTracker[0] === "BAA" || eventsTracker[0].length >= 4) {
                        // stop the multiple continuous undo and redo opeartions
                        return
                    // allow or stop range update after redo or undo row changes are handled by sheet event listener
                    } else if (undo_redo === true) {
                        if (eventsTracker[0] === "AB") {
                            // stop the AB case for redo row deletion
                            return
                        } else if (undo_redo === true && eventsTracker[0] === "ABA") {
                            // allow ABA case for redoing or undoing row addition
                        }
                    // allow or stop range update before row change are handled by sheet event listener
                    } else if (rowId_Mapping.length + 1 < tableRange.rowCount && thisTableChangeType === "RangeEdited") {
                        if (eventsTracker[0] === "BA") {
                            // stop the BA case for undoing row deletion
                            return
                        } else if (eventsTracker[0] === "BAB") {
                            // allow BAB case for normal row addition
                        }
                    }

                    let tableData = tableRange.values
                    let tableStartRow = tableRange.rowIndex;
                    let tableStartCol = tableRange.columnIndex;

                    let startRangeRowRelative
                    let startRangeColRelative
                    let endRangeRowRelative
                    let endRangeColRelative
                    let startRangeRowRelative_Coll
                    let startRangeColRelative_Coll
                    let endRangeRowRelative_Coll
                    let endRangeColRelative_Coll

                    if (isDiscontinuousRanges) {
                        startRangeRowRelative_Coll = range.items.map(item => item.rowIndex - tableStartRow)
                        startRangeColRelative_Coll = range.items.map(item => item.columnIndex - tableStartCol)
                        endRangeRowRelative_Coll = range.items.map(item => item.rowIndex - tableStartRow + item.rowCount - 1)
                        endRangeColRelative_Coll = range.items.map(item => item.columnIndex - tableStartCol + item.columnCount - 1)
                    } else {
                        startRangeRowRelative = range.rowIndex - tableStartRow;
                        startRangeColRelative = range.columnIndex - tableStartCol;
                        endRangeRowRelative = startRangeRowRelative + range.rowCount - 1
                        endRangeColRelative = startRangeColRelative + range.columnCount - 1
                    }

                    switch (thisTableChangeType) {
                        case 'RangeEdited':
                            if (isDiscontinuousRanges) {
                                startRangeRowRelative_Coll.forEach( (item, index) => {
                                    rangeChangeHandler(startRangeRowRelative_Coll[index], endRangeRowRelative_Coll[index], startRangeColRelative_Coll[index], endRangeColRelative_Coll[index], tableData, table.name, eventArgs)
                                })
                            } else {
                                // if case BBA, then do not sync
                                if (["BBA", "BB"].includes(eventsTracker[0]) && discontinuousRangeChange_undo_redo === false) { break }
                                // if range content is unchanged, then do not sync
                                if (eventArgs.details !== undefined && JSON.stringify(eventArgs.details.valueAsJsonAfter) === JSON.stringify(eventArgs.details.valueAsJsonBefore)) {break}
                                // if all okay, then sync
                                rangeChangeHandler(startRangeRowRelative, endRangeRowRelative, startRangeColRelative, endRangeColRelative, tableData, table.name, eventArgs)
                            }

                            console.log(`Range Updated: [${eventArgs.address}] in '${table.name}' table.`);
                            break;
                        case "RowInserted":
                            rowInsertedHandler(startRangeRowRelative, endRangeRowRelative, startRangeColRelative, endRangeColRelative, table.name, tableData)
                            console.log(`Row Inserted: [${eventArgs.address}] in '${table.name}' table.`)
                            break;
                        case "RowDeleted":
                            rowDeletedHandler(startRangeRowRelative, endRangeRowRelative, table.name)
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
            if (eventsTracker[1] === "Fulfilled") { eventsTracker = ['', ''] }  // refersh event tracker
            eventsTracker[0] += "A" // record event order
            multi_undo_redo = ''  // reset multi redo/undo tracker
            let previousTableData_copy = previousTableData !== undefined ? previousTableData : undefined //keep a copy of current table data
            sheetEventAddress = eventArgs.address

            Excel.run(function (ctx) {
                let table = ctx.workbook.tables.getItem(tableID);
                table.load("name")
                let tableRange = table.getRange()
                tableRange.load("rowIndex, columnIndex, rowCount, columnCount, values")
                let ABA_updatedRange
                if (eventsTracker[0] === "ABA") {
                    ABA_updatedRange = table.worksheet.getRange(tableEventAddress)
                    ABA_updatedRange.load("rowIndex, columnIndex, rowCount, columnCount, values")
                }

                return ctx.sync().then(() => {
                    if (multi_undo_redo) {return}

                    let rowId_Mapping = pp_eacb_rowIdMapping[workbookGUID + table.name]
                    eventsTracker[1] = "Fulfilled"
                    previousTableData = tableRange.values // update the previous data

                    function getRowChangeStartPosition(arrayA, arrayB) {
                        if (_.isEqual(arrayA,arrayB)) {
                            return "equivalent"
                        }

                        let aIndex = 0, bIndex = 0;

                        while (aIndex < arrayA.length || bIndex < arrayB.length) {
                            if (aIndex > arrayA.length || bIndex > arrayB.length) {
                                // where the row change occurs
                                return aIndex
                            } else if (JSON.stringify(arrayA[aIndex]) === JSON.stringify(arrayB[bIndex])) {
                                aIndex++;
                                bIndex++;
                            } else {
                                // where the row change occurs
                                return aIndex
                            }
                        }
                    }
                    let rowsGap = previousTableData_copy !== undefined ? tableRange.rowCount - previousTableData_copy.length : undefined
                    let currentTableData = tableRange.values

                    if (eventsTracker[0] === "ABA") {
                        // update the previous table first to prevent wrong start position of row change
                        let startPosition_rowsUpdated = ABA_updatedRange.rowIndex - tableRange.rowIndex
                        let rowsUpdatedCount = ABA_updatedRange.rowCount
                        previousTableData_copy.splice(startPosition_rowsUpdated, rowsUpdatedCount, ...currentTableData.slice(startPosition_rowsUpdated, startPosition_rowsUpdated + rowsUpdatedCount))
                    }
                    let rowChangeStartPosition = previousTableData_copy !== undefined ? getRowChangeStartPosition(previousTableData_copy, currentTableData) : undefined

                    ////////////////////////////////////////////
                    // Special Case Check
                    ////////////////////////////////////////////

                    // quit and let table event listener handle this if redo or undo multiple discontinuous range content change
                    if (sheetEventAddress.includes(",")) {
                        let numberOfRanges = sheetEventAddress.split(",").length
                        if (eventsTracker[0] === "B".repeat(numberOfRanges) + "A"
                            || "A" + eventsTracker[0] === "B".repeat(numberOfRanges)) {
                                return
                        }
                    } else if (['B', 'BB', 'BBA', 'AB', 'BAB'].includes(eventsTracker[0])) {
                        // normal row change operations cannot be multiple continuous undo/redo operations
                    } else if (['A','AB', 'BA'].includes(eventsTracker[0])) {
                        // reddo and undo case: A, AB and BA, cannot be multiple continuous undo/redo operations
                    }  else {
                        // continue and remark if multiple continuous undo/redo operations
                        function compareTables(arrayA, arrayB, startPosition, rowsGap) {
                            let arrayA_copy = _.cloneDeep(arrayA)
                            let endPosition = rowChangeStartPosition + Math.abs(rowsGap) - 1
                            if (rowsGap > 0) {
                                // add previous table
                                arrayA_copy.splice(startPosition, 0, ...arrayB.slice(startPosition, endPosition + 1))
                            } else if (rowsGap < 0) {
                                // delete previous table
                                arrayA_copy.splice(startPosition, Math.abs(rowsGap))
                            }
                            return _.isEqual(arrayA_copy, arrayB)
                        }

                        if (eventsTracker[0].length === 1) {
                            // single redo or undo operation: A
                        } else if (eventsTracker[0].length === 2) {
                            if (eventsTracker[0] === "AA") {
                                // must be undo or redo multiple continuous row changes if no change in tables content or tables row
                                if (rowChangeStartPosition === 'equivalent' || rowsGap === 0) {
                                    multi_undo_redo = true
                                // replicate the previous operation for previous table to match current table, if equal, then remark it as not multiple...
                                } else if (!compareTables(previousTableData_copy, currentTableData, rowChangeStartPosition, rowsGap)) {
                                    // if not equal, then must be multiple redo/undo
                                    multi_undo_redo = true
                                }
                            }
                        } else if (eventsTracker[0].length === 3) {
                            if (eventsTracker[0] === "ABA") {
                                if (rowChangeStartPosition === 'equivalent' || rowsGap === 0) {
                                    multi_undo_redo = true
                                } else if (!compareTables(previousTableData_copy, currentTableData, rowChangeStartPosition, rowsGap)) {
                                    // if not equal, then must be multiple redo/undo
                                    multi_undo_redo = true
                                }
                            } else if (["BAA", "AAA", "AAB"].includes(eventsTracker[0])) {
                                // these cases must be multiple....
                                multi_undo_redo = true
                            }
                        // if symbol length > 4, it must be multiple....
                        } else if (eventsTracker[0].length >= 4) {
                            multi_undo_redo = true
                        }
                    }

                    if (multi_undo_redo) {
                        // replace the existing table because multiple continuous redo/undo operations cannot be splited into individual operation
                        (async () => {
                            await rowDeletedHandler(1, rowId_Mapping.length, table.name)
                            rowInsertedHandler(1, currentTableData.length - 1, 0, currentTableData[0].length - 1, table.name, currentTableData)
                        })()
                        
                        console.log(`Multiple continuous undo or redo operations detected: The whole '${table.name}' table is updated.`);
                        return 
                    }

                    // stop if normal row changes case, allows only single undo or redo operation
                    if (rowId_Mapping.length + 1 === tableRange.rowCount) {
                        // stop the BBA case for normal row addition
                        return
                    } else if (rowId_Mapping.length + 1 < tableRange.rowCount &&
                            (eventsTracker[0] === "BAB" || eventsTracker[0] === "AB")) {
                        // stop the BAB and AB case for normal row addition;
                        return
                    } 

                    // sync undo or redo row changes
                    let startRangeRowRelative
                    let startRangeColRelative
                    let endRangeRowRelative
                    let endRangeColRelative

                    if (rowsGap > 0) {
                        undo_redo = true
                        // undo row deletion or redo row addition
                        startRangeRowRelative = rowChangeStartPosition
                        startRangeColRelative = 0
                        endRangeRowRelative = rowChangeStartPosition + Math.abs(rowsGap) - 1
                        endRangeColRelative = startRangeColRelative + tableRange.columnCount - 1

                        rowInsertedHandler(startRangeRowRelative, endRangeRowRelative, startRangeColRelative, endRangeColRelative, table.name, currentTableData)
                        console.log(`Row Inserted: [${tableRange.rowIndex + startRangeRowRelative + 1}:${tableRange.rowIndex + endRangeRowRelative + 1}] in '${table.name}' table.`)
                    } else if (rowsGap < 0) {
                        undo_redo = true
                        // undo row addition or redo row deletion
                        startRangeRowRelative = rowChangeStartPosition
                        endRangeRowRelative = rowChangeStartPosition + Math.abs(rowsGap) - 1

                        rowDeletedHandler(startRangeRowRelative, endRangeRowRelative, table.name)
                        console.log(`Row Deleted: [${tableRange.rowIndex + startRangeRowRelative + 1}:${tableRange.rowIndex + endRangeRowRelative + 1}] in '${table.name}' table.`)
                    }
                })
            })
        } catch (error) {
            errorHandler(error.message)
        }

    }

    async function hashString(string) {
        const encoder = new TextEncoder();
        const data = encoder.encode(string);
        const hashBuffer = await crypto.subtle.digest('SHA-256', data);
        const hashArray = Array.from(new Uint8Array(hashBuffer));
        const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
        return hashHex;
    }


})();

