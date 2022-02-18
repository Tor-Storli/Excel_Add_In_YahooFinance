/// <reference path="scripts/getdata.js" />
/// <reference path="scripts/getdata.js" />
/// <reference path="scripts/getdata.js" />
(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Excel 2016 +, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("Please select the desired time interval (range) from the droplist below. \n\nType in the ticker symbol in the textbox below.");
            $('#button-text').text("Get Data!");
            $('#button-desc').text("Highlights the largest number.");
            $('#txtTicker').text("ABBV");

            $("#btnFormatYahoodata").click(() => tryCatch(FormatData));
            $("#btnGetYahoodata").click(() => tryCatch(importJsonData));
            $("#btnRateofReturn").click(() => tryCatch(RateofReturn));
            $("#btnAddIcons").click(() => tryCatch(applyIconSetFormat));
            $("#btnAddBar").click(() => tryCatch(applyDataBarFormat));
            $("#btnClearConditional").click(() => tryCatch(clearAllConditionalFormats));
            $("#btnCandleStickChart").click(() => tryCatch(addCandleStickChart));
            $("#btnFilter").click(() => tryCatch(filterTable));
            $("#btnClearFilter").click(() => tryCatch(clearFilters));
            $("#btnRemoveChart").click(() => tryCatch(removeChart));


          //  $("#btnReturnsChart").click(() => tryCatch(addReturnsChart));

            //$("#btnFormatYahoodata").hide();
            //$("#btnRateofReturn").hide();
            //$("#btnAddIcons").hide();
            //$("#btnAddBar").hide();
            //$("#btnClearConditional").hide();
            //$("#btnCandleStickChart").hide();
            
           
        });
    };


    // Dynamically create an HTML SCRIPT element that obtains the details for the database ajax call.
    function attachDBScript() {
        var script = document.createElement("script");
        script.type = "text/javascript";
        script.src = "Scripts/GetData.js";

        // Use any selector
        $("head").append(script);
    };

    async function GetDatafromWebApi() {
        const response = await GetDataFromWebService();
        return response;
    };


    // Pivot each array into colums
    // If col index is zero (0) - datetime Unix timestamp - convert the value to a short date format.
    // Multiply by 1000 as JavaScript counts in milliseconds since epoch(which is 01/01/1970)
    function getCol(inputarray, col) {
        var column = [];
        for (var i = 0; i < inputarray.length; i++) {
            if (i == 0) {
                
                column.push(new Date(inputarray[0][col] * 1000).toLocaleDateString());
            }
            else {
                column.push(inputarray[i][col]);
            }
        }
        return column;
    }

 
    async function FormatData() {
        await Excel.run(async (context) => {

            const sheet = context.workbook.worksheets.getItem("Stock");
            let financeTable = sheet.tables.getItem("FinancialData");

            //Format Stock Prices to 2 decimals
            const formats = [["0.00"]];
            for (var i = 1; i <= 5; i++) {
                financeTable.columns.getItemAt(i).getDataBodyRange().numberFormat = formats;
            }

            //Sort by Datetime - Decending
            var sortRange = financeTable.getDataBodyRange();
            sortRange.sort.apply([
                {
                    key: 0,
                    ascending: false,
                },
            ]);

            //Add a new rate of return column to the table
           // financeTable.columns.add(null, [["=LN(F3/F2)"],["Return"]]);

            //console.log(data.length);
           // if (Office.context.requirements.isSetSupported("ExcelApi", "1.3")) {
                sheet.getUsedRange().format.autofitColumns();
                sheet.getUsedRange().format.autofitRows();
           // }
            sheet.activate();
            await context.sync();

        });
    }

    function RateofReturn() {
        Excel.run( (context) => {

            const sheet = context.workbook.worksheets.getItem("Stock");
            let financeTable = sheet.tables.getItem("FinancialData");

            //Sort by Datetime - Decending
            var sortRange = financeTable.getDataBodyRange();
            sortRange.sort.apply([
                {
                    key: 0,
                    ascending: false,
                },
            ]);

            //Add a new rate of return column to the table
            const CCRateofReturn = ["=LN(F2/F3)"];

            var financeTableRows = financeTable.rows;
            financeTableRows.load('items');

            return context.sync().then(function () {
                var rowCount = financeTableRows.count;

                const CCFormulaArray = [];
                CCFormulaArray.push(["Rate of Return"]);

                for (var i = 0; i < rowCount; i++) {
                    CCFormulaArray.push(CCRateofReturn);
                }

                const formats = [["0.00000%"]];

                const newCol = financeTable.columns.add(null, CCFormulaArray);

                newCol.getDataBodyRange().numberFormat = formats;
                
                sheet.getUsedRange().format.autofitColumns();
                sheet.getUsedRange().format.autofitRows();
            });


            sheet.activate();
            return context.sync();

        });
    }

    //Clear All Conditional Formats
    async function clearAllConditionalFormats() {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("Stock");
            const range = sheet.getRange();
            range.conditionalFormats.clearAll();

            await context.sync();

          //  $("#btnClearConditional").hide();
        });
    }

    async function applyDataBarFormat() {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("Stock");
            let financeTable = sheet.tables.getItem("FinancialData");

            var financeTableRows = financeTable.rows;
            financeTableRows.load('items');

            return context.sync().then(function () {
                var rowCount = financeTableRows.count;

                const range = sheet.getRange("F2:F" + rowCount);
                const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
                conditionalFormat.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;

                sheet.getUsedRange().format.autofitColumns();
                sheet.getUsedRange().format.autofitRows();

             //   $("#btnClearConditional").show();
            });


            sheet.activate;
            await context.sync();

         
        });
    }

    async function applyIconSetFormat() {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("Stock");
            let financeTable = sheet.tables.getItem("FinancialData");

            var financeTableRows = financeTable.rows;
            financeTableRows.load('items');

            return context.sync().then(function () {
                var rowCount = financeTableRows.count;

                const range = sheet.getRange("H2:H" + rowCount);
                const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
                const iconSetCF = conditionalFormat.iconSet;
                iconSetCF.style = Excel.IconSet.threeTriangles;

                /*
                   The iconSetCF.criteria array is automatically prepopulated with
                   criterion elements whose properties have been given default settings.
                   You can't write to each property of a criterion directly. Instead,
                   replace the whole criteria object.
       
                   With a "three*" icon set style, such as "threeTriangles", the third
                   element in the criteria array (criteria[2]) defines the "top" icon;
                   e.g., a green triangle. The second (criteria[1]) defines the "middle"
                   icon, The first (criteria[0]) defines the "low" icon, but it
                   can often be left empty as this method does below, because every
                   cell that does not match the other two criteria always gets the low
                   icon.            
               */
                iconSetCF.criteria = [
                    {},
                    {
                        type: Excel.ConditionalFormatIconRuleType.number,
                        operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
                        formula: "=0"
                    },
                    {
                        type: Excel.ConditionalFormatIconRuleType.number,
                        operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
                        formula: "=0.0000001"
                    }
                ];

                sheet.getUsedRange().format.autofitColumns();
                sheet.getUsedRange().format.autofitRows();

              //  $("#btnClearConditional").show();
            });

            sheet.activate;
            await context.sync();

        });
    }



    async function addCandleStickChart() {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("Stock");
           // const financeTable = sheet.tables.getItem("FinancialData");

            let rowCount = $("#divRowCount").text();
            let ticker = $('#txtTicker').val()
            let range = document.querySelector('#selRange').selectedOptions[0].value;

            let dataRange = sheet.getRange("A1:E" + rowCount);
            let chart = sheet.charts.add("StockOHLC", dataRange, "Auto");
            chart.style = 324;
            chart.apply = 8;

            chart.setPosition("J2", "Q20");
            chart.legend.position = "Right";
            chart.legend.format.fill.setSolidColor("white");

            chart.title.text = ticker + " - " + range + " - " + " Price Chart";

            await context.sync();
        });
    }

    async function filterTable() {
        await Excel.run(async (context) => {

            let selFilter = document.querySelector('#selFilter').selectedOptions[0].value;
            const sheet = context.workbook.worksheets.getItem("Stock");
            let financeTable = sheet.tables.getItem("FinancialData");

            let filter = financeTable.columns.getItem("Date").filter;
            filter.apply({
                filterOn: Excel.FilterOn.dynamic,
                dynamicCriteria: Excel.DynamicFilterCriteria = selFilter
            });

            showNotification(`Filter Applied: ` + selFilter, ``);

            //filter = financeTable.columns.getItem("Date").filter;
            //filter.apply({
            //    filterOn: Excel.FilterOn.values,
            //    values: ["Restaurant", "Groceries"]
            //});

            //var financeTableRows = financeTable.rows;
            //financeTableRows.load('items');

            //return context.sync().then(function () {

            //    var rowCount = financeTableRows.count;

            //    showNotification(`Filter Rows (` + rowCount + `): allDatesInPeriodDecember`, ``);

            //    //   $("#btnClearConditional").show();
            //});

            await context.sync();
        });
    }

    async function clearFilters() {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("Stock");

            const financeTable = sheet.tables.getItem("FinancialData");

            financeTable.clearFilters();

            messageBanner.hideBanner();

            await context.sync();
        });
    }

    async function removeChart() {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("Stock");

            let chart = sheet.charts.getItemAt(0);

            chart.delete();

            await context.sync();
        });
    }

    async function addReturnsChart() {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("Stock");
           // const financeTable = sheet.tables.getItem("FinancialData");

            let rowCount = $("#divRowCount").text();
            let ticker = $("#txtTicker").val();
            let range = document.querySelector('#selRange').selectedOptions[0].value;

          
            //let chartR = sheet.charts.add("line", "Auto");
            //chart.style = 324;
            //chart.apply = 8;

            let chartR = sheet.charts.getItemAt(0);

            const seriesCollection = sheet.charts.getItemAt(0).series;
            seriesCollection.load("count");

            await context.sync();

            if (seriesCollection.count > 0) {
                for (let i = 0; i < seriesCollection.count; i++) {

                    let series = seriesCollection.getItemAt(i);
                    series.delete();

                }
             
            }
            
            let xRangeSelection = sheet.getRange("A1:A" + rowCount);
            let rangeSelection = sheet.getRange("H2:H" + rowCount);


            // Add a series.
            let newSeries = seriesCollection.series.add("Returns");
            newSeries.setValues(rangeSelection);
            newSeries.setXAxisValues(xRangeSelection);

            await context.sync();

            chartR.setPosition("J24", "Q40");
            chartR.legend.position = "Right";
            chartR.legend.format.fill.setSolidColor("white");
            chartR.dataLabels.format.font.size = 15;
            chartR.dataLabels.format.font.color = "black";
            chartR.title.text = ticker + " - " + range + " - " +  " Returns Chart";

            await context.sync();
        });
    }

    async function importJsonData() {

        showNotification(`Retrieving data from Yahoo Finance,    this may take a minute...`,``);

        await Excel.run(async (context) => {
            context.workbook.worksheets.getItemOrNullObject("Stock").delete();
            const sheet = context.workbook.worksheets.add("Stock");

            let expensesTable = sheet.tables.add("A1:G1", true);
            expensesTable.name = "FinancialData";
            expensesTable.getHeaderRowRange().values = [["Date", "Open", "High", "Low", "Close", "AdjClose", "Volume"]];

            attachDBScript();

            let JsonOut = await GetDatafromWebApi();
            
            const jsonResponse = JSON.parse(JsonOut.d[0]);

            const datetime = jsonResponse.chart.result[0].timestamp;
            const open = jsonResponse.chart.result[0].indicators.quote[0].open;
            const close = jsonResponse.chart.result[0].indicators.quote[0].close;
            const high = jsonResponse.chart.result[0].indicators.quote[0].high;
            const low = jsonResponse.chart.result[0].indicators.quote[0].low;
            const adjclose = jsonResponse.chart.result[0].indicators.adjclose[0].adjclose;
            const volume = jsonResponse.chart.result[0].indicators.quote[0].volume;


            const arrayout = [];

            for (let i = 0; i < datetime.length; i++) {
                const arrsub = [];

                // datetime is a Unix timestamp - so we need to convert the value to a short date format.
                // How: Multiply by 1000 as JavaScript counts in milliseconds since epoch(which is 01/01/1970)
                arrsub.push(new Date(datetime[i] * 1000).toLocaleDateString());
                arrsub.push(open[i]);
                arrsub.push(high[i]);
                arrsub.push(low[i]);
                arrsub.push(close[i]);
                arrsub.push(adjclose[i]);
                arrsub.push(volume[i]);

                arrayout.push(arrsub);
            }

           // console.log(arr)


            //const data = [datetime, open, high, low, close, adjclose, volume];

           // var arrayout = []; //output array

            //for (let i = 0; i < datetime.length; i++) {
            //    const newcol = getCol(data, i); //Get next column from input array
            //    arrayout.push(newcol);
            //}

            $("#divRowCount").text(arrayout.length);

            expensesTable.rows.add(null, arrayout);

            sheet.getUsedRange().format.autofitColumns();
            sheet.getUsedRange().format.autofitRows();

            const selCell = sheet.getRange("H1:H1");

            selCell.select();
            sheet.activate();


            await context.sync();

            messageBanner.hideBanner();
        });
    }

    // Helper function for displaying notifications 
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    /** Default helper for invoking an action and handling errors. */
    async function tryCatch(callback) {
        try {
            await callback();
        } catch (error) {
            // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
            console.error(error);
        }
    }


})();
