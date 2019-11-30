/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {

 // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
      console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

// Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable;
    document.getElementById("download-dataset").onclick = downloadDataset;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

  }
});

function createTable() {
  Excel.run(function (context) {
    // TODO1: Queue table creation logic here.
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    // TODO2: Queue commands to populate the table with data.

    expensesTable.getHeaderRowRange().values =
        [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
      ["1/1/2017", "The Phone Company", "Communications", "120"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
      ["1/11/2017", "Bellows College", "Education", "350.1"],
      ["1/15/2017", "Trey Research", "Other", "135"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    ]);

    // TODO3: Queue commands to format the table.

    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();


    return context.sync();
  })
      .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
      });
}

//tim's tested function...

function downloadDataset() {
  Excel.run(function (context) {

    //retrieve dataset from dataHut
    var request = require("request");
    var options = { method: 'POST',
      url: 'https://hut34datahub.appspot.com/user/downloadFile',
      headers:
          { 'postman-token': '77d3bea5-f7dd-59be-7aee-9605aa7278ee',
            'cache-control': 'no-cache',
            'content-type': 'application/json' },
      body:
          { token: 'ey blah blah blah',
            accessToken: 'ya29.blah blah blah',
            datasetId: '3htIYCymXD8evnf84MfT' },
      json: true };
    request(options, function (error, response, body) {
      if (error) throw new Error(error);
      console.log(body);
    });

    //todo calculate size of dataset rows and columns to set table size?
    //todo find name of dataset to assign to hutDataset.name

    // TODO1: Queue table creation logic here.
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var hutDataset = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    hutDataset.name = "downloadedDataset";

    // TODO2: Queue commands to populate the table with data.
    hutDataset.getHeaderRowRange().values =
        [["Date", "Driver", "Track", "Race"]];

    hutDataset.rows.add(null /*add at the end*/, [
      ["1/1/2017", "The Phone Company", "Communications", "120"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
      ["1/11/2017", "Bellows College", "Education", "350.1"],
      ["1/15/2017", "Trey Research", "Other", "135"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    ]);

    // TODO3: Queue commands to format the table.
    hutDataset.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
    hutDataset.getRange().format.autofitColumns();
    hutDataset.getRange().format.autofitRows();

    return context.sync();

  })
      .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
      });
}
