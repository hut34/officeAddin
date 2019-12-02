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
          { token: 'eyJhbGciOiJSUzI1NiIsImtpZCI6IjRhOWEzMGI5ZThkYTMxNjY2YTY3NTRkZWZlZDQxNzQzZjJlN2FlZWEiLCJ0eXAiOiJKV1QifQ.eyJuYW1lIjoiVGltIE1jTmFtYXJhIiwicGljdHVyZSI6Imh0dHBzOi8vbGgzLmdvb2dsZXVzZXJjb250ZW50LmNvbS8tdVFSOXRjdTRUVTQvQUFBQUFBQUFBQUkvQUFBQUFBQUFBR28vVk03SUlJb2xKY1kvcGhvdG8uanBnIiwib3duZXIiOnRydWUsImFkbWluIjp0cnVlLCJpc3MiOiJodHRwczovL3NlY3VyZXRva2VuLmdvb2dsZS5jb20vaHV0MzRkYXRhaHViIiwiYXVkIjoiaHV0MzRkYXRhaHViIiwiYXV0aF90aW1lIjoxNTczMzc0ODU4LCJ1c2VyX2lkIjoieWhYQmdTb3VINGJBTGdYd282VngwbW9LSGpOMiIsInN1YiI6InloWEJnU291SDRiQUxnWHdvNlZ4MG1vS0hqTjIiLCJpYXQiOjE1NzUwNzI3ODksImV4cCI6MTU3NTA3NjM4OSwiZW1haWwiOiJ0aW1AaHV0MzQuaW8iLCJlbWFpbF92ZXJpZmllZCI6dHJ1ZSwiZmlyZWJhc2UiOnsiaWRlbnRpdGllcyI6eyJnb29nbGUuY29tIjpbIjExMzg3Nzk0Njc0Mjc1ODA3ODE0NyJdLCJlbWFpbCI6WyJ0aW1AaHV0MzQuaW8iXX0sInNpZ25faW5fcHJvdmlkZXIiOiJnb29nbGUuY29tIn19.RwfBBrZ2DKuk7rGCPF1EOCRbpW9kJ8NsLvnfp0OLMxZMpoFEVAJ3fXWifreGWIYHpcbH9b3iYszz0mrFOvxQIZxEsNZR6y78uYgKiZkwgxn8xhQbdVv19hVdZg89XCwUtre7Bkw7W_rAQpSDp1hEmarS9BrRYNtVtZTYaVmnnArJo5f3QCXBMNbbqGlIF5zFcxYMhvbAcfAJH2tm9TZ7zaJ7ewEZ09ejilnkXq3BwpozUznHxr85GcIEH0c4QmKIp5VqPCX_MjCEcYTCv6hdwruF2cYZvJW4WOPDTISJPULY1Qq-AVC6By52DRoHWcY_yTBPHvQ-9Uf1gBkqCqnlOQ',
            accessToken: 'ya29.ImCwB_nN3D8bqInzWEH5J-aPaGgotTxt3Y8ZQSO1RS7cxES5J-OT5XWRlQdyVcuv-gkc4ZqrJcbo6v-2fm46jFroml5yzCJIimyY7aXcLQDtpF-qK6ke5-TTYLWii2tJwss',
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
