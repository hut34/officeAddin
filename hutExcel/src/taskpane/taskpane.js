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
    document.getElementById("getDatasets").onclick = getDatasets;
    document.getElementById("getDataset").onclick = getDataset;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});


var expensesTable

var apiURL = "http://localhost:8080"

var idToken = "eyJhbGciOiJSUzI1NiIsImtpZCI6IjA0NjUxMTM5ZDg4NzUyYjY0OTM0MjUzNGE2YjRhMDUxMjVkNzhmYmIiLCJ0eXAiOiJKV1QifQ.eyJuYW1lIjoiUGV0ZXIgR29kYm9sdCIsInBpY3R1cmUiOiJodHRwczovL2xoNS5nb29nbGV1c2VyY29udGVudC5jb20vLXdaeElGR2F1RHVnL0FBQUFBQUFBQUFJL0FBQUFBQUFBQUFjL3M3dFF0OFdkXzV3L3Bob3RvLmpwZyIsIm93bmVyIjp0cnVlLCJhZG1pbiI6dHJ1ZSwiaXNzIjoiaHR0cHM6Ly9zZWN1cmV0b2tlbi5nb29nbGUuY29tL3N0YW5kYXJkZGF0YWh1YiIsImF1ZCI6InN0YW5kYXJkZGF0YWh1YiIsImF1dGhfdGltZSI6MTU3NTU4OTU2NSwidXNlcl9pZCI6IklkbWlKSTNPbXlSWTRkZ0txZzdpRXRMTG1ZcDEiLCJzdWIiOiJJZG1pSkkzT215Ulk0ZGdLcWc3aUV0TExtWXAxIiwiaWF0IjoxNTc1NTg5NTY2LCJleHAiOjE1NzU1OTMxNjYsImVtYWlsIjoicGV0ZXJAaHV0MzQuaW8iLCJlbWFpbF92ZXJpZmllZCI6dHJ1ZSwiZmlyZWJhc2UiOnsiaWRlbnRpdGllcyI6eyJnb29nbGUuY29tIjpbIjExMzI2ODUyODQ1Mzc2NjYyMzcyNyJdLCJlbWFpbCI6WyJwZXRlckBodXQzNC5pbyJdfSwic2lnbl9pbl9wcm92aWRlciI6Imdvb2dsZS5jb20ifX0.yQ599CmSAnCFSHEPxFQc6GSsupcBjves6weZN3jjUzJOaUzr0TKDYcUk-Z_OvcLgowFn-zIOmFHdSvZHfpPKDjkLIjUSFHMwA-h_R6QvT-1p2iaexDUbwQYvs0Unfj_U44Cuq7a-6quI639WWDY7W0RZ7cPPyj35Odj0esuh9Zf5sMqPGklHvQ40M_c8-SqEKc6XI2c2qBN2Ohg-pvM4rbLM22KsfsGGkbYypAaIkxdXJT6ayWLveD2Kmz-dJudQO6zU3npwgC5yYABI8gahRUkyUw91RHavylmYLYmOaEC4EFl1HggRhwlvDvDrBiKhpSW70XxOThJrU-po2--q4Q"
var accessToken = "ya29.ImG0B-64gC2L43v4-pOIxHNc_hvy8WdKb6fh2AiWkbvhFxk8AiGOAtmu3GyrYNNzoPvqNHkQTeJTc0fKdMpqa1HWLTRD_sz0fmw10uhVXAB_edU4f-inxII5pIzFEp7wSwtN"

async function getDatasets() {

    console.log('getting datasets')
    //gets all available datasets from the hub
    const response = await fetch('http://localhost:8080/user/getDatasets',
        {
            method: 'POST',
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                "accessToken": accessToken,
                "token": idToken
            })
        });
    const myJson = await response.json();
    console.log(JSON.stringify(myJson));

    /*
    const response = await fetch(apiURL+'/alive', {
        method: 'POST',
        mode: 'no-cors',
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json'
        }
    });

    const myJson = await response.json();
    console.log(JSON.stringify(myJson));
*/
}

async function getDataset() {

    var datasetId = 'U8VVNwv9UCGqGwlgwmgl'
    console.log('getting dataset ')

    //gets all available datasets from the hub
    const response = await fetch('http://localhost:8080/user/downloadFile',
        {
            method: 'POST',
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                "accessToken": accessToken,
                "token": idToken,
                "dataSetId": datasetId
            })
        });
    const myJson = await response.json();
    console.log(JSON.stringify(myJson));


    //the rows are in myJson.data
    Excel.run(async function (context) {
        // TODO1: Queue table creation logic here.
        var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        dataTable = currentWorksheet.tables.add("A1:G1", false /*hasHeaders*/);
        dataTable.name = "Hut34Data";

        /*
        dataTable.getHeaderRowRange().values =
            [["1", "2", "3", "4"]];
        */

        dataTable.rows.add(null /*add at the end*/, myJson.data);

        return context.sync();
    })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    return
}

function createTable() {


  Excel.run(async function (context) {
    // TODO1: Queue table creation logic here.
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    // TODO2: Queue commands to populate the table with data.


    expensesTable.getHeaderRowRange().values =
        [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
      ["1/1/2017", "The Phone Company-doodlebee", "Communications", "120"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
      ["1/11/2017", "Bellows College", "Education", "350.1"],
      ["1/15/2017", "Trey Research", "Other", "135"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    ]);


    //add a timestmap from the backend
/*

      const response = await fetch('http://localhost:8080/alive');
      const myJson = await response.json();
      console.log(JSON.stringify(myJson));
*/

      // TODO3: Queue commands to format the table.
      /*expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
      expensesTable.getRange().format.autofitColumns();
      expensesTable.getRange().format.autofitRows();*/


    return context.sync();
  })
      .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
      });
}

//tim's test function...


function downloadDataset() {

    let request = require("request");

    Excel.run(function (context) {

    //1. retrieve dataset from dataHut
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


  })
      .catch(function (error) {
        console.log("Error Tom: " + error);
        if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
      });
}

function webRequest(url, data) {
    return new Promise(function (resolve, reject) {
        fetch(url)
            .then(function (response){
                    return response.json();
                }
            )
            .then(function (json) {
                resolve(JSON.stringify(json.names));
            })
    })
}