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
    document.getElementById("getDatasets").onclick = getDatasetsToDownload;
      document.getElementById("getDatasetsToPurchase").onclick = getDatasetsToPurchase;
    document.getElementById("getDataset").onclick = getDataset;
      document.getElementById("deleteDataset").onclick = deleteDataset;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});


var expensesTable
var dataTable

var apiURL = "http://localhost:8080"

var accessToken = "ya29.a0Adw1xeWJyfcoqPpQeS_RTyEeXeWusirBxPb5CeE-OPnKvExrvSaGWAZUqunPAmARPo6x50X9U7yhbzOujT5I54mY81slm31NaqgISUuZjx7666xvliISNnaPhG5q-JbM517AeSy5ld2OniJ0yksmB1ugwXYW3SwjbL7K"
var idToken = "eyJhbGciOiJSUzI1NiIsImtpZCI6IjgyZTZiMWM5MjFmYTg2NzcwZjNkNTBjMTJjMTVkNmVhY2E4ZjBkMzUiLCJ0eXAiOiJKV1QifQ.eyJuYW1lIjoiUGV0ZXIgR29kYm9sdCIsInBpY3R1cmUiOiJodHRwczovL2xoNC5nb29nbGV1c2VyY29udGVudC5jb20vLWhUVXhvbUJzQ3lrL0FBQUFBQUFBQUFJL0FBQUFBQUFBQUFBL0FDSGkzcmN4b211dEM1NDNMNkpBWTROOFhNaVpsbHkwRUEvcGhvdG8uanBnIiwiYWRtaW4iOmZhbHNlLCJvd25lciI6ZmFsc2UsImFwcHJvdmVkVXNlciI6dHJ1ZSwiaXNzIjoiaHR0cHM6Ly9zZWN1cmV0b2tlbi5nb29nbGUuY29tL3NwZWVkZ2FzZGF0YWh1dCIsImF1ZCI6InNwZWVkZ2FzZGF0YWh1dCIsImF1dGhfdGltZSI6MTU4NDA3NDAyMCwidXNlcl9pZCI6IlFLU3RjS3VTUTJOd3BObzBxcGdsNU1VazFFSjMiLCJzdWIiOiJRS1N0Y0t1U1EyTndwTm8wcXBnbDVNVWsxRUozIiwiaWF0IjoxNTg0MDc0MDIxLCJleHAiOjE1ODQwNzc2MjEsImVtYWlsIjoicGV0ZXJnb2Rib2x0QGdtYWlsLmNvbSIsImVtYWlsX3ZlcmlmaWVkIjp0cnVlLCJmaXJlYmFzZSI6eyJpZGVudGl0aWVzIjp7Imdvb2dsZS5jb20iOlsiMTE1NTExNTkwNzE1MjQ5OTkyOTUyIl0sImVtYWlsIjpbInBldGVyZ29kYm9sdEBnbWFpbC5jb20iXX0sInNpZ25faW5fcHJvdmlkZXIiOiJnb29nbGUuY29tIn19.TW8dqOkRExBwJHmGVtD09if1hAI7qZd1_Dwg-Ik3DxxkUYulZs4cn0R-MVFFSmljkCpexRNtSjzSQ82dTXoTk3N4QPjaKHDbVPhGTmtHDHshhWo3BnH3w6eneBqYowxQsC93v5NXJFF6xpe5LAgDEK6Q6MvP4m0AyR9uPpHjd1_4MfjvcfTLibYOTlB9aBib4BrYr6QIAiqE9SYq6icRUR1q7iu6kS1eILIYKKQMQsf0s3TU91_VzTdqQkLlAtJmdTCm7w5ZFcpEDznqNpSuFf6SASMtnScjCOMQ12uvnr_x233Cx-x5725XEtHajCE2j1KCF6YOOqKiZZfb5cPhjA"

function deleteDataset() {
    console.log('deleting')

    Excel.run(async function (context) {
        dataTable.delete();
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

async function getDatasetsToDownload() {

    console.log('getting list of datasets available immediately to download')

    console.log('getting datasets')
    //gets all available datasets from the hub
    const response = await fetch('http://localhost:8080/user/getDatasetsToDownload',
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

    return

}

async function getDatasetsToPurchase() {

    console.log('getting list of datasets to purchase')
    //gets all available datasets from the hub
    const response = await fetch('http://localhost:8080/user/getDatasetsToPurchase',
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

    return

}

async function getDatasets() {

    console.log('getting datasets')
    //gets all available datasets from the hub
    const response = await fetch('http://localhost:8080/user/getDatasetsToDownload',
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

    var datasetId = '1d1o9ShF9pyACrugD6dH'
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
    //console.log(JSON.stringify(myJson));
    console.log('got the data')
    //console.log(myJson.header.length)

    let letter = myJson.header.length + 64

    //the rows are in myJson.data
    Excel.run(async function (context) {
        // TODO1: Queue table creation logic here.
        var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        dataTable = currentWorksheet.tables.add("A1:"+String.fromCharCode(letter)+"1", true /*hasHeaders*/);
        dataTable.name = "Hut34Data";

        let i = 0
        let headers = []
        while (i < myJson.header.length) {
            headers.push(myJson.header[i].name)
            i+=1
        }

        dataTable.getHeaderRowRange().values = [ headers ]
        dataTable.rows.add(null /*add at the end*/, myJson.data);

        dataTable.getRange().format.autofitColumns();
        dataTable.getRange().format.autofitRows();

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