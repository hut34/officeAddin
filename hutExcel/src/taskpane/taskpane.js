/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
 // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
      console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

// Assign event handlers and other initialization logic.

    document.getElementById("getDatasets").onclick = getDatasetsToDownload;
    document.getElementById("getDatasetButton").onclick = getDataset;
    document.getElementById("uploadDataset").onclick = uploadDataset;

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

var dataTable, datasetId, accessToken, idToken, apiURL

async function getDatasetsToDownload() {

    accessToken = document.getElementById("accessToken").value;
    idToken = document.getElementById("idToken").value;
    apiURL =  document.getElementById("apiURL").value;

    console.log('getting list of datasets available immediately to download')
    //gets all available datasets from the hub
    const response = await fetch(apiURL+'/user/getDatasetsToDownload',
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
    //console.log(JSON.stringify(myJson));

    var ul = document.getElementById("theList");

    myJson.forEach(function(dataset) {
        console.log(dataset.data.name)
        var li = document.createElement("li");
        li.className = "ms-ListItem is-unread"
        li.innerHTML = '' +
            '                <span class="ms-ListItem-primaryText" id="datasetName">'+ dataset.data.name +'</span>\n' +
            '                <!-- <span class="ms-ListItem-secondaryText">Meeting notes</span> -->\n' +
            '                <span class="ms-ListItem-tertiaryText">'+ dataset.id+'</span>\n' +
            '                <span class="ms-ListItem-metaText" id="time">'+ dataset.data.ENTRPPrice +'</span>\n' +
            '                <div class="ms-ListItem-selectionTarget"></div>\n' +
            '                <div class="ms-ListItem-actions">\n' +
            '                    <div class="ms-ListItem-action">\n' +
            '                        <i class="ms-Icon ms-Icon--Pinned"></i>\n' +
            '                    </div>\n' +
            '                </div>\n'
        ul.appendChild(li);
    })

    document.getElementById("listOfDatasets").style = "display:block;";

    //add to the list



    return

}

async function getDataset() {

    accessToken = document.getElementById("accessToken").value;
    idToken = document.getElementById("idToken").value;
    apiURL =  document.getElementById("apiURL").value;

    datasetId = document.getElementById("inputDatasetId").value;

    console.log('getting dataset ')

    //gets all available datasets from the hub
    const response = await fetch(apiURL+'/user/downloadFile',
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

async function uploadDataset() {

    let hutHeaders = []

    console.log('uploading')
    Excel.run(function (context) {

        var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        var uploadTable = currentWorksheet.tables.getItem("Hut34Data");

        // Get data from the header row
        var headerRange = uploadTable.getHeaderRowRange().load("values");

        // Get data from the table
        var bodyRange = uploadTable.getDataBodyRange().load("values");

        // Sync to populate proxy objects with data from Excel
        return context.sync()
            .then(function () {

                var headerValues = headerRange.values;
                var bodyValues = bodyRange.values;

                //console.log('we have the table, which is an update of '+datasetId)
                //console.log(headerValues)
                //console.log(bodyValues)

                //prep the data structure, and send it over to the hut;


                headerValues[0].forEach(function(header) {

                    let headObject = {}
                    headObject.name = header
                    headObject.description = header
                    headObject.type = "string"

                    hutHeaders.push(headObject)

                })

                sendToTheHut(hutHeaders, bodyValues)
            });
    }).catch();


}

function createTable() {


  Excel.run(async function (context) {

    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

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

async function sendToTheHut(cols, rows) {

    let data = {}


    data.name = "Dataset created from Excel, derived from "+datasetId
    data.header = cols
    data.data = rows
    data.coverImage = "https://images.unsplash.com/photo-1529078155058-5d716f45d604?ixlib=rb-1.2.1&ixid=eyJhcHBfaWQiOjEyMDd9&auto=format&fit=crop&w=1349&q=80"


    const response = await fetch(apiURL+'/user/createDataset',
        {
            method: 'POST',
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                "accessToken": accessToken,
                "token": idToken,
                "data": data
            })
        });
    const myJson = await response.json();
    console.log(myJson)
    return myJson
}