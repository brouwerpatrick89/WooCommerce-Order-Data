function syncProductData() {

    var sheet_name = "ProductData"//SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    //Logger.log(sheet_name)
    fetch_products(sheet_name)

}


function fetch_products(sheet_name) {

    var ck = "YOUR_API_KEY";

    var cs = "YOUR_API_SECRET";

    var website = "https://YOURWEBSITE.com";
  
    var manualDate = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getLastRow(); // Set your order start date in spreadsheet in cell B6

    var range = "B" + manualDate;
    
    var m = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange(range).getValue();

    var surl = website + "/wp-json/wc/v3/products?consumer_key=" + ck + "&consumer_secret=" + cs + "&page=1&per_page=100&order=asc&after=" + m; 

    var url = surl
    Logger.log(url)

    var options =

        {
            "method": "GET",
            "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
            "muteHttpExceptions": true,

        };

    var result = UrlFetchApp.fetch(url, options);

    Logger.log(result.getResponseCode())
    if (result.getResponseCode() == 200) {

        var params = JSON.parse(result.getContentText());
      //Logger.log(params);

    }

    var doc = SpreadsheetApp.getActiveSpreadsheet();

    var temp = doc.getSheetByName(sheet_name);

    var consumption = {};

    var arrayLength = params.length;
  
    for (var i = 0; i < arrayLength; i++) {
        var a, c, d;
        var container = [];
        //ProductID
        a = container.push(params[i]["id"]);
        //Date Created
        a = container.push(params[i]["date_created"]);
        //Date Modified
        a = container.push(params[i]["date_modified"]);
        //Name
        a = container.push(params[i]["name"]);
        //SKU
        a = container.push(params[i]["sku"]);
        //Price
        a = container.push(params[i]["price"]);
        //Starting Date
        var sku = params[i]["sku"];
        if (sku.endsWith("EB")){
          var day = sku.substr(-8,2);
          var month = sku.substr(-6,2);
          var year = "20" + sku.substr(-4,2);
          var mdate = month + "/" + day + "/" + year;
        } else if (sku.endsWith("-1")){
          var day = sku.substr(-8,2);
          var month = sku.substr(-6,2);
          var year = "20" + sku.substr(-4,2);
          var mdate = month + "/" + day + "/" + year;
        } else {
          var day = sku.substr(-6,2);
          var month = sku.substr(-4,2);
          var year = "20" + sku.substr(-2,2);
          var mdate = month + "/" + day + "/" + year;
        }
        a = container.push(mdate);

        var doc = SpreadsheetApp.getActiveSpreadsheet();

        var temp = doc.getSheetByName(sheet_name);

        temp.appendRow(container);
     
        Logger.log(params[i]);

        removeDuplicateProducts(sheet_name);
    }
}

function removeDuplicateProducts(sheet_name) {

    var doc = SpreadsheetApp.getActiveSpreadsheet();

    var sheet = doc.getSheetByName(sheet_name);

    var data = sheet.getDataRange().getValues();

    var newData = new Array();

    for (i in data) {

        var row = data[i];
      /*  TODO feature enhancement in de-duplication
        var date_modified =row[row.length-2];
      
        var order_key = row[row.length];
      
        var existingDataSearchParam = order_key + "/" + date_modified; 
       */

        var duplicate = false;

        for (j in newData) {
          
          var rowNewData = newData[j];
          
          var new_date_modified =rowNewData[rowNewData.length-2];
          
          var new_order_key = rowNewData[rowNewData.length];
          
          //var newDataSearchParam = new_order_key + "/" + new_date_modified; // TODO feature enhancement in de-duplication

          if(row.join() == newData[j].join()) {
                duplicate = true;

            }
          
          // TODO feature enhancement in de-duplication
          /*if (existingDataSearchParam == newDataSearchParam){
            duplicate = true;
          }*/

        }
        if (!duplicate) {
            newData.push(row);
        }
    }
    sheet.clearContents();
    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}
