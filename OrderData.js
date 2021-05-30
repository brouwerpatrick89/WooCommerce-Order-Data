function syncOrderData() {

    var sheet_name = "OrderData"//SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    //Logger.log(sheet_name)
    fetch_orders(sheet_name)

}


function fetch_orders(sheet_name) {

    var ck = "YOUR_API_KEY";

    var cs = "YOUR_API_SECRET";

    var website = "https://YOURWEBSITE.com";
  
    var manualDate = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getLastRow(); // Set your order start date in spreadsheet in cell B6

    var range = "B" + manualDate;
    
    var m = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange(range).getValue();

    var surl = website + "/wp-json/wc/v3/orders?consumer_key=" + ck + "&consumer_secret=" + cs + "&page=1&per_page=100&order=asc&after=" + m; 

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
        //OrderID
        a = container.push(params[i]["id"]);
        //Date Created
        a = container.push(params[i]["date_created"]);
        //Date Modified
        a = container.push(params[i]["date_modified"]);
        //Status
        a = container.push(params[i]["status"]);
        //First Name
        a = container.push(params[i]["billing"]["first_name"]);
        //Last Name
        a = container.push(params[i]["billing"]["last_name"]);
        //Company
        a = container.push(params[i]["billing"]["company"]);
        //Address_1
        a = container.push(params[i]["billing"]["address_1"]);
        //Address_2
        a = container.push(params[i]["billing"]["address_2"]);
        //City
        a = container.push(params[i]["billing"]["city"]);
        //State
        a = container.push(params[i]["billing"]["state"]);
        //Postcode
        a = container.push(params[i]["billing"]["postcode"]);
        //Country
        a = container.push(params[i]["billing"]["country"]);
        //Email
        a = container.push(params[i]["billing"]["email"]);
        //Phone
        a = container.push(params[i]["billing"]["phone"]);
        //Products, Quantity & SKU Loop
        c = params[i]["line_items"].length;

        var items = "";
        var skus = "";
        var total_line_items_quantity = 0;
        for (var k = 0; k < c; k++) {
          var item, item_f, qty, sku, meta;

            item = params[i]["line_items"][k]["name"];

            qty = params[i]["line_items"][k]["quantity"];
            
            sku = params[i]["line_items"][k]["sku"];

            item_f = qty + " x " + item;

            if (k == 0){
            items = items + item_f;
              if (sku == null){
                skus = "";
              } else {
                skus = skus + sku;
              }
          } else {
            items = items + ";\n" + item_f;
            if (sku == null){
              skus = "";
            } else {
              skus = skus + ";\n" + sku;
            }
          }        

            total_line_items_quantity += qty;
        }
        //Products
        a = container.push(items);
        //Quantity
        a = container.push(total_line_items_quantity);
        //SKU
        a = container.push(skus);
        //Total Price
        a = container.push(params[i]["total"]);
        //Discount
        a = container.push(params[i]["discount_total"]);
        //Refunds
        d = params[i]["refunds"].length;
      
        var refundItems = "";
      
        var refundValue = 0.0;
      
        for (var r = 0; r < d; r++) {
          var item, value;

            item = params[i]["refunds"][r]["reason"];

            value = params[i]["refunds"][r]["total"];
            
            refundValue += parseFloat(value);

            refundItems += item;

        }
      
        a = container.push(refundValue); //Refunded Value
      
        a = container.push(refundItems); //Refund Reason
        //Payment Method Title
        a = container.push(params[i]["payment_method_title"]);
        //Date Paid
        a = container.push(params[i]["date_paid"]);
        //Date Completed
        a = container.push(params[i]["date_completed"]);
        //Customer User Agent
        a = container.push(params[i]["customer_user_agent"]);
        //Customer Note
        a = container.push(params[i]["customer_note"]);
        //Order Key
        a = container.push(params[i]["order_key"]);

        var doc = SpreadsheetApp.getActiveSpreadsheet();

        var temp = doc.getSheetByName(sheet_name);

        temp.appendRow(container);
     
        Logger.log(params[i]);

        removeDuplicates(sheet_name);
    }
}

function removeDuplicates(sheet_name) {

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
