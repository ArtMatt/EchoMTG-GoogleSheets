function onOpen() {
  var menu = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.insertSheet("CurrentCollection")
    ss.insertSheet("Credentials")
  } catch (e) { Logger.log(e) } 
  
  menu.createMenu("EchoMTG")
  .addItem("List Account Collection", 'echoList')
  .addToUi();
  Browser.msgBox('Welcome! Please input email address and password to use the sheet');
}

function auth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var info = ss.getSheetByName("Credentials")
  try {
    info.getRange(1,1).setValue("Email Address")
    info.getRange(2,1).setValue("Password")
    var username = info.getRange(1,2).getValue();
    var passwd = info.getRange(2,2).getValue();
    var creds = "email="+username+"&password="+passwd
  } catch (e) { 
    ss.insertSheet("Credentials")
    var info = ss.getSheetByName("Credentials")
    info.getRange(1,1).setValue("Email Address:")
    info.getRange(2,1).setValue("Password:")
    Browser.msgBox('Please Enter User Information in Credentials Sheet')
  }
  
  var options = {
    'muteHttpExceptions': false, 
    "headers": {"Content-Type": "application/json"}, 
    "method": "POST"
  }
  try {
    var echoRequest = UrlFetchApp.fetch("https://www.echomtg.com/api/user/auth/?" + creds, options);
  } catch(e) { Browser.msgBox(e) }
  var authToken = JSON.parse(echoRequest).token;
  return authToken
}

function echoList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   try {
    ss.insertSheet("CurrentCollection")
  } catch (e) { Logger.log(e) } 
  var collection = ss.getSheetByName("CurrentCollection")
  collection.getRange(1,1).setValue("Name");
  collection.getRange(1,2).setValue("Foil");
  collection.getRange(1,3).setValue("Set Name");
  collection.getRange(1,4).setValue("Set Code");
  collection.getRange(1,5).setValue("Rarity");
  collection.getRange(1,6).setValue("Current Price");
  collection.getRange(1,7).setValue("TCG Mid");
  collection.getRange(1,8).setValue("TCG Low");
  collection.getRange(1,9).setValue("Price Acquired");
  collection.getRange(1,10).setValue("Inventory ID");
  collection.getRange(2,2).setValue("=sum(B3:B)");
  collection.getRange(2,6).setValue("=sum(F3:F)");
  collection.getRange(2,7).setValue("=sum(G3:G)");
  collection.getRange(2,8).setValue("=sum(H3:H)");
  collection.getRange(2,9).setValue("=sum(I3:I)");
  collection.setFrozenRows(2);
  var authToken = auth();
  var echoRequest = UrlFetchApp.fetch("https://www.echomtg.com/api/inventory/view/start=0&limit=1000000&auth=" + authToken, {'muteHttpExceptions': false});
  var inventory = JSON.parse(echoRequest).items;
  var currentSize = collection.getMaxRows()
  if (currentSize < inventory.length) {
    collection.insertRowsAfter(currentSize, inventory.length - currentSize)
  }
  Logger.log(currentSize);
  try { 
    for (c=0;c<inventory.length;c++){
      collection.getRange(c+3,1).setValue(inventory[c].name);
      collection.getRange(c+3,2).setValue(inventory[c].foil);
      collection.getRange(c+3,3).setValue(inventory[c].set);
      collection.getRange(c+3,4).setValue(inventory[c].set_code);
      collection.getRange(c+3,5).setValue(inventory[c].rarity);
      collection.getRange(c+3,6).setValue(inventory[c].current_price);
      collection.getRange(c+3,7).setValue(inventory[c].tcg_mid);
      collection.getRange(c+3,8).setValue(inventory[c].tcg_low);
      collection.getRange(c+3,9).setValue(inventory[c].price_acquired);
      collection.getRange(c+3,10).setValue(inventory[c].inventory_id);
    }
  } 
  catch (e) { 
    Browser.msgBox("Maybe retry your credentials? Error: " + e);
  } 
  collection.getRange(2,3).setValue('="Sets: "&counta(unique(C3:C))');
  collection.getRange(2,1).setValue('="Total: "&counta(A3:A)');
}
