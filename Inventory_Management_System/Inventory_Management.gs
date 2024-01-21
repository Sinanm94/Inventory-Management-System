function myReceiptView() {
  
  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE MENU SHEET          
  var menuSheet = ss.getSheetByName("Menu");
  var partSheet = ss.getSheetByName("Parts");
  var locationSheet = ss.getSheetByName("Locations");
  
  //CLEAR DATA
  menuSheet.getRange("A3:B8").clear();
  menuSheet.getRange("A10:C5000").clear();
  menuSheet.getRange("A3:B8").setDataValidation(null);
  
  //TITLES
  menuSheet.getRange(3,1).setValue(["Transaction Type"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(3,2).setValue(["Receipt"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(4,1).setValue(["Part Number"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(5,1).setValue(["Location"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(6,1).setValue(["Quantity"]).setFontSize(12).setFontWeight("bold");
  
  //PART DROPDOWN
  var partCell = menuSheet.getRange('B4'); 
  var partLastRow = partSheet.getLastRow();
  var partRange = partSheet.getRange('A2:A' + partLastRow);
  var partRule = SpreadsheetApp.newDataValidation().requireValueInRange(partRange).build();
  partCell.setDataValidation(partRule);
  
  //LOCATION DROPDOWN
  var locationCell = menuSheet.getRange('B5'); 
  var locationLastRow = locationSheet.getLastRow();
  var locationRange = locationSheet.getRange('A2:A' + locationLastRow);
  var locationRule = SpreadsheetApp.newDataValidation().requireValueInRange(locationRange).build();
  locationCell.setDataValidation(locationRule);
  
  //QUANTITY
  var quantityCell = menuSheet.getRange('B6');
  var quantityRule = SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(1).setAllowInvalid(false).build();
  quantityCell.setDataValidation(quantityRule);

}

function myShipmentView() {
  
  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE MENU SHEET          
  var menuSheet = ss.getSheetByName("Menu");
  var partSheet = ss.getSheetByName("Parts");
  var locationSheet = ss.getSheetByName("Locations");
  
  //CLEAR DATA
  menuSheet.getRange("A3:B8").clear();
  menuSheet.getRange("A10:C5000").clear();
  menuSheet.getRange("A3:B8").setDataValidation(null);
  
  //TITLES
  menuSheet.getRange(3,1).setValue(["Transaction Type"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(3,2).setValue(["Shipment"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(4,1).setValue(["Part Number"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(5,1).setValue(["Location"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(6,1).setValue(["Quantity"]).setFontSize(12).setFontWeight("bold");
  
  //PART DROPDOWN
  var partCell = menuSheet.getRange('B4'); 
  var partLastRow = partSheet.getLastRow();
  var partRange = partSheet.getRange('A2:A' + partLastRow);
  var partRule = SpreadsheetApp.newDataValidation().requireValueInRange(partRange).build();
  partCell.setDataValidation(partRule);
  
  //LOCATION DROPDOWN
  var locationCell = menuSheet.getRange('B5'); 
  var locationLastRow = locationSheet.getLastRow();
  var locationRange = locationSheet.getRange('A2:A' + locationLastRow);
  var locationRule = SpreadsheetApp.newDataValidation().requireValueInRange(locationRange).build();
  locationCell.setDataValidation(locationRule);
  
  //QUANTITY
  var quantityCell = menuSheet.getRange('B6');
  var quantityRule = SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(1).setAllowInvalid(false).build();
  quantityCell.setDataValidation(quantityRule);
  

}

function myTransferView() {
  
  //Browser.msgBox('Shipment');
  
  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE MENU SHEET          
  var menuSheet = ss.getSheetByName("Menu");
  var partSheet = ss.getSheetByName("Parts");
  var locationSheet = ss.getSheetByName("Locations");
  
  //CLEAR DATA
  menuSheet.getRange("A3:B8").clear();
  menuSheet.getRange("A10:C5000").clear();
  menuSheet.getRange("A3:B8").setDataValidation(null);
  
  //TITLES
  menuSheet.getRange(3,1).setValue(["Transaction Type"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(3,2).setValue(["Transfer"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(4,1).setValue(["Part Number"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(5,1).setValue(["From Location"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(6,1).setValue(["To Location"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(7,1).setValue(["Quantity"]).setFontSize(12).setFontWeight("bold");
  
  //PART DROPDOWN
  var partCell = menuSheet.getRange('B4'); 
  var partLastRow = partSheet.getLastRow();
  var partRange = partSheet.getRange('A2:A' + partLastRow);
  var partRule = SpreadsheetApp.newDataValidation().requireValueInRange(partRange).build();
  partCell.setDataValidation(partRule);
  
  //FROM LOCATION DROPDOWN
  var locationCell = menuSheet.getRange('B5'); 
  var locationLastRow = locationSheet.getLastRow();
  var locationRange = locationSheet.getRange('A2:A' + locationLastRow);
  var locationRule = SpreadsheetApp.newDataValidation().requireValueInRange(locationRange).build();
  locationCell.setDataValidation(locationRule);
  
  //TO LOCATION DROPDOWN
  var locationToCell = menuSheet.getRange('B6');
  locationToCell.setDataValidation(locationRule);
  
  //QUANTITY
  var quantityCell = menuSheet.getRange('B7');
  var quantityRule = SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(1).setAllowInvalid(false).build();
  quantityCell.setDataValidation(quantityRule);

}

function myInvAdjView() {
  
  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE MENU SHEET          
  var menuSheet = ss.getSheetByName("Menu");
  var partSheet = ss.getSheetByName("Parts");
  var locationSheet = ss.getSheetByName("Locations");
  
  //CLEAR DATA
  menuSheet.getRange("A3:B8").clear();
  menuSheet.getRange("A10:C5000").clear();
  menuSheet.getRange("A3:B8").setDataValidation(null);
  
  //TITLES
  menuSheet.getRange(3,1).setValue(["Transaction Type"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(3,2).setValue(["Inventory Adjustment"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(4,1).setValue(["Part Number"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(5,1).setValue(["Location"]).setFontSize(12).setFontWeight("bold");
  menuSheet.getRange(6,1).setValue(["New Quantity"]).setFontSize(12).setFontWeight("bold");
  
  //PART DROPDOWN
  var partCell = menuSheet.getRange('B4'); 
  var partLastRow = partSheet.getLastRow();
  var partRange = partSheet.getRange('A2:A' + partLastRow);
  var partRule = SpreadsheetApp.newDataValidation().requireValueInRange(partRange).build();
  partCell.setDataValidation(partRule);
  
  //LOCATION DROPDOWN
  var locationCell = menuSheet.getRange('B5'); 
  var locationLastRow = locationSheet.getLastRow();
  var locationRange = locationSheet.getRange('A2:A' + locationLastRow);
  var locationRule = SpreadsheetApp.newDataValidation().requireValueInRange(locationRange).build();
  locationCell.setDataValidation(locationRule);
  
  //QUANTITY
  var quantityCell = menuSheet.getRange('B6');
  var quantityRule = SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build();
  quantityCell.setDataValidation(quantityRule);
 

}

function submitResults() {
  
  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE MENU SHEET          
  var menuSheet = ss.getSheetByName("Menu");
  var inventorySheet = ss.getSheetByName("Inventory");
  
  //TRANSACTION
  var transaction = menuSheet.getRange(3,2).getValue();
  
  if(transaction == "Receipt")
  {
    //GET VALUES
    var partNumber = menuSheet.getRange(4,2).getValue();
    var location = menuSheet.getRange(5,2).getValue();
    var quantity = menuSheet.getRange(6,2).getValue();
    
    if(partNumber != '' && location != '' && quantity != '')
    { 
      //LAST ROW ON INVENTORY SHEET
      var lastRow = inventorySheet.getLastRow() + 1;
      
      var foundRecord = false;
      for (var j = 2; j < lastRow; j++)
      {
        // UPDATE EXISTING QUANTITY AT LOCATION
        if(inventorySheet.getRange(j,1).getValue() == partNumber  && inventorySheet.getRange(j,2).getValue() == location)
        {
          var newQuantity = Number(inventorySheet.getRange(j,3).getValue()) + quantity;
          inventorySheet.getRange(j,3).setValue([newQuantity]);
          foundRecord = true;
        }
        
      }
      
      if(foundRecord == false)
      {
        inventorySheet.getRange(lastRow,1).setValue([partNumber]);
        inventorySheet.getRange(lastRow,2).setValue([location]);
        inventorySheet.getRange(lastRow,3).setValue([quantity]);
        
      }  
       
    }
    else
    {
      Browser.msgBox('Enter Required Data');
      return;
    }
    
    Browser.msgBox('Complete');
    myReceiptView();
  }
  
  if(transaction == "Shipment")
  {
    //GET VALUES
    var partNumber = menuSheet.getRange(4,2).getValue();
    var location = menuSheet.getRange(5,2).getValue();
    var quantity = menuSheet.getRange(6,2).getValue();
    
    if(partNumber != '' && location != '' && quantity != '')
    { 
      //LAST ROW ON INVENTORY SHEET
      var lastRow = inventorySheet.getLastRow() + 1;
      
      var foundRecord = false;
      for (var j = 2; j < lastRow; j++)
      {
        // UPDATE EXISTING QUANTITY AT LOCATION
        if(inventorySheet.getRange(j,1).getValue() == partNumber  && inventorySheet.getRange(j,2).getValue() == location)
        {
          var newQuantity = Number(inventorySheet.getRange(j,3).getValue()) - quantity;
          
          if(newQuantity < 0)
          {
            Browser.msgBox('Not enough Quantity from location.  Transaction Denied');
            return;
          }
          else if(newQuantity == 0)
          {
            inventorySheet.deleteRow(j);
          }
          else
          {
            inventorySheet.getRange(j,3).setValue([newQuantity]);
          }
          
          foundRecord = true;

        }
        
      }
      
      if(foundRecord == false)
      {
        Browser.msgBox('No quantity to Ship from Location.  Transaction Denied');
        return;
      }
       
    }
    else
    {
      Browser.msgBox('Enter Required Data');
      return;
    }
    
    Browser.msgBox('Complete');
    myShipmentView();
  }
  
  if(transaction == "Transfer")
  {
    //GET VALUES
    var partNumber = menuSheet.getRange(4,2).getValue();
    var fromLocation = menuSheet.getRange(5,2).getValue();
    var toLocation = menuSheet.getRange(6,2).getValue();
    var quantity = menuSheet.getRange(7,2).getValue();
    
    if(partNumber != '' && toLocation != '' && fromLocation != '' && quantity != '')
    { 
      //LAST ROW ON INVENTORY SHEET
      var lastRow = inventorySheet.getLastRow() + 1;
      
      var foundRecord = false;
      for (var j = 2; j < lastRow; j++)
      {
        // UPDATE EXISTING QUANTITY AT FROM LOCATION
        if(inventorySheet.getRange(j,1).getValue() == partNumber  && inventorySheet.getRange(j,2).getValue() == fromLocation)
        {
          var newQuantity = Number(inventorySheet.getRange(j,3).getValue()) - quantity;
          
          if(newQuantity < 0)
          {
            Browser.msgBox('Not enough Quantity from location.  Transaction Denied');
            return;
          }
          else if(newQuantity == 0)
          {
            inventorySheet.deleteRow(j);
          }
          else
          {
            inventorySheet.getRange(j,3).setValue([newQuantity]);
          }
          foundRecord = true;
          
        }
        
      }
      
      if(foundRecord == false)
      {
        Browser.msgBox('No quantity to Transfer from Location.  Transaction Denied');
        return;
      }
      
      var foundRecord = false;
      for (var j = 2; j < lastRow; j++)
      {
        // UPDATE EXISTING QUANTITY AT TO LOCATION
        if(inventorySheet.getRange(j,1).getValue() == partNumber  && inventorySheet.getRange(j,2).getValue() == toLocation)
        {
          var newQuantity = Number(inventorySheet.getRange(j,3).getValue()) + quantity;         
          inventorySheet.getRange(j,3).setValue([newQuantity]);
          foundRecord = true;
        }
        
      }
      
      if(foundRecord == false)
      {
        inventorySheet.getRange(lastRow,1).setValue([partNumber]);
        inventorySheet.getRange(lastRow,2).setValue([toLocation]);
        inventorySheet.getRange(lastRow,3).setValue([quantity]);
        
      } 
       
    }
    else
    {
      Browser.msgBox('Enter Required Data');
      return;
    }
    
    Browser.msgBox('Complete');
    myTransferView();
  }
  
  if(transaction == "Inventory Adjustment")
  {
    //GET VALUES
    var partNumber = menuSheet.getRange(4,2).getValue();
    var Location = menuSheet.getRange(5,2).getValue();
    var quantity = menuSheet.getRange(6,2).getValue();
    
    if(partNumber != ''&& location != '' && quantity !== '')
    { 
      //LAST ROW ON INVENTORY SHEET
      var lastRow = inventorySheet.getLastRow() + 1;
      
      var foundRecord = false;
      for (var j = 2; j < lastRow; j++)
      {
        // UPDATE EXISTING QUANTITY AT FROM LOCATION
        if(inventorySheet.getRange(j,1).getValue() == partNumber  && inventorySheet.getRange(j,2).getValue() == Location)
        {          
          if(quantity < 0)
          {
            Browser.msgBox('Not enough Quantity must be greater than 0.  Transaction Denied');
            return;
          }
          else if(quantity == 0)
          {
            inventorySheet.deleteRow(j);
          }
          else
          {
            inventorySheet.getRange(j,3).setValue([quantity]);
          }
          foundRecord = true;
        }
        
      }
      if(foundRecord == false)
      {
        Browser.msgBox('No Location Found.  Transaction Denied');
        return;
      }
      
    }
    else
    {
      Browser.msgBox('Enter Required Data');
      return;
    }
    
    Browser.msgBox('Complete');
    myInvAdjView();
  }
  
}

function onEdit(e)
{
  var range = e.range;
  var spreadSheet = e.source;
  var spreadSheetName = spreadSheet.getActiveSheet().getName();
  var searchColumn = range.getColumn();
  var searchRow = range.getRow();
  
  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE MENU SHEET          
  var menuSheet = ss.getSheetByName("Menu");
  var inventorySheet = ss.getSheetByName("Inventory");
  
  Logger.log('Column: ' + searchColumn + ' Row: ' + searchRow + ' Value: ' + e.value + ' spreadSheetName: ' + spreadSheetName);
  
  if (searchColumn == 2 && searchRow == 4 && e.value != '' && spreadSheetName == 'Menu')
  {
    var partNumber = e.value;
    
    //LAST ROW ON INVENTORY SHEET
    var lastRow = inventorySheet.getLastRow() + 1;
    var foundRecord = false;
    
    for (var j = 2; j < lastRow; j++)
    {
      // UPDATE EXISTING QUANTITY AT FROM LOCATION
      if(inventorySheet.getRange(j,1).getValue() == partNumber)
      {  
        var nextRow = menuSheet.getLastRow() + 1;
        
        menuSheet.getRange(nextRow,1).setValue([inventorySheet.getRange(j,1).getValue()]).setFontSize(12).setFontWeight("bold");
        menuSheet.getRange(nextRow,2).setValue([inventorySheet.getRange(j,2).getValue()]).setFontSize(12).setFontWeight("bold");
        menuSheet.getRange(nextRow,3).setValue([inventorySheet.getRange(j,3).getValue()]).setFontSize(12).setFontWeight("bold");
        foundRecord = true;
      }
      
    }
    
    if(foundRecord == false)
    {
      menuSheet.getRange(10,1).setValue(['(NO RECORDS FOUND)']).setFontSize(12).setFontWeight("bold");
    }
    
    
  }
  
  
}