/*
    CODE.GS
*/
/** @OnlyCurrentDoc */
// Dev Ref: https://developers.google.com/apps-script/reference/spreadsheet

var loggingEnabled = true;
var editItem;

function startup() {
  try {  
    toast('Loading Vine config and menu','Vine Menu Status',3);

    var mainMenu = SpreadsheetApp.getUi().createMenu('🍃Vine');
    //mainMenu.addItem("Bulk Input","showBulkInput");
    var importMenu=SpreadsheetApp.getUi().createMenu('Import...');
    importMenu.addItem("...from Amazon Vine Itemized Report","showAVIRImport");
    importMenu.addItem("...from previous QwicksetTracking sheet","showQTImport");
    mainMenu.addSubMenu(importMenu);
    mainMenu.addItem("Current Item Review Form (Ctrl+Alt+Shift+0)","showReviewForm");
    mainMenu.addSeparator();
    mainMenu.addItem("Future Features","showFutureFeatures");
    mainMenu.addSeparator();
    mainMenu.addItem("About","showAbout");
    mainMenu.addToUi();
    menuLoaded=true;

    toast('Vine config and menu loaded. Ready for takeoff','Vine Menu Status',3);
  } catch (err){
    var msg=`Error encountered.\n\n${err.message}\n\nStackTrace:${err.stack}`;
    console.log(msg);
    var ui = SpreadsheetApp.getUi();
    this.alert('Hardstop Error',msg,ui.ButtonSet.OK);      
  }
}
function flushAll(){
  SpreadsheetApp.flush();
}
function getBase64Image(code){
  var b64 = new Base64Images();
  var imgSrc = b64.image(code);
  return imgSrc;
}
function getItemRowData(){
  var data=[];
  var error;
  var activeSheet=SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var range=activeSheet.getActiveRange();
  var row=range.getRowIndex();
  if (row==1){
    error={
      title:'Row error',
      message:'Please ensure you are on a row with item data.'
    };
  }else if (1==2 && activeSheet.getName().toUpperCase()!=='DATA'){
    error={
      title:'Sheet/tab error',
      message:'Please switch to "Data" tab and select a populated item row'
    };
  }else{
    var endColLetter='Y';
    var a1=`A${row}:${endColLetter}${row}`;
    var keyA1=`A1:${endColLetter}1`;
    var dataSheet= getSheetByName("Data");
    var dataRange=dataSheet.getRange(a1);
    var keyRange=dataSheet.getRange(keyA1);
    var values=JSON.parse(JSON.stringify(dataRange.getValues()))[0];
    var keyValues=JSON.parse(JSON.stringify(keyRange.getValues()))[0];
    var asin;
    for(var i=0;i<values.length;i++){
      var key = keyValues[i];
      key=key.replace(/\n/g,'').replace(' ','').toUpperCase();
      var value=values[i];
      if (key==='ASIN' && value) asin=value;
      var item = {
        row:row,
        column:i+1,
        key:key,
        value:value
      };
      data.push(item);
      console.log(`     ${JSON.stringify(item)}`);
    }
    if (!asin){
      error={
        title:'Missing ASIN',
        message:'The selected row does not appear to contain an ASIN'
      };
      data=[];
    }
  }
  return JSON.stringify({
    data:data,
    row:row,
    error:error
  });
}
function updateSheetWithItem(productData){
  console.log('updateSheetWithItem called.');
  console.log(`updateSheetWithItem(${JSON.stringify(productData)})`);
  var data = productData.data;
  var row = productData.row;
  var sheet = getSheetByName("Data");
  data.forEach(prop => {
    console.log(`setCellValue('${sheet.getName()}',\nrow:${row},\ncolumn:${prop.column},\nvalue:'${prop.value}')`);
    setCellValue(sheet,row,prop.column,prop.value);
  });
 return {
  isEdit:true
 }
}
function getCurrentItem(){
  var activeSheet=SpreadsheetApp.getActiveSheet();
  var activeSheetName=activeSheet.getName();
  var cell=activeSheet.getActiveCell();
  var range=activeSheet.getActiveRange();
  var row=range.getRowIndex();
  var col=range.getColumnIndex();
  var width=range.getWidth();
  var height=range.getHeight();
  var A1not = range.getA1Notation();
  return {
      activeSheet:activeSheet,
      activeSheetName:activeSheetName,
      cell:cell,
      range:range,
      row:row,
      col:col,
      width:width,
      height:height,
      A1not:A1not
  };        
}

function regexKeyValue(key,value){
  var result;
  if (key.toLowerCase()==='baseamzurl'){
    result = value.match(/(http|https):\/\/[a-z0-9\-._~%]+/gm);
  } else if (key.toLowerCase()==='asin'){
    result = value.match(/([0-9]{10})|B0([A-Z0-9]{8})/g);    
  } else if (key.toLowerCase()==='ordernum'){
    result = value.match(/[\d]{3}-[\d]{7}-[\d]{7}|[\d]{17}/g)
  }
  if (result && result.length==1) return result[0];
  return result;
}
function alert(title,message,buttons){
  var ui = SpreadsheetApp.getUi();
  if (!buttons) buttons=ui.ButtonSet.OK;
  ui.alert(title, message, buttons);
}
function showAbout() {
  var widget = HtmlService.createTemplateFromFile("About.html").evaluate().setWidth(500).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(widget, " ");
}
function showFutureFeatures() {
  var widget = HtmlService.createHtmlOutputFromFile("FutureFeatures.html").setWidth(500).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(widget, " ");
}
//function showBulkInput() {
//  var widget = HtmlService.createHtmlOutputFromFile("BulkInput.html").setWidth(1000).setHeight(1000);
//  SpreadsheetApp.getUi().showModalDialog(widget, " ");
//}
function showAVIRImport() {
  var widget = HtmlService.createHtmlOutputFromFile("AVIR_Import.html").setWidth(1000).setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(widget, " ");
}
function showQTImport() {
  var widget = HtmlService.createHtmlOutputFromFile("QT_Import.html").setWidth(1000).setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(widget, " ");
}
function showReviewForm() {
  var widget = HtmlService.createHtmlOutputFromFile("ReviewForm.html").setWidth(1000).setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(widget, " ");
}

function setCellValue(sheet, row, column, value) {
  var cellAddress=this.a1Notation(row,column)
  console.log(`setting cellAddress ${sheet.getName()}::${cellAddress} to "${value}"`);
  var valueRange = sheet.getRange(cellAddress);
  valueRange.setValue(value);
} 
function getAllASINS(){
  var ASINS =[];
  var sheet = getSheetByName("Data");
  var asinValues =sheet.getRange("F2:F").getValues();
  if (asinValues)
  {
    console.log(`ASINValues = ${JSON.stringify(asinValues)}`);
    asinValues = asinValues.map(function(asin){ if(asin[0] && asin[0].length>0) ASINS.push(asin[0]);}) ;
    console.log(`ASINValues = ${JSON.stringify(ASINS)}`);
    return ASINS;
  }
}
function nextNewRow(){
  var sheet = getSheetByName("Data");
  return sheet.getLastRow()+1;
  /*
  var asinValues = getAllASINS();
  console.log(`ASNValues=${asinValues}`);
  var lastRow = asinValues.filter(String).length;
  console.log(`nextNewRow = ${lastRow}`);
  return lastRow+2;
  */
}

function addItemToSheet(results){
  var item=results.items[results.index];
  var row=results.row-1;
  var stopIndex = Math.min(results.index+results.batchSize,results.items.length-1);
  console.log(`addItemToSheet() from ${results.index} to ${stopIndex} [results.items.length=${results.items.length}]`);
  for(var i=results.index;i<=stopIndex;i++){
    var item=results.items[i];
    row++;
    console.log(`     #${i} ASIN:${item.ASIN} row:${row}`);
    trySet(row,1,item,"Order Number");
    trySet(row,2,item,"Order Date");
    trySet(row,3,item,"Shipped Date");
    trySet(row,4,item,"Received Date");
    trySet(row,5,item,"Cancelled Date");
    trySet(row,6,item,"ASIN");
    trySet(row,7,item,"Product Name");
    trySet(row,8,item,"Category");
    trySet(row,9,item,"Estimated Tax Value");
    trySet(row,10,item,"MSRP");
    trySet(row,11,item,"Submitted Date");
    trySet(row,12,item,"Accepted Date");
    trySet(row,13,item,"Rejected Date");
    trySet(row,14,item,"Canceled Date");
    trySet(row,15,item,"Stars");
    trySet(row,16,item,"Photos");
    trySet(row,17,item,"Video");
    trySet(row,18,item,"Title");
    trySet(row,19,item,"Detail");
    trySet(row,20,item,"Notes");
  }

  return {
    items:results.items,
    batchSize:results.batchSize,
    index:stopIndex,
    row:row+1,
    savedASINSStartIndex:results.index,
    savedASINSEndIndex:stopIndex
  };
}

function trySet(row,column,item,field){
  var sheet = getSheetByName("Data");
  var value=item[field];
  if (value) sheet.getRange(row, column).setValue(value);
}

function getNamedRange(sheet, name) {
  console.log(`getNamedRange(sheet:${sheet.getName()},'${name}')`);
  if (!sheet) return;
  var namedRanges = sheet.getNamedRanges();
  for (var i = 0; i < namedRanges.length; i++) {
      if (namedRanges[i].getName().toLowerCase() == name.toLowerCase()) {
          return namedRanges[i].getRange();
      }
  }
  console.log(`     NO MATCHES`);
};

function a1Notation(row,col,fullHeight){
  console.log(`a1Notations(${row},${col},${fullHeight})`);
  var col = `${String.fromCharCode(col+64)}`;
  var a1= `${col}${row}`;
  if (fullHeight)
      return `${a1}:${col}`;
  else
      return a1;
}


function getCellValue(sheet, row, column) {
  var cellAddress = Utils.a1Notation(row,column)
  var valueRange = sheet.getRange(cellAddress);
  var value=valueRange.getValue();
  console.log(`getCellValue(sheet:${sheet.getName()},${row},${column})=${value}`);
  return value;
}
function getSheetByName(sheetName){
  return SpreadsheetApp.getActive().getSheetByName(sheetName); 
}
function tempLock(){
  var lock = LockService.getScriptLock();
  lock.waitLock(1000);
  lock.releaseLock();
}
function toast(message){
    SpreadsheetApp.getActiveSpreadsheet().toast(message);
}
function toast(message,title){
    SpreadsheetApp.getActiveSpreadsheet().toast(message,title);
}
function toast(message,title,timeoutSeconds){
    SpreadsheetApp.getActiveSpreadsheet().toast(message,title,timeoutSeconds);
}
