class Yoots {

    constructor(fresh=false,notif=false,enableLogging=true) {
        this.log(`Yoots called with fresh=${fresh}, notif=${notif}, enableLogging=${enableLogging}`);
        /* properties */
        if(!this._g) this._g={};
        if(!this.globalKeys) this.globalKeys=['config','apiInfo'];
        this.log(`init of _g`);
        this.enableLogging=enableLogging;
        this.loadGlobals(fresh,notif);
    }
    get config(){
        return this._g.config;
    }
    set globalKeys(value){
        this._globalKeys=value;
    }
    get globalKeys(){
        return this._globalKeys;
    }
    get apiKey(){
        //var apiInfo = this.getFromMemory("apiInfo");
        var apiInfo = this.retrieveFromMem("apiInfo");
        this.log(`get apiKey() -> this.retrieveFromMem('apiInfo') = ${JSON.stringify(apiInfo)}`);
        var apiKeyValue;
        if (apiInfo && apiInfo.success){
            apiKeyValue=apiInfo.accountInfo.apiKey;
            this.log(`     apiKey=${apiKeyValue} (from apiInfo)`);
        } else {
            apiKeyValue=this.getNamedCellValue(this.config.sheetNames.config,'Config_ASINAPIKEY');
            apiKeyValue=this.translateSUKey(apiKeyValue);
            this.log(`apiInfo not found so retrieving for ${apiKeyValue}`);
            this.apiInfo=this.retrieveAPIAccountInfo(apiKeyValue);
            apiKeyValue=this.apiInfo.success?this.apiInfo.accountInfo.apiKey:undefined;
            this.log(`     apiInfo=${JSON.stringify(this.apiInfo)}`);
        }
        this.log(`     apiKey=${apiKeyValue} (after translateSUKey)`);

        this.log(`              -> sending apiKey from get = ${JSON.stringify(apiKeyValue)}`);
        return apiKeyValue;
    }
    set apiKey(value){
        var apiInfo={};
        if (!value){
            apiInfo= {
                hasAPIKey:false,
                validAPIKey:false,
                accountInfo:{}
            };
        } else {
            var accountInfo = this.retrieveAPIAccountInfo(value);   // Validate apiKey
            if (accountInfo.success){
                apiInfo={
                    hasAPIKey:true,
                    validAPIKey:true,
                    accountInfo:accountInfo.accountInfo
                };
            } else {
                apiInfo={
                    hasAPIKey:true,
                    validAPIKey:false,
                    accountInfo:{}
                };
            }
        }
        
        this.setNamedCellValue(this.config.sheetNames.config,'Config_ASINAPIKEY',value);
    }
    get apiInfo(){
        return this.retrieveFromMem("apiInfo");
    }
    set apiInfo(value){
        this.assignToGlobalsAndCache("apiInfo",value);
    }
    get enableLogging(){
        return this._enableLogging;
    }
    set enableLogging(value){
        this._enableLogging=value;
    }
    translateSUKey(apiKey){
        return (apiKey==='.')?this.btoa("QkM0NTNEMDcwREQ5NDkyMEJFN0E3QTdBMzJCQjU0N0E"):apiKey;
    }

    loadConfig(){
        var config={
            sheetNames:{
                config: 'Config',
                data: 'Data',
            },
            namedRanges:{
                ASINAPIKEY:{sheet: 'config',name:'Config_ASINAPIKey',row:0,col:0}, //this.config.namedRanges.ASINAPIKEY.name
                //BASEAMZURL:{sheet: 'config', name:'Config_BaseAMZURL',row:0,col:0},
                ASIN:{sheet:'data', name:'Asin',col:0},
                ORDER:{sheet:'data', name:'Order',col:0},
                ETV:{sheet:'data', name:'ETV',col:0},
                MSRP:{sheet:'data', name:'MSRP',col:0},
                ITEM:{sheet:'data', name:'Item',col:0},
                CATEGORY:{sheet:'data', name:'Category',col:0},
                ORDEREDDATE:{sheet:'data', name:'Ordered',col:0},
                SHIPPEDDATE:{sheet:'data', name:'Shipped',col:0},
                RECEIVEDDATE:{sheet:'data', name:'Received',col:0},
                SUBMITTEDDATE:{sheet:'data', name:'Submitted',col:0},
                ACCEPTEDDATE:{sheet:'data', name:'Accepted',col:0},
                REJECTEDDATE:{sheet:'data', name:'Rejected',col:0},
                CANCELEDDATE:{sheet:'data', name:'Canceled',col:0},
                STARS:{sheet:'data', name:'Stars',col:0},
                PHOTOCOUNT:{sheet:'data', name:'Photos',col:0},
                VIDEOCOUNT:{sheet:'data', name:'Videos',col:0},
                TITLE:{sheet:'data', name:'Title',col:0},
                DETAIL:{sheet:'data', name:'Detail',col:0},
                NOTES:{sheet:'data', name:'Notes',col:0},
            },
            //ASINAPIKEY:'',
            BASEAMZURL:'https://www.amazon.com',
        };

        //Find named range values and locations
        for (const key in config.namedRanges){
            var namedRange = config.namedRanges[key];
            var sheet = this.getSheetByName(config.sheetNames[namedRange.sheet]);
            if (sheet) {
                if (namedRange.row>=0)namedRange.row= this.getRowByNamedRange(sheet, namedRange.name);
                if (namedRange.col>=0) namedRange.col=this.getColumnByNamedRange(sheet, namedRange.name);
                if (namedRange.row>=0 && namedRange.col>=0) config[key]=this.getNamedCellValue(config.sheetNames[namedRange.sheet], namedRange.name);
                if(key.toLowerCase()==='asin') this.log(`asin config setup = ${JSON.stringify(namedRange)}`);
            }
        }
        //this.log(`     this.config:${JSON.stringify(config)}`);
        return config;
    }
    getSheetByName(sheetName){
        return SpreadsheetApp.getActive().getSheetByName(sheetName); 
    }
    getSheetFromConfigKey(configSheetKey){
        return SpreadsheetApp.getActive().getSheetByName(this.config.sheetNames[configSheetKey]);
    }

    logTime(process, previous) {
        var latency = new Date().getTime() - previous.getTime();
        this.log('[TIMER] ' + process + ' (' + latency + ' seconds)');
        return new Date();
    }

    log(message) {
        if (!this.enableLogging)return;
        try{
            Logger.log(message);
        } catch {
            console.log(message);
        }
    }
    retrieveAPIAccountInfo(apiKey){
        this.log(`retrieveAPIAccountInfo('${apiKey}') called`);
        apiKey=this.translateSUKey(apiKey);
        var apiKeyFromSheet=this.getNamedCellValue(this.config.sheetNames.config, "Config_ASINAPIKey");
        if (this.apiInfo && apiKeyFromSheet===apiKey) {
            this.log(`found apiInfo and passed in apiKey(${apiKey} and sheet apiKey (${apiKeyFromSheet}) are equal so returning apiInfo (${JSON.stringify(this.apiInfo)})`);
            return this.apiInfo; //sheet and passed in identical so return existing
        }
        this.apiInfo=undefined;

        this.log(`different API Key detected....re-retrieving account info.`);

        var success=true;
        var message;
        this.log ('calling account API to get account Info');
        var accountURL = 'https://api.asindataapi.com/account?api_key='+apiKey;
        this.log(`accountURL = '${accountURL}'\n Getting content...`);
        var results = this.getContent(accountURL);  
        this.log(`content gotten.`);
        if (!results.success){
          return{
            success:false,
            message:results.message
          };
        }
        this.log(`was success.`);
      
        var response=results.response;
        this.log('results = '+response);
        var retrievedAccountInfo = JSON.parse(response);
        this.log('retrievedAccountInfo = '+JSON.stringify(retrievedAccountInfo));
        if (!retrievedAccountInfo.request_info.success) return undefined;
        var accountInfo = {
            autoTopupEnabled:retrievedAccountInfo.account_info.auto_top_up_enabled,
            topupCreditsRemaining:retrievedAccountInfo.account_info.topup_credits_remaining,
            apiKey:retrievedAccountInfo.account_info.api_key,
            name:retrievedAccountInfo.account_info.name,
            rateLimitPerMinute: retrievedAccountInfo.account_info.rate_limit_per_minute,
            plan:retrievedAccountInfo.account_info.plan,
            email:retrievedAccountInfo.account_info.email
        }
        this.log('apiInfo before = '+JSON.stringify(this.apiInfo));
        var retrievedAPIInfo= {
          success:success,
          message:message,
          accountInfo:accountInfo
        };
        this.apiInfo=retrievedAPIInfo;
        this.log('apiInfo after = '+JSON.stringify(this.apiInfo));
        return retrievedAPIInfo;
    }
    getCurrentItem(){
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
    
    getContent(url) {
    var success=false;
    var message='Issue retrieving Account object from ASIN Data API';
    var response;
    var contentText;
    var responseCode='N/A';
    //  try{
        let config = {
            muteHttpExceptions: true
        };
    
        response=UrlFetchApp.fetch(url,config);
        responseCode=response.getResponseCode();
        this.log(`getContent ResponseCode '${responseCode}`);
        if (responseCode===200)
        {
        success=true;
        contentText=response.contentText;
        } else if (responseCode===401){
        message='Invalid API Key';
        } else {
        message=response.request_info.message;
        }
    /*
    } catch (err){
        this.log(`getContent Catch '${JSON.stringify(err)}`);
        this.log(`getContent ResponseCode '${response.getResponseCode()}`);
    
        success=false;
        var request_infoStr = err.message.substring(err.message.indexOf('{'),err.message.lastIndexOf('}')+1)
        var request_info = JSON.parse(request_infoStr.trim()).request_info;
        message=request_info.message;
        responseCode=err.message.substring(err.message.indexOf('code')+5,err.message.indexOf('.',err.message.indexOf('code'))).trim();
    } finally{
    */
        var reply= {
        success:success,
        message:message,
        response:response,
        responseCode:responseCode
        };
        return reply;
    //  }
    }
            
    a1Notation(row,col,fullHeight){
        this.log(`a1Notations(${row},${col},${fullHeight})`);
        var col = `${String.fromCharCode(col+64)}`;
        var a1= `${col}${row}`;
        if (fullHeight)
            return `${a1}:${col}`;
        else
            return a1;
    }
    setASINAPIKeyTo(val)
    {
        this.setNamedCellValue(this.config.sheetNames.config,'Config_ASINAPIKEY',val);
    }
    
    getCurrentCacheValues(overrideKeys){
        var keys = overrideKeys??this.globalKeys;
        //this.log(`getCurrentCacheValues() keys=${JSON.stringify(keys)}`);
        var cacheValues={};
        //this.log(`cache.getAll(${JSON.stringify(keys)})=${JSON.stringify(CacheService.getUserCache().getAll(keys))}`);
        keys.forEach(key => {
            var value = this.getFromCache(key);
            if (value) {
                //this.log(`getCurrentCacheValues(${JSON.stringify(overrideKeys)}) found in ${key} in cache (${JSON.stringify(value)})`)
                cacheValues[key]=value;
            }
        });
        //this.log(`     returning cachedValues=${JSON.stringify(cacheValues)}`);
        return cacheValues;
    }
    getFromCache(key){
        var cache = CacheService.getUserCache();
        var value = cache.get(key);
        if (value && value!=="null" && value!=="{}"){
            //this.log(`getFromCache(${key}=${value})`);
            return JSON.parse(value);
        }
    }
    storeInCache(key,value){
        var cache=CacheService.getUserCache();
        cache.put(key,JSON.stringify(value));
    }
    nukeCache(){
        //this.log(`BEFORE nuking cache = ${JSON.stringify(this.getCurrentCacheValues())}`);
        //this.log(`     globalKeys=${JSON.stringify(this.globalKeys)}`);
        var oldGlobalKeys=this.globalKeys;
        var cache=CacheService.getUserCache();
        //this.log(`     GlobalKeys to remove = ${JSON.stringify(this.globalKeys)}`);
        cache.removeAll(this.globalKeys);
        this.globalKeys=[];
        this._g={};
        this.log(`AFTER nuking cache = ${JSON.stringify(this.getCurrentCacheValues(oldGlobalKeys))}`);
        this.log(`     globalKeys=${JSON.stringify(this.globalKeys)}`);
        this.log(`     GlobalKeys to remove = ${JSON.stringify(this.globalKeys)}`);
    }
    toast(message){
        SpreadsheetApp.getActiveSpreadsheet().toast(message);
    }
    toast(message,title){
        SpreadsheetApp.getActiveSpreadsheet().toast(message,title);
    }
    toast(message,title,timeoutSeconds){
        SpreadsheetApp.getActiveSpreadsheet().toast(message,title,timeoutSeconds);
    }
    assignToGlobalsAndCache(key,value)
    {
        //this.log(`assignToGlobalsAndCache() for ${key}}`);
        this._g[key]=value;
        this.storeInCache(key,value);
        if(this.globalKeys.indexOf(key)<0){ 
            this.globalKeys.push(key);
            //this.log(`   pushed ${key} into this.globalKeys.`)
        } else {
            //this.log(`   found ${key} already in this.globalKeys.`)
        }
    }
    getCacheKeyArray(){
        return [
                            "accountInfo",
                            "config",
                            "apiInfo"
        ];
    }
    
    loadGlobals(initAllFirst=false,notif=false){
        //this.log(`loadGlobals(initAllFirst=${initAllFirst})`);

        if(initAllFirst){
            //this.log('   clearing user cache...');
            this.nukeCache();
            if(notif) SpreadsheetApp.getUi().alert('Local Memory Refreshed.');
        }

        if (!this.retrieveFromMem("config"))  this.assignToGlobalsAndCache("config",this.loadConfig());    

        var hasAPIKey=false;
        var validAPIKey=false;
           
        if (this.apiKey && !this.retrieveFromMem("apiInfo")){                // see if apiInfo already retriewved if apiKey exists
            //this.log(`retrieving API Account info....(${this.apiKey})`);
            var accountInfo;
            try {
                accountInfo = this.retrieveAPIAccountInfo(this.apiKey);
                this.log(`accountInfo retrieved -> ${JSON.stringify(accountInfo)}`);
                if (accountInfo&&accountInfo.success){
                    hasAPIKey=true;
                    validAPIKey=true;
                } else {
                    hasAPIKey=true;
                    validAPIKey=false;
                }
                this.assignToGlobalsAndCache("apiInfo", {
                    hasAPIKey:hasAPIKey,
                    validAPIKey:validAPIKey,
                    accountInfo:accountInfo.accountInfo
                });
            } catch (err){
            this.log(`Hardstop Error: Error encountered retrieving accountInfo\n\nAccountInfo=${JSON.stringify(accountInfo)}\n\n${err.message}\n\nStackTrace:${err.stack}`);      
          }
        
        }
    }
    assignToGlobal(key,value){
        this._g[key]=value; 
        this.globalKeys.push(key);
    }
    //getFromMemory(key){
    //    if (this._g[key]) return this._g[key];
    //    if (this.getFromCache(key)) this.getFromCache(key);
    //}
    retrieveFromMem(key){
        if (this._g[key]) {
            this.log(`retrieveFromMem for ${key} -> found in _g, returning value (${JSON.stringify(this._g[key])})`);
            return this._g[key];     //globals has the value, all good
        }
        var fromCache = this.getFromCache(key);  
        if (fromCache){                     
            this.log(`retrieveFromMem for ${key} -> found in this.getFromCache(), updating _g and returning value`)
            this.assignToGlobal(key,fromCache); //globals doesn't have the value but cache does, assign to gobals, all good
            return fromCache;
        }
        this.log(`retrieveFromMem for ${key} -> not found, returning undefined.`)
        return undefined;                       //globals and cache do not have it, let calling code eval and assign
    }
    
    initAllFreshWithNotif(){
        this.loadGlobals(true);
        SpreadsheetApp.getUi().alert('Local Memory Refreshed.');
    }
    validAPIKey(){
        return (!(!this.apiKey));
    }
    apiKeyDisplay()
    {
        this.log(`apiKeyDisplay() -> ${JSON.stringify(this.apiKey)}`);
        return this.apiKey??'INVALID API KEY IN CONFIG';
    }    
    getNamedCellValue(sheetName, name) {
        var sheet=this.getSheetByName(sheetName);
        if (!sheet){
            this.log(`getNamedCellValue -> Sheet ${sheetName} not found!`);
        }
        var col = this.getColumnByNamedRange(sheet, name);
        var row = this.getRowByNamedRange(sheet, name);
        this.log(`getNamedCellValue '${sheetName}','${name}') -> sheet.getRange(${row},${col}).getValue()`);
        var value = sheet.getRange(row, col).getValue();
        return value;
    }
    setNamedCellValue(sheetName, name,value) {
        var sheet=this.getSheetByName(sheetName);
        var col = this.getColumnByNamedRange(sheet, name);
        var row = this.getRowByNamedRange(sheet, name);
        this.log(`setNamedCellValue '${sheetName}','${name}','${JSON.stringify(value)}) -> sheet.getRange(${row},${col})`);
        sheet.getRange(row, col).setValue(value);
    }
    getProductURL(asin) {
        var productURL = this._g.baseAMZURL + "/dp/" + asin;
        return productURL;
    }
    
    getNamedRange(sheet, name) {
        console.log(`getNamedRange(sheet:${sheet.getName()},'${name}')`);
        if (!sheet) return;
        var namedRanges = sheet.getNamedRanges();
        for (var i = 0; i < namedRanges.length; i++) {
            if (namedRanges[i].getName().toLowerCase() == name.toLowerCase()) {
                return namedRanges[i].getRange();
            }
        }
        console.log(`     NO MATCHES`);
    }
    
    getColumnByNamedRange(sheet, name) {
        if(!sheet) return;
        var namedRange = this.getNamedRange(sheet, name);
        if (namedRange != null){
            return namedRange.getColumn();
        } 
        this.log(`getColumnByNamedRange('${sheet.getName()}','${name}') did not find its data`);
        return -1;
    }
    
    getRowByNamedRange(sheet, name) {
        var namedRange = this.getNamedRange(sheet, name);
        if (namedRange != null){
            return namedRange.getRow();
        } 
        this.log(`getRowByNamedRange('${sheet.getName()}','${name}') did not find its data`);
        return -1;
    }
    
    nextNewRow(sheet){
        var asinNamedRange= this.getNamedRange(sheet,this.config.namedRanges.ASIN.name);
        var a1Not = this.a1Notation(asinNamedRange.getRow(),asinNamedRange.getColumn(),true);
        var asinValues = sheet.getRange(a1Not).getValues();
        var lastRow = asinValues.filter(String).length;
        this.log(`nextNewRow -> lastRow=${lastRow}, asinNamedRange.getRow()=${asinNamedRange.getRow()}`);
        return lastRow+asinNamedRange.getRow();
    }
    
    flush() {
        SpreadsheetApp.flush();
    }
    atob(text) {
        return Utilities.base64Encode(text, Utilities.Charset.UTF_8);
    }
    btoa(text) {
        return Utilities.newBlob(Utilities.base64Decode(text, Utilities.Charset.UTF_8)).getDataAsString();
    }
    getQueryStringFromParams(params) {
        return '?' + Object.keys(params).map(key => `${key}=${encodeURIComponent(params[key])}`).join('&');
    }
/*      
    cellUpdate(e) {
        this.log(`cellUpdate(e) eJSON='${JSON.stringify(e)}`);
        var editedSheet = e.range.getSheet().getName();
        var editedCol = e.range.getColumn();
        var editedRow = e.range.getRow();
    
        var sheetName = this.configsheet.name;//'Config/Help';
        var configSheet=this.getSheetByName(sheetName);
    
        this.log(`onEdit() triggered for ${editedSheet}:[${editedCol},${editedRow}]`);
        this.log(`onEdit() looking for match on ${sheetName}:[${config.namedRanges.ASINAPIKEY.col},${config.namedRanges.ASINAPIKEY.row}]`);
    
        let sheetCheck = (editedSheet == sheetName);
        let cellCheck = (editedCol===config.namedRanges.ASINAPIKEY.col && editedRow===config.namedRanges.ASINAPIKEY.row)
    
        if (!(sheetCheck && cellCheck)) {
            return
        }
        else {
            var ui = SpreadsheetApp.getUi();
            var performAccountCheck = ui.alert('API Key Edit','Changing the API Key requires it be checked for validity. Continue?',ui.ButtonSet.OK_CANCEL) ;
            if (performAccountCheck===ui.Button.OK){
                var editedAPIKey=this.getNamedCellValue(configSheet, this.config.namedRanges.ASINAPIKEY.name);
                var accountInfo =this.retrieveAPIAccountInfo(editedAPIKey);
                if (!accountInfo.success)
                {
                this.setASINAPIKeyTo(e.oldValue);
                ui.alert('API Key Validation',`${editedAPIKey} does not appear to be valid.\n\n${accountInfo.message}\n\nPlease log into your ASIN Data API account to double-check.`,ui.ButtonSet.OK);
                } else {
                ui.alert('API Key Validation','Valid API key detected.',ui.ButtonSet.OK);
                }
            } else {
                this.setASINAPIKeyTo(e.oldValue);
                ui.alert('API Key Validation','The value has been reset.',ui.ButtonSet.OK);
            }
            }
    }
*/

}