var ss = SpreadsheetApp.getActiveSpreadsheet();
var startColumn = 'summary';
var maxResults = 300;

function onOpen(){
  var menuEntries = [{
    name: "Complete info of tickets",
    functionName: "fillTicketsInfo"
  },{
    name: "Bring tickets from filter",
    functionName: "getTicketsFromFilter"
  }];
  
  ss.addMenu("Auto Fill Menu", menuEntries);
}

function getTicketsFromFilter() {
  var actualSheet = ss.getActiveSheet();
  var columnHeads = actualSheet.getSheetValues(1, 1, 1, actualSheet.getLastColumn())[0];
  var fields = getFiledsFromFisrtRow(columnHeads);
  var filterId = getFilterId(actualSheet.getName());
  var issues = paginateResults(filterId);
  
  for (var i=0; i<issues.length; i++) { 
    writeInSheet(actualSheet, i+2, issues[i], columnHeads, fields);
  }
}

function fillTicketsInfo(){
  var actualSheet = ss.getActiveSheet();
  var allJiraIds = actualSheet.getSheetValues(2, 2, actualSheet.getLastRow()-1, 1);
  var columns = actualSheet.getSheetValues(1, 1, 1, actualSheet.getLastColumn())[0];
  var fieldsToLoad = getFiledsFromFisrtRow(columns);
  
  for (var i=0; i<allJiraIds.length; i++) {
    var jiraId = allJiraIds[i];
    
    if(jiraId.length > 0 && jiraId[0] == "") {
      continue;
    }
    
    writeInSheet(actualSheet, i+2, infoByTicketId(fieldsToLoad, "issue" + jiraId), columns, fieldsToLoad);
  }
}

function infoByTicketId(fields, jiraId) {  
  return hitJira(jiraId + '?fields=' + fields.join());
}

function writeInSheet(actualSheet, rowIndex, issue, columns, fieldsToLoad){
  for (var j=0;j<fieldsToLoad.length;j++) {
    var columnIndex = getColumnIndex(columns, fieldsToLoad[j]) + 1;
    
    actualSheet.getRange(rowIndex, columnIndex).setValue(normalizeField(issue, fieldsToLoad[j]));     
  }
}

function normalizeField(issue, fieldName) {
  var value = fieldName == 'key' ? processValue(issue.key) : processValue(accssesJsonNestedProperty(issue.fields, fieldName));

  //Unifica los datos de los 2 proyectos para la grafica.
  if(fieldName.indexOf('project') == 0 && (value == "EDP" || value == "Checkout" || value == "CHECK")) {
    value = "EDP-Checkout";
  }
   
  if(fieldName == 'aggregatetimespent' && value != null) {
    value = value/3600;
  }
 
  return value;
}

function processValue(value) {
  var valueType = typeof value;
  
  if (value == null) {
    return "";
  } else if(valueType === 'string' || valueType === 'number') {
    return value;
  } else if (Array.isArray(value)) {
    var arrayValues = [];
    for (var i=0; i<value.length; i++){
      arrayValues.push(processValue(value[i]));
    }
    return arrayValues.join();
  } else if (valueType === 'object') {
    var name = value.name;
    if (name == null) {
      name =value.value;
    }
    
    return name;
  }
  
  return "";
}

function getColumnIndex(firstRow, columnName) {
  return firstRow.indexOf(columnName);
}

function getFiledsFromFisrtRow(firstRow) {
  var indexOfStartColumn = getColumnIndex(firstRow, startColumn);
  var fields = [];
  
  for (var i=indexOfStartColumn;i<firstRow.length;i++){
    fields.push(firstRow[i])
  }
  
  return fields;
}

function paginateResults(filterId){
  var totalIssues = getTotalCount(filterId);
  var results = new Array();
  
  if(totalIssues > maxResults) {
    var pages = Math.floor(totalIssues/maxResults);
    var startIndex = 0;
    for (var i = 0; i <= pages; i++) {
      results = results.concat(hitJira('search?jql=filter+%3D+' + filterId + '&startAt='+ startIndex +'&maxResults=' + maxResults).issues);
      startIndex += maxResults;
    }
  }
  else {
    results = hitJira('search?jql=filter+%3D+' + filterId + '&maxResults=' + maxResults).issues
  }
  
  return results;
}

function getTotalCount(filterId) {
  return hitJira('search?jql=filter+%3D+' + filterId + '&maxResults=1&fields=*none').total;
}

function hitJira(apiToHit, payload){
  //Add your jira url and authentication
  var url = '/jira/rest/api/2/' + apiToHit;
  var options = {
    headers: {
      'Authorization': ''
    },
    contentType: 'application/json',
    method: "get"
  };
  
  if(payload != null){
    options.method = "put";
    options.payload = payload;
  }
  
  var response = UrlFetchApp.fetch(url, options);
  
  if(response.getContentText() == null || response.getContentText() == ""){
    return {};
  }
  
  return JSON.parse(response.getContentText());
}

function getFilterId(sheetToGetTheFilter){
  var filterSheet = ss.getSheetByName("Filters");
  var filters = filterSheet.getSheetValues(1, 1, filterSheet.getLastRow(), 1);

  for (var i = 0; i<filters.length;i++) {
    if(filters[i] == sheetToGetTheFilter){
      return filterSheet.getSheetValues(i+1, 2, 1, 1)[0];
    }
  } 
}

function accssesJsonNestedProperty(instance, property) {
    property = property.replace(/\[(\w+)\]/g, '.$1'); // convert indexes to properties
    property = property.replace(/^\./, '');           // strip a leading dot
    var a = property.split('.');
    for (var i = 0; i<a.length;i++) {
        var k = a[i];
        if (k in instance) {
            instance = instance[k];
        } else {
            return;
        }
    }
    return instance;
}

function updateProgramValue(){
  var actualSheet = ss.getActiveSheet();
  var allJiraIds = actualSheet.getSheetValues(2, 1, actualSheet.getLastRow()-1, 1);
  
   for (var i=0; i<allJiraIds.length; i++) {
     var jiraId = allJiraIds[i];
    
     if(jiraId.length > 0 && jiraId[0] == "") {
       continue;
     }
     
     var payload = {
       "update": {
         "customfield_10891": [{"add":{
                  "value": ""
         }
                               }]
       }
     }
     
     var ticket = hitJira(jiraId[0], JSON.stringify(payload));
     Logger.log(ticket);
   }
}
