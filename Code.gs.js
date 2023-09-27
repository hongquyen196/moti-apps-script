/*
# CREATED BY: BPWEBS.COM
# URL: https://www.bpwebs.com
*/

function doGet(request) {
  return HtmlService.createTemplateFromFile('Index').evaluate();
}


/* DEFINE GLOBAL VARIABLES, CHANGE THESE VARIABLES TO MATCH WITH YOUR SHEET */
function globalVariables() {
  var firstRow = 'A3';                                              //** CHANGE !!!
  var lastCol = 'HP';                                                //** CHANGE !!!
  var varArray = {
    //spreadsheetId : '1_-47ZdLK69Dq5N1M33MpSB3Gsbp4PT9C1gH4qXR66ec',  //** CHANGE !!!
    spreadsheetId : '1xnyb6QcRL3Wvm6ZOlVNshwB_Oco3E1N7iESY5PrfD-k',  //** CHANGE !!!
    dataRage      : 'Data!' + firstRow + ':' + lastCol,              //** NO CHANGE !
    idRange       : 'Data!' + firstRow + ':A',                       //** NO CHANGE !
    firstRow      : firstRow,                                        //** NO CHANGE !
    lastCol       : lastCol,                                         //** NO CHANGE !
    insertRange   : 'Data!A1:' + lastCol + '1',                      //** NO CHANGE !
    sheetID       : '0'                                              //** CHANGE !!! .../edit#gid=0  Ref:https://developers.google.com/sheets/api/guides/concepts#sheet_id  
  };
  return varArray;
}

/*
# PROCESSING FORM ---------------------------------------------------------------------------------
*/


/* PROCESS FORM */
function processForm(formObject) {
  var validation = null;
  if (formObject.RecId && checkID(formObject.RecId)) {//Execute if form passes an ID and if is an existing ID
    validation = checkCanUpdate(formObject);
    if (validation) {
      //RELOAD LASTEST ROWS
      getLastTenRows();
      return validation;
    }
    updateData(getFormValues(formObject), globalVariables().spreadsheetId, getRangeByID(formObject.RecId)); // Update Data
  } else { //Execute if form does not pass an ID
    // validation = checkOldCustomer(formObject);
    // if (validation) {
    //   return validation;
    //}
    appendData(getFormValues(formObject), globalVariables().spreadsheetId, globalVariables().insertRange); //Append Form Data
  }
  return getLastTenRows();//Return lastest rows
}

/* GET FORM OBJECT CELL GENERATE FROM 0_0 TO 10_20 */
function getformObjectCells(formObject) {
  var objectCells = [];
  for(var i = 0; i < 10; i++) {
    for(var j = 0; j < 20; j++)
    objectCells.push(formObject[`cell${i}_${j}`])
  }
  return objectCells;
}

/* GET FORM VALUES AS AN ARRAY */
function getFormValues(formObject) {
  const formInfo = [
    formObject.fullName,
    formObject.phone,
    formObject.address,
    formObject.city,
    formObject.district,
    formObject.page].map(v => v.trim());
  const formData = [
    null,
    formObject.quantity,
    formObject.total,
    formObject.ship,
    formObject.deposit,
    formObject.note,
    null,
    null,
    formObject.prepair,
    null,
    null,
    null,
    null,
    null,
    null,
    ...getformObjectCells(formObject)
  ].map(v => isNumeric(v) ? parseInt(v) : v);
  /* ADD OR REMOVE VARIABLES ACCORDING TO YOUR FORM*/
  if (formObject.RecId && checkID(formObject.RecId)) {
    var values = [[formObject.RecId.toString(), formObject.DateCreated, formObject.CreatedBy , ...formInfo, ...formData]];
  } else {
    var values = [[new Date().getTime().toString(), getCurrentDate(), getCurrentUser(), ...formInfo, ...formData]];
  }
  return values;
}

function isNumeric(str) {
  if (typeof str != "string") return false;
  return !isNaN(str) &&
         !isNaN(parseFloat(str));
}

/*
## CURD FUNCTIONS ----------------------------------------------------------------------------------------
*/


/* CREATE/ APPEND DATA */
function appendData(values, spreadsheetId, range) {
  var valueRange = Sheets.newRowData();
  valueRange.values = values;
  var appendRequest = Sheets.newAppendCellsRequest();
  appendRequest.sheetID = spreadsheetId;
  appendRequest.rows = valueRange;
  var results = Sheets.Spreadsheets.Values.append(valueRange, spreadsheetId, range, { valueInputOption: "USER_ENTERED" });
}


/* READ DATA */
function readData(spreadsheetId, range) {
  var result = Sheets.Spreadsheets.Values.get(spreadsheetId, range);
  return result.values;
}

function getLastIndexRow() {
  var lastIndexRow = Sheets.Spreadsheets.Values.get(globalVariables().spreadsheetId, "A5");
  var lastIndex = lastIndexRow.values[0][0];
  console.log(lastIndex);
  return lastIndex;
}


/* UPDATE DATA */
function updateData(values, spreadsheetId, range) {
  var valueRange = Sheets.newValueRange();
  valueRange.values = values;
  var result = Sheets.Spreadsheets.Values.update(valueRange, spreadsheetId, range, { valueInputOption: "USER_ENTERED" });
}


/*DELETE DATA*/
function deleteData(ID) {
  //https://developers.google.com/sheets/api/guides/batchupdate
  //https://developers.google.com/sheets/api/samples/rowcolumn#delete_rows_or_columns
  var startIndex = getRowIndexByID(ID);

  var deleteRange = {
    "sheetId": globalVariables().sheetID,
    "dimension": "ROWS",
    "startIndex": startIndex,
    "endIndex": startIndex + 1
  }

  var deleteRequest = [{ "deleteDimension": { "range": deleteRange } }];
  Sheets.Spreadsheets.batchUpdate({ "requests": deleteRequest }, globalVariables().spreadsheetId);

  return getLastTenRows();//Return lastest rows
}



/* 
## HELPER FUNCTIONS FOR CRUD OPERATIONS --------------------------------------------------------------
*/


/* CHECK FOR EXISTING ID, RETURN BOOLEAN */
function checkID(ID) {
  var idList = readData(globalVariables().spreadsheetId, globalVariables().idRange).reduce(function (a, b) { return a.concat(b); });
  return idList.includes(ID);
}


/* GET DATA RANGE A10 NOTATION FOR GIVEN ID */
function getRangeByID(id) {
  if (id) {
    var startRow = parseInt(globalVariables().firstRow.replace(/^\D+/g, ''));
    var idList = readData(globalVariables().spreadsheetId, globalVariables().idRange);
    for (var i = 0; i < idList.length; i++) {
      if (id == idList[i][0]) {
        return 'Data!A' + (i + startRow) + ':' + globalVariables().lastCol + (i + startRow);
      }
    }
  }
}


/* GET RECORD BY ID */
function getRecordById(id) {
  if (id && checkID(id)) {
    var result = readData(globalVariables().spreadsheetId, getRangeByID(id));
    return result;
  }
}

/* GET RECORD BY PHONE */
function getRecordByPhone(phone) {
  var data = readData(globalVariables().spreadsheetId, 'Data!E:E') || [];
  var result = data.find(record => record == phone);
  return result;
}

/* GET ROW NUMBER FOR GIVEN ID */
function getRowIndexByID(id) {
  if (id) {
    var idList = readData(globalVariables().spreadsheetId, globalVariables().idRange);
    for (var i = 0; i < idList.length; i++) {
      if (id == idList[i][0]) {
        var rowIndex = parseInt(i + 1);
        return rowIndex;
      }
    }
  }
}

/* GET ALL RECORDS */
function getAllData() {
  var data = readData(globalVariables().spreadsheetId, globalVariables().dataRage) || [];
  return data.filter(record => record && record[0]);
}

/*GET LAST 5 RECORDS */
function getLastFiveRows() {
  var lastIdx = getLastIndexRow();
  var range = 'Data!A' + (lastIdx - 5) + ':' + globalVariables().lastCol + lastIdx;
  var data = readData(globalVariables().spreadsheetId, range) || [];
  return data.filter(record => record && record[0]);
}

/*GET LAST 10 RECORDS */
function getLastTenRows(data) {
  var lastTenRows = data || getAllData();
  var lastRow = lastTenRows.length;
  if (lastRow > 10) {
    lastTenRows = lastTenRows.slice(lastRow - 10, lastRow);
  }
  return lastTenRows;
}

/* FULL TEXT SEARCH */
function getFullTextSearch(value = "") {
  var data = getAllData().filter(record => JSON.stringify(record).includes(value));
  return getLastTenRows(data);
}

/* CHECK CAN UPDATE SG, SNT */
function checkCanUpdate(formObject) {
  var message = "";
  var { RecId } = formObject;
  var record = getRecordById(RecId);
  if (record && record[0]) {
    record = record[0];
    if (!["SG","SNT"].includes(record[17])) {
      message = "Dữ liệu 'Chuẩn bị' đã bị thay đổi. Không thể cập nhật."
    }
  }
  Logger.log(message);
  return message;
}

/* CHECK OLD CUSTOMER */
function checkOldCustomer(formObject) {
  var message = "";
  var { phone, page } = formObject;
  var phoneCheck = !!getRecordByPhone(phone);
  var oldCustomer = [
    "Khách_cũ_Vàng",
    "Khách_cũ_Nâu",
    "Khách_cũ_Hồng",
    "Khách_cũ_Đỏ",
    "TH",
    "BH"
  ]
  var pageCheck = !oldCustomer.includes(page);
  if (phoneCheck && pageCheck) {
    message = "Số điện thoại [" + phone + "], Page [" + page + "] bị trùng với đơn của khách cũ. Cần chỉnh sửa lại Page phù hợp."
  }
  Logger.log(message);
  return message;
}

/*
## OTHER HELPERS FUNCTIONS ------------------------------------------------------------------------
*/

/*GET DROPDOWN LIST */
function getDropdownList(range) {
  var list = readData(globalVariables().spreadsheetId, range);
  return list;
}

/*GET DROPDOWN LISTS */
function getDropdownLists(ranges) {
  return ranges.map(f => ({id: f, values: getDropdownList(f)}));
}

/* INCLUDE HTML PARTS, EG. JAVASCRIPT, CSS, OTHER HTML FILES */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getCurrentUser() {
  var userInfo = getNameOfCurrentEditor();
  Logger.log(userInfo)
  return userInfo;
}

function getNameOfCurrentEditor() {
  var userData = "";
  // try {
  //   const idToken = ScriptApp.getIdentityToken();
  //   const body = idToken.split('.')[1];
  //   const decoded = Utilities.newBlob(Utilities.base64Decode(body)).getDataAsString();
  //   const { given_name: firstName, family_name: lastName } = JSON.parse(decoded);
  //   userData = firstName;
  // } catch (e) {
    userData = Session.getActiveUser().getUsername();
  // }
  return userData;
}

function getCurrentDate() {
  return Utilities.formatDate(new Date(), "GMT+7", "M/d")
}

function getDistrictOfCity(range, value) {
  var cities = readData(globalVariables().spreadsheetId, range);
  var disctrict = [];
  for (var i = 0, length = cities.length; i < length; i++) {
    if (cities[i][0] === value) {
      disctrict = readData(globalVariables().spreadsheetId, cities[i][1]);
      break;
    }
  }
  return disctrict;
}


