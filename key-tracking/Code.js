/*
Object Oriented Design
Goals:
-Alumni/Peoples who have reached expiration date
-Renew key?
-Edit currentFormToClass to deal with wrong dates
-Edit currentFormToClass to deal with no advisor
-1st Year PhD  advisor is Heather
*/

class keyInfo{
  constructor(andrewID,keyNumber,roomNumber,givenDate,expDate){
    this.id         = hash_id(keyNumber+andrewID)
    this.keyNumber  = keyNumber          //A string due to the dash in the middle
    this.roomNumber = roomNumber         //String to designate building name
    this.givenDate  = givenDate;
    this.expDate    = expDate;
    this.status     = true
  }
  getId(){
    return this.id
  }
  getKey(){
    return this.keyNumber
  }
  getRoom(){
    return this.roomNumber
  }
  getGivenDate(){
    return this.givenDate
  }
  getExpirationDate(){
    return this.expDate
  }
  getStatus(){
    return this.status
  }
}

class keyRecord {
  constructor(first,last,andrewID,advisor,dept,key,room,givenDate,expDate)  {
    this.firstName = first;
    this.lastName  = last;
    this.andrewID  = (andrewID == "") ? this.firstName +"_"+ this.lastName+ "_no_andrew_id" : andrewID; 
    this.advisor   = advisor;
    this.dept      = dept;
    this.key       = [new keyInfo(andrewID,key,room,givenDate,expDate)];
  }
  //Basic constructor functions
  getFirstName(){
    return this.firstName
  }
  getLastName(){
    return this.lastName
  }
  getName() {
    return this.firstName +" " +this.lastName
  }
  getAndrewID(){
    return this.andrewID
  }
  getAdvisor(){
    return this.advisor
  }
  getDepartment(){
    return this.dept
  }
  getKeys(){
    return this.key
  }
  setKey(newKeySet){
    this.key = newKeySet
  }
 
  addKey(andrewID,key,room,givenDate,expDate) {
    var newKey = new keyInfo(andrewID,key,room,givenDate,expDate)
    var keys = this.key
    keys.push(newKey)
    this.key = keys
  }
  removeKey(remKey){
    this.key.forEach((keyInfo) =>{
      if(keyInfo.getKey() == remKey){
        keyInfo.deactivate()
      }
    });
  }
}

//Murmer Hash3 function set up
function hash_id(str){
  var hash =  0x811c9dc5 // FNV offset basis
  for(var i = 0; i < str.length; i++){
    hash ^= str.charCodeAt(i)
    hash  = (hash * 0x01000193) >>> 0
  }
  return hash >>> 0
}

/**(
 * Capture changes to active spreadsheet (Key Main Sheet)
 * @param {*} e - event object
 */
function onEdit(e) {
  const sheet           = e.range.getSheet();
  const unverifiedSheet = "Unverified Input";

  if (sheet.getName() === unverifiedSheet) {
    const row    = e.range.getRow()
    const col    = e.range.getColumn()
    const values = e.range.getValues()

    // Only trigger if row >= 2 and column is between 3 (C) and 11 (K)
    if (row >= 2 && col >= 3 && col <= 11) {
      for(var r = 0; r < values.length; r++){
        for(var c = 0; c < values[0].length; c++){
          var temp_r = row + r
          var temp_c = col + c
          var value = values[r][c]
          submitUnverifedData(temp_r,temp_c,value)
        }
      }      
    }
  }
}

function onFormSubmit(e){
  var sheet = e.range.getSheet()
  var sheetName = sheet.getName()
  var data  = e.values

  if(sheetName === "Key Check-In Form"){
    var date_returned = data[0];
    var email         = data[1];
    var firstName     = data[2];
    var lastName      = data[3];
    var advisor       = data[4];
    var andrewID      = data[5];
    var key           = data[6];
    var room          = data[7];
    manualCheckIn(andrewID,firstName,lastName,advisor,key,room)
  }

  if (sheetName == " Key Check-Out Form"){
    var date_returned = data[0]
    var email         = data[1]
    var firstName     = data[2]
    var lastName      = data[3]
    var advisor       = data[4]
    var andrewID      = data[5]
    //figure out how to deal with 6
    var key           = data[7]
    var room          = data[8]
    var givenDate     = data[9]
    var expDate       = data[10]




    //add to log and add to unverifed!!!!!
    //1. append to unverifeid
    //2. append to add to log

    //FIX LOG ENTRIES!!!!!
  }
}


/**
 * Returns all active entries from the log     //////////////////////////////////////////////////////NOT IN USE
 */
function activeEntries(){
  const keySS    = SpreadsheetApp.getActiveSpreadsheet()
  const logSheet = keySS.getSheetByName('Log')
  const range    = logSheet.getRange(2,1,logSheet.getLastRow(),logSheet.getLastColumn()) 
  const log_values = range.getValues()

  var active_entries = new Map()
  for(var log_row in log_values){
    var status    = log_row[0]
    var andrewID  = log_row[2]
    var lastName  = log_row[3]
    var firstName = log_row[4]
    var advisor   = log_row[5]
    var dept      = log_row[6]
    var key       = log_row[7]
    var room      = log_row[8]
    var expDate   = log_row[9]
    var givenDate = log_row[10]
    if(status == "Active"){
      var newKeyRec = new keyRecord(firstName,lastName,andrewID,advisor,dept,key,room,givenDate,expDate)
      active_entries.set(andrewID,newKeyRec)
    }
  }
  return active_entries
}

/**
 * Returns all inactive entries from the log         /////////////////////////////////////////NOT IN USE
 */
function inactiveEntries(){
  const keySS    = SpreadsheetApp.getActiveSpreadsheet()
  const logSheet = keySS.getSheetByName('Log')
  const range    = logSheet.getRange(2,1,logSheet.getlastRow(),logSheet.getLastColumn()) 
  const log_values = range.getValues()

  var inactive_entries = new Map()
  for(var log_row in log_values){
    var status    = log_row[0]
    var andrewID  = log_row[2]
    var lastName  = log_row[3]
    var firstName = log_row[4]
    var advisor   = log_row[5]
    var dept      = log_row[6]
    var key       = log_row[7]
    var room      = log_row[8]
    var expDate   = log_row[9]
    var givenDate = log_row[10]
    if(status == "Inactive"){
      var newKeyRec = new keyRecord(firstName,lastName,andrewID,advisor,dept,key,room,givenDate,expDate)
      inactive_entries.set(andrewID,newKeyRec)
    }
  }
  return inactive_entries
}





















/**
 * Returns all active entries that are verified as valid inputs
 */
function verifiedEntries(keySS){
  const logSheet = keySS.getSheetByName('Log')
  const range    = logSheet.getRange(2,1,logSheet.getLastRow(),logSheet.getLastColumn()) 
  const log_values = range.getValues()

  var verifiedEntries = new Map()
  for(var log_row of log_values){
    const id        = log_row[0]
    const status    = log_row[1]
    const approval  = log_row[2]
    const andrewID  = log_row[3]
    const lastName  = log_row[4]
    const firstName = log_row[5]
    const advisor   = log_row[6]
    const dept      = log_row[7]
    const key       = log_row[8]
    const room      = log_row[9]
    const expDate   = log_row[10]
    const givenDate = log_row[11]
    if((status == "Active") && (approval == "Approved")){
      var newKeyRec = new keyRecord(firstName,lastName,andrewID,advisor,dept,key,room,givenDate,expDate)
      verifiedEntries.set(andrewID,newKeyRec)
    }
  }
  return verifiedEntries
}

/***************Helper Functions used for safety checks********/
function validKey(key) {
  //Some error with Key formating
  if(!key.includes("4501-")){
    //Key doesnt have dash, need to correct
    if(key.includes("4501")){
      end_i = key.length - 1;
      i = end_i;
      key_copy = key;
      base = "";
      add = "";
      while((i != -1) || (key_copy.slice(0,i) != "4501")){//(key[i] != "-")){
        base = key[i].concat(base);
        i--;
     }
    } 
    else return "invalid key";
  } return key;
}

function validRoomNum(num){
  const floorOpt = ["D","C","B","A","1","2","3","4","a","b","c","d"]; //find better way to deal with no Cap
  floor = num[0];
  digits = num.slice(1,num.length);
  validFloor = false;
  for(opt in floorOpt){if(opt == floor) validFloor = true;}
  //Is the first digit the floor number
  if(!validFloor) return false;
  //valid lenght of digits (3 digits to be room num in Doherty)
  else if(digits.length != 3) return false;
  //Are the digits(remaining room num) a valid number
  else if(parseInt(digits) == NaN) return false;
  else return true;
}

//make sure the second half is actually a number!!!!
function validRoom(room){
  roomNum = 0
  if(room.includes("DH ")){
    return  validRoomNum(room.slice(3,room.length)) ? room : "invalid room";
  } 
  else if(room.includes("Doherty")){
    roomNum = room.slice(7,room.length);
    while (roomNum[0] == " "){
      roomNum = roomNum.slice(1,roomNum.length);
    }
    return validRoomNum(roomNum) ? "DH ".concat(roomNum): "invalid room";
  } 
  else if(room.includes("DH")) {
    roomNum = room.slice(2,room.length);
    return validRoomNum(roomNum)? "DH ".concat(roomNum): "invalid room";
  } 
  else return "invalid room"
}

function validDate(date){
  //1. Date object
  if(Object.prototype.toString.call(date) === "[object Date]"){
    //if NaN, the date is not possible
    if(isNaN(date.getTime())){return "invalid date"}
    return date 
  }
  //String formatted as date (reformat to date object)
  if(typeof date === "string"){
    //YYYY-MM-DD
    var iso_date_regex = /^\d{4}-\d{2}-\d{2}$/; 
    //MM/DD/YYYY
    var  us_date_regex = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
    var split;
    if(iso_date_regex.test(date)){
      split = date.split("-")
    }
    else if(us_date_regex.test(date)){
      split = date.split("/")
    } else {return "invalid date"}
    var year  = parseInt(split[0],10)
    var month = parseInt(split[1],10)
    var day   = parseInt(split[2],10)

    //Note: Month in date option starts at 0 (Jan = 0, Feb = 1,Mar = 2, ...)
    var full_date = new Date(year,month-1,day)

    const year_valid  = full_date.getFullYear()  === year
    const month_valid = full_date.getMonth() + 1 === month
    const day_valid   = full_date.getDate()      === day
    if(year_valid && month_valid && day_valid){
      return full_date
    }  
  }
  return "invalid Date"
}

function confirmUser(first,last,advisor,andrew,key,room,entry){
  const key_rooms = entry.getKeys() 
  var key_room_status = false
  for(var pairs in key_rooms){
    if((pairs.getKey() === key) &&(pairs.getRoom() ===room)){
      key_room_status = true
    }
  }
  //need a better way to access rooms and keys!!!
  if((first == entry.getFirstName()) && (last == entry.getLastName()) 
    && (advisor == entry.getAdvisor())&& (andrew == entry.getAndrewID()) 
    && key_room_status) {
      return true
    } return false
}

/*****************Parsing entry data***********************/
/**
 * Parsing the sheets for entries 
 * */
function parseKeySheet(allEntries,id){
  var newSS = SpreadsheetApp.openById(id)
  var sheets = newSS.getSheets()
  
  //All the data
  for(var i = 0; i < sheets.length; i++){
    //Figure out the end range <--or how to avoid going over empty setion
    var sheet = sheets[i];
    var firstNameArr = sheet.getRange('D2:D').getValues(); 
    var lastNameArr  = sheet.getRange('C2:C').getValues();
    var andrewIDArr  = sheet.getRange('F2:F').getValues();
    var advisorArr   = sheet.getRange('G2:G').getValues();
    var deptArr      = sheet.getRange('I2:I').getValues();
    var keysArr      = sheet.getRange('J2:J').getValues();
    var roomArr      = sheet.getRange('K2:K').getValues();
    var givenArr     = sheet.getRange('M2:M').getValues();
    var expArr       = sheet.getRange('N2:N').getValues();

    for(var i = 0; i < 110; i++){ //Replace with length of the rows.
      var firstName = firstNameArr[i][0];
      var lastName = lastNameArr[i][0];
      var andrewID = andrewIDArr[i][0];
      var advisor = advisorArr[i][0];
      var dept = deptArr[i][0];
      var keys = keysArr[i][0];
      var room = roomArr[i][0];
      var given0 = givenArr[i][0];
      var exp0 = expArr[i][0];

      //No AndrewIDs
      if(andrewID == ""){
        var bEntry = new keyRecord(firstName,lastName,andrewID, advisor, dept,keys,room,given0,exp0);
        allEntries.set(firstName.concat(lastName,"-no-andrewID"),bEntry);
      }
      //No Entries or New AndrewID
      else if((!allEntries.has(andrewID)) || (allEntries.size== 0)) {
        var newEntry = new keyRecord(firstName,lastName,andrewID, advisor, dept,keys,room,given0,exp0); //Given and exp date not given
        allEntries.set(andrewID,newEntry);
      } 
      //Adding a key to existing record
      else {//<---Overwriting old values
        var entry = allEntries.get(andrewID); 
        allEntries.delete(andrewID); 
        entry.addKey(keys,room,given0,exp0);
        allEntries.set(andrewID,entry); 
      }
    }
  }
  //Eventually, will add the new form once it starts to populate
  return allEntries;
}

/**
 * Translate form response to entries (with keyReponse class) )
 * */
function checkoutFormToEntries(allEntries){
  //Array of form responses
  var firstName,lastName, advisor,andrewID, key,room,givenDate,expDate,ques,answ,dept;
  const form1 = FormApp.openByUrl(
    'https://docs.google.com/forms/d/1fPmkuLoWQXsgwz1ruQw3rkGO93eN1PrUEUINBaV4MBc/edit');
  var allResp = form1.getResponses();

  //Individual responses
  for(const resp of allResp) {
    //All the questions and response stores in an item
    for(item of resp.getItemResponses()){
      ques = item.getItem().getTitle();
      answ = item.getResponse();
      if(ques == "First Name:"){
        firstName = answ;
      } else if(ques == "Last Name:") {
        lastName = answ;
      } else if(ques == "Advisor:") {
        advisor = answ;
      } else if(ques == "andrewID:") {
        andrewID = answ;
      } else if(ques == "Key Number:") {
        key = validKey(answ);
      } else if(ques == "Room (Include Building and Room Number) Ex: DH 3213A") {
        room = validRoom(answ);
      } else if(ques == "What date were you given the key/key access?") {
        givenDate = validDate(answ);
      } else if(ques == "What date will you lose acess? (Typically expected graduation date)") {
        expDate = validDate(answ);
      } else if(ques == "Are you a part of the Chemical Engineering Department?") {
        if(answ == "Yes"){
          dept = "Chemical Engineering";
        } else if("No"){
          dept = "Other Department";
        } else{
          dept = answ;
        }
      }
    }
    if(!allEntries.has(andrewID)){
      var newEntry = new keyRecord(firstName,lastName,andrewID,advisor,
                                      dept,key,room,givenDate,expDate);
      allEntries.set(andrewID,newEntry);
    } else {
        var Entry = allEntries.get(andrewID);
        allEntries.delete(andrewID);
        Entry.addKey(key,room,givenDate,expDate);
        allEntries.set(andrewID,Entry);
    }
  }
  return allEntries
}

/*****************Manipulating the log sheet*************/
/**
 * Read log data and turn ACTIVE values into entries
 */
function logToEntries(){
  var keySS = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = keySS.getSheetByName('Log');
  var allEntries = new Map();

  var logRange  = logSheet.getRange('B:L');
  var logValues = logRange.getValues();

  for(var i = 1; i < logValues.length; i++){
    var row = logValues[i];
    if(row.length == 0 || row.length < 11 || row[0] == "Inactive" || row[0] == "" || row[0] == NaN) continue;
    var andrewID  = row[2];
    var lastName  = row[3];
    var firstName = row[4];
    var advisor   = row[5];
    var dept      = row[6];
    var key       = row[7];
    var room      = row[8];
    var expDate   = row[9];
    var givenDate = row[10];
    if((andrewID == '') && (lastName =='') && (firstName == '') && 
       (advisor == '') && (dept == '') && (key == '') && (room == '') &&
       (expDate == '') && (givenDate == '')){break;}

    var newKeyRec = new keyRecord(firstName,lastName,andrewID,advisor,dept,key,room,givenDate,expDate);
    allEntries.set(andrewID,newKeyRec);
  }
  return allEntries;
}

/**
 * Adds any value to the log based on the input to the function
 */
function addToLog(andrewID,keyRecord,logSheet,logEntries,activity){
  var keys = keyRecord.getKeys()
  for(var i = 0; i < keys.length; i++){
    var key = keys[i]
    if(!logEntries.has(andrewID)){
      logSheet.appendRow([
        key.getId(),
        activity,
        'Unverified',
        keyRecord.getAndrewID(),
        keyRecord.getLastName(),
        keyRecord.getFirstName(),
        keyRecord.getAdvisor(),
        keyRecord.getDepartment(),
        key.getKey(),
        key.getRoom(),
        key.getExpirationDate(),
        key.getGivenDate()
      ])
    }
  }
}

/**
 * Update approval status of a log (based on what happens in the unverifed sheet)
 */
function updateLogApproval(id,andrewID,key,approval,status){
  const keySS    = SpreadsheetApp.getActiveSpreadsheet()
  const logSheet = keySS.getSheetByName('Log')

  //Find all instances of the andrewID and the key value in the spreadsheet
  if (id != -1){
    var found = logSheet.createTextFinder(id).findAll()[0].getRow()
  } else {
    var andrew_found = logSheet.createTextFinder(andrewID).findAll()
    var key_found    = logSheet.createTextFinder(key).findAll()  
    //Key and andrewid both exist somewhere in the sheet
    if(andrew_found && key_found){
      var andrew_rows = []
      for(var i = 0; i < andrew_found.length;i++){ 
        andrew_rows.push(andrew_found[i].getRow())
      } 
      var key_rows = [] 
      for(var j = 0; j < key_found.length; j++){
        key_rows.push(key_found[j].getRow())
      }
      //1.Now that the value is found, get the full range
      //  matching column value is found (andrewid and key are on the same column)  
      var found   = andrew_rows.find(a => key_rows.includes(a)) 
    }
  } 
  var fullRow = logSheet.getRange(found,1,1,logSheet.getLastColumn())
  var row1    = fullRow.getValues()[0]
  row1[1] = status
  row1[2] = approval 
  fullRow.setValues([row1])
}

/************************ Manipulating the unverifed sheet ****************/

/**
 * Take all new input values and add them to the unverifid sheet in 'Key Sheet Main;
 */
function unverifiedValueCollection(){
  var keySS    = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = keySS.getSheetByName('Log');
  var approvalAndAndrew = logSheet.getRange('C2:D').getValues();
  
  var logEntries = logToEntries();
  var allEntries = new Map();

  //Read spreasheets with data. Should be in 'Key Inputs' folder
  var inputFolder = null
  const folders = DriveApp.getFolders()
  while(folders.hasNext()){
    inputFolder = folders.next()
    var name = inputFolder.getName()
    if(name == "Key Inputs"){
      break;
    }
  }
  const inputFiles = inputFolder.getFilesByType(MimeType.GOOGLE_SHEETS)
  while(inputFiles.hasNext()){
    var file = inputFiles.next()
    allEntries = parseKeySheet(allEntries,file.getId()) 
  }
  allEntries = checkoutFormToEntries(allEntries);
  var unverifiedEntries = new Map();

  //if not in log or unverified in the log
  for(const [andrewID,keyRecord] of allEntries){
    //check if in log or if unverifed in log
    var arr = ["Unverified",andrewID]

    //In log sheet as unverified
    if(approvalAndAndrew.includes(arr)) {
      unverifiedEntries.set(andrewID,keyRecord);
    } 
    //Not in log (add to unverified and add to log)
    else if(!logEntries.has(andrewID)){
      unverifiedEntries.set(andrewID,keyRecord)
      addToLog(andrewID,keyRecord,logSheet,logEntries)
    }    
  }
  return unverifiedEntries;
}

/**
 * Create dropdown that determines status of unverified values
 */
function createApprovalDropdown() { 
  var keySS = SpreadsheetApp.getActiveSpreadsheet();
  var unverifiedSheet = keySS.getSheetByName('Unverified Input');
  var rule = SpreadsheetApp.newDataValidation()
              .requireValueInList(['Select','Approved', 'Denied'],true)
              .setHelpText("Select an option")
              .build();  

  var goalRange = unverifiedSheet.getRange(1,2,rowTotal,1); //check if i need to switch the 1st and 2nd value
  goalRange.setDataValidation(rule);

  for (var i = 0; i < rowTotal; i++) {
    var cell = unverifiedSheet.getRange(i + 2, 1); // A2, A3, ...
    cell.setValue('Select');
  }

  var currRules = unverifiedSheet.getConditionalFormatRules();
  var newRules = currRules.filter(function(rule) {
    var ranges = rule.getRanges();
    return !ranges.some(r => r.getA1Notation().startsWith("B"));
  });

  newRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Approved')
      .setBackground('#b7e1cd') // light green
      .setRanges([goalRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Denied')
      .setBackground('#e06666') // light red
      .setRanges([goalRange])
      .build()
  );
  unverifiedSheet.setConditionalFormatRules(newRules);
}

/**
 * Format unverified sheet to have input entries and the dropdown values
 */
function entryToUnverifiedInput(){
  var keySS           = SpreadsheetApp.getActiveSpreadsheet();
  var unverifiedSheet = keySS.getSheetByName('Unverified Input');
  var rule            = SpreadsheetApp.newDataValidation()
                          .requireValueInList(['Select','Approved', 'Denied'],true)
                          .setHelpText("Select an option")
                          .build();  

  var dropdownRange = unverifiedSheet.getRange('A2:A');
  dropdownRange.setDataValidation(rule);

  var unverifiedEntries = unverifiedValueCollection()
  unverifiedEntries.forEach((entryRecord) => {
    var keys = entryRecord.getKeys()
    for(i = 0; i < keys.length; i++){
      var key = keys[i];
      unverifiedSheet.appendRow([
        'Select',
        key.getId(),
        entryRecord.getAndrewID(),
        entryRecord.getLastName(),
        entryRecord.getFirstName(),
        entryRecord.getAdvisor(),
        entryRecord.getDepartment(),
        key.getKey(), 
        key.getRoom(),
        key.getExpirationDate(),
        key.getGivenDate()
      ])
    }
  });

  var currRules = unverifiedSheet.getConditionalFormatRules();
  var newRules = currRules.filter(function(rule) {
    var ranges = rule.getRanges();
    return !ranges.some(r => r.getA1Notation().startsWith("A"));
  });
  newRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Approved')
      .setBackground('#b7e1cd') // light green
      .setRanges([dropdownRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Denied')
      .setBackground('#e06666') // light red
      .setRanges([dropdownRange])
      .build()
  );
  unverifiedSheet.setConditionalFormatRules(newRules);
}

/**
 * If change in the unverified section, use id to update the log 
 * @param (*) value
 * @param {*} row 
 * @param {*} col 
 */
function submitUnverifedData(row,col,value){
  const unverfiedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Unverified Input')
  const id_loc             = unverfiedSheet.getRange(row,2)
  const id = id_loc.getValue()

  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log')
  var log_found  = logSheet.createTextFinder(id).findAll()[0]
  var log_row    = log_found.getRow()
  var logRange = logSheet.getRange("A2:L")
  
  switch(col){
    case 3:
      //andrewid --> also change 
      var keyNum = logRange.getCell(log_row-1,9).getValue()
      var new_id = hash_id(keyNum+ value) 
      logRange.getCell(log_row-1,4).setValue(value)      
      logRange.getCell(log_row-1,1).setValue(new_id)
      id_loc.setValue(new_id)
      break
    case 4:
      //last name
      logRange.getCell(log_row-1,5).setValue(value)
      break
    case 5:
      //first name
      logRange.getCell(log_row-1,6).setValue(value)
      break
    case 6:
      //advisor
      logRange.getCell(log_row-1,7).setValue(value)
      break
    case 7:
      //department
      logRange.getCell(log_row-1,8).setValue(value)
      break
    case 8:
      //key
      var old_andrew = logRange.getCell(log_row-1,4).getValue()
      var new_id = hash_id(value+old_andrew)
      logRange.getCell(log_row-1,9).setValue(value)
      logRange.getCell(log_row-1,1).setValue(new_id)
      id_loc.setValue(new_id)
      break
    case 9:
      //room
      logRange.getCell(log_row-1,10).setValue(value)
      break
    case 10:
      //expiration date
      logRange.getCell(log_row-1,11).setValue(value)
      break
    case 11:
      //given date
      logRange.getCell(log_row-1,12).setValue(value)
      break
  }
}

/**
 * Given what is in the approval tab, update what is in the unverifeid tab. Approve Selected - Button
 */
function submitSelectedData(){
  var keySS            = SpreadsheetApp.getActiveSpreadsheet();
  var unverifiedSheet  = keySS.getSheetByName('Unverified Input');
  var approveEntries   = new Map()
  var deletedEntires   = new Map()
  var remainingEntries = new Map()

  var val = true
  var i = 0
  //row,col,numrows,numcol
  var entry_raw = unverifiedSheet.getRange(2+i,1,1,12) //A2:L
  var entry     = entry_raw.getValues()[0]
  while(val){
    var approval  = entry[0] 
    var id        = entry[1]   
    var andrewID  = entry[2]
    var lastName  = entry[3]
    var firstName = entry[4]
    var advisor   = entry[5]
    var dept      = entry[6]
    var key       = entry[7]
    var room      = entry[8]
    var expDate   = entry[9]
    var givenDate = entry[10]

    //Add to note if there are invalid values (conjoin message values)
    var msg = ""
    var key_msg = ""
    var room_msg = ""
    var given_date_msg = ""
    var exp_date_msg  = ""

    //Keys
    if(key == 'invalid key' || key == ''){
      key_msg = "invalid key"
    }
    msg = msg + key_msg
    
    //Rooms
    if(room == 'invalid room' || room == ""){
      if(msg == ""){room_msg = "invalid room"}
      else{room_msg = ", "+"invalid room"}
    }
    msg = msg + room_msg
    
    //Given Date
    if(givenDate == "invalid date" || givenDate == ""){
      if(msg == ""){given_date_msg = "invalid given date"}
      else {given_date_msg = ", " + "invalid given date"}
    }
    msg = msg + given_date_msg

    //Expiration Date
    if(expDate == "invalid date" || expDate == ""){
      if(msg == ""){exp_date_msg = "invalid expiration date"}
      else{exp_date_msg = ", " + "invalid expiration date"}    
    }
    msg = msg + exp_date_msg
    var keyRec = new keyRecord(firstName,lastName,andrewID,advisor,dept,key,room,givenDate,expDate);
    if(msg == "" || approval == "Denied"){  
      //Add 'Approve' or 'Denied' to own set. ignore 'Selected'
      if(approval == "Approved"){
        if(approveEntries.has(andrewID)){
          var entry = approveEntries.get(andrewID)
          approveEntries.delete(andrewID)
          entry.addKey(key,room,givenDate,expDate)
          approveEntries.set(andrewID,entry)
        } else{
          approveEntries.set(andrewID,keyRec)
        }
        entry_raw.clear()
      } 
      else if(approval == "Denied"){
        if(deletedEntires.has(andrewID)){
          var entry = deletedEntires.get(andrewID)
          deletedEntires.delete(andrewID)
          entry.addKey(key,room,givenDate,expDate)
          deletedEntires.set(andrewID.entry)
        } else{
          deletedEntires.set(andrewID,keyRec)
        }
        entry_raw.clear()
      }
      else{
        if(remainingEntries.has(andrewID)){
          var entry = remainingEntries.get(andrewID)
          remainingEntries.delete((andrewID))
          entry.addKey(key,room,givenDate,expDate)
          remainingEntries.set(andrewID,entry)
        } else{
          remainingEntries.set(andrewID,keyRec)
        }
        entry_raw.clear()
      }
    } 
    else{
        if(remainingEntries.has(andrewID)){
          var entry = remainingEntries.get(andrewID)
          remainingEntries.delete((andrewID))
          entry.addKey(key,room,givenDate,expDate)
          remainingEntries.set(andrewID,entry)
        } else{
          remainingEntries.set(andrewID,keyRec)
        }      
        entry_raw.clear()
    }
    //Update loop conditions 
    i = i + 1
    entry_raw = unverifiedSheet.getRange(2+i,1,1,12)  
    entry     = entry_raw.getValues()[0]
    //check if next row has at least one value 
    val = !(entry.every(cell => (cell === "" || cell === null)))
  }

  //Update the log  
  approveEntries.forEach((entryRecord) => {
    var keys = entryRecord.key
    for(var i = 0; i < keys.length; i++){
      var key = keys[i]
      updateLogApproval(key.getId(),entryRecord.getAndrewID(),key.getKey(),"Approved","Active")
    }
  });

  deletedEntires.forEach((entryRecord) => {
    var keys = entryRecord.key
    for(var i = 0; i < keys.length; i++){
      var key = keys[i]
      updateLogApproval(key.getId(),entryRecord.getAndrewID(),key.getKey(),"Denied","Inactive")
    }
  });
  
  //first value is undefined
  remainingEntries.forEach((entryRecord) => {
    var keys = entryRecord.key
    for(var i = 0; i < keys.length; i++) {
      unverifiedSheet.appendRow([
        'Select',
        keys[i].getId(),
        entryRecord.getAndrewID(),
        entryRecord.getLastName(),
        entryRecord.getFirstName(),
        entryRecord.getAdvisor(),
        entryRecord.getDepartment(),
        keys[i].getKey(),
        keys[i].getRoom(),
        keys[i].getExpirationDate(),
        keys[i].getGivenDate(),
        msg
       ])
    }
  });
  analysis();
}

/**
 * Approve all unverified input, regardless of what is in the approval tab. Approve All - Button ///////////////////////////////////update once other function is complete 
 
function approveAllData(){
  //clear all the data in the unverifeid
  var keySS            = SpreadsheetApp.getActiveSpreadsheet();
  var unverfiedSheet   = keySS.getSheetByName('Unverified Input');
  var allEntries       = new Map(); 
  var remainingEntries = new Map()

  var val = true // this needs to be updated!!!!!!!!!!!!!!!!!1!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  var i = 0;
  var entry_raw = unverfiedSheet.getRange(2+i,1,1,12); //one row
  var entry     = entry_raw.getValues()[0] //check this!!!!! [0]  
  while(val){
    var approval  = entry[0]
    var id        = entry[1]
    var andrewID  = entry[2]
    var lastName  = entry[3]
    var firstName = entry[4]
    var advisor   = entry[5]
    var dept      = entry[6]
    var key       = entry[7]
    var room      = entry[8]
    var expDate   = entry[9]
    var givenDate = entry[10]
    
    var msg = ""
    var key_msg = ""
    var room_msg = ""
    var given_date_msg = ""
    var exp_date_msg  = ""
    //Add to note if there are invalid values (conjoin message values)

    //Keys
    if(key == 'invalid key' || key == ""){
      key_msg = "invalid key"
    }
    msg = msg + key_msg
    
    //Rooms
    if(room == 'invalid room' || room == ""){
      if(msg == ""){room_msg = "invalid room"}
      else{room_msg = ", "+"invalid room"}
    }
    msg = msg + room_msg
    
    //Given Date
    if(givenDate == "invalid date"||givenDate ==""){
      if(msg == ""){given_date_msg = "invalid date"}
      else {given_date_msg = ", " + "invalid date"}
    }
    msg = msg + given_date_msg

    //Expiration Date
    if(expDate == "invalid date" || expDate == ""){
      if(msg == ""){exp_date_msg = "invalid date"}
      else{exp_date_msg = ", " + "invalid date"}    
    }
    msg = msg + exp_date_msg

    var keyRec = new keyRecord(firstName,lastName,andrewID,advisor,dept,key,room,givenDate,expDate)
    if((key != 'invalid key') && (room != 'invalid room') && (expDate != 'invalid date') && (givenDate != 'invalid date')){
      //All 'Approve'. All can be added to the map for entries      
      if(allEntries.has(andrewID)){
        var entry = allEntries.get(andrewID)
        allEntries.delete(andrewID)
        entry.addKey(key,room,givenDate,expDate)
        allEntries.set(andrewID,entry)
      } else {
        allEntries.set(andrewID,keyRec)
      }

    }
    else {

      if(remainingEntries.has(andrewID)){
        var entry = remainingEntries.get(andrewID)
        remainingEntries.delete(andrewID)
        entry.addKey(key,room,givenDate,expDate)
        remainingEntries.set(andrewID,entry)
      } else{
        remainingEntries.set(andrewID,keyRec)
      }      
    }
    entry_raw.clear()
    //Uppdate loop conditions
    i = i + 1
    entry_raw = unverfiedSheet.getRange(2+i,1,1,12)
    entry     = entry_raw.getValues()[0]
    //Check if next row is empty
    val = !(entry.every(cell => (cell === "" || cell === null)))
  }
  //Update the log
  var logSheet       = keySS.getSheetByName('Log');
  var logEntries     = logSheet.getRange("A2:K").getValues();
  
  for(var i = 0; i < logEntries.length; i++){
    var entry_row  = logEntries[i]
    var andrewID1  = entry_row[3]
    var key1       = entry_row[8]
    var id1        = entry_row[0]

    var found_entry = allEntries.get(andrewID1) //undefined if not there
    if(found_entry != undefined){
      var keys = found_entry.key
      for(var i = 0; i  < keys.length; i++){
        var k1 = keys[i]
        if(k1.keyNumber == key1){
          updateLogApproval(id1,andrewID1,key1,"Approved","Active")
        }
      }
    }
  }
  remainingEntries.forEach((entryRecord) => {
    var keys = entryRecord.key
    for(var i = 0; i < keys.length; i++){
      unverfiedSheet.appendRow([
        'Select',
        keys[i].getId(),
        entryRecord.getAndrewID(),
        entryRecord.getLastName(),
        entryRecord.getFirstName(),
        entryRecord.getAdvisor(),
        entryRecord.getDepartment(),
        keys[i].getKey(),
        keys[i].getRoom(),
        keys[i].getExpirationDate(),
        keys[i].getGivenDate(),
        msg
      ])
    }
  });
}
*/



/************ Check in values in the sheets********/
/**
 * 
 * @param {*} allEntries 
 * @param {*} andrewID 
 * @param {*} firstName 
 * @param {*} lastName 
 * @param {*} advisor 
 * @param {*} key 
 * @param {*} room 
 * @returns 
 */
function manualCheckIn(andrewID,firstName,lastName,advisor,key,room){
  var logEntries = logToEntries()
  if(logEntries.has(andrewID)){      
    var entry     = logEntries.get(andrewID)
    //var confirmed = confirmUser(firstName,lastName,advisor,andrewID,key,room,entry) 
    var id = -1
    var key_count = entry.getKeys().length
    //1.Remove specific key
    if(key_count == 1){
      id = entry.key[0].id
    } else{
      entry.key.forEach((keyDetails) => {
        if(keyDetails.getKey() == key){
          id = keyDetails.id
        } 
      });
    }
    //2.Update the log to show key has been removed
    updateLogApproval(id,andrewID,key,"Approved","Inactive")
    analysis()
  }
}

/**
 * Return a key and rerun analysis(populate sheets and main sheet)
 * @param {*} allEntries 
 * @returns 
 */
function checkInForm(){
  var firstName,lastName, advisor,andrewID, key,room;
  const checkInForm = FormApp.openByUrl("https://docs.google.com/forms/d/1t6IxYbw-evVopJd3XGKHRxb9HfWWke0ozHA39XT-1z8/")
  var allResponses = checkInForm.getResponses()
 
  for(const resp of allResponses) {
    //All the questions and response stores in an item
    for(item of resp.getItemResponses()){
      ques = item.getItem().getTitle();
      answ = item.getResponse();
      if(ques == "First Name:"){
        firstName = answ;
      } else if(ques == "Last Name:") {
        lastName = answ;
      } else if(ques == "Advisor:") {
        advisor = answ;
      } else if(ques == "andrewID:") {
        andrewID = answ;
      } else if(ques == "Key Number:") {
        key = validKey(answ);
      } else if(ques == "Room (Include Building and Room Number) Ex: DH 3213A") {
        room = validRoom(answ);
      } 
    }
    manualCheckIn(andrewID,firstName,lastName,advisor,key,room)
  }
}






function scheduleReload(){
////////////////////////////

//Check status ofvalues.remove them if checked in. Email or add to list if expired/near expiration (notification to return the values)
}





/**
 * Fill the sheets that represent each expiration sheet 
 * @param {*} dataSS - the active spreadsheet
 * @returns - AllEntries 
 */
function fillSheets(dataSS){
  const allSheets      = dataSS.getSheets()
  const template_sheet = allSheets[allSheets.length - 1] //Template is always the last sheet
  var allEntries       = verifiedEntries(dataSS)  

  //Delete all the previous year sheets
  allSheets.forEach((sheet) => {
    if((sheet.getSheetName() != "Main") 
      && (sheet.getSheetName() != "Template") 
      && (sheet.getSheetName() != "Unverified Input")
      && (sheet.getSheetName() != "Log")  
      && (sheet.getSheetName() != "Key Check-In Form") 
      && (sheet.getSheetName() != "Key Check-Out Form")){
      dataSS.deleteSheet(sheet)
    }
  });

  var today = new Date()
  today.setHours(0,0,0,0)

  //Get the years from all the entries (map) in an array
  const years = []//"Expired"]
  allEntries.forEach((entryRecord) => {
    var keys = entryRecord.getKeys()
    for(i = 0; i < keys.length; i++){
      var key = keys[i]
      var date = new Date(key.getExpirationDate())
      var entry_yr = date.getFullYear()

      //Year not added and current day is after today
      if(!years.includes(entry_yr) && date > today){
        years.push(entry_yr)
      }
    }
  });

  years.sort()
  years.push("Expired")
  years.reverse()//sort years array in descending order
  
  //Create sheets with the given years
  for(i = 0; i < years.length; i++ ){
    //Create new sheet
    if (i == 0)
      {var new_sheet = dataSS.insertSheet((years[i]), i+1, {template: template_sheet})}
    else
      {var new_sheet = dataSS.insertSheet((years[i]).toFixed(0), i+1, {template: template_sheet})}
    //Name the new sheet
    new_sheet.getRange("A1").setValue((`Expiration: ${years[i]} `))
  }
  //Add entry to the different sheets
  allEntries.forEach((entryRecord) => {
    var keys = entryRecord.getKeys()
    for(i = 0; i < keys.length; i++){
      var key = keys[i]
      var date = new Date(key.getExpirationDate())
      if( date <= today){
        var yr1 = "Expired"
      } else {
        var yr1  = date.getFullYear()
      }
      var new_sheet = dataSS.getSheetByName(yr1)
      new_sheet.appendRow([
        key.getExpirationDate(),entryRecord.getAndrewID(),
        entryRecord.getLastName(),entryRecord.getFirstName(),
        entryRecord.getAdvisor(),entryRecord.getDepartment(),
        key.getKey(),key.getRoom(),key.getGivenDate()
      ])
    }
  });
  return allEntries
}

/**
 *  Add a time function  to regularily refresh this function!!!!!
 */
function analysis(){
  const dataSS    = SpreadsheetApp.getActiveSpreadsheet()
  var allEntries = fillSheets(dataSS)

  var currDate = new Date()

  var oneDate  = new Date(currDate)
  oneDate.setMonth(oneDate.getMonth() + 1) 

  var weekDate = new Date(currDate)
  weekDate.setDate(weekDate.getDate() + 7)

  var dayDate  = new Date(currDate)
  dayDate.setDate(dayDate.getDate() + 1)

  var andrew_day   = []
  var andrew_week  = []
  var andrew_one   = []
  var expired_list = []

  allEntries.forEach((entryRecord) => {
    var keys = entryRecord.getKeys()
    for(i = 0; i < keys.length; i++){
      var key = keys[i]
      var expiration = new Date(key.expDate)

      if(isDateInFrame(currDate,dayDate,expiration)){
        andrew_day.push(entryRecord)
      } else if(isDateInFrame(currDate,weekDate,expiration)){
        andrew_week.push(entryRecord)
      } else if(isDateInFrame(currDate,oneDate,expiration)){
        andrew_one.push(entryRecord)
      } else if(isExpired(currDate,expiration)){
        expired_list.push(entryRecord)
      } 
    }  
  });
  const sheets    = dataSS.getSheets()
  const mainSheet = sheets[0]

  //One month
  mainSheet.getRange("B8:B").clearContent()
  mainSheet.getRange("C8:C").clearContent()
  var one        = mainSheet.getRange("C8:C")
  var one_values = one.getValues()
  mainSheet.getRange(7,3).setValue('1 Month')
  for(var i = 0; i < one_values.length; i++){
    if(i < andrew_one.length){
      var keys = andrew_one[i].getKeys()
      for(var j = 0; j < keys.length; j++){
        mainSheet.getRange(8+i,2).setValue(keys[i].getId())
        mainSheet.getRange(8+i,3).setValue(andrew_one[i].getAndrewID())
      } 
    } else{break;}

  }

  //1 week
  mainSheet.getRange("D8:D").clearContent()
  mainSheet.getRange("E8:E").clearContent()
  var week        = mainSheet.getRange("E8:E")
  var week_values = week.getValues()
  mainSheet.getRange(7,5).setValue('1 Week')
  for(var i = 0; i < week_values.length; i++){
    if(i < andrew_week.length){
      var keys = andrew_week[i].getKeys()
      for(var j = 0; j < keys.length; j++){
        mainSheet.getRange(8+i,4).setValue(keys[i].getId())
        mainSheet.getRange(8+i,5).setValue(andrew_week[i].getAndrewID())
      } 
    } else {break;}

  }

  //1 day
  mainSheet.getRange("F8:F").clearContent()
  mainSheet.getRange("G8:G").clearContent()
  var day        = mainSheet.getRange("G8:G")
  var day_values = day.getValues()
  mainSheet.getRange(7,7).setValue('1 Day')
  for(var i = 0; i < day_values.length; i++){
    if(i <andrew_day.length){    
      var keys = andrew_day[i].getKeys()
      for(var j = 0; j < keys.length; j++){
        mainSheet.getRange(8+i,6).setValue(keys[i].getId())
        mainSheet.getRange(8+i,7).setValue(andrew_day[i].getAndrewID())
      }      
    } else {break;}

 }

  //Expired
  mainSheet.getRange("H8:H").clearContent()
  mainSheet.getRange("I8:I").clearContent()
  var expired        = mainSheet.getRange("I8:I")
  var expired_values = expired.getValues()
  mainSheet.getRange(7,9).setValue('Expired')
  for(var i = 0; i < expired_values.length; i++){
    if(i < expired_list.length){    
      var keys = expired_list[i].getKeys()
      for(var j = 0; j < keys.length; j++){
        mainSheet.getRange(8+i,8).setValue(keys[i].getId())
        mainSheet.getRange(8+i,9).setValue(expired_list[i].getAndrewID())
      }      
    } else {break;}

  }
}

function expiration_month(){
  const dataSS    = SpreadsheetApp.getActiveSpreadsheet()
  const mainSheet = dataSS.getSheetByName("Main")
  const folder    = DriveApp.getFoldersByName("Keys Project").next() //original is a iterator. need next
  const files     = folder.getFiles()
  var allEntries  = verifiedEntries(dataSS)  

  var index = 0

  var andrew_one      = []
  var id_list         = []
  var andrew_one_temp = mainSheet.getRange("C8:C").getValues()
  var id_list_temp    = mainSheet.getRange("B8:B").getValues()
  var one             = andrew_one_temp[index][0]
  var id              = id_list_temp[index][0]

  while(one != ""){
    var one_value = allEntries.get(one) 
    var keys      = one_value.key
    for(var i = 0; i < keys.length; i++){
      andrew_one.push(one_value)      
      id_list.push(keys[i].getId())
    }

    index = index + 1
    one = andrew_one_temp[index][0]
    id  = id_list_temp[index][0]
  }

  while(files.hasNext()){
    var file = files.next()
    if ((file.getMimeType() === MimeType.GOOGLE_DOCS) && (file.getName() == "Month Till Expiration")){
      var doc_month = DocumentApp.openById(file.getId())
      expire_msg(id_list,andrew_one,doc_month,file.getName()) 
    }
  }
}

function expiration_week(){
  const dataSS    = SpreadsheetApp.getActiveSpreadsheet()
  const mainSheet = dataSS.getSheetByName("Main")
  const folder    = DriveApp.getFoldersByName("Keys Project").next()
  const files     = folder.getFiles()
  var allEntries  = verifiedEntries(dataSS)  
  
  var index = 0

  var andrew_week       = []
  var id_list           = []
  var andrew_week_temp  = mainSheet.getRange("E8:E").getValues()
  var id_list_temp      = mainSheet.getRange("D8:D").getValues()
  var week              = andrew_week_temp[index][0]
  var id                = id_list_temp[index][0]

  while(week != ""){
    var week_value = allEntries.get(week)
    var keys       = week_value.key
    for(var i = 0; i < keys.length; i++){
      andrew_week.push(week_value)
      id_list.push(keys[i].getId())
    }

    index = index + 1
    week = andrew_week_temp[index][0]
    id   = id_list_temp[index][0]
  }

  while(files.hasNext()){
    var file = files.next()
    if ((file.getMimeType() === MimeType.GOOGLE_DOCS) && (file.getName() == "Week Till Expiration")){
      var doc = DocumentApp.openById(file.getId())
      expire_msg(id_list,andrew_week,doc,file.getName()) 
    }
  }
}

function expiration_day(){
  const dataSS    = SpreadsheetApp.getActiveSpreadsheet()
  const mainSheet = dataSS.getSheetByName("Main")
  const folder    = DriveApp.getFoldersByName("Keys Project").next()
  const files     = folder.getFiles()
  var allEntries  = verifiedEntries(dataSS) 

  var index = 0

  var andrew_day      = []
  var id_list         = []
  var andrew_day_temp = mainSheet.getRange("G8:G").getValues()
  var id_list_temp    = mainSheet.getRange("F8:F").getValues()
  var day             = andrew_day_temp[index][0]
  var id              = id_list_temp[index][0]

  while(day != ""){
    var day_value = allEntries.get(day)
    var keys      = day_value.key
    for(var i = 0; i < keys.length; i++){
      andrew_day.push(day_value)
      id_list.push(keys[i].getId())
    }

    index = index + 1
    day   = andrew_day_temp[index][0]
    id    = id_list_temp[index][0]
  }

  while(files.hasNext()){
    var file = files.next()
    if ((file.getMimeType() === MimeType.GOOGLE_DOCS) && (file.getName() == "Day Till Expiration")){
      var doc = DocumentApp.openById(file.getId())
      expire_msg(id_list,andrew_day,doc,file.getName()) 
    }
  }
}

function expiration_exp_11(){
  const dataSS    = SpreadsheetApp.getActiveSpreadsheet()
  const mainSheet = dataSS.getSheetByName("Main")
  const folder    = DriveApp.getFoldersByName("Keys Project").next() //original is a iterator. need next
  const files     = folder.getFiles()
  var allEntries = verifiedEntries(dataSS) 

  var index = 0
  var expired_list      = []
  var expired_list_temp = mainSheet.getRange("I8:I").getValues()  
  var exp               = expired_list_temp[index][0]
  
  while(exp != ""){
    var exp_value = allEntries.get(exp)
    expired_list.push(exp_value)
    index = index + 1
    exp = expired_list_temp[index][0]
  }

  while(files.hasNext()){
    var file = files.next()
    if ((file.getMimeType() === MimeType.GOOGLE_DOCS) && (file.getName() == "Expired")){
      var doc = DocumentApp.openById(file.getId())
      expire_msg(-1,expired_list,doc,file.getName()) 
    }
  }
}

function expiration_exp(){
  const dataSS    = SpreadsheetApp.getActiveSpreadsheet()
  const mainSheet = dataSS.getSheetByName("Main")
  const folder    = DriveApp.getFoldersByName("Keys Project").next() //original is a iterator. need next
  const files     = folder.getFiles()
  var allEntries = verifiedEntries(dataSS) 

  var index = 0
  
  var expired_list      = []
  var id_list           = []
  var expired_list_temp = mainSheet.getRange("I8:I").getValues()  
  var id_list_temp      = mainSheet.getRange("H8:H").getValues()
  var exp               = expired_list_temp[index][0]
  var id                = id_list_temp[index][0]
  
  while(exp != ""){
    var exp_value = allEntries.get(exp)
    var keys = exp_value.key
    for(var i = 0; i < keys.length; i++){
      expired_list.push(exp_value)
      id_list.push(keys[i].getId())
    }
    index = index + 1
    exp = expired_list_temp[index][0]
    id  = id_list_temp[index][0]
  }

  while(files.hasNext()){
    var file = files.next()
    if ((file.getMimeType() === MimeType.GOOGLE_DOCS) && (file.getName() == "Expired")){
      var doc = DocumentApp.openById(file.getId())
      expire_msg(id_list,expired_list,doc,file.getName()) 
    }
  }
}

/**
 * Send emails to individuals in the expiration range using templates in 
 * the Keys Project folder 
 
function expiration_check(){
  const dataSS    = SpreadsheetApp.getActiveSpreadsheet()
  const mainSheet = dataSS.getSheetByName("Main")
  const folder    = DriveApp.getFoldersByName("Keys Project").next() //original is a iterator. need next
  const files     = folder.getFiles()
  
  var allEntries = verifiedEntries(dataSS)  

  var index = 0
  var andrew_one          = []
  var andrew_one_temp = mainSheet.getRange("B8:B").getValues()
  var one             = andrew_one_temp[index][0]
  while(one != ""){
    var one_value = allEntries.get(one) 
    andrew_one.push(one_value)
    index = index + 1
    one = andrew_one_temp[index][0]
  }

  index = 0
  var andrew_week           = []
  var andrew_week_temp  = mainSheet.getRange("C8:C").getValues()
  var week              = andrew_week_temp[index][0]
  while(week != ""){
    var week_value = allEntries.get(week)
    andrew_week.push(week_value)
    index = index + 1
    week = andrew_week_temp[index][0]
  }

  index = 0
  var andrew_day          = []
  var andrew_day_temp = mainSheet.getRange("D8:D").getValues()
  var day             = andrew_day_temp[index][0]
  while(day != ""){
    var day_value = allEntries.get(one)
    andrew_day.push(day_value)
    index = index + 1
    day = andrew_day_temp[index][0]
  }
  index = 0
  var expired_list          = []
  var expired_list_temp = mainSheet.getRange("E8:E").getValues()  
  var exp               = expired_list_temp[index][0]
  while(exp != ""){
    var exp_value = allEntries.get(exp)
    expired_list.push(exp_value)
    index = index + 1
    exp = expired_list_temp[index][0]
  }
  
  while(files.hasNext()){
    var file = files.next()
    if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
      switch (file.getName()){
        case "Month Till Expiration":
          var doc_month = DocumentApp.openById(file.getId())
          expire_msg(-1,andrew_one,doc_month,file.getName()) 
          break
        case "Week Till Expiration":
          var doc_week = DocumentApp.openById(file.getId())
          expire_msg(-1,andrew_week,doc_week,file.getName()) 
          break
        case "Day Till Expiration":
          var doc_day  = DocumentApp.openById(file.getId())
          expire_msg(-1,andrew_day,doc_day,file.getName()) 
          break
        case "Expired":
          var doc_exp  = DocumentApp.openById(file.getId())
          expire_msg(-1,expired_list,doc_exp,file.getName()) 
          break
      }
    }
  }
}*/

/**
 * Send emails to all people in an expiration range list using the doc as the email body
 * and the subj as the subjected of the email
 * 
 * @param {*} list - Speciic expiration range list 
 * @param {*} doc  - Google Doc that has the template for the email sent to people
 *                   in expiration range
 * @param {*} subj - Subject of the email to the people in the expiration range list
 */
function expire_msg(id,list,doc,subj){
  var doc_string = doc.getBody().getText()

  //for(var entry_record of list){
  for(var j =0; j < list.length; j++){
    var entry_record = list[j]
    var keys = entry_record.getKeys()
    for(var i = 0; i < keys.length; i++){
      if(id[j] == keys[i].getId()){
        var recipient       = entry_record.getAndrewID() + "@andrew.cmu.edu"
        var doc_string_name = doc_string.replace("[First]",entry_record.getFirstName()) 
                                        .replace("[Last]",entry_record.getLastName())
                                        .replace("[keyNumber]",keys[i].getKey())      
                                        .replace("[Room]",keys[i].getRoom())
                                        .replace("[expDate]",keys[i].getExpirationDate())     
        MailApp.sendEmail(recipient,subj,doc_string_name)                                         
      }
    }

  }
}

function isDateInFrame(start, end,date){
  if(date == null || date == undefined) return false
  return start.getTime() <= date.getTime() 
      && date.getTime()  <= end.getTime()
}

function isExpired(curr,date){
  if(date == null || date == undefined) return false
  return curr.getTime() > date.getTime()
}


