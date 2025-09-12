/*
Object Oriented Design
Goals:
-Alumni/Peoples who have reached expiration date
-Renew key?
-Edit currentFormToClass to deal with wrong dates
-Edit currentFormToClass to deal with no advisor
-1st Year PhD  advisor is Heather
*/

var E = null;
var log;
var currentRecord;

class keyInfo{
  constructor(keyNumber,roomNumber,givenDate,expDate){
    this.keyNumber= keyNumber //A string due to the dash in the middle
    this.roomNumber = roomNumber //String to designate building name
    this.givenDate = givenDate;
    this.expDate = expDate;
    this.status = true
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

  //////////////////////////////////////////////
  deactivate(){
    this.status = false
  }
  activate(){
    this.status = true
  }
  
  active(){
    var curr = new Date()
    if((this.givenDate instanceof Date)&& (this.givenDate < curr)){
      return false
    } else {
      return true
    }
  }
  expired(){
    return !this.active()
  }
}

class keyRecord {
  constructor(first,last,andrewID,advisor,dept,key,room,givenDate,expDate)  {
    this.firstName = first;
    this.lastName  = last;
    this.andrewID  = andrewID;
    this.advisor   = advisor;
    this.dept      = dept;
    this.key       = [new keyInfo(key,room,givenDate,expDate)];
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
    if(this.andrewID == ""){
      return this.firstName +"_"+ this.lastName+ "_no_andrew_id"
    }
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
  /////////////////////////////////////
  setKey(newKeySet){
    this.key = newKeySet
  }
  getActiveKeys(){
    const allKeys = this.key
    var activeKeys
    allKeys.forEach((key)=>{
      if(key.active()) activeKeys.push(key)
    });
    return activeKeys
  }
  getInactiveKeys(){
    const allKeys = this.key
    var inactiveKeys
    allKeys.forEach((key) => {
      if(key.expired()) inactiveKeys.push(key)
    });
    return inactiveKeys
  }
  addKey(key,room,givenDate,expDate) {
    var newKey = new keyInfo(key,room,givenDate,expDate)
    var keys = this.key
    keys.push(newKey)
    this.key = keys
    //this.key.push(newKey)//make sure it is an array
  }
  removeKey(remKey){
    //??Can one key open multiple rooms???
    this.key.forEach((keyInfo) =>{
      if(keyInfo.getKey() == remKey){
        keyInfo.deactivate()
      }
    });
  }
}

function activeEntries(){
  const keySS    = SpreadsheetApp.getActiveSpreadsheet()
  const logSheet = keySS.getSheetByName('Log')
  const range    = logSheet.getRange(2,1,logSheet.getlastRow(),logSheet.getLastColumn()) 
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

function verifiedEntries(){
  const keySS    = SpreadsheetApp.getActiveSpreadsheet()
  const logSheet = keySS.getSheetByName('Log')
  const range    = logSheet.getRange(2,1,logSheet.getlastRow(),logSheet.getLastColumn()) 
  const log_values = range.getValues()

  var verifiedEntries = new Map()
  for(var log_row in log_values){
    const status    = log_row[0]
    const approval  = log_row[1]
    const andrewID  = log_row[2]
    const lastName  = log_row[3]
    const firstName = log_row[4]
    const advisor   = log_row[5]
    const dept      = log_row[6]
    const key       = log_row[7]
    const room      = log_row[8]
    const expDate   = log_row[9]
    const givenDate = log_row[10]

    if((status == "Active") && (approval == "Approval")){
      var newKeyRec = new keyRecord(firstName,lastName,andrewID,advisor,dept,key,room,givenDate,expDate)
      verifiedEntries.set(andrewID,newKeyRec)
    }
  }
  return verifiedEntries

}

/**
function activeEntries(entries){
  var active = new Map()
  Object.entries(entries).forEach(([andrewID,key]) => {
    //Create new key record w/ only inactive keys
    var newKeyRec = keyRecord(key.getFirstName(),key.getLastName(),andrewID,null,null,null,null)
    newKeyRec.setKey(key.getActiveKeys)
    //Set to active dictionary
    active.set(andrewID,newKeyRec) //change key record valyues
  });
  return active
}
function inactiveEntries(entries) {
  var inactive = new Map()
  Object.entries(entries).forEach(([andrewID,key]) => {
    //Create new key record w/ only active keys
    var newKeyRec = keyRecord(key.getFirstName(),key.getLastName(),andrewID,null,null,null,null)
    newKeyRec.setKey(key.getInactiveKeys)
    //Set to inactive dictionary
    inactive.set(andrewID,newKeyRec)
  });
  return inactive
}
*/

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////Helper Functions used for safety checks
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

//Used in checkInForm <--- check if working properly
function confirmUser(first,last,advisor,andrew,key,room,entry){
  const key_rooms = entry.getKeys() //???????????????
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
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////Parsing entry data
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
        entry.addKey(keys,room,advisor);
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
        Entry.addKey(key,room,advisor);
        allEntries.set(andrewID,Entry);
    }
  }
  return allEntries
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////Manipulating the log sheet
/**
 * Read log data and turn them into entries
 */
function logToEntries(){
  var keySS = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = keySS.getSheetByName('Log');
  var allEntries = new Map();

  var logRange = logSheet.getRange('A:J');
  var logValues = logRange.getValues();

  for(var i = 1; i < logValues.length; i++){
    var row = logValues[i];
    if(row.length == 0 || row.length < 11 || row[0] == "Inactive") break;
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
 * !!!!CHange so it can distingush active and non-active entries
 * 
 * 
 * Initially adds a value to the log if it is not already in the log (should be right after initially parsed)
 */
function addAllToLog(){
  var keySS    = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = keySS.getSheetByName('Log');
  
  var logEntries = logToEntries();
  var allEntries = new Map();
  allEntries = checkoutFormToEntries(allEntries);
  allEntries = checkInForm(allEntries);

  for(const [andrewID, keyRecord] of allEntries){
    addToLog(andrewID,keyRecord,logSheet,logEntries,"Active")
  }
}

function test_update_log(){
  updateLog("bnyaga","4501-000","Approved")
}

/**
 * Update approval status of a log (based on what happens in the unverifed sheet)
 */
function updateLog(andrewID,key,approval){
  const keySS    = SpreadsheetApp.getActiveSpreadsheet()
  const logSheet = keySS.getSheetByName('Log')

  //Find all instances of the andrewID and the key value in the spreadsheet
  var andrew_found = logSheet.createTextFinder(andrewID).findAll()
  var key_found    = logSheet.createTextFinder(key).findAll()  
  //Key and andrewid both exist somewhere in the sheet
  if(andrew_found && key_found){
    var andrew_rows = []
    for(var i = 0; i < andrew_found.length;i++){ //?????????????????????????
      andrew_rows.push(andrew_found[i].getRow())
    } 
    var key_rows = [] 
    for(var j = 0; j < key_found.length; j++){
      key_rows.push(key_found[j].getRow())
    }
    //1.Now that the value is found, get the full range
    //  matching column value is found (andrewid and key are on the same column)  
    var found   = andrew_rows.find(a => key_rows.includes(a)) 
    //fullRow = location of 'found' column
    var fullRow = logSheet.getRange(found,1,1,logSheet.getLastColumn())
    var row1    = fullRow.getValues()[0]
    //2.replace the approval values I was looking for
    row1[1] = approval
    fullRow.setValues([row1]) //debug these values
  } 
}
// function searchLog(){
//   var keySS = SpreadsheetApp.getActiveSpreadsheet();
//   var logSheet = keySS.getSheetName('Log');
//   var andrewIDList = logSheet.getRange('A2:A').getValues();  
// }
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////Manipulating the unverifed sheet

/**
 * Take all new input values and add them to the unverifid sheet in 'Key Sheet Main;
 */
function unverifiedValueCollection(){
  var keySS    = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = keySS.getSheetByName('Log');
  var approvalAndAndrew = logSheet.getRange('B2:C').getValues();
  
  var logEntries = logToEntries() //log values to entry
  //Need to change to active entries //////////////////////////////////////////////
  var allEntries = new Map();

  //Read spreasheets with data. Should be in 'Key Inputs' folder
  var inputFolder = null
  const folders = DriverApp.getFolders()
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
  allEntries = checkInForm(allEntries);         //////////////////////////////////////////////is this necessary??


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

  //var goalRange = unverifiedSheet.getRange('A2:A')
  // //unverifiedSheet.getRange(1,2,rowTotal,1); //check if i need to switch the 1st and 2nd value
  // goalRange.setDataValidation(rule);
  // var allEntries = new Map();
  // allEntries = checkoutFormToEntries(allEntries);
  // allEntries = checkInForm(allEntries);

  var unverifiedEntries = unverifiedValueCollection()
  unverifiedEntries.forEach((entryRecord) => {
    var keys = entryRecord.getKeys()
    for(i = 0; i < keys.length; i++){
      var key = keys[i];
      //var date = new Date(key.getExpirationDate());
      unverifiedSheet.appendRow([
        'Select',
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
 * Given what is in the approval tab, update what is in the unverifeid tab. Approve Selected - Button
 * 
 * 
 * Check i values are valid
 * 
 */
function submitSelectedData(){
  var keySS           = SpreadsheetApp.getActiveSpreadsheet();
  var unverifiedSheet = keySS.getSheetByName('Unverified Input');
  var allEntries      = new Map()
  var deletedEntires  = new Map()

  var val = true
  var i = 0
  var entry_raw = unverifiedSheet.getRange(2+i,1,1,10)  //One row
  var entry     = entry_raw.getValues()[0] //check if  [0] is necessary
  while(val){
    var approval  = entry[0]    
    var andrewID  = entry[1]
    var lastName  = entry[2]
    var firstName = entry[3]
    var advisor   = entry[4]
    var dept      = entry[5]
    var key       = entry[6]
    var room      = entry[7]
    var expDate   = entry[8]
    var givenDate = entry[9]
    //change the color for the values!!!!!!!!!!!!!!!!!!

    //Only add values if they are no invalid values
    if((key != 'invalid key') && (room != 'invalid room') && (givenDate != "invalid date") && (expDate != "invalid date")){
      var keyRec = new keyRecord(firstName,lastName,andrewID,advisor,dept,key,room,givenDate,expDate);

      //Add 'Approve' or 'Denied' to own set. ignore 'Selected'
      if(approval == "Approved"){
        allEntries.set(andrewID,keyRec)
        entry_raw.clear()
      } 
      else if(approval == "Denied"){
        deletedEntires.set(andrewID,keyRec)
        entry_raw.clear()
      }

    }
    //Update loop conditions 
    i = i + 1
    entry_raw = unverifiedSheet.getRange(2+i,1,1,10)  
    entry     = entry_raw.getValues()[0]
    //check if next row is empty
    val = entry.every(cell => (cell != "" && cell != null))
  }
  //Update the log
  var logSheet = keySS.getSheetByName('Log')
  var logEntries_raw = logSheet.getRange("A2:K");
  var logEntries = logEntries_raw.getValues()

  for(var i = 0; i < logEntries.length; i++){
    var entry_row = logEntries[i]
    var andrewID1 = entry_row[2]
    var key1      = entry_row[7]

    //For all log values, check if it matches value in allEntries (approved entries)
    var found_entry = allEntries.get(andrewID1)
    if(found_entry != undefined){
      var keys = found_entry.key
      for(var i = 0; i < keys.length; i++){
        k = keys[i]
        keyNum = k.keyNumber
        //If andrew ID(above) and key match, say it is approved in log
        if(keyNum == key1){
          updateLog(andrewID1,key1,"Approved")
        }
      }
    }

    //For all log vales, check if it matches value in deletedEntries (deleted entries)
    var found_entry1 = deletedEntires.get(andrewID1)
    if(found_entry1 != undefined){
      var keys1 = found_entry1.key
      for(var j = 0; j < keys1.length; j++){
        k1 = keys1[j]
        keyNum1 = k1.keyNumber
        //If andrewID(above) and key match, say it is denied in log
        if(keyNum1 == key1){
          updateLog(andrewID1,key1,"Denied")
        }
      }
    }
  }
  //return the entries value. call this in analysis
  return allEntries ///////////////////////////////may not need to return a value since log updates
}

/**
 * Approve all unverified input, regardless of what is in the approval tab. Approve All - Button
 */
function approveAllData(){
  //clear all the data in the unverifeid
  var keySS          = SpreadsheetApp.getActiveSpreadsheet();
  var unverfiedSheet = keySS.getSheetByName('Unverified Input');
  var allEntries     = new Map(); 

  var val = true // this needs to be updated!!!!!!!!!!!!!!!!!1!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!11
  var i = 0;
  var entry_raw = unverfiedSheet.getRange(2+i,1,1,10); //one row
  var entry = entry_raw.getValues()[0] //check this!!!!! [0]  
  while(val){
    //var approval  = entry[0]
    var andrewID  = entry[1]
    var lastName  = entry[2]
    var firstName = entry[3]
    var advisor   = entry[4]
    var dept      = entry[5]
    var key       = entry[6]
    var room      = entry[7]
    var expDate   = entry[8]
    var givenDate = entry[9]

    if((key != 'invalid key') && (room != 'invalid room') && (expDate != 'invalid date') && (givenDate != 'invalid date')){
      var keyRec = new keyRecord(firstName,lastName,andrewID,advisor,dept,key,room,givenDate,expDate)

      //All 'Approve'. All can be added to the map for entries
      allEntries.set(andrewID,keyRec)
      entry_raw.clear()
    }
    //Uppdate loop conditions
    i = i + 1
    entry_raw = unverfiedSheet.getRange(2+i,1,1,10)
    entry     = entry_raw.getValues()[0]
    //Check if next row is empty
    val = entry.every(cell => (cell != "" && cell != null))
  }
  //Update the log
  var logSheet       = keySS.getSheetByName('Log');
  var logEntries_raw = logSheet.getRange("A2:K");
  var logEntries     = logEntries_raw.getValues();
  
  for(var i = 0; i < logEntries.length; i++){
    var entry_row  = logEntries[i]
    var andrewID1  = entry_row[2]
    var key1       = entry_row[7]
    
    var found_entry = allEntries.get(andrewID1) //undefined if not there
    if(found_entry != undefined){
      var keys = found_entry.key
      for(var i = 0; i  < keys.length; i++){
        k = keys[i]
        keyNum = k.keyNumber
        if(keyNum == key1){
          updateLog(andrewID1,key1,"Approved")
        }
      }
    }
  }
  return allEntries ////////////////////////may not need to return a value since log updates
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////Check in values in the sheets 


function manualCheckIn(allEntries,andrewID,firstName,lastName,advisor,key,room){
    if(allEntries.has(andrewID)){      
      var entry     = allEntries.get(andrewID)
      var confirmed = confirmUser(firstName,lastName,advisor,andrewID,key,room,entry) 
      if(confirmed){
        var key_count = entry.getKeys().length
        //1.Remove specific key
        if(key_count == 1){allEntries.delete(andrewID)}   
        else{
          keys = []
          entry.key.forEach((keyDetails) => {
            if(keyDetails.getKey() == key){
              keyDetails.deactivate()       //This does not seem to work. deactiveate isnt doing anything but sorting to active and non active fkeys <--may need to remove this
            } else{
              keys.push(keyDetails)
            }
          });
          entry.setKey(keys)
          allEntries.set(andrewID,entry)   
        }
        //2.Update the log to show key has been removed
        updateLog(andrewID,key,"Inactive")
      }
    }
  return allEntries
}

function checkInForm(allEntries){
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
    allEntries = manualCheckIn(allEntries,andrewID,firstName,lastName,advisor,key,room)
  }
  return allEntries
}

function scheduleReload(){
////////////////////////////

//Check status ofvalues.remove them if checked in. Email or add to list if expired/near expiration (notification to return the values)
}

function currentKeys(){
  ///////////////////////////////////////
}


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function fillSheets(allEntries){
  /** 
  //Input SS Files
  var   inputFolder = null
  const folders = DriveApp.getFolders()
  while (folders.hasNext()){
    inputFolder = folders.next()
    var name   = inputFolder.getName()
    if(name == "Key Inputs"){
      break; //leave the loop when folder is found
    }
  }
  const inputFiles = inputFolder.getFilesByType(MimeType.GOOGLE_SHEETS)
  while(inputFiles.hasNext()){
    var file = inputFiles.next()
    var id = file.getId() 
    var file_name = file.getName() //////////////////
    //var allEntries = parseKeySheet(allEntries,id)
  }

  //INSERT !VERIFIED! ENTRIES HERE!!!!
  // allEntries = checkoutFormToEntries(allEntries) //??MAY not be necessary
  // allEntries = checkInForm(allEntries)        
  var allEntries = unverifiedEntries()       
  */
  var allEntries = verifiedEntries()
  

  //Recalculate when ever there is a change (change in what?????)
  // const interval = dataSS.setRecalculationInterval(
  //   SpreadsheetApp.RecalculationInterval.ON_CHANGE,
  // )
  const dataSS = SpreadsheetApp.getActiveSpreadsheet() //'Keys Sheet Main'
  const allSheets = dataSS.getSheets()
  const template_sheet = allSheets[allSheets.length - 1]

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

  //Get the years from all the entries (map) in an array
  const years = []
  allEntries.forEach((entryRecord) => {
    var keys = entryRecord.getKeys()
    for(i = 0; i < keys.length; i++){
      var key = keys[i]
      var date = new Date(key.getExpirationDate())
      var entry_yr = date.getFullYear()
      if(!years.includes(entry_yr)){
        years.push(entry_yr)
      }
    }
  });

  years.sort().reverse()//sort years array in descending order

  //Create sheets with the given years
  for(i = 0; i < years.length; i++ ){
    //Create new sheet
    var new_sheet = dataSS.insertSheet((years[i]).toFixed(0), i+1, {template: template_sheet})
    //Name the new sheet
    if(i == 0) 
      {new_sheet.getRange("A1").setValue('Unknown Expiration')} 
    else 
      {new_sheet.getRange("A1").setValue((`Expiration: ${years[i]} `))}
  }

  //!!!!!!!Create condition for unknown years


  //////////////////////////

  //Add entry to the different sheets
  allEntries.forEach((entryRecord) => {
    var keys = entryRecord.getKeys()
    for(i = 0; i < keys.length; i++){
      var key = keys[i]
      var date = new Date(key.getExpirationDate())
      var yr1  = date.getFullYear() 
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

//On the main sheet
function analysis(){
  const dataSS = SpreadsheetApp.getActiveSpreadsheet()

  //Add more spreadsheets here! Add the spreadsheet folder!!!!
  //var allEntries = checkoutFormToEntries(null) //CHANGE TO THE ACTUAL ENTRIES
  var allEntries = new Map()
  allEntries = fillSheets(allEntries)
  var E = allEntries
  //ADD CONDIITON FOR ALL ENTRIES
  var currDate   = new Date()

  var sixDate = new Date(currDate)
  sixDate.setMonth(sixDate.getMonth() + 6)
  
  var threeDate = new Date(currDate)
  threeDate.setMonth(threeDate.getMonth() + 3)
  
  var oneDate = new Date(currDate)
  oneDate.setMonth(oneDate.getMonth() + 1) 

  var weekDate = new Date(currDate)
  weekDate.setDate(weekDate.getDay() + 7)

  var dayDate = new Date(currDate)
  dayDate.setDate(dayDate.getDate() + 1)

  var andrew_day   = []
  var andrew_week  = []
  var andrew_one   = []
  var andrew_three = []
  var andrew_six   = []
  var expired_list = []
  var unknown_list = []


  allEntries.forEach((entryRecord) => {

    //error in 1 and 3

    var keys = entryRecord.getKeys()
    for(i = 0; i < keys.length; i++){
      var key = keys[i]
      var expiration = new Date(key.getExpirationDate())
      if(isDateInFrame(currDate,dayDate,expiration)){
        andrew_day.push(entryRecord.getAndrewID())

      } else if(isDateInFrame(currDate,weekDate,expiration)){
        andrew_week.push(entryRecord.getAndrewID())

      } else if(isDateInFrame(currDate,oneDate,expiration)){
        andrew_one.push(entryRecord.getAndrewID())
      

      } else if (isDateInFrame(currDate,threeDate,expiration)){
        andrew_three.push(entryRecord.getAndrewID())
      } else if(isDateInFrame(currDate,sixDate,expiration)){
        andrew_six.push(entryRecord.getAndrewID())
      } 
      
      
      else if(isExpired(currDate,expiration)){
        ////////////////////////////////
        expired_list.push(entryRecord.getAndrewID())
      
      } else {
        //////////////////////////////
        unknown_list.push(entryRecord.getAndrewID()) 
      }
    }  
  })
  
  const sheets = dataSS.getSheets()
  const mainSheet = sheets[0]

  var six     = mainSheet.getRange("B8:B")
  var six_values = six.getValues()
  mainSheet.getRange(7,2).setValue('6 Months')
  for(var i = 0; i < six_values.length; i++){
    if(i < andrew_six.length){
      mainSheet.getRange(8+i,2).setValue(andrew_six[i])
    }
  }

  var three   = mainSheet.getRange("D8:D")
  var three_values = three.getValues()
  mainSheet.getRange(7,3).setValue('3 Month')
  for(var i = 0; i < three_values.length; i++){
    if(i < andrew_three.length){
      mainSheet.getRange(8+i,3).setValue(andrew_three[i])
    }
  }

  var one     = mainSheet.getRange("D8:D")
  var one_values = one.getValues()
  mainSheet.getRange(7,4).setValue('1 Month')
  for(var i = 0; i < one_values.length; i++){
    if(i < andrew_one.length){
      mainSheet.getRange(8+i,4).setValue(andrew_one[i])
    }
  }


  //1 week
  var week    = mainSheet.getRange("E8:E")
  var week_values = week.getValues()
  mainSheet.getRange(7,5).setValue('1 Week')
  for(var i = 0; i < week_values.lenght; i++){
    if(i < andrew_week.length){
      mainSheet.getRange(8+i,5).setValue(andrew_week[i])
    }
  }

  //1 day
  var day        = mainSheet.getRange("F8:F")
  var day_values = day.getValues()
  mainSheet.getRange(7,6).setValue('1 Day')
 for(var i = 0; i < day_values.length; i++){
  if(i <andrew_day.length){
    mainSheet.getRange(8+i,6).setValue(andrew_day[i])
  }
 }

  var expired = mainSheet.getRange("G8:G")
  var expired_values = expired.getValues()
  mainSheet.getRange(7,7).setValue('Expired')
  for(var i = 0; i < expired_values.length; i++){
    if(i < expired_list.length){
      mainSheet.getRange(8+i,7).setValue(expired_list[i])
    }
  }

  var unk     = mainSheet.getRange("H8:H")
  var unk_values = unk.getValues()
  mainSheet.getRange(7,8).setValue('Unknown')
  for(var i = 0; i < unk_values.length; i++){
    if(i < unknown_list.length){
      mainSheet.getRange(8+i,8).setValue(unknown_list[i])
    }
  }
  //If it is within 6 months of expiration


  //Sort through all the sections
  //Later->Add new Sheet at the end of each year
  //1.Add all the new form responses
  //2.Add form checked out to each student
  //3.Add form to show outstanding keys
  //4.Add form to show current keys
  /**Toast message should be sent if deadline 
   * is being approachd **/

}

////////////////////////////////////////////////////////////////////////////////////////
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Key Menu')  
      .addItem('Show sidebar', 'sidebarHome') //        
      .addToUi();
}

function sidebarHome() {
  var html = HtmlService.createHtmlOutputFromFile('home_sidebar').setTitle('Keys Project Home');
  SpreadsheetApp.getUi().showSidebar(html);
}

function sidebarModify(andrewID){
  var temp = HtmlService.createTemplateFromFile('modify_sidebar')
  var entry = setEntry(andrewID)
  temp.firstName = entry.getFirstName()
  var html = temp.evaluate().setTitle('Modify Sidebar');
  SpreadsheetApp.getUi().showSidebar(html);
}

function setEntry(andrewID){
  //find a way to pass in all entries
  if(E.has(andrewID)){
    return E.get(andrewID)
  } 
  else {return null}
}

function processInputs(fname, lname, advisor, andrewID, 
                      keyNum, roomNum, givenDate, loseDate) {
  // Process the inputs here
  Logger.log('Input 1: ' + fname);
  Logger.log('Input 2: ' + lname);
  Logger.log('Input 3: ' + advisor);
  Logger.log('Input 4: ' + andrewID);
  Logger.log('Input 5: ' + keyNum);
  Logger.log('Input 6: ' + roomNum);
  Logger.log('Input 7: ' + givenDate);
  Logger.log('Input 8: ' + loseDate);
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
