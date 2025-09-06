/*
Fix 1969 Dates tab!

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

/////////////////////////////////////
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
///////////////////////////////////////////////////////////
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

//Valid date!!!!!
function validDate(){
/////////////////////////////////////
  return
}


//Parsing the code sheets
function parseKeySheet(allEntries,id){
  var newSS = SpreadsheetApp.openById(id)
  var sheets = newSS.getSheets()
  
  //All the data
  for(var i = 0; i < sheets.length; i++){
    //Figure out the end range <--or how to avoid going over empty setion
    var sheet = sheets[i];
    var firstNameArr = sheet.getRange('D2:D').getValues(); //sheet.getRange('D2:D113').getValues(); ////////Standardize!!!!
    var lastNameArr  = sheet.getRange('C2:C').getValues();//sheet.getRange('C2:C113').getValues();
    var andrewIDArr  = sheet.getRange('F2:F').getValues();//sheet.getRange('F2:F113').getValues();
    var advisorArr   = sheet.getRange('G2:G').getValues();//sheet.getRange('G2:G113').getValues();
    var deptArr      = sheet.getRange('I2:I').getValues();//sheet.getRange('I2:I113').getValues();
    var keysArr      = sheet.getRange('J2:J').getValues();//sheet.getRange('J2:J113').getValues();
    var roomArr      = sheet.getRange('K2:K').getValues();//sheet.getRange('K2:K113').getValues();
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

/*
Translate form response to spreadsheet format (with keyReponse class) )
*/
function checkoutFormToEntries(allEntries){//function currentFormToClass(allEntries) { 
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
        givenDate = answ;
      } else if(ques == "What date will you lose acess? (Typically expected graduation date)") {
        expDate = answ;
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

//Safety check
function confirmUser(first,last,advisor,andrew,key,room,entry){
  if((first != entry.getFirstName()) 
    || (last != entry.getLastName()) 
    || (advisor != entry.getAdvisor())
    || (andrew != entry.getAndrewID())) {
      return true
    } return false
}

//Read log data and turn them into entries
function logToEntries(){
  var keySS = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = keySS.getSheetByName('Log');
  var allEntries = new Map();

  var logRange = logSheet.getRange('A:J');
  var logValues = logRange.getValues();

  for(var i = 1; i < logValues.length; i++){
    var row = logValues[i];
    if(row.length == 0 || row.length < 10) break;
    //var approval  = row[0];
    var andrewID  = row[1];
    var lastName  = row[2];
    var firstName = row[3];
    var advisor   = row[4];
    var dept      = row[5];
    var key       = row[6];
    var room      = row[7];
    var expDate   = row[8];
    var givenDate = row[9];
    if((andrewID == '') && (lastName =='') && (firstName == '') && 
       (advisor == '') && (dept == '') && (key == '') && (room == '') &&
       (expDate == '') && (givenDate == '')){break;}

    var newKeyRec = new keyRecord(firstName,lastName,andrewID,advisor,dept,key,room,givenDate,expDate);
    allEntries.set(andrewID,newKeyRec);
  }
  return allEntries;
}

function addToLog(){
  var keySS    = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = keySS.getSheetByName('Log');
  
  var logEntries = logToEntries();
  var allEntries = new Map();
  allEntries = checkoutFormToEntries(allEntries);
  allEntries = checkInForm(allEntries);

  for(const [andrewID, keyRecord] of allEntries){
    var keys = keyRecord.getKeys()
    for(i = 0; i < keys.length; i++){
      var key = keys[i];
      //
      if(!logEntries.has(andrewID)){
        logSheet.appendRow([
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

}


function updateLog(andrewID,key){
  const keySS    = SpreadsheetApp.getActiveSpreadsheet()
  const logSheet = keySS.getSheetByName('Log')

  var andrew_found = logSheet.createTextFinder(andrewID).findAll() //???
  var key_found    = logSheet.createTextFinder(key).findAll()


  //find a value in andrew and key where row matches
  
  if(found){
    ///////
  }


}
// function searchLog(){
//   var keySS = SpreadsheetApp.getActiveSpreadsheet();
//   var logSheet = keySS.getSheetName('Log');
//   var andrewIDList = logSheet.getRange('A2:A').getValues();  
// }

function unverifiedValueCollection(){
  var keySS = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = keySS.getSheetByName('Log');
  var approvalAndAndrew = logSheet.getRange('A2:B').getValues();
  
  var logEntries = logToEntries()

  var allEntries = new Map();
  allEntries = checkoutFormToEntries(allEntries);
  allEntries = checkInForm(allEntries);
  var unverifiedEntries = new Map();


  //if not in log, unverified list. add to log as unverified
  for(const [andrewID,keyRecord] of allEntries){

    //check if in log or if unverifed in log
    var arr = ["Unverified",andrewID]
    var value = approvalAndAndrew.includes(arr);
    if(!logEntries.has(andrewID) || approvalAndAndrew.includes(arr)) {
      unverifiedEntries.set(andrewID,keyRecord);
    }
    //add to the log! specifiy it is unverfied
  }
  return unverifiedEntries;
}

//Check if in the log. If not, add to the entry
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
    if(allEntries.has(andrewID)){
      var entry = allEntries.get(andrewID)
      var confirmed = confirmUser(firstName,lastName,advisor,andrewID,key,room,entry) //checking keys/rooms??????

      if(confirmed){
        allEntries.delete(andrewID)
        keys = []
        //Can be more efficent!!!
        entry.key.forEach((keyDetails) => {
          if(keyDetails.getKey() == key){
            keyDetails.deactivate()
          }
          keys.push(keyDetails)
        });
        entry.setKey(keys)
        allEntries.set(andrewID,entry)        
      }
    }
  }
  return allEntries
}

function correctUnknown(){
////////////////////////////
}

function manualCheckIn(){
///////////???????
}

function scheduleReload(){
////////////////////////////
}

function scheduledDeleteCheck(){
///????????????
}


function currentKeys(){
  ///////////////////////////////////////
}

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


//Approve Selected - Button
function submitSelectedData(){
  //clear values that are selected
  var keySS           = SpreadsheetApp.getActiveSpreadsheet();
  var unverifiedSheet = keySS.getSheetByName('Unverified Input');
  var sheetEntries_raw = unverifiedSheet.getRange("A2:J")
  var sheetEntries     = sheetEntries_raw.getValues()
  var allEntries = new Map()
  var deletedEntires = new Map()
  for(var i = 0; i < data.length; i++){

    //loop through specific range, not data value!!!!!

    //Changes to Approve,Denied, --keep selected
    var entry = sheetEntries[i]
    var approval = entry[0]
    
    var andrewID  = entry[1]
    var lastName  = entry[2]
    var firstName = entry[3]
    var advisor   = entry[4]
    var dept      = entry[5]
    var key       = entry[6]
    var room      = entry[7]
    var expDate   = entry[8]
    var givenDate = entry[9]
    var keyRec = new keyRecord(firstName,lastName,andrewID,advisor,dept,key,room,givenDate,expDate);
    if(approval == "Approve"){
      allEntries.set(andrewID,keyRec)
      //Clear and update the log
    } 
    if(approval == "Denied"){
      deletedEntires.set(andrewID,keyRec)
      //clear and update the low
    }
    ///Ignore the 'Selected' option

    return allEntries
  }
  //return the entries value. call this in analysis
  return null //CHange
}

//Approve All - Button
function approveAllData(){
  //clear all the data in the unverifeid
  var keySS = SpreadsheetApp.getActiveSpreadsheet();
  var unverfiedSheet = keySS.getSheetByName('Unverified Input');
  //var sheetEntries_raw = unverfiedSheet.getRange("A2:J")
  //var sheetEntries = sheetEntries_raw.getValues()
  var allEntries   = new Map(); 

  //next row is not empty
  //for(var i = 0; i < data.length; i++){

  var val = true // this needs to be updated!!!!!!

  while(val){
    var entries_raw = unverfiedSheet.getRange(2+i,1,1,10);
    var entries = entries_raw.getValues()
    //var entry = sheetEntries[i]
    var entry = entries[i]
    var entry_len = entry.length
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
    var keyRec = new keyRecord(firstName,lastName,andrewID,advisor,dept,key,room,givenDate,expDate);
    allEntries.set(andrewID,keyRec)
    entries_raw.clear()
  }
  //UPDATE THE LOG!!!!!
  var logSheet = keySS.getSheetByName('Log');
  var logEntries_raw = logSheet.getRange("A2:J");
  var logEntries = logEntries_raw.getValues();
  
  for(var i = 0; i < logEntries.length; i++){
    var entry_row    = logEntries[i]
    var andrewID1 = entry_row[1]
    var key      = entry_row[6]

    var found_entry = allEntries.get(andrewID1) //undefined if not there
    if(found_entry != undefined){
      var keys = found_entry.getKeys()
      if(keys.includes(key)){
        //update the log to have the proper approval value
        //????????????
        //how do I find the specifice value there are at??????????????
      }
    }
  }


  sheetEntries_raw.clear()
  //return the entries value. call this in analysis
  return allEntries
}

function fillSheets(allEntries){
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
    var allEntries = parseKeySheet(allEntries,id)
  }

  //Form
  allEntries = checkoutFormToEntries(allEntries) //??MAY not be necessary
  allEntries = checkInForm(allEntries)               

  const dataSS = SpreadsheetApp.getActiveSpreadsheet() //'Keys Sheet Main'

  //Recalculate when ever there is a change (change in what?????)
  // const interval = dataSS.setRecalculationInterval(
  //   SpreadsheetApp.RecalculationInterval.ON_CHANGE,
  // )

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
  E = allEntries
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
