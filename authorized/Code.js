/**
 * Captuers changes to actice spreadsheet
 * @param {*} e  - event object
 */
function onEdit(e){
  const buttonCell = 'D2';
  const authorizeSS = e.source;
  const button = authorizeSS.getRange(buttonCell);

  const edit = e.range;
  const sheet = edit.getSheet();
  const changeCol = edit.getColumn();
  const changeRow = edit.getRow();

  //Change in the emails
  if (changeCol === 1){
    const button = sheet.getRange("D2");
    button.setValue("Data Changed")
    button.setBackground("#ffcccc");
    condenseEmails();
  }


  //Do not change the error line
  if((changeCol == button.getColumn()) && (changeRow == button.getRow())){
    //revert to the original value and give a warning
    SpreadsheetApp.getActive().toast("Do not edit the error line");
    e.range.setValue(e.oldValue ?? '');
  }
}

/**
 * Strategy to ensure all emails are compact, especially after a value has been deleted.
 */
//https://stackoverflow.com/questions/4009085/checking-if-an-email-is-valid-in-google-apps-script
function condenseEmails(){
  var rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(['Select Permissions','Editing', 'Viewing', 'Commenting'],true)
            .setHelpText("Select an option")
            .build();  

  const accessSheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const emails_column = accessSheet.getRange('A2:A').getValues();
  const access_column = accessSheet.getRange('B2:B').getValues();
  const email_regex =  /\S+@\S+\.\S+/;
  
  //Only get valid emails from A column
  var rowTotal = 0;
  var email_access = {};
  for(var i = 0; i < emails_column.length; i++){
    //non-empty string and valid email
    if((emails_column[i][0] !== "") && (email_regex.test(emails_column[i][0]))){
      var email = emails_column[i][0];
      var access = access_column[i][0];
      if(access == ""){
        email_access[email] = null;
      } else {
        email_access[email] = access;
      }
      rowTotal++;
    }
  }

  //Remove old, unformated version of A colum
  accessSheet.getRange('A2:A').clear();
  var deleteDropdown = accessSheet.getRange('B2:B');
  deleteDropdown.clearContent();
  deleteDropdown.clearDataValidations();

  //Add valid emails back to A column and add drop down menu to B column
  var index = 0;
  for(const [key,value] of Object.entries(email_access)){
    var email_cell  = accessSheet.getRange(index+2,1); // A2,A3,A4...
    email_cell.setValue(key);
    var access_cell = accessSheet.getRange(index+2,2); // B2,B3,B4...
    access_cell.setDataValidation(rule);
    //No dropdown menu
    if(value == null) 
      {access_cell.setValue('Select Permissions');} 
    //Already dropdown menu
    else             
      {access_cell.setValue(value)}
    index++;
  }

  //Set diffent colors for the values in the drop down menu
  var goalRange = accessSheet.getRange(2,2,rowTotal,1);
  var currRules = accessSheet.getConditionalFormatRules();
  var newRules = currRules.filter(function(rule) {
    var ranges = rule.getRanges();
    return !ranges.some(r => r.getA1Notation().startsWith("B"));
  });

  newRules.push(
  SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Editing')
    .setBackground('#b7e1cd') // light green
    .setRanges([goalRange])
    .build(),
  SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Viewing')
    .setBackground('#fffee0') // light yellow
    .setRanges([goalRange])
    .build(),
  SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Commenting')
    .setBackground('#fce5cd') // light orange
    .setRanges([goalRange])
    .build()
  );
  accessSheet.setConditionalFormatRules(newRules);
}

/**
 * Submititng the emails to formally be the authorized used for the key sheet
 */
function submitData(){
  const accessSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();  
  //('https://docs.google.com/spreadsheets/d/1vQtW4KrtQg0T16zBT5GRi9oMe3q009HKhgctaHoyt-E/edit?gid=0#gid=0');
  const keySS       = DriveApp.getFileById('1vQtW4KrtQg0T16zBT5GRi9oMe3q009HKhgctaHoyt-E');
  const file        = DriveApp.getFileById(keySS.getId());

  const emails = accessSheet.getRange('A2:A').getValues();
  const access = accessSheet.getRange('B2:B').getValues();
  const permanentAccess = [ //IT staff only
    "anyaga@andrew.cmu.edu",
    "sarahhug@andrew.cmu.edu",
    "nyaga.angel@outlook.com"
  ]; 

  permanentAccess.forEach(function(email){
    keySS.addEditor(email);
  });

  //Error message location (D2)
  var errorRow = 2;
  const errorCol = 5;
  const errorRange = accessSheet.getRange(errorRow,errorCol,accessSheet.getMaxRows() - errorRow + 1,1);
  errorRange.clearContent().setBackground(null);

  //Length of non-empty emails (assumed already condensed) --> can add a break or make while loop
  var rowTotal = 0
  for (var i = 0; i < emails.length; i++) {
    if (emails[i][0] !== ""){
      rowTotal++;
    }
  }

  const keepUsers = new Set([...permanentAccess]);
  for(let i = 0; i < rowTotal; i++){
    var currEmail = emails[i][0];
    var currAccess = access[i][0];
    try {
      keepUsers.add(currEmail)
      //remove curr permissions //???
      const editors = keySS.getEditors().map(user => user.getEmail()); 
      const viewers = keySS.getViewers().map(user => user.getEmail());
      // Remove all users not in keep list
      const allUsers = new Set([...editors, ...viewers]);
      for (const email of allUsers) {
        if (!keepUsers.has(email)) {
          try {
            file.removeEditor(email);
          } catch (e1) {
            Logger.log(`Failed to remove editor ${email}: ${e1.message}`);
          }
          try {
            file.removeViewer(email);
          } catch (e2) {
            Logger.log(`Failed to remove viewer ${email}: ${e2.message}`);
          }
        }
      }

      if(currAccess === "Editing"){
        keySS.addEditor(currEmail);
      } else if (currAccess === "Viewing"){
        keySS.addViewer(currEmail);
      } else if (currAccess === "Commenting")
        keySS.addCommenter(currEmail)
    } catch (e) {
      Logger.log(`Error processing ${currEmail}: ${e.message}`);
      const errorCell = accessSheet.getRange(errorRow,errorCol);
      errorCell.setValue(`Error (${currEmail}): ${e.message}`).setBackground("#f4cccc");
      errorRow++;
    }
  }
  const button = accessSheet.getRange("D2");
  button.setValue("Submitted");
  button.setBackground("#ccffcc");
}

https://developers.google.com/apps-script/reference/spreadsheet/protection
function protectButtonCell() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const buttonCell = sheet.getRange('D2');

  //Create section to protect
  const protection = buttonCell.protect();
  protection.setDescription("Protect Submit Button");
  protection.removeEditors(protection.getEditors());

  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}
