/*
ABOUT THIS SCRIPT-
This is a Google Form builder application, with some key added features for handling emails, attachements, and formatting that aren't in standard Google Forms.
It was written by Wesley Keller, basically on a dare from Marty Byro.
Please contact Wes (me) if you have any questions about the code.  I tried to comment it as best as I could so that someone relatively new to Javascript could understand the program.
If a comment does not have a signature, assume I made it.
*/

//global variables
  var ssKey = '<<<YOUR SPREADSHEET KEY HERE>>>>';  
  var reloadLink = '<<<<YOUR RELOAD WEB SERVICE HERE>>>>';
  
  //PLEASE NOTE - YOU WILL NEED TO REPLACE THE SPREADSHEET KEY AND RELOAD LINK FOR THIS SCRIPT TO WORK, AND DO THE SAME IN THE RELOAD SCRIPT.
  //THIS IS THE ONLY TIME YOU SHOULD NEED TO UPDATE THIS CODE.  READ THE DEPLOYMENT GUIDE!!

  /*
  These are for technical debugging:
  level 0: no logging.
  level 1: log function calls.
  level 2: log class calls.
  level 3: log all non-loop variables.
  level 4: log all loop calls and loop variables (aka, shit be messed up real bad.)
  
  */
  var logFilesOn = false;
  var logFileLevel = 2;

/*
URL Construction:
https://[WEBAPPLINK]?CaseID=###&Random=#####

*/



//doGet is the function that creates the web form.  It is the main script, and what executes immediately upon opening the Form.
//Note that 'doGet' is a Google function requirement for web apps.  You cannot change the name of this function, or 'doPost'.

function doGet() {

  //here, we define some basic stuff we need to quickly work with the Google UI methods.
  var app = UiApp.createApplication();
  var formP = app.createFormPanel();
  var headerP = app.createFlowPanel().setId('headerP');
  var vertP = app.createVerticalPanel().setId('vertP1');
  var vertP2 = app.createVerticalPanel().setId('vertP2');
  var vertP3 = app.createVerticalPanel().setId('vertP3');
  var colPanel = app.createHorizontalPanel().setId('colPanel');
  var allPanel = app.createVerticalPanel().setId('allPanel');
  
  //here, we grab the values of the fields we're going to create on the form from the spreadsheet
  var ss = SpreadsheetApp.openById(ssKey);
  var sheet = ss.getSheetByName('Fields');
  var range = sheet.getDataRange();
  var values = range.getValues();
  var lastRow = range.getLastRow();
  
  //...and the styles we'll apply to the form...
  var styleSheet = ss.getSheetByName('Styles');
  var styles = styleSheet.getDataRange().getValues();
  var stLen = styleSheet.getDataRange().getLastRow();
  
  //...and some general settings...
  var setSheet = ss.getSheetByName('Settings');
  var settings = setSheet.getDataRange().getValues();
  
  var submit_handler = app.createServerHandler("RequiredFieldHandler");  //a special handler for 'required' fields
  
  //now it is time to create our form.  The form elements are just being added in what is called a Vertical Panel - meaning they will stack in the order you add them to the sheet.
 
   var shortname = Session.getActiveUser();

  //here, we add our default styles to the panels themselves.

   addStyle(headerP,settings[7][1],styles,stLen);
   addStyle(vertP,settings[8][1],styles,stLen);
   addStyle(vertP2,settings[9][1],styles,stLen);
   addStyle(vertP3,settings[10][1],styles,stLen);
   addStyle(colPanel,settings[11][1],styles,stLen);
  
  //now the good stuff.  This long loop will generate all the elements you've described in the Fields sheet, and pull data from any ListBox or RadioBox tabs you have created.
  //this might run slightly faster as a case/switch test, and be cleaner.  It's fine right now.
  
  for (var i = 1; i < lastRow; i++) {
    switch (values[i][0])  {
      
     case 'TextBox':
      var visible = true;
      if (values[i][4]=='No') {visible = false;} 
      var lbl = app.createLabel(values[i][1]).setVisible(visible);
      var tb = app.createTextBox().setName(values[i][2]).setVisible(visible).setValue(values[i][5]).addFocusHandler(submit_handler).setId(values[i][2]);
      addStyle(lbl,values[i][6],styles,stLen);
      addStyle(tb,values[i][7],styles,stLen);
      submit_handler.addCallbackElement(tb);
      LabelAndColumnLocationHandler(values[i][9],tb,values[i][3],lbl);
      break;
     case 'Label':
      var lbl = app.createLabel(values[i][1]);
      addStyle(lbl,values[i][6],styles,stLen);
      LabelAndColumnLocationHandler(values[i][9],lbl,values[i][3]);
      break;
     case 'TextArea':
      var lbl = app.createLabel(values[i][1]);
      var ta = app.createTextArea().setName(values[i][2]).addFocusHandler(submit_handler).setId(values[i][2]);
      addStyle(lbl,values[i][6],styles,stLen);
      addStyle(ta,values[i][7],styles,stLen);
      submit_handler.addCallbackElement(ta);
      LabelAndColumnLocationHandler(values[i][9],ta,values[i][3],lbl);
      break;
     case 'DateBox':
      var lbl = app.createLabel(values[i][1]);
      var db = app.createDateBox().setName(values[i][2]).addValueChangeHandler(submit_handler).setId(values[i][2]);
      addStyle(lbl,values[i][6],styles,stLen);
      addStyle(db,values[i][7],styles,stLen);
      submit_handler.addCallbackElement(db);
      LabelAndColumnLocationHandler(values[i][9],db,values[i][3],lbl);
      break;
     case 'FileUpload':
      var lbl = app.createLabel(values[i][1]);
      var fu = app.createFileUpload().setName("%$FILE$%"+values[i][2]).addChangeHandler(submit_handler);  //The special string name is used to detect the existence of a file in the Post action.
      addStyle(lbl,values[i][6],styles,stLen);
      addStyle(fu,values[i][7],styles,stLen);
      submit_handler.addCallbackElement(fu);
      LabelAndColumnLocationHandler(values[i][9],fu,values[i][3],lbl);
      break;
     case 'CheckBox':
      var cb = app.createCheckBox(values[i][1]).setName(values[i][2]).setFormValue(values[i][5]).addFocusHandler(submit_handler).setId(values[i][2]);
      addStyle(cb,values[i][7],styles,stLen);
      LabelAndColumnLocationHandler(values[i][9],cb,values[i][3]);
      break;
     case 'RadioButton':
      var rb = app.createRadioButton(values[i][2], values[i][1]).setFormValue(values[i][5]).addFocusHandler(submit_handler).setId(values[i][2]);
      addStyle(rb,values[i][7],styles,stLen);
      submit_handler.addCallbackElement(rb);
      LabelAndColumnLocationHandler(values[i][9],rb,values[i][3]);
      break;
     case 'Image':
      var img = app.createImage(values[i][1]);
      LabelAndColumnLocationHandler(values[i][9],img,values[i][3]);
      break;
     case 'UserName':
      var visible = true;
      if (values[i][4]=='No') {visible =false;}
      var lbl_Shortname = app.createLabel(values[i][1]).setVisible(visible);
      var tb_Shortname = app.createTextBox().setValue(shortname).setVisible(visible).setName(values[i][2]).addFocusHandler(submit_handler).setId(values[i][2]);
      addStyle(lbl_Shortname,values[i][7],styles,stLen);
      addStyle(tb_Shortname,values[i][7],styles,stLen);
      submit_handler.addCallbackElement(tb_Shortname);
      LabelAndColumnLocationHandler(values[i][9],tb_Shortname,values[i][3],lbl_Shortname);
      break;
     case 'ListBox':
      var lbl = app.createLabel(values[i][1]);
      var lb = app.createListBox().setName(values[i][2]).addFocusHandler(submit_handler).setId(values[i][2]);
      addStyle(lbl,values[i][6],styles,stLen);
      addStyle(lb,values[i][7],styles,stLen);
      var lSheet = ss.getSheetByName('Lists');
      var lRange = lSheet.getDataRange();
      var lValues = lRange.getValues();
      var lLastRow = lRange.getLastRow();
      
      for (var j = 0; j < lLastRow; j++) {
         if (lValues[j][0] == values[i][2]) {
           lb.addItem(lValues[j][1]);
          }
          }
      submit_handler.addCallbackElement(lb);
      LabelAndColumnLocationHandler(values[i][9],lb,values[i][3],lbl);
      break;
     default:
      break;
    }
//More types to come.  Next small block adds a Submit button and builds the app.

   }
  
  //******DEPRECIATED.  I moved all the validations on blank fields to a focus or change handler, but you could recover this if performance suffers.
  //Here we add a validation rule.  Long term, I want this added to each field.
 
 /*
 
  var btn_Validate = app.createButton("Validate").addClickHandler(submit_handler);
  vertP.add(btn_Validate);
 
 */
 
  //Almost done!  This just throws a submit button on the whole thing, and then gives you the form.  Note the submit button comes from a parameter.
        
  var btn_Submit = app.createSubmitButton(settings[2][1]).setId('Submit').setEnabled(false);
  vertP.add(btn_Submit);
  colPanel.add(vertP).add(vertP2).add(vertP3);
  allPanel.add(headerP).add(colPanel);
  formP.add(allPanel);
  app.add(formP);

  return app;
}

/*

OK, now we're ready to post the info.  When we created the fields above, we assigned field names to everything.  Those names get passed to the event linked to the Submit button we created
in the last part of the doGet function, and we can now work with the name/value pairs that got passed in the 'eventInfo' variable.

*/


//*********************  THE POST PROGRAM STARTS HERE   ***********************************

function doPost(eventInfo) {
  
//again, just some standard stuff we need to work with the UI and spreadsheet methods quickly.
  var app = UiApp.getActiveApplication();
  var par = eventInfo.parameter;
  var ss = SpreadsheetApp.openById(ssKey);
  var curDate = new Date();
  var caseData = new Object();
 
//now we grab all the field names from the spreadsheet, so we can search for them.

  var fSheet = ss.getSheetByName('Fields');
  var fDataRange = fSheet.getDataRange();
  var fLastRow = fDataRange.getLastRow();
  var fType = fSheet.getRange(1, 1, fLastRow, 1).getValues();
  var fData = fSheet.getRange(1, 3, fLastRow, 1).getValues();
  
  //Logger.log(fData);  //loggers are fun, and can help you identify problems.  If you write code, use them!  But you can delete this line if you want.

//next, we grab all the destination sheet headers, so we can search for them, match them to the field names, and write them to the next open line.

  var dSheet = ss.getSheetByName('RequestData');
  var dRange = dSheet.getDataRange();
  var dLastRow = dRange.getLastRow();
  var dLastColumn = dRange.getLastColumn();
  var dKeys = dSheet.getRange(2, 1, 1, dLastColumn).getValues();
  
  dLastRow++;
  
  //Logger.log(dKeys);  //another logger.  These were added mainly for debugging the program - I should really clean them up once I'm done.
  
  var wSheet = ss.getSheetByName('Workflow');
  var wRange = wSheet.getDataRange();
  var wLastRow = wRange.getLastRow();
  var wValues = wRange.getValues();
  
  var workflowObject = WorkflowDetermination(par,wValues,dKeys,wLastRow);
  
  
  //lastly, we've got to grab those settings again.
  
  var setSheet = ss.getSheetByName('Settings');
  var settings = setSheet.getDataRange().getValues();
  var gFolder = DocsList.getFolder(settings[4][1]);
  var counter = settings[1][1];
  
  par['caseID'] = counter;
  setSheet.getDataRange().getCell(2, 2).setValue(counter+1);
  
  
  //This is where we actually commit the Form Post data to the spreadsheet
 
  var myFiles = new Array();  //we have to create an array for sending multiple files!
  //dSheet.getRange(dLastRow,1).setValue(par.shortname);
  dSheet.getRange(dLastRow,1).setValue(curDate);


  for (var i = 0; i < fLastRow; i++) {
     //we start by looking at attachments, since we need to pass these to email and google.
       if (fType[i] == 'FileUpload') {
         var file = par[("%$FILE$%"+fData[i])];
         //Logger.log('LOOK AT ME ->');
         //Logger.log(par);
           var gDocName = file.getName();
           if (settings[3][1] == "X") {
              if (par['file'] !== '' || undefined) {
                var gDoc = DocsList.createFile(file);
                if (gFolder !== '') {gDoc.addToFolder(gFolder)};
              var docLink = gDoc.getUrl();
             }
           }
         myFiles.push(file);
         
      }
    
     for (var j = 0; j < dLastColumn; j++) {
      //now, we update each spreadsheet column with the pair value.
       caseData[fData[i]] = par[fData[i]];
       if (fData[i] == dKeys[0][j]) {
         var value = par[fData[i]];
         if (fType[i] == 'FileUpload') {value = docLink;}
         var column = j+1;
         dSheet.getRange(dLastRow, column).setValue(value);
       }
     }
    }
  
  var random = Math.ceil(((Math.random()*100000)));  //the random number is used to protect the other script (which reloads the case) from being spoofed.
  var caseViewURL = reloadLink+'?CaseID='+counter+'&Random='+random;
  var jsonCase = JSON.stringify(caseData);
  
    dSheet.getRange(dLastRow,3).setValue(jsonCase);
    
    dSheet.getRange(dLastRow,2).setValue(counter);
    
    dSheet.getRange(dLastRow,4).setValue(random);
    
    var json1 = JSON.stringify(workflowObject);
    dSheet.getRange(dLastRow,5).setValue(json1);
    
    dSheet.getRange(dLastRow,6).setValue(caseViewURL);
 
  //lastly, we handle emails.

  SendEmails(myFiles, counter);

  //MailApp.sendEmail('keller.wesley@gmail.com', 'DPC test', 'this is a test', {attachments: myFiles});
  
  
  //...and the styles we'll apply to the form...
  var styleSheet = ss.getSheetByName('Styles');
  var styles = styleSheet.getDataRange().getValues();
  var stLen = styleSheet.getDataRange().getLastRow();
  
  
  
  app.close();
  Utilities.sleep(100);
  
  //throwaway code to show case ID.  Consider replacing this later.
  
  var submitMessage = settings[5][1];
  
  var submitLines = submitMessage.split('~');
  var submitLen = submitLines.length;
  var subVertP = app.createVerticalPanel();
  
  
  for (var l = 0; l < submitLen; l++) {
    var submitLabel = app.createLabel(submitLines[l]);
    addStyle(submitLabel,settings[6][1],styles,stLen);  
    subVertP.add(submitLabel);
    }
    
  var counterLabel = app.createLabel('Your Case ID is '+counter+'.  Keep this for your reference.');
  addStyle(counterLabel, settings[6][1],styles,stLen);
  subVertP.add(counterLabel);
  
  app.add(subVertP);
  
  
  return app;
  
  //a quick call to another function, which runs some standard updates.
  ScheduledUpdater();
  
}


/*this is a mini-function that adds styles to any element.  It splits the values from the 'Fields' table by commas and applies every style listed.
'obj' is the UI element.
'style' is the style name'
'styles' is the entire style table.  If this caused performance issues, it could be cached.
'stLen' is the length of the style table, used to set a length for the for loop.

*/

function addStyle(obj, style, styles, stLen) {
    
    for (var i=0; i < stLen; i++) {
      var splitStyle = style.split(',');
      var splitLen = splitStyle.length;
       for (var j = 0; j < splitLen; j++) {
        if (splitStyle[j] == styles[i][0]) {
          obj.setStyleAttribute(styles[i][1],styles[i][2]);
         }
        }
      }
    return obj;
}


/*
This function just handles sending a quick email to say "a form was submitted".
Some methods may not work - I have not spent a ton of time on it yet.

*/


function SendEmails(attach, counter) {
  var ss = SpreadsheetApp.openById(ssKey);
  var sheet = ss.getSheetByName('Email_Notification');
  var range = sheet.getDataRange();
  var values = range.getValues();
  var lastRow = range.getLastRow();

  
  for (var i=1; i < lastRow; i++) {
       var noreply = false;
       if (values[i][4] == "Yes") {noreply = true};
       
              
       var sub = counter + values[i][1];
       
    MailApp.sendEmail({
  
      to: values[i][0],
      subject: sub,
      body: values[i][2], 
      attachments: attach,
      name: values[i][5],
      noReply: noreply,
      replyTo: values[i][6]
      });
  
  }
  
}


/*
This function is used to make sure any required field is populated with some data.
The function tests every field to see if it meets the requirement each time the focus changes.  Invoked from the main script.

At some point, this could become a more robust test for multiple condidions.

It should also be replaced with Client Handlers instead of a server handler.  Client handlers will perform much better and have more options, but present
additional challenges that Server Handlers do not.

*/


function RequiredFieldHandler(eventInfo) {
   var app = UiApp.createApplication();
   var par = eventInfo.parameter;
   var submitButton = app.getElementById('Submit');
   
   var ss = SpreadsheetApp.openById(ssKey);
   var sheet = ss.getSheetByName('Fields');
   var range = sheet.getDataRange();
   var values = range.getValues();
   var lastRow = range.getLastRow();
   
   var allOK = true;
   
   for (var i = 1; i < lastRow; i++) {
      if (values[i][8] == "X") {
         var test = par[values[i][2]];
         var badInput = app.getElementById(values[i][2]);
         badInput.setStyleAttribute('borderColor', 'LightGray').setStyleAttribute('borderWidth', '1px');
           if (test == "" || undefined) { 
              allOK = false;
              badInput.setStyleAttribute('borderColor', 'red').setStyleAttribute('borderWidth', '3px');
            
           }
         }
       }
   
   
   if (allOK == true) {submitButton.setEnabled(true);}
   else if (allOK == false) {submitButton.setEnabled(false);}
   
   return app;
  
}


/*
This function is actually meant to be a stand-alone function that runs on a time trigger.  It basically just runs through a set of instructions
to update fields in the 'RequestData' table if certain conditions are met, and kicks off actions based on the results.  

Functionality is somewhat limited because of the risk of a user setting up a loop command, which is still possible using multiple lines.
I would consider refactoring this at some point for better control.


*/



function ScheduledUpdater() {
  var ss = SpreadsheetApp.openById(ssKey);
  var sheet = ss.getSheetByName('Updater');
  var range = sheet.getDataRange();
  var values = range.getValues();
  var lastRow = range.getLastRow();
  
  var reqSheet = ss.getSheetByName('RequestData');
  var reqRange = reqSheet.getDataRange();
  var reqValues = reqRange.getValues();
  var reqLastRow = reqRange.getLastRow();
  var reqLastCol = reqRange.getLastColumn();
  
  var reqData = new Object();
  var col = new Object();
  
  //we work with the deletion rows separately, because you need to delete the newest rows first.  Basically, if we delete an older row, it shifts the new rows, and we start deleting the wrong rows.
  var deleteRow = new Array();
  
  for (var i=2; i < reqLastRow; i++) {
 
   for (var j=0; j < reqLastCol; j++) {
     reqData[reqValues[1][j]] = reqValues[i][j];
     col[reqValues[1][j]] = j;
     
    }
     reqData['columns'] = col;


  
  for (var k=1; k < lastRow; k++) {
    var doAction = false;
    var operator = values[k][1];
    var Action = values[k][4];
    
    
    
    
    switch (operator) {
      
      case '=':
        if (values[k][2] == values[k][3]) {Logger.log('Your update value cannot match your initial value, as this would cause loops.  This is an error.');  return;}
        if (reqData[values[k][0]] == values[k][2]) {
           reqSheet.getRange(i+1,reqData.columns[values[k][0]]+1).setValue(values[k][3]);
           doAction = true;
          }
        break;
      case '<>':
        if (values[k][2] !== values[k][3]) {Logger.log('Your update value must match your initial value, as this would cause loops otherwise.  This is an error.');  return;}
        if (reqData[values[k][0]] !== values[k][2]) {
           reqSheet.getRange(i+1,reqData.columns[values[k][0]]+1).setValue(values[k][3]);
           doAction = true;
        }
        break;
    /* Putting greater than and less than operators on hold for now.  Getting a little crazy with this thing.
        */
      case '<':
        if (values[k][2] <= reqData[values[k][0]]) {
           reqSheet.getRange(i+1,reqData.columns[values[k][0]]+1).setValue(values[k][3]);
           doAction = true;
        }
        break;
      case '>':
        if (values[k][2] >= reqData[values[k][0]]) {
           reqSheet.getRange(i+1,reqData.columns[values[k][0]]+1).setValue(values[k][3]);
           doAction = true;
        }
        break;
      default:
      Logger.log('Invalid Operator');
      break;
        
        
      }
    
    if (doAction == true) {
      switch (Action) {
        
        case 'SendEmail':  //this one just sends an email to one static person.  Useful for case errors, aging, etc.
           var to = values[k][5];
           var subject = (reqData[values[k][6]]+' '+values[k][7]);
           var body = values[k][8];
           body = body+('\n');
           var addFields = values[k][9].split(',');
           var addsLen = addFields.length;
             for (var l = 0; l < addsLen; l++) {
                body = body+('\n');
                body = body+(addFields[l]+': '+reqData[addFields[l]]);
             }
           MailApp.sendEmail(to, subject, body, {cc:'keller.wesley@gmail.com'});
        break;  
        case 'SendEmailToCaseValue':  //this will send emails to any value from the case.  Probably the most useful.
           var to = reqData[values[k][5]];
           var subject = (reqData[values[k][6]]+' '+values[k][7]);
           var body = values[k][8];
           body = body+('\n');
           var addFields = values[k][9].split(',');
           var addsLen = addFields.length;
             for (var l = 0; l < addsLen; l++) {
                body = body+('\n');
                body = body+(addFields[l]+': '+reqData[addFields[l]]);
             }
           MailApp.sendEmail(to, subject, body, {cc:'keller.wesley@gmail.com'});
        break;
        case 'DeleteLine':  //this will delete the record.  Useful if the table is becoming too big.
           deleteRow.push((i+1));
           break;
        
        
        default:
        Logger.log('no action selected');
        break;
      }
    }
   }



  }
 deleteRow.reverse();  //remember what I said above about deleting the newest (furthest down) rows first?
 var rowsToDelete = deleteRow.length;
 
 for (var l = 0; l < rowsToDelete; l++) { reqSheet.deleteRow(deleteRow[l])}
 
 
 }
/*
This function handles both where the label appears next to an input element, and where the element and label appear on the form.

Although I really only want to allow one location (left) and one label location (top).  Anything else looks messy to me.

panel: where the element goes
inputField: the actual UI element we are working with.
location: where the label goes. optional.  Doesn't look optional, but it should be.
label: the label for the UI element.  Optional.


*/



function LabelAndColumnLocationHandler(panel,inputField,location,label) {
    var app = UiApp.getActiveApplication();
    
    
    var headerP = app.getElementById('headerP');
    var vertP = app.getElementById('vertP1');
    var vertP2 = app.getElementById('vertP2');
    var vertP3 = app.getElementById('vertP3');
    var colPanel = app.getElementById('colPanel');
    
    Logger.log(vertP);
    Logger.log(inputField);
    Logger.log(label);
    
    var addToPanel = vertP;
    
    switch (panel) {
      case 'left':
        addToPanel = vertP;
        break;
      case 'middle':
        addToPanel = vertP2;
        break;
      case 'right':
        addToPanel = vertP3;
        break;
      case 'header':
        addToPanel = headerP;
        break;
      default:
        addToPanel = vertP;
        break;
       }
        
    switch (location) {
      
      case 'top':
        if (label !== undefined) {addToPanel.add(label)};
        addToPanel.add(inputField);
        break;
      case 'inline':
        var horiz = app.createHorizontalPanel();
        horiz.add(inputField);
        if (label !== undefined) {horiz.add(label)};
        addToPanel.add(horiz);
        break;
      case 'inlineLeft':
        var horiz = app.createHorizontalPanel();
        if (label !== undefined) {horiz.add(label)};
        horiz.add(inputField);
        addToPanel.add(horiz);
        break;
      default:
        if (label !== undefined) {addToPanel.add(label)};
        addToPanel.add(inputField);
        break;
   
   return;
  }
}

/*
This particular beast is used to create the JSON that holds all the workflow approval steps.  Note that the workflow is created once,
when the case is first submitted.  Later changes to the roles are noted (in case some one leaves the company), but the route
itself remains unchanged if the workflow tab is updated.

This makes a lot of calls to the methods of the classes in the 'WorkflowUpdate.gs' script.  It's best to read through it, it is probably the most advanced 
javascript stuff in this project.


*/


function WorkflowDetermination(par,wValues,dKeys,wLastRow) {


    var workflowObj = new UpdateWorkflowObject().setState(0)
                                                .setStatus('In Process');
    /*  Never know when you'll need those loggers...  
    Logger.log('MY WORKFLOW OBJECT -> '+JSON.stringify(workflowObj));
    Logger.log('My Parameters : '+JSON.stringify(par));
    Logger.log('My workflow Values : '+wValues);
    Logger.log('My key Values : '+dKeys);
    Logger.log('My last row : '+wLastRow);
    */
    
    
    var workOrder = new Array();
    var workflowSteps = new Array();
    
    
    //valid workflow states are 'Approved', 'Rejected', 'In Process'.
    //valid workflow step states are 'Approved', 'Rejected', 'Active' and 'Pending'.
    
    var ss = SpreadsheetApp.openById(ssKey);
    var tSheet = ss.getSheetByName('Tests');
    var tRange = tSheet.getDataRange();
    var tValues = tRange.getValues();
    var tLastRow = tRange.getLastRow();
    
    
    
    for (var i = 1; i < wLastRow; i++) {
     
     var thisStateTest = wValues [i][3];
     var addState = true;
     
     var tAddState = true;
     
     for (var j = 1; j < tLastRow; j++) {
       
      var thisTest = tValues[j][0];
      Logger.log('Test from workflows : '+thisStateTest);
      Logger.log('Test from Tests : '+thisTest);
      Logger.log('Is it a test? : '+tValues[j][1]);
      
      if (thisStateTest == thisTest && tValues[j][1] == 'Test') {
      
      
      
      
      var operator = tValues[j][3];
      var testAndOr = tValues[j][5];
      //Logger.log(operator+'--'+testAndOr);
       
      switch (operator) {
            
       case '=':
        //Logger.log('Compare : '+par[tValues[j][2]]+' <-> '+tValues[j][4]);
        if (par[tValues[j][2]] !== tValues[j][4]) {
           //Logger.log('Test Passed!');
           tAddState = false;
          }
        break;
       case '<>':
        if (par[tValues[j][2]] == tValues[j][4]) {
           tAddState = false;
          }
        break;
       default:
        tAddState = true;
        break;
        
        }
        
       }
      /*
      Logger.log('Add State : '+addState);
      Logger.log('Test to Add State : '+tAddState);
      */
      
      if (tAddState == false) {addState = false;}
      }
      /*
      Logger.log('After: Add State : '+addState);
      Logger.log('After: Test to Add State : '+tAddState);
      */
      if (addState == true) {
      
         var curState = wValues[i][0];
         var workflowStep = new UpdateWorkflowStep().setStepState(curState)
                                                    .setStepStatus('Pending');
         var json1 = JSON.stringify(workflowStep);
         
         workflowSteps.push(json1);
         
         
         workOrder.push(wValues[i][0]);
      
      }
        
        workflowObj.setWorkflow(workOrder)
                   .setWorkflowSteps(workflowSteps);
        
     }   
        
    return workflowObj;
       
}


/*  
THIS GROUP OF CLASSES IS FOR WORKING WITH THE WORKFLOW OBJECT AT THE HEADER OR MASTER LEVEL.

These are a collection of Methods for working with the JSON Class that contains workflow items.  

This stuff is advanced javascript.  Either you know what you're looking at and don't need comments, or you should probably just back away slowly.


*/
function UpdateWorkflowObject() {

  this.setState = function (state) {this.state = state; return this;};
  this.getState = function () {return this.state;};
  this.setStatus = function (status) {this.status = status; return this;};
  this.getStatus = function () {return this.status;};
  this.setWorkflow = function (workflow) {this.workflow = workflow; return this;};
  this.getWorkflow = function () {return this.workflow;};
  this.setWorkflowSteps = function (steps) {this.steps = steps; return this;};
  this.getWorkflowSteps = function () {return this.steps;};  

}

function ReloadWorkflowObject(workflowObj) {

   var workflowObject = JSON.parse(workflowObj);
   
   UpdateWorkflowObject.call(workflowObject);
   
   return workflowObject;
}

/*
THIS GROUP IS FOR WORKING WITH YOUR WORKFLOW AT THE STEP OR ACTION LEVEL.
*/

function UpdateWorkflowStep() {

  this.setStepState = function (state) {this.stepState = state; return this;};
  this.getStepState = function () {return this.stepState;};
  this.setStepStatus = function (status) {this.stepStatus = status; return this;};
  this.getStepStatus = function () {return this.stepStatus;};
  this.setStepActionBy = function (user) {this.stepActionBy = user; return this;};
  this.getStepActionBy = function () {return this.stepActionBy;};
  this.setStepActionTime = function (time) {this.stepActionTime = time; return this;};
  this.getStepActionTime = function () {return this.stepActionTime;};

}

function ReloadWorkflowObjectStep(workflowStep) {

   var workflowStepObj = JSON.parse(workflowStep);
   
   UpdateWorkflowStep.call(workflowStepObj);
   
   return workflowStepObj;
}






/*

------------DEPRECIATED----------------

function AdvanceWorkflowStatus (workflowObj) {

   //a dumb function, made to manually advance workflows to the next active state.
   //Really a Proof of Concept for methods for re-loading and changing the workflow JSON.  I don't think I use it anywhere in the final project, and I'm going to comment it out to make sure.
   
  Logger.log(workflowObj);
  var workflow = ReloadWorkflowObject(workflowObj);
  Logger.log(workflow);
  var state = workflow.getState();
  var workflowSteps = workflow.getWorkflowSteps();
  var order = workflow.getWorkflow();
  var orderLen = order.length;

  var newStepsArray = new Array();

  if (state == 0) {workflow.setState(order[0]);}
  
  //DEBUGGER VALUE
  state = 100;
  //END DEBUGGER VALUE
  
  
  for (var i = 0; i < orderLen; i++) {
     if (order[i] == state) {
     
       Logger.log(workflowSteps[i]);
       var step = ReloadWorkflowObjectStep(workflowSteps[i]);
       Logger.log(step);
       var stepState = step.getStepState();
       var stepStatus = step.getStepStatus();
       Logger.log('LOOOOOOOOOOOK ---->: '+stepState+'  '+stepStatus);
       var nextStep = 999;
       
       if (order[(i+1)] !== undefined) {nextStep = order[(i+1)];};
       
       switch (stepStatus) {
       
         case 'Pending':
            break;
         case 'Approved':
            workflow.setState(nextStep);
            if (nextStep == 999) {workflow.setStatus('Approved');};
            step.setStepStatus('Approved');
            break;
         case 'Active':
            break;
         case 'Rejected':
            workflow.setState(nextStep);
            if (nextStep == 999) {workflow.setStatus('Rejected');};
            step.setStepStatus('Rejected');
            break;
         default:
            Logger.log('default');
            break;
           }
           
           newStepsArray.push(step);
       }
       
      }   
        
    workflow.setWorkflowSteps(newStepsArray); 
        
  return workflow;
  
}

*/


//This guy just fetches the workflow object for a particular case ID.  He's not used anywhere else here, and I'm going to comment him out and see if it breaks anything.


/*

----------------DEPRECIATED-------------------

function FindAndReloadWorkflowObject(CaseID) {

  var ss = SpreadsheetApp.openById(ssKey);
  var dSheet = ss.getSheetByName('RequestData');
  var dRange = dSheet.getDataRange();
  var dValues = dRange.getValues();
  var dLastRow = dRange.getLastRow();
  var dLastColumn = dRange.getLastColumn();
  var dKeys = dSheet.getRange(2, 1, 1, dLastColumn).getValues();
  
  for (var i = 2; i < dLastRow; i++) {
     if (CaseID == dValues[i][1]) {
     
     var workflowObj = dValues[i][4];
     
     UpdateWorkflowObject.call(workflowObj);
     
     }
   }

  

  return workflowObj;
  
}

*/