//My reload page function. See the main function or the readme for details. Author - Wes Keller

//global variables
  var ssKey = '<<<YOUR SPREADSHEET KEY HERE>>>>';  



  //PLEASE NOTE - YOU WILL NEED TO REPLACE THE SPREADSHEET KEY FOR THIS SCRIPT TO WORK.
  //THIS IS THE ONLY TIME YOU SHOULD NEED TO UPDATE THIS CODE.
  
  /*these are for technical debugging.
  level 0: no logging.
  level 1: log function calls.
  level 2: log class calls.
  level 3: log all non-loop variables.
  level 4: log all loop calls and loop variables (aka, things be messed up real bad.)
  
  */
  var logFilesOn = false;
  var logFileLevel = 2;

/*
URL Construction:
https://[WEBAPPLINK]?CaseID=###&Random=#####

*/

function doGet(e) {

  if (logFilesOn == true && logFileLevel >= 1) {Logger.log('start function: doGet')};

    //*************   Here is where we actually reload the case values!   *******************

 //Since you can't load a dynamic web app like this without passing the values in the URL, sometimes I use the section below to do it manually.

 //TEMP EXCLUDED DEBUG VALUES!!!!!!!!  You can't really debug code called from a constructed URL as a web service, but you can put dummy values in the code to make it work.  So that's what we do here.
 var par = e.parameter;
 var caseID = par.CaseID;
 var caseRandom = par.Random;
 //END TEMP EXCLUDED DEBUG VALUES!!!!!!!!
 
 //TEMPORARY DEBUG VALUES!!!!!!!
 //var caseID = 192;
 //var caseRandom = 20469;
 //END TEMPORARY DEBUG VALUES!!!!!!!
 
 //and before we even start, we're gonna check the status.
 ApprovalStatusCheck(caseID,caseRandom);

 Utilities.sleep(200);
  
 var theCase = CaseLoader(caseID,caseRandom);
 var theWorkflow = WorkflowLoader(caseID,caseRandom);
 var theSteps = theWorkflow.getWorkflowSteps();
 var caseRow = CaseRowFinder(caseID,caseRandom);
 
 Logger.log(caseID+' : '+caseRandom);



  //some basic stuff we need to quickly work with the UI methods.
  var app = UiApp.createApplication();
  var formP = app.createAbsolutePanel();
  var headerP = app.createFlowPanel().setId('headerP');
  var vertP = app.createVerticalPanel().setId('vertP1');
  var vertP2 = app.createVerticalPanel().setId('vertP2');
  var vertP3 = app.createVerticalPanel().setId('vertP3');
  var colPanel = app.createHorizontalPanel().setId('colPanel');
  var allPanel = app.createVerticalPanel().setId('allPanel');
  
  //Gotta grab our query data:
  
  
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
  
  //...and our workflow values...
  var wSheet = ss.getSheetByName('Workflow');
  var wRange = wSheet.getDataRange();
  var wLastRow = wRange.getLastRow();
  var wValues = wRange.getValues();
  
  //...AND our requests...
  var dSheet = ss.getSheetByName('RequestData');
  var dRange = dSheet.getDataRange();
  var dLastRow = dRange.getLastRow();
  var dLastColumn = dRange.getLastColumn();
  var dKeys = dSheet.getRange(2, 1, 1, dLastColumn).getValues();
  
  //...and some general settings...
  var setSheet = ss.getSheetByName('Settings');
  var settings = setSheet.getDataRange().getValues();
  
   addStyle(headerP,settings[7][1],styles,stLen);
   addStyle(vertP,settings[8][1],styles,stLen);
   addStyle(vertP2,settings[9][1],styles,stLen);
   addStyle(vertP3,settings[10][1],styles,stLen);
   addStyle(colPanel,settings[11][1],styles,stLen);
  

 

 
  
  
  
 var radioGroup = '';
  
  //now the good stuff.  This long loop will generate all the elements you've described in the Fields sheet, and pull data from any ListBox or RadioBox tabs you have created.
  
  for (var i = 1; i < lastRow; i++) {
    if (values[i][0] == 'TextBox') {
      var lbl = app.createLabel(values[i][1]);
      var tb = app.createTextBox().setName(values[i][2]).setValue(theCase[values[i][2]]).setId(values[i][2]);
      addStyle(lbl,values[i][6],styles,stLen);
      addStyle(tb,values[i][7],styles,stLen);
      LabelAndColumnLocationHandler(values[i][9],tb,values[i][3],lbl);
      }
    else if (values[i][0] == 'Label') {
      var lbl = app.createLabel(values[i][1]);
      addStyle(lbl,values[i][6],styles,stLen);
      LabelAndColumnLocationHandler(values[i][9],lbl,values[i][3]);
      }
    else if (values[i][0] == 'TextArea') {
      var lbl = app.createLabel(values[i][1]);
      var ta = app.createTextArea().setName(values[i][2]).setId(values[i][2]).setValue(theCase[values[i][2]]);
      addStyle(lbl,values[i][6],styles,stLen);
      addStyle(ta,values[i][7],styles,stLen);
      LabelAndColumnLocationHandler(values[i][9],ta,values[i][3],lbl);
      }
    else if (values[i][0] == 'DateBox') {
      var lbl = app.createLabel(values[i][1]);
      var db = app.createTextBox().setName(values[i][2]).setId(values[i][2]).setValue(theCase[values[i][2]]);
      addStyle(lbl,values[i][6],styles,stLen);
      addStyle(db,values[i][7],styles,stLen);
      LabelAndColumnLocationHandler(values[i][9],db,values[i][3],lbl);
      }
    else if (values[i][0] == 'FileUpload') {
      var lbl = app.createLabel(values[i][1]);
      //var fu = app.createTextBox().setValue('');
      for (var f = 0; f < dLastColumn; f++) {
         if (dKeys[0][f] == values[i][2]) {
             var fu = app.createAnchor('see attachment', dRange.getCell((caseRow+1),(f+1)).getValue())
             //fu.setValue(dRange.getCell((caseRow+1),(f+1)).getValue())
           }
          }
      addStyle(lbl,values[i][6],styles,stLen);
      addStyle(fu,values[i][7],styles,stLen);
      LabelAndColumnLocationHandler(values[i][9],fu,values[i][3],lbl);
      }
    else if (values[i][0] == 'CheckBox') {
      var lbl = app.createLabel(values[i][1]);
      var tb = app.createTextBox().setName(values[i][2]).setValue(theCase[values[i][2]]).setId(values[i][2]);
      addStyle(lbl,values[i][6],styles,stLen);
      addStyle(tb,values[i][7],styles,stLen);
      LabelAndColumnLocationHandler(values[i][9],tb,values[i][3],lbl);
      }
    else if (values[i][0] == 'RadioButton') {
      if (radioGroup !== values[i][2]) {
        radioGroup = values[i][2];
        var lbl = app.createLabel(values[i][1]);
        var tb = app.createTextBox().setName(values[i][2]).setValue(theCase[values[i][2]]).setId(values[i][2]);
        addStyle(lbl,values[i][6],styles,stLen);
        addStyle(tb,values[i][7],styles,stLen);
        LabelAndColumnLocationHandler(values[i][9],tb,values[i][3],lbl);
       }
      }
    else if (values[i][0] == 'Image') {
      var img = app.createImage(values[i][1]);
      LabelAndColumnLocationHandler(values[i][9],img,values[i][3]);
      }
    else if (values[i][0] == 'UserName') {
      var visible = true;
      if (values[i][4]=='No') {visible =false;}
      var lbl_Shortname = app.createLabel(values[i][1]).setVisible(visible);
      var tb_Shortname = app.createTextBox().setVisible(visible).setName(values[i][2]).setId(values[i][2]).setValue(theCase[values[i][2]]);
      addStyle(lbl_Shortname,values[i][7],styles,stLen);
      addStyle(tb_Shortname,values[i][7],styles,stLen);
      LabelAndColumnLocationHandler(values[i][9],tb_Shortname,values[i][3],lbl_Shortname);
      }
    else if (values[i][0] == 'ListBox') {
      var lbl = app.createLabel(values[i][1]);
      var tb = app.createTextBox().setName(values[i][2]).setValue(theCase[values[i][2]]).setId(values[i][2]);
      addStyle(lbl,values[i][6],styles,stLen);
      addStyle(tb,values[i][7],styles,stLen);
      LabelAndColumnLocationHandler(values[i][9],tb,values[i][3],lbl);
      }
    
//More types to come.  Next small block adds a Submit button and builds the app.

   }
  
  //******DEPRECIATED.  I moved all the validations on blank fields to a focus or change handler, but you could recover this if performance suffers.
  //Here we add a validation rule.  Long term, I want this added to each field.
 
 /*
 
  var btn_Validate = app.createButton("Validate").addClickHandler(submit_handler);
  vertP.add(btn_Validate);
 
 */
 
 var caseIDLbl = app.createLabel('Case ID: ');
 var caseStatusLbl = app.createLabel('Case Status: ');
 var caseIDVal = app.createLabel(caseID);
 var caseStatusVal = app.createLabel(theWorkflow.getStatus());
 
 addStyle(caseIDLbl,settings[16][1],styles,stLen);
 addStyle(caseStatusLbl,settings[16][1],styles,stLen);
 addStyle(caseIDVal,settings[17][1],styles,stLen);
 addStyle(caseStatusVal,settings[17][1],styles,stLen);
 
 var approvalsHeader = app.createLabel('Approvals');
 addStyle(approvalsHeader,settings[18][1],styles,stLen);
 
 vertP2.add(approvalsHeader)
       .add(app.createHorizontalPanel().add(caseIDLbl).add(caseIDVal))
       .add(app.createHorizontalPanel().add(caseStatusLbl).add(caseStatusVal));
 
 
 //NOW, we're going to start getting the workflow data.  Let's start by loading all the approval steps into the right hand panel.
 
  var stepCount = theSteps.length;
  
  for (var j = 0; j < stepCount; j++) {
     var thisStep = ReloadWorkflowObjectStep(theSteps[j]);
     var thisState = thisStep.getStepState();
     var thisStatus = thisStep.getStepStatus();
     var thisActionBy = thisStep.getStepActionBy();
     var thisActionTime = thisStep.getStepActionTime();
     var thisStepName = 'Unknown';
     var thisStepApprover = 'Unknown';
     
     if (thisActionBy == undefined) {thisActionBy = ''}
     if (thisActionTime == undefined) {thisActionTime = ''}
     
     for (var k = 1; k < wLastRow; k++) {
        if (thisState == wValues[k][0]) {
           thisStepName = wValues [k][1];
           thisStepApprover = wValues [k][2];
         }
       }  
       
     var thisStepRoleUsers = GetRoleMembers(thisStepApprover);   
     
     
     var thisStepRoleUserArray = ConvertArrayToObject(thisStepRoleUsers);
     
     
     var stepPanel = app.createFlowPanel();
                                          //.setStyleAttribute('border', '5px').setStyleAttribute('borderColor', 'black').setStyleAttribute('padding', '10px');
       
       switch (thisStatus) {
         
         case 'Pending':
           stepPanel.setStyleAttribute('backgroundColor', 'LightGray')
           break;
         case 'Active':
           stepPanel.setStyleAttribute('backgroundColor', 'CornflowerBlue')
           break;
         case 'Rejected':
           stepPanel.setStyleAttribute('backgroundColor', 'Crimson')
           break;
         case 'Approved':
           stepPanel.setStyleAttribute('backgroundColor', 'SeaGreen')
           break;
         default:
           break;
       }
     
     var user = Session.getActiveUser().toString();
     var hidUserName = app.createTextBox().setValue(user).setName('UserName').setVisible(false);
     
     var hidStepState = app.createTextBox().setValue(thisState).setName('State').setVisible(false);
     var hidCaseId = app.createTextBox().setValue(caseID).setName('CaseID').setVisible(false);
     var hidRandom = app.createTextBox().setValue(caseRandom).setName('Random').setVisible(false);
     
     
     var dispStepNameLabel = app.createLabel('Step Name: ');
     var dispStepName = app.createLabel(thisStepName);
     var dispStepStatusLabel = app.createLabel('Step Status: ');
     var dispStepStatus = app.createLabel(thisStatus);
     var dispStepApproverLabel = app.createLabel('Step Approver: ');
     var dispStepApprover = app.createLabel(thisStepApprover);
     var dispStepActionByLabel = app.createLabel('Action taken by: ');
     var dispStepActionBy = app.createLabel(thisActionBy);
     var dispStepActionTimeLabel = app.createLabel('Action Time: ');
     var dispStepActionTime = app.createLabel(thisActionTime);
     
     addStyle(dispStepNameLabel,settings[16][1],styles,stLen);
     addStyle(dispStepName,settings[17][1],styles,stLen);
     addStyle(dispStepStatusLabel,settings[16][1],styles,stLen);
     addStyle(dispStepStatus,settings[17][1],styles,stLen);
     addStyle(dispStepApproverLabel,settings[16][1],styles,stLen);
     addStyle(dispStepApprover,settings[17][1],styles,stLen);
     addStyle(dispStepActionByLabel,settings[16][1],styles,stLen);
     addStyle(dispStepActionBy,settings[17][1],styles,stLen);
     addStyle(dispStepActionTimeLabel,settings[16][1],styles,stLen);
     addStyle(dispStepActionTime,settings[17][1],styles,stLen);
     
     var approveHandler = app.createServerHandler('ApproveActionButtonHandler').addCallbackElement(hidStepState)
                                                                             .addCallbackElement(hidCaseId)
                                                                             .addCallbackElement(hidRandom)
                                                                             .addCallbackElement(hidUserName);
                                                                             
     //var approveClientHandler = app.createClientHandler().forEventSource().setEnabled(false);
     
                                                                             
     var rejectHandler = app.createServerHandler('RejectActionButtonHandler').addCallbackElement(hidStepState)
                                                                             .addCallbackElement(hidCaseId)
                                                                             .addCallbackElement(hidRandom)
                                                                             .addCallbackElement(hidUserName);                                                                       
                                                                             
                                                                             
     var approveButton = app.createButton('Approve').addClickHandler(approveHandler).setId('appBtn'+j);
     var rejectButton = app.createButton('Reject').addClickHandler(rejectHandler).setId('rejBtn'+j);
     
     approveButton.addClickHandler(app.createClientHandler().forEventSource().setEnabled(false).forTargets(rejectButton).setEnabled(false).setVisible(false));
     rejectButton.addClickHandler(app.createClientHandler().forEventSource().setEnabled(false).forTargets(approveButton).setEnabled(false).setVisible(false));
     
     
     
     if (user in thisStepRoleUserArray == false || thisStatus !== 'Active' || theWorkflow.getStatus() == 'Rejected') {
           approveButton.setEnabled(false).setVisible(false);
           rejectButton.setEnabled(false).setVisible(false)
       }
     
     //addStyle(vertP2,settings[13][1],styles,stLen);
     addStyle(stepPanel,settings[13][1],styles,stLen);
     
     
     stepPanel.add(hidStepState)
              .add(hidUserName)
              .add(hidCaseId)
              .add(hidRandom)
              .add(app.createHorizontalPanel().add(dispStepNameLabel).add(dispStepName))  //start label changes
              .add(app.createHorizontalPanel().add(dispStepStatusLabel).add(dispStepStatus))
              .add(app.createHorizontalPanel().add(dispStepApproverLabel).add(dispStepApprover))
              .add(app.createHorizontalPanel().add(dispStepActionByLabel).add(dispStepActionBy))
              .add(app.createHorizontalPanel().add(dispStepActionTimeLabel).add(dispStepActionTime))  //end label changes
              .add(approveButton)
              .add(rejectButton);
     
     vertP2.add(stepPanel);
     
     }
     
     
  /*   
  Well, I'm stuck in an airport, so let's add a comments box!
  The goal here is to get the user name, timestamp, and a text area.  We're gonna need a bigger boat, so let's add a class and methods to store all this crap
  in a custom object.
  */
 
  var commentFlow = app.createVerticalPanel();

  var commentHeader = app.createLabel('Comments and Attachments');
  addStyle(commentHeader,settings[18][1],styles,stLen);
  
  commentFlow.add(commentHeader);

  
  //This section handles loading any existing comments.
  
  var caseCommentsJSON = dSheet.getRange((caseRow+1), 8,1,1).getValue();
  
  
  if (caseCommentsJSON !== '') {
  
  var caseComments = ReloadAllComments(caseCommentsJSON);
  
  var allComments = caseComments.getComment();
  var commentLength = caseComments.getCommentCount();
  
  for (var l = 0; l < commentLength; l++) {
  
     var thisCommentFlow = app.createVerticalPanel();
     var thisComment = ReloadOneComment(allComments[l]);
     var commentTextArea = app.createTextArea().setValue(thisComment.getCommentText()).setEnabled(false);
     
     if (thisComment.getAttachmentPath() == undefined) {var commentAttachment = app.createLabel('No attachment');}
     else {var commentAttachment = app.createAnchor('see attachment', thisComment.getAttachmentPath());};
     
     var commentTimeLabel = app.createLabel('Entered on: ');
     var commentTimeValue = app.createLabel(thisComment.getTimeDate());
     var commentUser = app.createLabel('Entered by: ');
     var commentUserValue = app.createLabel(thisComment.getUser());
     
     addStyle(thisCommentFlow,settings[13][1],styles,stLen);
     addStyle(commentTextArea,settings[15][1],styles,stLen);
     addStyle(commentTimeLabel,settings[16][1],styles,stLen);
     addStyle(commentTimeValue,settings[17][1],styles,stLen);
     addStyle(commentUser,settings[16][1],styles,stLen);
     addStyle(commentUserValue,settings[17][1],styles,stLen);
     
     thisCommentFlow.add(commentTextArea)
                .add(commentAttachment)
                .add(app.createHorizontalPanel().add(commentTimeLabel).add(commentTimeValue))
                .add(app.createHorizontalPanel().add(commentUser).add(commentUserValue));
     
     commentFlow.add(thisCommentFlow);
     }
     
  addStyle(commentFlow,settings[12][1],styles,stLen);
  
  }
  
  
  
  //This section of comments covers adding new comments. :)
  
  var commentFormPanel = app.createFormPanel();
  var commentAddPanel = app.createVerticalPanel();
  
  var addLabel = app.createLabel('Add a new Comment:');
  addStyle(addLabel,settings[19][1],styles,stLen);
  var addCommentTextArea = app.createTextArea().setName('commentText').setId('commentText');
  addStyle(addCommentTextArea,settings[20][1],styles,stLen);
  var addCommentAttachment = app.createFileUpload().setName('fileLink');
  var addCaseRowBox = app.createTextBox().setValue(caseRow).setVisible(false).setId('theRow').setName('theRow');
  
  /*
 
  var commentHandler = app.createServerHandler('AddCommentHandler').addCallbackElement(addCommentTextArea)
                                                                 .addCallbackElement(addCaseRowBox)
                                                                 .addCallbackElement(addCommentAttachment)
                                                                 .addCallbackElement(hidUserName);
   */
   
  //var disableHandler = app.createClientHandler().forEventSource().setEnabled(false);
 
  var addCommentButton = app.createSubmitButton().setText('Add Comment');
                                                 //.addClickHandler(commentHandler)
                                                 //.addClickHandler(disableHandler);
                                                 
  var hidUserName2 = app.createTextBox().setValue(user).setName('UserName').setVisible(false); //fix line

  commentAddPanel.add(addLabel)
                 .add(addCommentTextArea)
                 .add(addCommentAttachment)
                 .add(addCommentButton)
                 .add(addCaseRowBox)
                 .add(hidUserName2);
                  
  commentFormPanel.add(commentAddPanel);                  
  commentFlow.add(commentFormPanel);
 
  vertP3.add(commentFlow);
  
 
  //Almost done!  This just throws a submit button on the whole thing, and then gives you the form.  Note the submit button comes from a parameter.
        
  var btn_Submit = app.createSubmitButton(settings[2][1]).setId('Submit').setEnabled(false);
  vertP.add(btn_Submit);
  colPanel.add(vertP).add(vertP2).add(vertP3);
  allPanel.add(headerP).add(colPanel);
  formP.add(allPanel);
  app.add(formP);

  if (logFilesOn == true && logFileLevel >= 1) {Logger.log('end function: doGet')};

  return app;
  
  
  
}


function doPost (eventInfo) {

  var curTime = new Date();
  var ss = SpreadsheetApp.openById(ssKey);
  var dSheet = ss.getSheetByName('RequestData');
  var dRange = dSheet.getDataRange();
  var app = UiApp.createApplication();
  var par = eventInfo.parameter;
  
  
  //Logger.log('Parameters: '+JSON.stringify(par));
  
  var commentText = par.commentText;
  var fileSource = par.fileLink;
  var theRow = par.theRow;
  var user = par.UserName;
  
  //Logger.log('File Source value is: '+fileSource.length);
  
  var fileLen = fileSource.length;
  
  theRow = ++theRow;
  
  //I'm gonna go ahead and create the Google Doc here:
  var setSheet = ss.getSheetByName('Settings');
  var settings = setSheet.getDataRange().getValues();
  var gFolder = DocsList.getFolder(settings[4][1]);
  if (fileLen > 0) {
     var gDoc = DocsList.createFile(fileSource);
     if (gFolder !== '') {gDoc.addToFolder(gFolder)};
     var fileLink = gDoc.getUrl();
   }
  
  var commentNumber = 1;
  
  
  var oldCommentsJSON = dSheet.getRange(theRow,8).getValue();
  
  var allCommentsArray = new Array();
  
  Logger.log('old comments JSON: '+oldCommentsJSON);
  

  if (oldCommentsJSON !== '') {
  
    //Logger.log('Any existing comments?  The answer is yes.');
    var oldComments = ReloadAllComments(oldCommentsJSON);
    var numOfComments = oldComments.getCommentCount();
    //Logger.log('num of comments: '+numOfComments);
    ++numOfComments;
    commentNumber = numOfComments;
    allCommentsArray = oldComments.getComment();
   }
   
  
   
  var thisComment = new UpdateOneComment().setUser(user)
                                    .setCommentNumber(commentNumber)
                                    .setTimeDate(curTime)
                                    .setCommentText(commentText)
                                    .setAttachmentPath(fileLink);
                                    
  var thisCommentString = JSON.stringify(thisComment);
                                    
  
  allCommentsArray.push(thisCommentString);
  
  
  var newComments = new UpdateAllComments().setCommentCount(commentNumber)
                                           .setComment(allCommentsArray);
  
  var newCommentsString = JSON.stringify(newComments);
  
  //Logger.log('new comment string: '+newCommentsString);

  
  
  dSheet.getRange(theRow,8).setValue(newCommentsString);
  
  
  return app;
  


}



/*
this is a mini-function that adds styles to any element.  It splits the values from the 'Fields' table by commas and applies every style listed.
'obj' is the UI element.
'style' is the style name'
'styles' is the entire style table.  If this caused performance issues, it could be cached.
'stLen' is the length of the style table, used to set a length for the for loop.

*/



function addStyle(obj, style, styles, stLen) {
    
    if (logFilesOn == true && logFileLevel >= 1) {Logger.log('call function: addStyle')};
    
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
This function handles both where the label appears next to an input element, and where the element and label appear on the form.

Although I really only want to allow one location (left) and one label location (top).  Anything else looks messy to me.

panel: where the element goes
inputField: the actual UI element we are working with.
location: where the label goes. optional.  Doesn't look optional, but it should be.
label: the label for the UI element.  Optional.


*/



function LabelAndColumnLocationHandler(panel,inputField,location,label) {

    if (logFilesOn == true && logFileLevel <= 1) {Logger.log('call function: LabelAndColumnLocationHandler')};
   
    var app = UiApp.getActiveApplication();
    
    
    var headerP = app.getElementById('headerP');
    var vertP = app.getElementById('vertP1');
    var vertP2 = app.getElementById('vertP2');
    var vertP3 = app.getElementById('vertP3');
    var colPanel = app.getElementById('colPanel');
    
    //Logger.log(vertP);
    //Logger.log(inputField);
    //Logger.log(label);
    
    var addToPanel = vertP;
    
    switch (panel) {
      case 'left':
        addToPanel = vertP;
        break;
      case 'middle':
        addToPanel = vertP;
        break;
      case 'right':
        addToPanel = vertP;
        break;
      case 'header':
        addToPanel = headerP;
        break;
      case 'left':
        addToPanel = vertP;
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
I'm not going to waste a lot of mips talking about this guy, because he should be gutted and just use the result from the CaseRowFinder Class.

Basically, he finds the case and loads the JSON containing all the case data.  The find part should be replaced with CaseRowFinder.

Case: the CaseID
Random: the random number that gets assigned to the case.  Prevents simple URL spoofing to look at other cases.

*/



function CaseLoader(Case,Random) {

  if (logFilesOn == true && logFileLevel <= 1) {Logger.log('start function: CaseLoader')};
  
  //first, we need the entered cases:
  var ss = SpreadsheetApp.openById(ssKey);
  var dSheet = ss.getSheetByName('RequestData');
  var dRange = dSheet.getDataRange();
  var dLastRow = dRange.getLastRow();
  var dLastColumn = dRange.getLastColumn();
  var dKeys = dSheet.getRange(2, 1, 1, dLastColumn).getValues();
  var dValues = dRange.getValues();
  
  //and an empty object for the case:
  
  var theCase = new Object;
 
  //Logger.log('I was called!');
 
  //next, let's find the case:

  for (var i = 2; i < dLastRow; i++) {
    
    /*
    Logger.log('Compare -> '+Case+' : to : '+dValues[i][1]);
    Logger.log('Compare -> '+Random+' : to : '+dValues[i][3]);
    */
    
    
     if (Case == dValues[i][1] && Random == dValues[i][3]) {
        theCase = JSON.parse(dValues[i][2]);
        
        
        //Can I insert a break here?  Look into that.
       }
     }
  return theCase;
  
  if (logFilesOn == true && logFileLevel <= 1) {Logger.log('end function: CaseLoader')};
}



/*  
THIS GROUP OF CLASSES IS FOR WORKING WITH THE WORKFLOW OBJECT AT THE HEADER OR MASTER LEVEL.

These are a collection of Methods for working with the JSON Class that contains workflow items.  

This stuff is advanced javascript.  Either you know what you're looking at and don't need my comments, or you should probably just back away slowly.


*/

function UpdateWorkflowObject() {
  if (logFilesOn == true && logFileLevel <= 2) {Logger.log('start class: UpdateWorkflowObject')};

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
   if (logFilesOn == true && logFileLevel <= 2) {Logger.log('start class: ReloadWorkflowObject')};
   
   var workflowObject = JSON.parse(workflowObj);
   
   UpdateWorkflowObject.call(workflowObject);
   
   return workflowObject;
}

/*
THIS GROUP IS FOR WORKING WITH YOUR WORKFLOW AT THE STEP OR ACTION LEVEL.
*/

function UpdateWorkflowStep() {

  if (logFilesOn == true && logFileLevel <= 2) {Logger.log('start class: UpdateWorkflowStep')};

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

   if (logFilesOn == true && logFileLevel <= 2) {Logger.log('start class: ReloadWorkflowStep')};

   var workflowStepObj = JSON.parse(workflowStep);
   
   UpdateWorkflowStep.call(workflowStepObj);
   
   return workflowStepObj;
}





/*

-----------------DEPRECIATED---------------------


function AdvanceWorkflowStatus (workflowObj) {

   if (logFilesOn == true && logFileLevel <= 1) {Logger.log('start function: AdvanceWorkflowStatus')};

   //a dumb function, made to manually advance workflows to the next active state.
   //Really a Proof of Concept for methods for re-loading and changing the workflow JSON.
   
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
  //state = 100;
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


/*

----------DEPRECIATED-----------------

function FindAndReloadWorkflowObject(CaseID) {

  if (logFilesOn == true && logFileLevel <= 1) {Logger.log('start function: FindAndReloadWorkflowObject')};

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
     
     var json = JSON.parse(workflowObj);
     
     var theWorkflow = UpdateWorkflowObject.call(json);
     
     }
   }

  

  return theWorkflow;
  
}


*/


//YES, I KNOW I'M BEING LAZY HERE.  I should not have a separate CaseLoader and Workflow Loader.  I'm doubling the resources to call the Case.

/*
I'm not going to waste a lot of mips talking about this guy, because he should be gutted and just use the result from the CaseRowFinder Class.

Basically, he finds the case and loads the JSON containing all the case data.  The find part should be replaced with CaseRowFinder.

Case: the CaseID
Random: the random number that gets assigned to the case.  Prevents simple URL spoofing.

*/

function WorkflowLoader(Case,Random) {

  if (logFilesOn == true && logFileLevel <= 1) {Logger.log('start function: WorkflowLoader')};

  //first, we need the entered cases:
  var ss = SpreadsheetApp.openById(ssKey);
  var dSheet = ss.getSheetByName('RequestData');
  var dRange = dSheet.getDataRange();
  var dLastRow = dRange.getLastRow();
  var dLastColumn = dRange.getLastColumn();
  var dKeys = dSheet.getRange(2, 1, 1, dLastColumn).getValues();
  var dValues = dRange.getValues();
  
  //and an empty object for the case:
 
  Logger.log('workflow was called!');
 
  //next, let's find the case:

  for (var i = 2; i < dLastRow; i++) {
    
    Logger.log('Compare -> '+Case+' : to : '+dValues[i][1]);
    Logger.log('Compare -> '+Random+' : to : '+dValues[i][3]);
    
     if (Case == dValues[i][1] && Random == dValues[i][3]) {
        var theWorkflow = ReloadWorkflowObject(dValues[i][4]);
        
        
        //Can I insert a break here?  Look into that.
       }
     }
  return theWorkflow;
  
}


function ApproveActionButtonHandler(eventInfo) {
 
   if (logFilesOn == true && logFileLevel <= 1) {Logger.log('start handler: ApproveActionButtonHandler')};
  
   var app = UiApp.createApplication();
   
   var curTime = new Date();
   
   var par = eventInfo.parameter;
   
   var caseID = par.CaseID;
   var random = par.Random;
   var state = par.State;
   var user = par.UserName;
   
   var theWorkflow = WorkflowLoader(caseID,random); 
   var theSteps = theWorkflow.getWorkflowSteps();
   var theFlow = theWorkflow.getWorkflow();
   var stepsLen = theSteps.length;
   
   var theRow = CaseRowFinder(caseID,random);
   var newSteps = new Array();
   
   for (var i = 0; i < stepsLen; i++) {
      var thisStep = ReloadWorkflowObjectStep(theSteps[i]);
      var stepState = thisStep.getStepState();
      
      var updateStep = false;
      if (stepState == state) {updateStep = true;}
      
      switch (updateStep) {
         
         case true:
      
         var newStep = new UpdateWorkflowStep().setStepState(state)
                                               .setStepStatus('Approved')
                                               .setStepActionBy(user)
                                               .setStepActionTime(curTime);
                                               
         var jsonStep = JSON.stringify(newStep);
         var newCaseState = theFlow[(i+1)];
         //this bit right here is shitty.  I should never test for undefined.
         if (newCaseState == undefined) {newCaseState = 999;
                                         theWorkflow.setStatus('Approved');
                                        }
         break;
         
         case false:
         
         var jsonStep = JSON.stringify(thisStep);
         break;
         
         default:
         var jsonStep = JSON.stringify(thisStep);
         break;
         
         
         
         }
         
         
       
       newSteps.push(jsonStep);
       
       }
   theWorkflow.setWorkflowSteps(newSteps).setState(newCaseState);
   
   var ss = SpreadsheetApp.openById(ssKey);
   var dSheet = ss.getSheetByName('RequestData');
   
   //NEW SECTION FOR UPDATES TO USERS
   
   //Logger.log('new case state: '+newCaseState);
   //Logger.log('THE ROW: '+theRow);
   
   NotifyNextCaseUser(newCaseState,theRow,'Approved');
   UpdateCaseValues(state,theRow);
   
   
   var json1 = JSON.stringify(theWorkflow);

  
   dSheet.getRange(theRow+1, 5).setValue(json1);

   
   return app;
  
}





function RejectActionButtonHandler(eventInfo) {

   if (logFilesOn == true && logFileLevel <= 1) {Logger.log('start handler: RejectActionButtonHandler')};

   var app = UiApp.createApplication();
   
   
   var curTime = new Date();
   
   var par = eventInfo.parameter;
   
   var caseID = par.CaseID;
   var random = par.Random;
   var state = par.State;
   var user = par.UserName;
   
   var theWorkflow = WorkflowLoader(caseID,random);
   var theSteps = theWorkflow.getWorkflowSteps();
   var theFlow = theWorkflow.getWorkflow();
   var stepsLen = theSteps.length;
   
   var theRow = CaseRowFinder(caseID,random);
   var newSteps = new Array();
   
   for (var i = 0; i < stepsLen; i++) {
      var thisStep = ReloadWorkflowObjectStep(theSteps[i]);
      var stepState = thisStep.getStepState();
      
      var updateStep = false;
      if (stepState == state) {updateStep = true;}
      
      switch (updateStep) {
         
         case true:
      
           var newStep = new UpdateWorkflowStep().setStepState(state)
                                                 .setStepStatus('Rejected')
                                                 .setStepActionBy(user)
                                                 .setStepActionTime(curTime);
                                               
           var jsonStep = JSON.stringify(newStep);
           var newCaseState = theFlow[(i+1)];
           if (newCaseState == undefined) {newCaseState = 999;
                                           theWorkflow.setStatus('Rejected');
                                          }
         break;
         
         case false:
         
           var jsonStep = JSON.stringify(thisStep);
         break;
         
         default:
           var jsonStep = JSON.stringify(thisStep);
         break;
         
         
         
         }
         
         
       
       newSteps.push(jsonStep);
       
       }
   theWorkflow.setWorkflowSteps(newSteps).setState(newCaseState);
   
   var ss = SpreadsheetApp.openById(ssKey);
   var dSheet = ss.getSheetByName('RequestData');

   NotifyNextCaseUser(newCaseState,theRow,'Rejected');
   UpdateCaseValues(state,theRow);
   
  
   var json1 = JSON.stringify(theWorkflow);
   
  
   dSheet.getRange(theRow+1, 5).setValue(json1);

   
   return app;
  
}




//TAKE NOTE - I need to replace all the case lookups with this bad boy.

/*
A little Class that finds the spreadsheet row where a case lives, based on the CaseID and the random number.

*/

function CaseRowFinder(Case,Random) {

  if (logFilesOn == true && logFileLevel <= 1) {Logger.log('start function: CaseRowFinder')};

  //first, we need the entered cases:
  var ss = SpreadsheetApp.openById(ssKey);
  var dSheet = ss.getSheetByName('RequestData');
  var dRange = dSheet.getDataRange();
  var dLastRow = dRange.getLastRow();
  var dLastColumn = dRange.getLastColumn();
  var dKeys = dSheet.getRange(2, 1, 1, dLastColumn).getValues();
  var dValues = dRange.getValues();
  

  for (var i = 2; i < dLastRow; i++) {
    
    
     if (Case == dValues[i][1] && Random == dValues[i][3]) {
        var theRow = i;
        
        
        //Can I insert a break here?  Look into that.
       }
     }
  return theRow;
  
}
/*
The purpose of this script escapes me at the moment, but I'm sure it was important.

Um...it looks like it is called at the beginning of the reload step, to see if the case is approved from an overall standpoint.  
It also forces some stuff to get the case moving to an active state.

At least I used the CaseRowFinder here.  Although I should be loading the row once, not calling the whole spreadsheet each time.
*/


function ApprovalStatusCheck(caseID,random) {

   if (logFilesOn == true && logFileLevel <= 1) {Logger.log('start function: ApprovalStatusCheck')};

   var app = UiApp.createApplication();
   var theWorkflow = WorkflowLoader(caseID,random);
   var theSteps = theWorkflow.getWorkflowSteps();
   
   var stepsLen = 0;  //start fix
   if (theSteps !== undefined) {
      stepsLen = theSteps.length;
     }  //end fix
   var theRow = CaseRowFinder(caseID,random);
   var newSteps = new Array();
   
   var workflowState = theWorkflow.getState();
   var workflowPath = theWorkflow.getWorkflow();
   
   var lastState = 999;  //start fix
   if (workflowPath !== undefined) {
      var lastState = workflowPath[workflowPath.length-1];
   
        if (workflowPath[0] > workflowState) {theWorkflow.setState(workflowPath[0]);
                                              workflowState = workflowPath[0]};
     }
     else if (workflowPath == undefined) {theWorkflow.setState('999').setStatus('No Workflow');};   //end fix
   
  
   for (var i = 0; i < stepsLen; i++) {
      var thisStep = ReloadWorkflowObjectStep(theSteps[i]);
      var stepState = thisStep.getStepState();
      var stepStatus = thisStep.getStepStatus();
      
      if (stepState == workflowState && stepStatus == 'Pending') {thisStep.setStepStatus('Active')}
      if (stepStatus == 'Rejected') {theWorkflow.setStatus('Rejected').setState(999);}
      if (stepState == 'Approved' && stepState == lastState) {theWorkflow.setStatus('Approved').setState(999)}
         
      var jsonStep = JSON.stringify(thisStep);
         
      newSteps.push(jsonStep);           
         
     }
         
       
       
   theWorkflow.setWorkflowSteps(newSteps);
   
   var ss = SpreadsheetApp.openById(ssKey);
   var dSheet = ss.getSheetByName('RequestData');
  
   var json1 = JSON.stringify(theWorkflow);

  
   dSheet.getRange(theRow+1, 5).setValue(json1);
   
   if (logFilesOn == true && logFileLevel <= 1) {Logger.log('end function: ApprovalStatusCheck')};

  
}

/*
This sends an automatic email to the next guy in the approval flow.

It uses a lot of the same stuff in the 'ScheduleHandler.gs' function Class from the main script to add variables from the case back to the email.

*/

function NotifyNextCaseUser(newCaseState,theRow,theAction) {

   if (logFilesOn == true && logFileLevel <= 1) {Logger.log('start function: NotifyNextCaseUser')};


   var ss = SpreadsheetApp.openById(ssKey);
   
   var wSheet = ss.getSheetByName('Workflow');
   var wRange = wSheet.getDataRange();
   var wLastRow = wRange.getLastRow();
   var wValues = wRange.getValues();
   
   var dSheet = ss.getSheetByName('RequestData');
   var dRange = dSheet.getDataRange();
   var dValues = dRange.getValues();
   var dLastRow = dRange.getLastRow();
   var dLastColumn = dRange.getLastColumn();
   var dKeys = dSheet.getRange(2, 1, 1, dLastColumn).getValues();
   
   
   var theCaseID = dValues[theRow][1];
   
   if (theAction == 'Approved') {
   
   for (var i = 1; i < wLastRow; i++) {
      if (wValues[i][0] == newCaseState) {
        
         var to = GetRoleMembers(wValues[i][2]);  //still only seems to send to the first person in the array.  Need more testing.
         
           var subject = (theCaseID+' '+wValues[i][6]);
           var body = wValues[i][7];
           body = body+('\n');
           var addFields = wValues[i][5].split(',');
           var addsLen = addFields.length;
             for (var l = 0; l < addsLen; l++) {
                //key index lookup
                for (var m = 0; m < dLastColumn; m++) {
                  if (dKeys[0][m] == addFields[l]) {
                      body = body+('\n');
                      body = body+(addFields[l]+': '+dValues[theRow][m]);
             }
            }
           }
           MailApp.sendEmail(to, subject, body);
               
        }    
   
     }
   }
   
  
   if (theAction == 'Rejected') {
   
   for (var i = 1; i < wLastRow; i++) {
      if (wValues[i][0] == newCaseState) {
        
         var to = GetRoleMembers(wValues[i][2]);  //still only seems to send to the first person in the array.  Need more testing.
         
           var subject = (theCaseID+' '+wValues[i][10]);
           var body = wValues[i][11];
           body = body+('\n');
           var addFields = wValues[i][9].split(',');
           var addsLen = addFields.length;
             for (var l = 0; l < addsLen; l++) {
                //key index lookup
                for (var m = 0; m < dLastColumn; m++) {
                  if (dKeys[0][m] == addFields[l]) {
                      body = body+('\n');
                      body = body+(addFields[l]+': '+dValues[theRow][m]);
             }
            }
           }
           MailApp.sendEmail(to, subject, body);
               
        }    
      }

   }
  
  
  
  
  
  
  if (logFilesOn == true && logFileLevel <= 1) {Logger.log('end function: NotifyNextCaseUser')};
  
}

/*
I guess part of clearly naming your functions and variables is that it makes comments silly.  Well, this updates values within the case on the spreadsheet.



*/



function UpdateCaseValues(ApprovedCaseState,theRow,theAction) {

   if (logFilesOn == true && logFileLevel <= 1) {Logger.log('start function: UpdateCaseValues')};

   var ss = SpreadsheetApp.openById(ssKey);
   
   var wSheet = ss.getSheetByName('Workflow');
   var wRange = wSheet.getDataRange();
   var wLastRow = wRange.getLastRow();
   var wValues = wRange.getValues();
   
   var dSheet = ss.getSheetByName('RequestData');
   var dRange = dSheet.getDataRange();
   var dLastRow = dRange.getLastRow();
   var dLastColumn = dRange.getLastColumn();
   var dKeys = dSheet.getRange(2, 1, 1, dLastColumn).getValues();
   
   var tSheet = ss.getSheetByName('Tests');
   var tRange = tSheet.getDataRange();
   var tValues = tRange.getValues();
   var tLastRow = tRange.getLastRow();
   
   if(theAction == 'Approved') {
   
   for (var i = 1; i < wLastRow; i++) {  //for each value in the Workflow spreadsheet
      if (wValues[i][0] == ApprovedCaseState) {
         for (var j = 1; j < tLastRow; j++) {
            if (wValues[i][4] == tValues[j][0] && tValues[j][1] == 'Update') {
                  for (var k = 0; k < dLastColumn; k++) {
                     if(dKeys[0][k] == tValues[j][2]) {
                        dRange.getCell((theRow+1), (k+1)).setValue(tValues[j][4]);
                       }
                     }
                  }
                }
              }
            }
     }
   
   if(theAction == 'Rejected') {
   
   for (var i = 1; i < wLastRow; i++) {  //for each value in the Workflow spreadsheet
      if (wValues[i][0] == ApprovedCaseState) {
         for (var j = 1; j < tLastRow; j++) {
            if (wValues[i][8] == tValues[j][0] && tValues[j][1] == 'Update') {
                  for (var k = 0; k < dLastColumn; k++) {
                     if(dKeys[0][k] == tValues[j][2]) {
                        dRange.getCell((theRow+1), (k+1)).setValue(tValues[j][4]);
                       }
                     }
                  }
                }
              }
            }
       }
   
   

  if (logFilesOn == true && logFileLevel <= 1) {Logger.log('end function: UpdateCaseValues')};
  
}

function GetRoleMembers(role) {
  
  var roleMembers = new Array();
  
  var ss = SpreadsheetApp.openById(ssKey);
  var rSheet = ss.getSheetByName('Roles');
  var rRange = rSheet.getDataRange();
  var rLastRow = rRange.getLastRow();
  var rValues = rRange.getValues();
  
  for (var i = 1; i < rLastRow; i++) {
    if (role == rValues[i][0]) {
      
      roleMembers.push(rValues[i][1]);
    }
  }
  
  if(roleMembers.length == 0) {roleMembers = role}
  
  return roleMembers;
  
}


function ConvertArrayToObject(array)
{
  var object = {};
  for(var i = 0 ; i < array.length; i++)  {
    object[array[i]]='';
   } 
  return object;
}


/*
These are the classes and methods we use for working with the JSONs related to comments.

All of the comments are stored in a JSON, which is in turn stored in a larger JSON that holds all the comments and keeps a little metadata on the comments.

Right now there is no method for deleting comments.

*/


function UpdateOneComment() {

  this.setCommentNumber = function (commentNumber) {this.commentNumber = commentNumber; return this;};
  this.getCommentNumber = function () {return this.commentNumber;};
  this.setUser = function (user) {this.user = user; return this;};
  this.getUser = function () {return this.user;};
  this.setTimeDate = function (timeDate) {this.timeDate = timeDate; return this;};
  this.getTimeDate = function () {return this.timeDate;};
  this.setCommentText = function (commentText) {this.commentText = commentText; return this;};
  this.getCommentText = function () {return this.commentText;};
  this.setAttachmentPath = function (attachmentPath) {this.attachmentPath = attachmentPath; return this;};
  this.getAttachmentPath = function () {return this.attachmentPath;};
  
}

function ReloadOneComment(comment) {

   var thisComment = JSON.parse(comment);
   
   UpdateOneComment.call(thisComment);
   
   return thisComment;
}


function UpdateAllComments() {
  this.setComment = function (comment) {this.comment = comment; return this;};
  this.getComment = function () {return this.comment;};
  this.setCommentCount = function (commentCount) {this.commentCount = commentCount; return this;};
  this.getCommentCount = function () {return this.commentCount;};
  
}

function ReloadAllComments(allComments) {

   var commentObject = JSON.parse(allComments);
   
   UpdateAllComments.call(commentObject);
   
   return commentObject;
}





function AddCommentHandler (eventInfo) {

  var curTime = new Date();
  var ss = SpreadsheetApp.openById(ssKey);
  var dSheet = ss.getSheetByName('RequestData');
  var dRange = dSheet.getDataRange();
  var app = UiApp.createApplication();
  var par = eventInfo.parameter;
  
  
  Logger.log('Parameters: '+JSON.stringify(par));
  
  var commentText = par.commentText;
  var fileSource = par.fileLink;
  var theRow = par.theRow;
  var user = par.UserName;
  
  Logger.log(fileSource);
  
  theRow = ++theRow;
  
  //I'm gonna go ahead and create the Google Doc here:
  var setSheet = ss.getSheetByName('Settings');
  var settings = setSheet.getDataRange().getValues();
  var gFolder = DocsList.getFolder(settings[4][1]);
  //var gDoc = DocsList.createFile(fileSource);
  //if (gFolder !== '') {gDoc.addToFolder(gFolder)};
  //var fileLink = gDoc.getUrl();
  
  
  var commentNumber = 1;
  
  
  var oldCommentsJSON = dSheet.getRange(theRow,8).getValue();
  
  var allCommentsArray = new Array();
  
  Logger.log('old comments JSON: '+oldCommentsJSON);
  

  if (oldCommentsJSON !== undefined) {
  
    Logger.log('Any existing comments?  The answer is yes.');
    var oldComments = ReloadAllComments(oldCommentsJSON);
    var numOfComments = oldComments.getCommentCount();
    Logger.log('num of comments: '+numOfComments);
    ++numOfComments;
    commentNumber = numOfComments;
    allCommentsArray = oldComments.getComment();
   }
   
  
   
  var thisComment = new UpdateOneComment().setUser(user)
                                    .setCommentNumber(commentNumber)
                                    .setTimeDate(curTime)
                                    .setCommentText(commentText);
                                    //.setAttachmentPath(fileLink);
                                    
  var thisCommentString = JSON.stringify(thisComment);
                                    
  
  allCommentsArray.push(thisCommentString);
  
  
  var newComments = new UpdateAllComments().setCommentCount(commentNumber)
                                           .setComment(allCommentsArray);
  
  var newCommentsString = JSON.stringify(newComments);
  
  Logger.log('new comment string: '+newCommentsString);

  
  
  dSheet.getRange(theRow,8).setValue(newCommentsString);
  
  
  return app;
  


}