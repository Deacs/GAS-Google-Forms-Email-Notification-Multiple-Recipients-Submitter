// Author Shawn Hodgson
// http://etwasshawn.shawnjhodgson.com/2016/01/forms-send-form-submission-to-multiple
// Updated 8/9/2016 
// Code.gs START
// menu added on open

function onOpen() {
  FormApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Settings')   
      .addItem('Authorize', 'authorize')
      .addItem('Set Email', 'setEmailInfo')     
      .addToUi();
}

function testReturn(){
  return 'Hello';
}

function notherTest(){

 var body = 'Hello ' +  'Link to edit form submission: ' +
        '<a href='+ '"' + link + '">Click Here</a>'+ ' ' +
         msgBodyTable;  
 
}
//easily authorize the script to run from the menu
function authorize(){
var respEmail = Session.getActiveUser().getEmail();
MailApp.sendEmail(respEmail,"Form Authorizer", "Your form has now been authorized to send you emails");
}

//function setEmailInfo(){
  //need to use create template from file
//  var html = HtmlService.createHtmlOutputFromFile('Index').setHeight(500);
//  FormApp.getUi().showModalDialog(html, "Email Settings");
//}

function setEmailInfo(){
  //need to use create template from file
  var html = HtmlService.createTemplateFromFile('Index').evaluate().setHeight(750);
  FormApp.getUi().showModalDialog(html, "Email Settings");
}


//saving administrative settings as properties
function processForm(myform){  
  var rec = myform.toEmailTB;
  var cc = myform.ccEmailTB;
  var msg =  myform.messageTA;  
  var sub = myform.subjectTB;
  var addSubmitter = myform.addSubmitterCB;
  var secHeader = myform.secHeaderCB;
  var responLink = myform.responLinkCB;
  var includeEmpty = myform.includeEmptyCB;     
  
  setProperty('EMAIL_ADDRESS',rec);
  setProperty('CC_ADDRESS',cc);
  setProperty('EMAIL_SUBJECT',sub);
  setProperty('EMAIL_MESSAGE',msg);
  setProperty('ADD_SUBMITTER',addSubmitter);
  setProperty('SECTION_HEADER',secHeader);
  setProperty('RESPONSE_LINK',responLink);
  setProperty('INCLUDE_EMPTY',includeEmpty);
}

//setting script properties
function setProperty(key,property){
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty(key);
  if (!property){
    property = "";
  }
  scriptProperties.setProperty(key,property);
}

//getting script properties
function getProperty(property){
  var scriptProperties = PropertiesService.getScriptProperties();
  var savedProp = scriptProperties.getProperty(property);
  return  savedProp;
}



//function to put it all together
function controller(e){ 
  var response = e.response;   
  var emailTo = getProperty('EMAIL_ADDRESS');    
  var emailSubject = getProperty('EMAIL_SUBJECT'); 
  var message = getProperty('EMAIL_MESSAGE');  
  var cc =   getProperty('CC_ADDRESS');  
  var secHeader = getProperty('SECTION_HEADER'); 
  var includeEmpty = getProperty('INCLUDE_EMPTY'); 
  var addSubmitter = getProperty('ADD_SUBMITTER'); 
  var responseLink = getProperty('RESPONSE_LINK');      
  //get questions and responses   
  var resp = getResponse(response,secHeader,includeEmpty);
  //find email submitter
  var submitter = getSubmitter(response,addSubmitter);
  //format with html   
  var msgBodyTable = formatHTML(resp);
  //get response link
  var link = getLink(response,responseLink);
  //email  
  var body = emailBody(message, msgBodyTable, link);  
  sendEmail(emailTo,submitter,emailSubject,body,cc);

 
}

//get edit link
function getLink(response,responseLink){
   var link = " ";
   if (responseLink == 'true'){  
     link = response.getEditResponseUrl();
     return link;
   }
  else{
    return link;
  }  
}

//Get submitter if submitter exists and addSubmitter is true
function getSubmitter(response,addSubmitter){
  // var response = e.response
    Logger.log("In getsubmitter1");
   Logger.log(" e.response " +  response);
 // var formRespID = response.getID();
   // Logger.log("formRespID " + formRespID);
  //var formResp = getRespone(formRespID);
    //  Logger.log("formResp " + formResp);
  var noEmail = "";
  var email = "";

  if (addSubmitter == 'true'){  
 var itemRes = response.getItemResponses();
// var itemRes = formResp.getItemResponses();
  for (var i = 0; i < itemRes.length; i++){
    var respQuestion = itemRes[i].getItem().getTitle();
      Logger.log("respQuestion " + respQuestion);
    var index = itemRes[i].getItem().getIndex(); 
    var emailTest = respQuestion.toLowerCase();
    var regex = /.*email.*/;    
    if(regex.test(emailTest) == true){    
      email = email + "," + itemRes[i].getResponse(); 
      break; 
        Logger.log("email " + email);
    }   
  }     
    return email; 
  }
  return noEmail;  
};


function getResponseLink(){
 var destID = FormApp.getActiveForm().getDestinationId();
 var ssURL = SpreadsheetApp.openById(destID).getUrl();  
 return ssURL;
};

function emailBody(message, msgBodyTable, link){
  var body;
  
  if(link != ' '){ 
    body = message +  '<br> Link to edit form submission: ' + '<a href='+ '"' + link + '">Click Here</a>'+ '<br>' + msgBodyTable;
  }
  else{
    body = message +  msgBodyTable;
  }
  
  return body;

};
//function to send out mail
function sendEmail(emailRecipient,submitter,emailSubject,body,ccRecipient){
  //if submitter is not empty then append it to the emailRecipient values
  if (submitter != ""){
  emailRecipient = emailRecipient + "," + submitter;
  }
  MailApp.sendEmail(emailRecipient,emailSubject,"", {htmlBody: body, cc: ccRecipient}); 
}

//Function get form items and form responses. Builds and and returns an array of quesions: answer. 
function getResponse(response,secHeader,includeEmpty){
  var form = FormApp.getActiveForm();
  var items = form.getItems(); 
  var response = response;
  var itemRes = response.getItemResponses();
  var array = [];     

  for (var i = 0; i < items.length; i++){
    var question = items[i].getTitle();    
    var answer = "";   
    //include section headers and description in email only runs when user sets setHeader to true
      if (items[i].getType() == "SECTION_HEADER" && secHeader == 'true' ){
        var description = items[i].getHelpText();
        var title = items[i].getTitle();
        var regex = /^\s*(?:[\dA-Z]+\.|[a-z]\)|•)\s+/gm;
        description = description.replace(regex,"<br>");        
        array.push("<strong>" + title + "</strong><br>" + description);
        continue; 
      }
      
    //loop through to see if the form question title and the response question title matches. If so push to array, if not answer is left as "" 
    for (var j = 0; j < itemRes.length; j++){   
      var respQuestion = itemRes[j].getItem().getTitle();
      //itemRes[j].getResponse()
      if (question == respQuestion){        
        if(items[i].getType() == "CHECKBOX"){          
          var answer =  formatCheckBox(itemRes[j].getResponse());        
          break;
        }
        else{
        var answer = itemRes[j].getResponse();
        break;
        }
      }  
    } 
 
    //run this block of code if no empty responses are included
    if(includeEmpty != 'true'){
      if(answer != ""){ 
        array.push("<strong>" + question + "</strong>" + ": " + answer);
      }      
    }
    //run this block of clode if empty responses are included 
    else{
       array.push("<strong>" + question + "</strong>" + ": " + answer);
     } 
  }
  
  return array;
}

function formatCheckBox(chkBoxArray){
  
   for (var i = 0; i < chkBoxArray.length; i++){
     chkBoxArray[i] = "<br>" + chkBoxArray[i];
   }
  
  return chkBoxArray.join(" ");

}
//formats an array as a table
function formatHTML(array){
   
  var style = '"border-collapse: separate; border-spacing: 0 1em;"'; 
  //var style = '"color: green;"';
  var tableStart = '<br><br><html><body><table style='+style+'>';
  var tableEnd = "</table></body></html>";
  var rowStart = "<tr>";
  var rowEnd = "</tr>";
  var cellStart = "<td>";
  var cellEnd = "</td>";
   for (i in array){ 
     array[i] = rowStart + cellStart + array[i] + cellEnd + rowEnd;
     }  
  array  = array.join('');
  array = tableStart + array + tableEnd;
 
  return array;
}
//Code.gs STOP