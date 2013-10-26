// What is it? ... A macro for Gmail written in Google Apps Script
// What is it for? ... Those who, for what ever reason, prefer to save their messages in Gdrive instead of the Gmail folders
// How does it work? ... Every minute the code searches your Gmail threads for messages that you want save to Gdrive. 
//                       It then converts your threads into the PDF or GDOC format and saves them to your specified folder of your Gdrive.
// How do I use it? ... 
//      1) create the following Gmail labels: "RootDoc", "RootPDF", "ReferenceDOC", "ReferencePDF"
//      2) open Google Apps Script (https://script.google.com/macros/) and create a "blank project" 
//      3) copy and paste the following code, then hit save (nb: give it whatever name you want)
//      4) run "createTriggers" 
//      5) it will ask you to authorize, after you authorize run "createTriggers" again just to be sure that it worked


function RootDoc() {
  var gLabel = "RootDoc";
  
  // get messages
  var thread = GmailApp.search("label:" + gLabel);
  for (var x = 0; x < thread.length; x++) {
    var messages = thread[x].getMessages();
    for (var y = 0; y < messages.length; y++) {
      
      var messageId = messages[y].getId();
      var messageDate = Utilities.formatDate(messages[y].getDate(), Session.getTimeZone(), "dd/MM/yyyy | HH:mm:ss");
      var messageFrom = messages[y].getFrom();
      var messageSubject = messages[y].getSubject();
      var messageBody = messages[y].getBody();
      var messageAttachments = messages[y].getAttachments();
      var newname = messageDate;
      
      // get message's text
      var html = messageBody
      html=html.replace(/<\/div>/ig, '\n');
      html=html.replace(/<\/li>/ig, '\n');
      html=html.replace(/<li>/ig, ' *');
      html=html.replace(/<\/ul>/ig, '\n');
      html=html.replace(/<\/p>/ig, '\n');
      html=html.replace(/<br\/?>/ig, '\n');
      html=html.replace(/<[^>]+>/ig, '');
      
      
      // create document
      var doc = DocumentApp.create(newname);
      var body = doc.getBody();
      var header = doc.addHeader();
      
      // transpose the text
      var ExportedHeader = header.appendParagraph('DATE: ' + messageDate + ' · FROM: ' + messageFrom + ' · SUB: ' + messageSubject);     
      ExportedHeader.setFontFamily(DocumentApp.FontFamily.ARIAL);
      ExportedHeader.setFontSize(10);
      ExportedHeader.setBold(true);
      ExportedHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      ExportedHeader.setForegroundColor('#666666');
      ExportedHeader.setSpacingBefore(0)
      
      var divider = header.appendHorizontalRule();
      
      var ExportedBody = body.appendParagraph(html);
      ExportedBody.setFontFamily(DocumentApp.FontFamily.ARIAL);
      ExportedBody.setFontSize(10);
      ExportedBody.setForegroundColor('#262626');
      ExportedBody.setSpacingBefore(10);
      ExportedBody.setBold(false);
      ExportedBody.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      
      // save any attachments 
      if(messageAttachments.length>0){
        var parent = DocsList.createFolder('DATE: ' + messageDate).getId();
      }
      
      for(var i = 0; i < messageAttachments.length; i++) {
        var attachmentName = messageAttachments[i].getName();
        var attachmentContentType = messageAttachments[i].getContentType();
        var attachmentBlob = messageAttachments[i].copyBlob();
        DocsList.getFolderById(parent).createFile(attachmentBlob);
      }
    }
    
    // archive the processed message
    GmailApp.getUserLabelByName(gLabel).removeFromThread(thread[x]);
    thread[x].moveToArchive();
  }
}


//////////////////////////////////////////////////////////////////////////////ℱℳ
function RootPDF() {
  var gLabel = "RootPDF";
  
  // get messages
  var thread = GmailApp.search("label: " + gLabel);
  for (var x = 0; x < thread.length; x++) {
    var messages = thread[x].getMessages();
    for (var y = 0; y < messages.length; y++) {
      
      var messageId = messages[y].getId();
      var messageDate = Utilities.formatDate(messages[y].getDate(), Session.getTimeZone(), "dd/MM/yyyy | HH:mm:ss"); 
      var messageFrom = messages[y].getFrom();
      var messageSubject = messages[y].getSubject();
      var messageBody = messages[y].getBody();
      var messageAttachments = messages[y].getAttachments();
      var body = 'DATE: ' + messageDate + ' · FROM: ' + messageFrom + ' · SUB: ' + messageSubject + ' ' + messageBody
      
      // convert and save file to folder
      var rootFolder = DocsList.getRootFolder();
      var newname = messageDate;
      var destination = DocsList.getRootFolder().createFolder(newname);
      var htmlBodyFile = destination.createFile('body.html', body, "text/html");
      var pdfBlob = htmlBodyFile.getAs('application/pdf');
      pdfBlob.setName(newname + ".pdf");
      destination.createFile(pdfBlob);
      htmlBodyFile.setTrashed(true);
      
      // save any attachments 
      for(var i = 0; i < messageAttachments.length; i++) {
        var attachmentName = messageAttachments[i].getName();
        var attachmentContentType = messageAttachments[i].getContentType();
        var attachmentBlob = messageAttachments[i].copyBlob();
        destination.createFile(attachmentBlob);
      }
    }
    
    // archive the processed message
    GmailApp.getUserLabelByName(gLabel).removeFromThread(thread[x]);
    thread[x].moveToArchive();
  }
}

//////////////////////////////////////////////////////////////////////////////ℱℳ
function ReferenceDOC() {
  var gLabel = "ReferenceDOC";
  // the macro finds the "Gmail References" folder based on name not location, so you can move it wherever you want as long as the name doesn't change
  // get messages
  var thread = GmailApp.search("label:" + gLabel);
  for (var x = 0; x < thread.length; x++) {
    var messages = thread[x].getMessages();
    for (var y = 0; y < messages.length; y++) {
      
      var messageId = messages[y].getId();
      var messageDate = Utilities.formatDate(messages[y].getDate(), Session.getTimeZone(), "dd/MM/yyyy | HH:mm:ss");
      var messageFrom = messages[y].getFrom();
      var messageSubject = messages[y].getSubject();
      var messageBody = messages[y].getBody();
      var messageAttachments = messages[y].getAttachments();
      var newname = messageDate;
      
      // get message's text
      var html = messageBody
      html=html.replace(/<\/div>/ig, '\n');
      html=html.replace(/<\/li>/ig, '\n');
      html=html.replace(/<li>/ig, ' *');
      html=html.replace(/<\/ul>/ig, '\n');
      html=html.replace(/<\/p>/ig, '\n');
      html=html.replace(/<br\/?>/ig, '\n');
      html=html.replace(/<[^>]+>/ig, '');
      
      // create document
      var refFolder = Utilities.formatDate(new Date(), Session.getTimeZone(), "yyyy MM");
      
      try{
        var parent = DocsList.getFolder('Gmail References');
      }catch(err){
        var parent = DocsList.createFolder('Gmail References');
      }
      try{
        var parent = DocsList.getFolder('Gmail ' + refFolder)
        }catch(err){
          var parent = parent.createFolder('Gmail ' + refFolder)
          }
      var doc = parent.create(newname)
      var body = doc.getBody();
      var header = doc.addHeader();
      
      // transpose the text
      var ExportedHeader = header.appendParagraph('DATE: ' + messageDate + ' · FROM: ' + messageFrom + ' · SUB: ' + messageSubject);     
      ExportedHeader.setFontFamily(DocumentApp.FontFamily.ARIAL);
      ExportedHeader.setFontSize(10);
      ExportedHeader.setBold(true);
      ExportedHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      ExportedHeader.setForegroundColor('#666666');
      ExportedHeader.setSpacingBefore(0)
      
      var divider = header.appendHorizontalRule();
      
      var ExportedBody = body.appendParagraph(html);
      ExportedBody.setFontFamily(DocumentApp.FontFamily.ARIAL);
      ExportedBody.setFontSize(10);
      ExportedBody.setForegroundColor('#262626');
      ExportedBody.setSpacingBefore(10);
      ExportedBody.setBold(false);
      ExportedBody.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      
      // save any attachments 
      for(var i = 0; i < messageAttachments.length; i++) {
        var attachmentName = messageAttachments[i].getName();
        var attachmentContentType = messageAttachments[i].getContentType();
        var attachmentBlob = messageAttachments[i].copyBlob();
        parent.createFile(attachmentBlob);
      }
    }
    
    // archive the processed message
    GmailApp.getUserLabelByName(gLabel).removeFromThread(thread[x]);
    thread[x].moveToArchive();
  }
}


//////////////////////////////////////////////////////////////////////////////ℱℳ
function ReferencePDF() {
  var gLabel = "ReferencePDF";
  
  // get messages
  var thread = GmailApp.search("label: " + gLabel);
  for (var x = 0; x < thread.length; x++) {
    var messages = thread[x].getMessages();
    for (var y = 0; y < messages.length; y++) {
      
      var messageId = messages[y].getId();
      var messageDate = Utilities.formatDate(messages[y].getDate(), Session.getTimeZone(), "d"); //"dd/MM/yyyy | HH:mm:ss"
      var messageFrom = messages[y].getFrom();
      var messageSubject = messages[y].getSubject();
      var messageBody = messages[y].getBody();
      var messageAttachments = messages[y].getAttachments();
      var body = 'DATE: ' + messageDate + ' · FROM: ' + messageFrom + ' · SUB: ' + messageSubject + ' ' + messageBody
      
      
      // convert and save file to folder
      var refFolder = Utilities.formatDate(new Date(), Session.getTimeZone(), "yyyy MM");
      
      try{
        var parent = DocsList.getFolder('Gmail References');
      }catch(err){
        var parent = DocsList.createFolder('Gmail References');
      }
      try{
        var parent = DocsList.getFolder('Gmail ' + refFolder)
        }catch(err){
          var parent = parent.createFolder('Gmail ' + refFolder)
          }
      var newname = messageDate + ' | ' + messageSubject;
      var destination = parent.createFolder(newname);
      var htmlBodyFile = destination.createFile('body.html', body, "text/html");
      var pdfBlob = htmlBodyFile.getAs('application/pdf');
      pdfBlob.setName(newname + ".pdf");
      destination.createFile(pdfBlob);
      htmlBodyFile.setTrashed(true);
      
      // save any attachments 
      for(var i = 0; i < messageAttachments.length; i++) {
        var attachmentName = messageAttachments[i].getName();
        var attachmentContentType = messageAttachments[i].getContentType();
        var attachmentBlob = messageAttachments[i].copyBlob();
        destination.createFile(attachmentBlob);
      }
    }
    
    // archive the processed message
    GmailApp.getUserLabelByName(gLabel).removeFromThread(thread[x]);
    thread[x].moveToArchive();
  }
}

//////////////////////////////////////////////////////////////////////////////ℱℳ
function createTriggers(){
  var triggers = ScriptApp.getProjectTriggers();
  for(var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  var RootDocTrig = ScriptApp.newTrigger("RootDoc").timeBased().everyMinutes(1)
  var RootPDFTrig = ScriptApp.newTrigger("RootPDF").timeBased().everyMinutes(1)
  var ReferenceDocTrig = ScriptApp.newTrigger("ReferenceDoc").timeBased().everyMinutes(1)
  var ReferencePDFTrig = ScriptApp.newTrigger("ReferencePDF").timeBased().everyMinutes(1)
  
  }