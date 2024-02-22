var eventName = "Machine Learning Crash Course";
//Change only one time
var slideTemplateId = "19_fW9N3Yr9dQbFxqpEh9HaT-jOr6NqeFa9bxbkvs_GI"; // Example: https://docs.google.com/presentation/d/1u_7EtmmmZ_cdGHno2PrAme4AjWa40Cm8rcHgDo1PIpo
var sheetId = "1rErYVz3l7CydNe7Zwur6sAg7su0CuXnGLdZCFUOwTog"; // Example: https://docs.google.com/spreadsheets/d/1VY3_SsdomBnLhfQ2NR-aiVRbyi41yuF64l7UQZaSgZo
var tempFolderId = "1F0Xe3JCd67odDWylCsQrkMXDbzsy_uRv"; // Example: https://drive.google.com/drive/folders/12cRJ-Jf2KFjaAkAmrPKNv1_XfcM594ei
var leadName = "John Aziz";
var title = "Microsoft Learn Student Ambassador";
var teamName = "Microsoft Learn Student Ambassadors";

// Create Slides with the data form the spread sheet and update the status once created
function createCertificates() {
  var template = DriveApp.getFileById(slideTemplateId);
  
  var sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var nameIndex = headers.indexOf("Name");
  var emailIndex = headers.indexOf("Email");
  var dateIndex = headers.indexOf("Date");
  var descriptionIndex = headers.indexOf("Description");
  var slideIndex = headers.indexOf("Slide ID");
  var statusIndex = headers.indexOf("Status");
  
  for (var i = 1; i < values.length; i++) {
    var rowData = values[i];
    var name = rowData[nameIndex];
    var date = rowData[dateIndex];
    var description = rowData[descriptionIndex];
        
    var tempFolder = DriveApp.getFolderById(tempFolderId);
    var slideId = template.makeCopy(tempFolder).setName(name).getId();        
    var slide = SlidesApp.openById(slideId).getSlides()[0];
   
    slide.replaceAllText("Receiver Name", name);
    slide.replaceAllText("Description", description);
    slide.replaceAllText("Date Issued", "Date Issued: " + Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM dd, yyyy"));
    slide.replaceAllText("Your Name", leadName);
    slide.replaceAllText("Title", title);
    slide.replaceAllText("Team Name", teamName);
    
    sheet.getRange(i + 1, slideIndex + 1).setValue(slideId);
    sheet.getRange(i + 1, statusIndex + 1).setValue("CREATED");
    SpreadsheetApp.flush();
  }
}

// Send Email with the pdf version of the slide attached and a message down below then updates status in the sheet
function sendCertificates() {
  var sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var nameIndex = headers.indexOf("Name");
  var emailIndex = headers.indexOf("Email");
  var dateIndex = headers.indexOf("Date");
  var slideIndex = headers.indexOf("Slide ID");
  var statusIndex = headers.indexOf("Status");
  
  for (var i = 1; i < values.length; i++) {
    var rowData = values[i];
    var name = rowData[nameIndex];
    var email = rowData[emailIndex];
    var date = rowData[dateIndex];
    var slideId = rowData[slideIndex];
    
    var attachment = DriveApp.getFileById(slideId);
    var senderName = teamName;
    var subject = name + ", You're awesome!";
    var body = "On behalf of Microsoft Learn Student Ambassador, "+
               "we would like to thank you for participating with us in the "+eventName+".\n\n"+
               "This certificate is for people who attended the session.\n\n"+
               "Thank you again for being part of this journey, and we encourage you to stay updated with our events.\n\n"+
               "We wish you the best of luck in your future endeavors.\n\n"+
               "Kindly find your certificate attached.\n\n"+
               "Sincerely,\n" + teamName + " team";
    GmailApp.sendEmail(email, subject, body, {
      attachments: [attachment.getAs(MimeType.PDF)],
      name: senderName
    });
    sheet.getRange(i + 1, statusIndex + 1).setValue("SENT");
    SpreadsheetApp.flush();
  }
}
