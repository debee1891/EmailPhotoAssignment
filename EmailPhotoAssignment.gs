function EmailPhotoAssignment(e) {
  //update the value below if the sheet is changed. enter the sheet ID. the sheet id is found between d/ and /edit within the sheet URL
  var ss = SpreadsheetApp.openById("1zwNTXJKmxXsxGAKD1BilLdnmDuTqz7R8fRAaDRgfvoY"); 
  //ensure that the sheet/tab name is correct
  var sheet = ss.getSheetByName("Form responses 1");  
  var lock = LockService.getPublicLock();
  var range = sheet.getRange("A3:11");
  
  
   // sort the sheet to from old to newest updates/entries based on registered timestamp
 range.sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);
  lock.waitLock(30000);  // 30 secs
    var data = sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn()).getValues(); 
    lock.releaseLock();
  //change the recipient email below. To check if the email is generating or the script is running without submitting the form, click Run on the menu. 
  //to add a recipient put comma (,) and the email address)
    var emailAddress = "servicedesk@aap.com.au";  
  //update this portion if sheet is changed. 
  //  var blurb ="https://docs.google.com/spreadsheets/d/1qrWVIkC909Tw7rW-N5PZy3eKpDI5cx51nheBzMFjbvQ/edit?resourcekey#gid=1400925612"
   // answers are stored in an array. at the moment arrays are from 0-12 (for rows) go to sheet2 of the file for a quick  visual reference of each array. This is where you edit the fields/data in the email notification:
  var body = "<HTML><BODY>"
    + "<P>" + "<b>AAP Photo Assignments Form</b>"
      +  "<table border='1' CELLSPACING='1' style='border-color:lightblue;border-collapse:separate;'>"
      +    "<tr>" 
      +      "<td><b> Who </b></td>" 
      +      "<td>"+ data[0][2] +"</td>"
      +     "</tr>"
      +     "<tr>"
      +      "<td><b>What</b></td>" 
      +      "<td>"+ data[0][3] +"</td>"
      +     "</tr>"
      +     "<tr>"
      +      "<td><b>Where</b></td>" 
      +      "<td>"+ data[0][4] +"</td>"
      +     "</tr>"
      +     "<tr>"
      +     "<td><b>When</b></td>" 
      +      "<td>"+ data[0][5] +"<br></td>"
      +     "</tr>"
      +     "<tr>"
      +     "<td><b>Assignment Contact Details</b></td>" 
      +      "<td>"+ data[0][6] +"<br></td>"
      +     "</tr>"
      +     "<tr>"
      +      "<td><b>Assignment Description </b></td>" 
      +      "<td>"+ data[0][7] +"</td>"
      +     "</tr>"
      +     "<tr>"
      +      "<td><b>Photo Description</b></td>" 
      +      "<td>"+ data[0][8] +"</td>"
      +     "</tr>"
 

      +   "</table>"
  //    +  "<P>" + '<a href=\"'+blurb+ '">Check the entries on Google Sheet</a>'
      + "</HTML></BODY>";
  var message = "Body Message";    
  //the subject line can be customised here:
  var subject = "AAP Photo Assignments from " + data[0][01];
    MailApp.sendEmail(emailAddress, subject, message, {replyTo: data[0][01], htmlBody:body}); 
    sheet.getRange(sheet.getLastRow(),10).setValue('yes'); 
       
 } 
