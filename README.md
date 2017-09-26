# googleScripts

## Collection of scripts from Google spreadsheets and docs

```
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Send Forms to Candidates')
      .addItem('Email Forms Now', 'emailForms')
      .addToUi();
}

function emailForms() {
  SpreadsheetApp.getUi()
  getRowAsArray();
}

// Google Doc id from the contract template
var sourceId = "1u33IUn0Yz-pH_KaXMiK_yWqODJHKGjhE5Y6QEBT4paQ";

// In which Google Drive we save the new contracts
var contractsFolder = "0B4wtdGtu5r3HamRkUmxCNmRXQ3M";

// create timestamp to mark when communication was sent
var timestamp = new Date(); 

// Chatbot spreadsheet and sheet where data is coming from
var ss = SpreadsheetApp.openById("1AF07mCT2_g4GP6L3ATEbEJfPSoX-0AHaCjAbRUudXzM");
var sheet = ss.getSheetByName("responses");
var lastRow = sheet.getLastRow();
var headings = ss.getRange('A1:BO1').getValues();
var range = sheet.getRange(2,1,lastRow-1,73).getValues();
Logger.log(range.length); 

//Get row data of new candidate
function getRowAsArray() {
  // loop over range and send communication if "Yes" option chosen
  for (var i = 0; i < range.length; i++) {
    if (range[i][0] === "Yes") {
      var firstName = range[i][9];
      var lastName = range[i][10];
      var fullName = firstName + " " + lastName;
      createNewContract(fullName);
      //createGoogleDoc(headings, range[i]); 
     }
  };
}

/* Duplicates a Google Apps doc and return a new document with a given name from the orignal */
function createNewContract (fullName) {
  //Get the template from the drive
  var source = DriveApp.getFileById(sourceId);
  //Make a copy using the candidates full name as the document title
  var newFile = source.makeCopy(fullName);
  //Save it in the specified folder
  var targetFolder = DriveApp.getFolderById(contractsFolder);
  //Add the file to the contracts folder
  targetFolder.addFile(newFile);
  replaceWords(newFile);
};
  

//Search a key word in the document and replaces it with the text from the spreadsheet
function replaceWords(newFile) {
  var doc = DocumentApp.openById(newFile.getId());
  var body = doc.getBody();
  var ss = SpreadsheetApp.openById("1AF07mCT2_g4GP6L3ATEbEJfPSoX-0AHaCjAbRUudXzM");
  var sheet = ss.getSheetByName("responses");
  var lastRow = sheet.getLastRow();
  var headings = ss.getRange('A1:BO1').getValues();
  var range = sheet.getRange(2,1,lastRow-1,73).getValues();
  for (var i = 0; i < range.length; i++) {
    if (range[i][0] === "Yes") {
      var position = range[i][3];
      var payrate = range[i][4];
      var startdate = range[i][5];
      var jobdescription = range[i][6];
      var supervisor = range[i][7];
      var responsibilities = range[i][8];
      var firstName = range[i][9];
      var lastName = range[i][10];
      var fullName = firstName + " " + lastName;
      var email = range[i][12];
      var mobile = range[i][13];
      var homeph = range[i][14];
      var address = range[i][15];
      var preferStart = range[i][25];
      var residency = range[i][27];
      body.replaceText("FULLNAME", fullName);
      body.replaceText("DATE", timestamp);
      body.replaceText("POSITION", position);
      body.replaceText("PAYRATE", payrate);
      body.replaceText("START", startdate);
      body.replaceText("JOBDESCRIPTION", jobdescription);
      body.replaceText("SUPERVISOR", supervisor);
      body.replaceText("RESPONSIBILITIES", responsibilities);
      doc.saveAndClose();
    }
          // reset first column to 'No' so no repeated emails
      sheet.getRange(i+1, 1, 1, 1).setValue('No');
      // add timestamp to Date Registered column to show when communication was sent
      sheet.getRange(i+1, 1, 2, 1).setValue(timestamp);
    
  };
   var irdFileId = '0B25-feUG_tBZQ085c3dFMVdLYlU';
   var accFileId = '0B4wtdGtu5r3HNXlqallhSFJFcEE';
   var incidentFileId = '0B25-feUG_tBZeWdPSGNzdmFSMkE';
   var contractFileId = (doc.getId());
   var irdFile = DriveApp.getFileById(irdFileId);
   var accFile = DriveApp.getFileById(accFileId);
   var incidentFile = DriveApp.getFileById(incidentFileId);
   var contractFile = DriveApp.getFileById(contractFileId);
   var apexEmail = "janie@manu.net.nz";

    MailApp.sendEmail({
      to: email,
      cc: apexEmail,
      subject: "Apex - application for work",
      attachments:
        [irdFile.getAs(MimeType.PDF), accFile.getAs(MimeType.PDF), incidentFile.getAs(MimeType.PDF), contractFile.getAs(MimeType.PDF) ],
      htmlBody: 
      "Hi " + fullName +",<br><br>" +
      "Thank you for registering for work with Apex Recruitment." +  
      "<br><br>Please check the following information you provided is correct." +
      "<br>Full Name: " + fullName +
      "<br>Email: " + email +
      "<br>Mobile phone: " + mobile +
      "<br>Home phone: " + homeph +
      "<br>Address: " + address +
      "<br>Preferred start date: " + preferStart +
      "<br>Residency status: " + residency +
      "<br><br>I have attached an Employment Contract - please read carefully and sign.  Also complete the IR330 form with your tax details." +
      "<br>If you are a NZ resident and have claimed for ACC in the past, please complete the ACC Claims history form.  " +
      "<br>The other two attachments are for your use while working with Apex: an incident report form, and a timesheet form.  " +
      "<br>You will need to provide all the required information and signed documents, before we can offer you employment." +
      "<br><br>If there is anything you want to change let me know by replying to this email. Any questions - give me a ring on 021558508 or (09)3903947. I look forward to hearing from you." +
      "<br><br>Kind regards, Georgia Butt"
    });

    
}
    

//Extra variables in case we make a new PDF for the summary instead and have a password to open it
      //var license = range[i][16];
      //var ownTransport = range[i][19];
      //var boots = range[i][21];
      //var hat = range[i][22];
      //var vest = range[i][23];
      //var startNow = range[i][24];
      //var preferStart = range[i][25];
      //var notAvail = range[i][26];
      //var residency = range[i][27];
      //var emergency = range[i][30];
      //var emerPh = range[i][31];
      //var ird = range[i][32];
      //var bankName = range[i][33];
      //var bankNumber = range[i][34];      
            



/*
function createGoogleDoc(headings, data){
  var firstname = data[3];
  var lastname = data[4];
  var timestamp = new Date();
  var doc = DocumentApp.create(firstname+' '+lastname+ '\'s Application');
  var H1 = doc.appendParagraph('Application of '+ firstname+' '+lastname );
  H1.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  //Personnel Information
  var H2_1 = doc.appendParagraph('Personnel Information');
  H2_1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  for(i = 2; i < 10; i++){
    doc.getBody().appendParagraph(headings[0][i]+': '+data[i]);
  }  
  //Drivers License Info
  var H2_2 = doc.appendParagraph('Drivers License Information');
  H2_2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  for(i = 10; i < 15; i++){
    doc.getBody().appendParagraph(headings[0][i]+': '+data[i]);
  }  
  doc.saveAndClose()
   //create pdf
   var newFile = DriveApp.createFile(doc.getAs('application/pdf'))
  //find the document id and move it to the specified folder
  //var docFile = DriveApp.getFileById( doc.getId() );
  //DriveApp.getFolderById('0B25-feUG_tBZVGFJTjY5c3BLc2M').addFile( docFile );  
}
*/

```
