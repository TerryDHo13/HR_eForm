//Custom menu of Spreadsheet
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('AutoFill Docs');
    menu.addItem('Create New Docs', 'createNewGoogleDocs')
    menu.addToUi();
  }
  
  //Creates Document links to Google Docs for each row of data
  function createNewGoogleDocs() {
  
    //Get template and folder ID
    const SHEETID = '1UPi073wOUf4uzO0VEDHY4zU1DT5CrQpbAnwIUvu4Xo0';
    const data_sheet = SpreadsheetApp.openById(SHEETID).getSheetByName('data');
    const data = data_sheet.getDataRange().getValues();
  
    //!!IMPORTANT!! Ensure that the template and folder ID are entered in the row and column of the 'data' sheet
    const googleDocTemplateID = data[0][4];
    const folderID = data[1][4];
  
    const googleDocTemplate = DriveApp.getFileById(googleDocTemplateID);
    const destinationFolder = DriveApp.getFolderById(folderID)
  
    //Store the sheet as variable
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Form Responses 1')
  
    //Get all of the values as a 2D array
    const rows = sheet.getDataRange().getValues();
  
    //Processing each spreadsheet row
    rows.forEach(function (row, index) {
  
      //Check if this row is the headers. If so, skip it
      if (index === 0) return;
  
      //Check if document link has been generated. If so, skip it
      if (row[22]) return;
  
      //Title of the document in our destinationFolder
      const copy = googleDocTemplate.makeCopy(`${row[1]}, ${row[2]} MANPOWER REQUISITION FORM`, destinationFolder)
  
      //Once we have the copy, we then open it using the DocumentApp
      const doc = DocumentApp.openById(copy.getId())
      //All of the content lives in the body, so we get that for editing
      const body = doc.getBody();
  
      //In these lines, replace our replacement tokens with values from our spreadsheet row
      body.replaceText('{Timestamp}', row[0]);
      body.replaceText('{Email Address}', row[16]);
      body.replaceText('{No. of staff requested:}', row[3]);
      body.replaceText('{Required Date:}', row[4]);
      body.replaceText('{Position:}', row[5]);
      body.replaceText('{Grade:}', row[6]);
      body.replaceText('{Department:}', row[7]);
      body.replaceText('{Division:}', row[8]);
      body.replaceText('{Report to:}', row[9]);
      body.replaceText('{Is this position:}', row[10]);
      body.replaceText('{Is this position:}', row[11]);
      body.replaceText('{Salary Range:}', row[12]);
      body.replaceText('{Qualification:}', row[13]);
      body.replaceText('{Working experience:}', row[14]);
      body.replaceText('{Practical skill:}', row[15]);
      body.replaceText('{Name:}', row[1]);
      body.replaceText('{Date:}', row[2]);
      body.replaceText('{_uid}', row[17]);
      body.replaceText('{_status}', row[18]);
  
      //Approver details for approver 1
      var _approver_1 = JSON.parse(row[20])
      body.replaceText('{_approver_1."name"}', _approver_1.name);
      body.replaceText('{_approver_1."title"}', _approver_1.title);
      body.replaceText('{_approver_1."status"}', _approver_1.status);
  
      if (_approver_1.comments === null) {
        body.replaceText('{_approver_1."comments"}', "");
      } else {
        body.replaceText('{_approver_1."comments"}', _approver_1.comments);
      }
  
      const formatted_approver_1_timestamp = new Date(_approver_1.timestamp).toLocaleDateString();
      body.replaceText('{_approver_1."timestamp"}', formatted_approver_1_timestamp);
  
      //Approver details for approver 2
      var _approver_2 = JSON.parse(row[21])
      body.replaceText('{_approver_2."name"}', _approver_2.name);
      body.replaceText('{_approver_2."title"}', _approver_2.title);
      body.replaceText('{_approver_2."status"}', _approver_2.status);
  
      if (_approver_2.comments === null) {
        body.replaceText('{_approver_2."comments"}', "");
      } else {
        body.replaceText('{_approver_2."comments"}', _approver_2.comments);
      }
  
      const formatted_approver_2_timestamp = new Date(_approver_2.timestamp).toLocaleDateString();
      body.replaceText('{_approver_2."timestamp"}', formatted_approver_2_timestamp);
  
      //We make our changes permanent by saving and closing the document
      doc.saveAndClose();
      //Store the url of our new document in a variable
      const url = doc.getUrl();
      //Write that value back to the 'Document Link' column in the spreadsheet. 
      sheet.getRange(index + 1, 23).setValue(url)
    })
  }
  