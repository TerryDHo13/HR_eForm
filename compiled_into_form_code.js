const SHEETID = '1UPi073wOUf4uzO0VEDHY4zU1DT5CrQpbAnwIUvu4Xo0';
const sheet = SpreadsheetApp.openById(SHEETID).getSheetByName('data');
const data = sheet.getDataRange().getValues();

const testing_0 = data[0][0];
const testing_1 = data[0][1];
const testing_2 = data[0][2];
const testing_3 = data[1][0];
const testing_4 = data[1][1];
const testing_5 = data[1][2];

// Define the approval flows in this object
const FLOWS = {
  defaultFlow: [
    {
      email: testing_0,
      name: testing_1,
      title: testing_2,
    },
    {
      email: testing_3,
      name: testing_4,
      title: testing_5,
    },
  ],
};

function App() {
  this.form = FormApp.getActiveForm();
  this.formUrl = this.form.getPublishedUrl();
  this.url =
    "https://script.google.com/macros/s/AKfycbxMGDGGVOkE1Icvd7aYWhudBSjHIZi45auB-Z_1sjhZ0Nb3yac9SDh-Q3DmyOcF6H7F/exec"; // IMPORTANT - copy the web app url after deploy
  this.title = this.form.getTitle();
  this.sheetname = "Form Responses 1"; // DO NOT change - the default google form responses sheet name
  this.flowHeader = "Department"; // IMPORTANT - key field for your flows
  this.uidHeader = "UID";
  this.uidPrefix = "UID-";
  this.uidLength = 5;
  this.statusHeader = "Status";
  this.responseIdHeader = "_response_id";
  this.emailHeader = "Email Address"; // DO NOT CHANGE - make sure email collection is enabled in Google Form

  this.pending = "Pending";
  this.approved = "Approved";
  this.rejected = "Rejected";
  this.waiting = "Waiting";


  //Get the details of the spreadsheet
  this.sheet = (() => {

    let sheet;
    try {
      const id = this.form.getDestinationId();
      sheet = SpreadsheetApp.openById(id);

      //If error occurs, create new spreadsheet
    } catch (e) {
      const id = this.form.getId();
      const parentFolder = DriveApp.getFileById(id).getParents().next();

      const spreadsheet = SpreadsheetApp.create(this.title + " (Responses)");
      const spreadsheet_id = spreadsheet.getId();

      this.form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);

      DriveApp.getFileById(spreadsheet_id).moveTo(parentFolder);

      sheet = spreadsheet;
    }
    return sheet.getSheetByName(this.sheetname);
  })();

  //Converts the data from spreadsheet to JSON
  this.parsedValues = () => {

    const parsedValues = this.sheet.getDataRange().getDisplayValues().map((value) => {
      return value.map((cell) => {
        try {
          return JSON.parse(cell);

          //If error occurs, returns the original cell content
        } catch (e) {
          return cell;
        }
      });
    });
    return parsedValues;
  };

  //Get task from spreadsheet based on ID
  this.getTaskById = (id) => {

    const values = this.parsedValues();
    const record = values.find((value) => value.some((cell) => cell.taskId === id));
    const row = values.findIndex((value) => value.some((cell) => cell.taskId === id)) + 1;

    const headers = values[0];
    const statusColumn = headers.indexOf(this.statusHeader) + 1;

    let task, approver, nextApprover, column, approvers, email, status, responseId;

    if (record) {
      task = record.slice(0, headers.indexOf(this.statusHeader) + 1).map((item, i) => {
        return {
          label: headers[i],
          value: item
        };
      });

      email = record[headers.indexOf(this.emailHeader)];
      status = record[headers.indexOf(this.statusHeader)];
      responseId = record[headers.indexOf(this.responseIdHeader)];
      approver = record.find((item) => item.taskId === id);
      column = record.findIndex((item) => item.taskId === id) + 1;
      nextApprover = record[record.findIndex((item) => item.taskId === id) + 1];
      approvers = record.filter((item) => item.taskId);
    }
    return { email, status, responseId, task, approver, nextApprover, approvers, row, column, statusColumn };
  };

  //Get response from spreadsheet based on ID
  this.getResponseById = (id) => {
    
    const values = this.parsedValues();
    const record = values.find((value) => value.some((cell) => cell === id));

    const headers = values[0];

    let task, approvers, status;

    if (record) {
      task = record.slice(0, headers.indexOf(this.statusHeader) + 1).map((item, i) => {
        return {
          label: headers[i],
          value: item,
        };
      });

      status = record[headers.indexOf(this.statusHeader)];
      approvers = record.filter((item) => item.taskId);
    }
    return { task, approvers, status };
  };

  //Create a unique ID for each form
  this.createUid = () => {

    const properties = PropertiesService.getDocumentProperties();
    let uid = Number(properties.getProperty(this.uidHeader));
    if (!uid) uid = 1;

    properties.setProperty(this.uidHeader, uid + 1);

    return (
      this.uidPrefix +
      (uid + 10 ** this.uidLength).toString().slice(-this.uidLength)
    );
  };

  //Creates and sends the approval email to approver
  this.sendApproval = ({ task, approver, approvers }) => {

    const template_approval = HtmlService.createTemplateFromFile("approval_email.html");

    template_approval.title = this.title;
    template_approval.task = task;
    template_approval.approver = approver;
    template_approval.approvers = approvers;
    template_approval.actionUrl = `${this.url}?taskId=${approver.taskId}`;
    template_approval.formUrl = this.formUrl;

    template_approval.approved = this.approved;
    template_approval.rejected = this.rejected;
    template_approval.pending = this.pending;
    template_approval.waiting = this.waiting;

    const subject = "Approval Required - " + this.title;

    const options = { htmlBody: template_approval.evaluate().getContent() };

    GmailApp.sendEmail(approver.email, subject, "", options);
  };

  //Sends a notification 
  this.sendNotification = (taskId) => {
 
    const { email, responseId, status, task, approvers } = this.getTaskById(taskId);
    console.log({ email, status, task, approvers });

    const template_notification = HtmlService.createTemplateFromFile(
      "notification_email.html"
    );

    template_notification.title = this.title;
    template_notification.task = task;
    template_notification.status = status;
    template_notification.approvers = approvers;
    template_notification.formUrl = this.formUrl;
    template_notification.approvalProgressUrl = `${this.url}?responseId=${responseId}`;

    template_notification.approved = this.approved;
    template_notification.rejected = this.rejected;
    template_notification.pending = this.pending;
    template_notification.waiting = this.waiting;

    const subject = `Approval ${status} - ${this.title}`;

    const options = { htmlBody: template_notification.evaluate().getContent() };

    GmailApp.sendEmail(email, subject, "", options);
  };

  //IMPORTANT PROCESS
  // add addtional data to form response when update
  //Handles the form submission and approval workflow
  this.onFormSubmit = () => {

    const values = this.parsedValues();
    const headers = values[0];
    let lastRow = values.length;
    let startColumn = headers.indexOf(this.uidHeader) + 1;
    if (startColumn === 0) startColumn = headers.length + 1;

    const responses = this.form.getResponses();
    const lastResponse = responses[responses.length - 1];
    const responseId = lastResponse.getId();
    const newHeaders = [this.uidHeader, this.statusHeader, this.responseIdHeader];
    const newValues = [this.createUid(), this.pending, responseId];

    const flowKey = values[lastRow - 1][headers.indexOf(this.flowHeader)];
    const flow = FLOWS[flowKey] || FLOWS.defaultFlow;
    let taskId;
    flow.forEach((item, i) => {
      newHeaders.push("_approver_" + (i + 1));

      item.comments = null;
      item.taskId = Utilities.base64EncodeWebSafe(Utilities.getUuid());
      item.timestamp = new Date();
      if (i === 0) {
        item.status = this.pending;
        taskId = item.taskId;
      } else {
        item.status = this.waiting;
      }
      if (i !== flow.length - 1) {
        item.hasNext = true;
      } else {
        item.hasNext = false;
      }
      newValues.push(JSON.stringify(item));
    });

    this.sheet
      .getRange(1, startColumn, 1, newHeaders.length)
      .setValues([newHeaders])
      .setBackgroundColor("#34A853")
      .setFontColor("#FFFFFF")

    this.sheet
      .getRange(lastRow, startColumn, 1, newValues.length)
      .setValues([newValues]);

    this.sendNotification(taskId);
    const { task, approver, approvers } = this.getTaskById(taskId);
    this.sendApproval({ task, approver, approvers });
  };

  //Approver approves the form
  this.approve = ({ taskId, comments }) => {
    const { task, approver, approvers, nextApprover, row, column, statusColumn } = this.getTaskById(taskId);

    if (!approver) return;
    approver.comments = comments;
    approver.status = this.approved;
    approver.timestamp = new Date();
    this.sheet.getRange(row, column).setValue(JSON.stringify(approver));

    if (approver.hasNext) {
      nextApprover.status = this.pending;
      nextApprover.timestamp = new Date();
      this.sheet
        .getRange(row, column + 1)
        .setValue(JSON.stringify(nextApprover));
      this.sendApproval({ task, approver: nextApprover, approvers });
    } else {
      this.sheet.getRange(row, statusColumn).setValue(this.approved);
      this.sendNotification(taskId);
    }
  };

  //Approver rejects the form
  this.reject = ({ taskId, comments }) => {
    const { approver, row, column, statusColumn } = this.getTaskById(taskId);

    if (!approver) return;
    approver.comments = comments;
    approver.status = this.rejected;
    approver.timestamp = new Date();
    this.sheet.getRange(row, column).setValue(JSON.stringify(approver));
    this.sheet.getRange(row, statusColumn).setValue(this.rejected);
    this.sendNotification(taskId);
  };
}


function _onFormSubmit() {
  const app = new App();
  app.onFormSubmit();
}


function approve({ taskId, comments }) {
  const app = new App();
  app.approve({ taskId, comments });
}


function reject({ taskId, comments }) {
  const app = new App();
  app.reject({ taskId, comments });
}


function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}


//Handles the display of the forms
function doGet(_event) {
  const { taskId, responseId } = _event.parameter;
  const app = new App();
  let template;

  if (taskId) {
    template = HtmlService.createTemplateFromFile("index");
    const { task, approver, approvers, status } = app.getTaskById(taskId);
    template.task = task;
    template.status = status;
    template.approver = approver;
    template.approvers = approvers;
    template.url = `${app.url}?taskId=${taskId}`;

  } else if (responseId) {
    template = HtmlService.createTemplateFromFile("approval_progress");
    const { task, approvers, status } = app.getResponseById(responseId);
    template.task = task;
    template.status = status;
    template.approvers = approvers;

  }
  else {
    template = HtmlService.createTemplateFromFile("404.html");
  }

  template.title = app.title;
  template.pending = app.pending;
  template.approved = app.approved;
  template.rejected = app.rejected;
  template.waiting = app.waiting;

  //Configuration for HTML output
  const htmlOutput = template.evaluate();
  htmlOutput
    .setTitle(app.title)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return htmlOutput;
}

//Create a trigger for the form submission
function createTrigger() {
  const functionName = "_onFormSubmit";
  const triggers = ScriptApp.getProjectTriggers();

  const triggerExist = triggers.some(
    (trigger) => trigger.getHandlerFunction() === functionName
  );

  if (triggerExist) return;

  return ScriptApp.newTrigger(functionName)
    .forForm(FormApp.getActiveForm())
    .onFormSubmit()
    .create();
}

function onOpen(){
  const ui = FormApp.getUi();
  const menu = ui.createMenu('AutoFill Docs');
  menu.addItem('Create New Docs', 'createNewGoogleDocs')
  menu.addToUi();
}


function createNewGoogleDocs(){
  //!!IMPORTANT!! Ensure that the template and folder ID are entered in the row and column of the 'data' sheet
  const googleDocTemplateID = data[0][4];
  const folderID = data[1][4];

  const googleDocTemplate = DriveApp.getFileById(googleDocTemplateID);
  const destinationFolder = DriveApp.getFolderById(folderID)

  //Store the sheet as variable
  const sheet = SpreadsheetApp.openById(SHEETID).getSheetByName('Form Responses 1')

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