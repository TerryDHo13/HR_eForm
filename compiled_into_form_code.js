//!!IMPORTANT!! Ensure SHEETID is correct 
//!!IMPORTANT!! Ensure that 'data' sheet contains the approver details, template ID and Folder ID
const SHEETID = '1OXSB4nOj8Oh5jvjOxzONw50GNAc8X0gU4nY6QyhSyyc';
const sheet = SpreadsheetApp.openById(SHEETID).getSheetByName('data');
const data = sheet.getDataRange().getValues();
numOfApprovers = 0;

const headers = data[0];

//Columns for email, name, and title
const emailColumn = headers.indexOf('Email');
const nameColumn = headers.indexOf('Name');
const titleColumn = headers.indexOf('Title');

//Extract approval flows from the "data" sheet
function extractFlows(data) {
  const flows = [];
  var numOfApproverRequired = data[1][6]

  for (let i = 1; i <= numOfApproverRequired; i++) {
    const flow = {
      email: data[i][emailColumn],
      name: data[i][nameColumn],
      title: data[i][titleColumn],
    };
    flows.push(flow);
    numOfApprovers += 1;
  }
  return flows;
}

const FLOWS = {
  defaultFlow: extractFlows(data),
};

function App() {
  this.form = FormApp.getActiveForm();
  this.formUrl = this.form.getPublishedUrl();
  this.url =
    "https://script.google.com/macros/s/AKfycbwY3IDPzi1DvqRfxpGaf5-9bLwgY6gBQvkVOANMVcSGYXnfzCqxWUnL5xBCpS-ayr0d/exec";
  //!!IMPORTANT!! - copy the web app url after deploy
  this.title = this.form.getTitle();
  this.sheetname = "Form Responses 1"; // DO NOT change - the default google form responses sheet name
  this.flowHeader = "Department"; //!!IMPORTANT!! - key field for your flows
  this.uidHeader = "UID";
  this.uidPrefix = "UID-";
  this.uidLength = 5;
  this.statusHeader = "Status";
  this.responseIdHeader = "_response_id";
  this.documentLink = "Document Link";
  this.emailHeader = "Email Address"; // DO NOT CHANGE - make sure email collection is enabled in Google Form

  this.pending = "Pending";
  this.approved = "Approved";
  this.rejected = "Rejected";
  this.waiting = "Waiting";

  console.log(FLOWS);

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

    //If record is obtained, extract details
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

    //If record is obtained, extract details
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

  this.resetUid = () => {
    const props = PropertiesService.getDocumentProperties();
    props.deleteProperty(this.uidHeader);
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

  //!!IMPORTANT PROCESS!!
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

    //Add "Document Link" header at the end
    newHeaders.push(this.documentLink);
    newValues.push("");

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

function resetUid() {
  const app = new App();
  app.resetUid();
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

//Custom menu on Google Form
function onOpen(e) {
  FormApp.getUi()
    .createMenu('Other features')
    .addItem('Create New Docs', 'createNewGoogleDocs')
    .addItem('Reset UID', 'resetUid')
    .addItem('Convert Google Docs to PDFs', 'convertGoogleDocsToPDFs')
    .addItem('Sync with Google Sheet', 'syncFormWithSheet')
    .addToUi();
}

//Replace the placeholders in the template with data from spreadsheet
function replacePlaceholdersInDocument(body, placeholderMap) {
  for (const placeholder in placeholderMap) {
    if (placeholderMap.hasOwnProperty(placeholder)) {
      body.replaceText(placeholder, placeholderMap[placeholder]);
    }
  }
}

//Replace the approver placeholders in the template with data from spreadsheet
function replaceApproverPlaceholders(body, approverData, placeholderPrefix) {
  const replacements = {
    name: approverData.name,
    title: approverData.title,
    status: approverData.status,
    comments: approverData.comments === null ? "" : approverData.comments,
    timestamp: new Date(approverData.timestamp).toLocaleDateString(),
  };

  for (const [placeholder, value] of Object.entries(replacements)) {
    const fullPlaceholder = `{${placeholderPrefix}.${placeholder}}`;
    body.replaceText(fullPlaceholder, value);
  }
}

//Creates the Google Docs using data from spreadsheet
function createNewGoogleDocs() {
  //!!IMPORTANT!! Ensure that the template and folder ID are entered in the row and column of the 'data' sheet
  const googleDocTemplateID = data[1][3];
  const folderID = data[1][4];

  const googleDocTemplate = DriveApp.getFileById(googleDocTemplateID);
  const destinationFolder = DriveApp.getFolderById(folderID)

  //Store the sheet as variable
  const sheet = SpreadsheetApp.openById(SHEETID).getSheetByName('Form Responses 1')

  //Get all of the values as a 2D array
  const rows = sheet.getDataRange().getValues();

  const headers = rows[0];

  //Processing each spreadsheet row
  rows.forEach(function (row, index) {

    //Extract headers from spreadsheet
    const placeholderMap = {};
    headers.forEach((header, index) => {
      const placeholder = `{${header}}`;
      placeholderMap[placeholder] = row[index];
    });
    const numberOfPlaceholders = Object.keys(placeholderMap).length;

    //Check if this row is the headers. If so, skip it
    if (index === 0) return;

    //Check if document link has been generated. If so, skip it
    if (row[numberOfPlaceholders - 1]) return;

    //Title of the document in our destinationFolder
    const copy = googleDocTemplate.makeCopy(`${placeholderMap['{Name:}']}, ${new Date(placeholderMap['{Timestamp}']).toLocaleString()} MANPOWER REQUISITION FORM`, destinationFolder)

    const doc = DocumentApp.openById(copy.getId())
    const body = doc.getBody();
    replacePlaceholdersInDocument(body, placeholderMap);

    //Get details of the approvers
    for (let i = 1; i <= numOfApprovers; i++) {
      const approverHeader = `_approver_${i}`;
      const _approver_ = JSON.parse(placeholderMap[`{${approverHeader}}`]);
      replaceApproverPlaceholders(body, _approver_, approverHeader);
    }

    doc.saveAndClose();
    const url = doc.getUrl();

    sheet.getRange(index + 1, numberOfPlaceholders).setValue(url)
  })
}

function convertGoogleDocsToPDFs() {
  //Initialize Data folder ID and PDF folder ID
  const datafolderID = data[1][4];
  const pdfFolderID = data[1][5];

  const dataFolder = DriveApp.getFolderById(datafolderID)
  const pdfFolder = DriveApp.getFolderById(pdfFolderID);

  console.log('PDF Folder ID: ' + pdfFolder);
  console.log('Data Folder ID: ' + dataFolder);

  const invoices = dataFolder.getFiles();
  console.log('Files: ' + invoices);

  while (invoices.hasNext()) {
    var invoice = invoices.next();
    var id = invoice.getId();
    var file = DriveApp.getFileById(id);
    var PDFblob = file.getAs(MimeType.PDF);
    var PDF = pdfFolder.createFile(PDFblob);
  }
}

function syncFormWithSheet() {
  //When pressed, sync the question from Google Form with Spreadsheet header
}
