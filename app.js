//!!IMPORTANT!! Ensure SHEETID is correct 
//!!IMPORTANT!! Copy the web app url after deploy
const SHEETID = 'YOUR_SHEET_ID';
const urlLink = "YOUR_WEB_URL";

// Get data from sheet
const sheet = SpreadsheetApp.openById(SHEETID).getSheetByName('data');
const data = sheet.getDataRange().getValues();
numOfApprovers = 0;

const logoId = data[1][4]; // Image Logo Id 
const emailAddressColumn = "Email Address"; // DO NOT CHANGE - make sure email collection is enabled in Google Form
const validDomain = 'YOUR_DOMAIN_NAME'; // Enter valid domain 

// Get form title
const form = FormApp.getActiveForm();
const formTitle = form.getTitle();

const headers = data[0];

// Columns for email, name, and title
const emailColumn = headers.indexOf('Email');
const nameColumn = headers.indexOf('Name');
const titleColumn = headers.indexOf('Title');

// Extract approval flows from the "data" sheet
function extractFlows(data) {
  const flows = [];
  var numOfApproverRequired = data[1][5]

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

// GAS Retry function
function call(func, optLoggerFunction) {
  for (var n = 0; n < 6; n++) {
    try {
      return func();
    } catch (e) {
      if (optLoggerFunction) {
        optLoggerFunction("GASRetry " + n + ": " + e);
      }
      if (n == 5) {
        throw e;
      }
      Utilities.sleep((Math.pow(2, n) * 1000) + (Math.round(Math.random() * 1000)));
    }
  }
}

const FLOWS = {
  defaultFlow: extractFlows(data),
};

function App() {
  //If unable to get form detail and formUrl, retry 
  try {
    this.form = call(FormApp.getActiveForm);
    console.log(this.form);

    this.formUrl = call(this.form.getPublishedUrl);
    console.log(this.formUrl);
  } catch (e) {
    Logger.log("Error: " + e.toString());
  }

  this.url = urlLink;
  this.title = this.form.getTitle();
  this.sheetname = "Form Responses 1"; // DO NOT change - the default google form responses sheet name
  this.uidHeader = "UID";
  this.uidPrefix = "UID-";
  this.uidLength = 5;
  this.statusHeader = "Status";
  this.responseIdHeader = "_response_id";
  this.documentLink = "Document Link";
  this.emailHeader = emailAddressColumn;

  this.pending = "Pending";
  this.approved = "Approved";
  this.rejected = "Rejected";
  this.waiting = "Waiting";

  console.log(FLOWS);

  // Get the details of the spreadsheet
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

  // Converts the data from spreadsheet to JSON
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

  // Get task from spreadsheet based on ID
  this.getTaskById = (id) => {

    const values = this.parsedValues();
    const record = values.find((value) => value.some((cell) => cell.taskId === id));
    const row = values.findIndex((value) => value.some((cell) => cell.taskId === id)) + 1;

    const headers = values[0];
    const statusColumn = headers.indexOf(this.statusHeader) + 1;

    let task, approver, nextApprover, column, approvers, email, status, responseId;

    // If record is obtained, extract details
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

  // Get response from spreadsheet based on ID
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

  // Create a unique ID for each form
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

  // Creates and sends the approval email to approver
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

    MailApp.sendEmail(approver.email, subject, "", options);
  };

  // Sends a notification 
  this.sendNotification = (taskId) => {

    const { email, responseId, status, task, approvers } = this.getTaskById(taskId);

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

    MailApp.sendEmail(email, subject, "", options);
  };

  //!!IMPORTANT PROCESS!!
  // Handles the form submission and approval workflow
  this.onFormSubmit = () => {

    const values = this.parsedValues();
    const headers = values[0];
    let lastRow = values.length;
    let startColumn = headers.indexOf(this.uidHeader) + 1;
    if (startColumn === 0) startColumn = headers.length + 1;

    const responses = this.form.getResponses();
    const lastResponse = responses[responses.length - 1];
    const responseId = lastResponse.getId();

    const newHeaders = [this.uidHeader, this.statusHeader, this.responseIdHeader, this.documentLink];
    const newValues = [this.createUid(), this.pending, responseId, ""];

    const flow = FLOWS.defaultFlow;

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

  // Approver approves the form
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

  // Approver rejects the form
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


// Handles the display of the forms
function doGet(_event) {
  const { taskId, responseId } = _event.parameter;
  const app = new App();
  let template;

  const userEmail = Session.getActiveUser().getEmail();

  const isValid = validate(userEmail, validDomain);

  if (!isValid) {
    const html = HtmlService.createTemplateFromFile("non_user.html")
      .evaluate()
      .setTitle("Unauthorized Access")
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    return html;
  }

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
  template.waiting = app.waiting

  // Configuration for HTML output
  const htmlOutput = template.evaluate();
  htmlOutput
    .setTitle(app.title)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return htmlOutput;
}

function validate(email, domain) {
  const userEmailDomain = email.split('@')[1];
  return userEmailDomain === domain;
}

// Create a trigger for the form submission
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

// Replace the placeholders in the template with data from spreadsheet
function replacePlaceholdersInDocument(body, placeholderMap) {
  for (const placeholder in placeholderMap) {
    if (placeholderMap.hasOwnProperty(placeholder)) {
      body.replaceText(placeholder, placeholderMap[placeholder]);
    }
  }
}

// Replace the approver placeholders in the template with data from spreadsheet
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

// Check if folder exist in parent folder
function folderExists(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext();
}

// Creates the Google Docs using data from spreadsheet
function createNewGoogleDocs() {

  var googleDocFolder = "Document";
  var parentFolderId = data[1][3];

  // Get the parent folder
  var parentFolder;
  if (parentFolderId) {
    try {
      parentFolder = DriveApp.getFolderById(parentFolderId);
    } catch (error) {
      console.log("Folder '" + parentFolder + "' does not exist.");
      return;
    }
  }

  // Check and create Google Doc folder within the parent folder
  if (!folderExists(parentFolder, googleDocFolder)) {
    parentFolder.createFolder(googleDocFolder);
  }

  var docFolderIterator = parentFolder.getFoldersByName(googleDocFolder);
  var docFolder;
  if (docFolderIterator.hasNext()) {
    docFolder = docFolderIterator.next();
  } else {
    console.log("Folder '" + googleDocFolder + "' does not exist.");
    return;
  }

  const gDocFolderId = docFolder.getId();
  var googleDocFolderId = DriveApp.getFolderById(gDocFolderId);

  // Store the sheet as variable
  const sheet = SpreadsheetApp.openById(SHEETID).getSheetByName('Form Responses 1')

  // Get all of the values as a 2D array
  const rows = sheet.getDataRange().getValues();

  const headers = rows[0];

  const headerMap = {};

  headers.forEach((header, columnIndex) => {
    headerMap[header] = columnIndex;
  });

  // Processing each spreadsheet row
  rows.forEach(function (row, index) {

    // Get headers from spreadsheet
    const placeholderMap = {};
    headers.forEach((header, index) => {
      // Check if headers contains any special characters
      const containsSpecialChars = /[.*+?^${}()|[\]\\]/.test(header);

      // Escape special characters only if necessary
      const escapedHeader = containsSpecialChars ? header.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') : header;

      // Construct the placeholder
      const placeholder = `{${escapedHeader}}`;

      // Assign the value to the placeholder map
      placeholderMap[placeholder] = row[index];
      console.log(placeholderMap[placeholder]);
    });

    const documentLinkColumnIndex = headerMap['Document Link'];
    const documentLinkValue = row[documentLinkColumnIndex];


    // Check if this row is the headers. If so, skip it
    if (index === 0) return;

    // Check if document link has been generated. If so, skip it
    if (documentLinkValue) return;

    // Check if 'Status' is 'Pending'. If so, skip it
    if (placeholderMap['{Status}'] == 'Pending') return;

    // Title of the document in our destinationFolder
    var emailAddress = placeholderMap[`{${emailAddressColumn}}`]
    var username = emailAddress.substring(0, emailAddress.indexOf('@'))

    // Create Doc Template
    var doc = DocumentApp.create(`${username}, ${new Date(placeholderMap['{Timestamp}']).toLocaleString()} ${formTitle}`);
    var file = DriveApp.getFileById(doc.getId()); // Get the file corresponding to the newly created document
    file.moveTo(googleDocFolderId); // Move the file to the target folder

    const body = doc.getBody();

    var img = DriveApp.getFileById(logoId).getBlob();

    // Insert the image into the document
    var image = body.appendImage(img);
    var imageStyles = {};
    imageStyles[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
    imageStyles[image.setWidth(180).setHeight(80)];
    image.getParent().setAttributes(imageStyles);

    var titleParagraph = body.appendParagraph(formTitle);
    titleParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    titleParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    titleParagraph.setFontFamily('Arial');

    var dataTable = body.appendTable();
    var headersForDoc = headers.slice(0, headerMap['Document Link'] - 1);

    var placeholderForDoc = headersForDoc.map(function (headerForDoc) {
      return '{' + headerForDoc + '}';
    });

    var numofRowsData = headersForDoc.length;
    var numofColumnsData = 2;

    for (var row = 0; row < numofRowsData; row++) {
      var tableRow = dataTable.appendTableRow();
      for (var col = 0; col < numofColumnsData; col++) {
        if (col == 0) {
          var headerCell = tableRow.appendTableCell();
          headerCell.setText(headersForDoc[row]);
        }
        else {
          var headerCell = tableRow.appendTableCell();
          headerCell.setText(placeholderForDoc[row]);
        }
      }
    }
    var titleParagraph2 = body.appendParagraph("Approver Details");
    titleParagraph2.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    titleParagraph2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    titleParagraph2.setFontFamily('Arial');

    var approverTable = body.appendTable();
    var headers2 = ['Approver', 'Title', 'Status', 'Comment', 'Timestamp'];
    var placeHolders2 = ['{_approver_1.name}', '{_approver_1.title}', '{_approver_1.status}', '{_approver_1.comments}', '{_approver_1.timestamp}'];

    // Add header row
    var headerRow = approverTable.appendTableRow();
    headers2.forEach(function (header) {
      headerRow.appendTableCell().setText(header);
    });

    // Add rows for approver details
    for (var i = 1; i <= numOfApprovers; i++) {
      var approverRow = approverTable.appendTableRow();
      placeHolders2.forEach(function (placeholder) {
        var approverDetails = placeholder.replace('_approver_1', `_approver_${i}`);
        approverRow.appendTableCell().setText(approverDetails);
      });
    }

    replacePlaceholdersInDocument(body, placeholderMap);

    // Get details of the approvers
    for (let i = 1; i <= numOfApprovers; i++) {
      const approverHeader = `_approver_${i}`;
      const approverJSONString = placeholderMap[`{${approverHeader}}`];

      // Check if the JSON string is not empty
      if (approverJSONString && approverJSONString.trim() !== '') {
        try {
          const _approver_ = JSON.parse(approverJSONString);
          replaceApproverPlaceholders(body, _approver_, approverHeader);
        } catch (error) {
          console.error(`Error parsing JSON for ${approverHeader}: ${error}`);
          // Handle the error as needed (e.g., log, ignore, or take corrective action)
        }
      } else {
        // Handle the case where the JSON string is empty
        console.warn(`JSON string for ${approverHeader} is empty.`);
      }
    }

    doc.saveAndClose();
    const url = doc.getUrl();

    sheet.getRange(index + 1, headerMap['Document Link'] + 1).setValue(url)
  })
}

// Create PDF from Google Docs
function convertGoogleDocsToPDFs() {
  var pdfFolderName = "Pdf"; // Names of the child folders
  var googleDocFolderName = "Document";
  var parentFolderId = data[1][3];

  // Get the parent folder
  if (parentFolderId) {
    try {
      parentFolder = DriveApp.getFolderById(parentFolderId);
    } catch (error) {
      console.log("Folder with ID '" + parentFolderId + "' could not be retrieved.");
      return;
    }
  }

  // Check and create Pdf folder within the parent folder
  if (!folderExists(parentFolder, pdfFolderName)) {
    parentFolder.createFolder(pdfFolderName);
  }

  var pdfFolderIterator = parentFolder.getFoldersByName(pdfFolderName);
  var pdfFolder;
  if (pdfFolderIterator.hasNext()) {
    pdfFolder = pdfFolderIterator.next();
  } else {
    console.log("Folder '" + pdfFolderName + "' does not exist.");
    return;
  }

  const pdfFolderId = pdfFolder.getId();

  // Check if the data folder exist
  if (!folderExists(parentFolder, googleDocFolderName)) {
    console.log("Folder '" + googleDocFolderName + "' does not exist.");
    return;
  } else {
    var googleDocFolderIterator = parentFolder.getFoldersByName(googleDocFolderName);
    var googleDocFolder;
    if (googleDocFolderIterator.hasNext()) {
      googleDocFolder = googleDocFolderIterator.next();
    }
  }

  const gDocFolderId = googleDocFolder.getId();

  const googleDoc_Folder = DriveApp.getFolderById(gDocFolderId);
  const pdf_folder = DriveApp.getFolderById(pdfFolderId);


  const invoices = googleDoc_Folder.getFiles();

  const invoicesPDF = pdf_folder.getFiles();

  var pdfNameArray = [];


  while (invoicesPDF.hasNext()) {
    var file = invoicesPDF.next();
    var title = file.getName().replace('.pdf', '');
    pdfNameArray.push(title);

  }

  while (invoices.hasNext()) {
    var invoice = invoices.next();

    var fileName = invoice.getName().replace('.pdf', '');
    console.log(fileName);

    if (pdfNameArray.includes(fileName)) {
    }
    else {
      console.log("The ID doesnt exist");
      var id = invoice.getId();
      var file = DriveApp.getFileById(id);

      var PDFblob = file.getAs(MimeType.PDF);
      pdfFolder.createFile(PDFblob);
    }
  }
}