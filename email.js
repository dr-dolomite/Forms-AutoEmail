/**
  Replace the "<DOCID>" with your document ID, or the entire URL per say. Should be something like:
  var EMAIL_TEMPLATE_DOC_URL = 'https://docs.google.com/document/d/asdasdakvJZasdasd3nR8kmbiphqlykM-zxcrasdasdad/edit?usp=sharing';
*/

var EMAIL_TEMPLATE_DOC_URL = 'https://docs.google.com/document/d/1Iheq5wLb4PYYQ26K7StByCCChkkjgOsLa7PrE37JGNs/edit?usp=sharing';
var EMAIL_SUBJECT = 'Webinar Invitation';

/**
 * Sends a customized email for every response on a form.
 *
 * @param {Object} e - Form submit event
 */
function onFormSubmit(e) {
  var responses = e.namedValues;

  // If the question title is a label, it can be accessed as an object field.
  // If it has spaces or other characters, it can be accessed as a dictionary.
  
  /** 
    NOTE: One common issue people are facing is an error that says 'TypeError: Cannont read properties of undefined'
    That usually means that your heading cell in the Google Sheet is something else than exactly 'Email address'.
    The code is Case-Sesnsitive so this HAS TO BE exactly the same on line 25 and your Google Sheet.
  */
  var email = responses['Email Address'][0].trim();
  var fullName = responses['Fullname  (First Name, MI, Last Name)'][0].trim();
   var headerImageUrl = 'https://fastnetphbyrus.com/zoombackground/eheader.png'; // Replace with the URL of your header image

  Logger.log('; responses=' + JSON.stringify(responses));

  MailApp.sendEmail({
    to: email,
    subject: EMAIL_SUBJECT,
    htmlBody: createEmailBody(fullName, headerImageUrl),
  });
  Logger.log('email sent to: ' + email);

  // Append the status on the spreadsheet to the responses' row.
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  var column = e.values.length + 1;
  sheet.getRange(row, column).setValue('Email Invitation Sent.');
}

/**
 * Creates email body and includes the links based on topic.
 *
 * @param {string} name - The recipient's name.
 * @param {string} headerImageUrl - The URL of the header image.
 * @return {string} - The email body as an HTML string.
 */
function createEmailBody(fullName, headerImageUrl) {
  // Make sure to update the emailTemplateDocId at the top.
  var docId = DocumentApp.openByUrl(EMAIL_TEMPLATE_DOC_URL).getId();
  var emailBody = docToHtml(docId, fullName, headerImageUrl);
  return emailBody;
}

/**
 * Downloads a Google Doc as an HTML string.
 *
 * @param {string} docId - The ID of a Google Doc to fetch content from.
 * @param {string} fullName - The Full Name to insert into the document.
 * @param {string} headerImageUrl - The URL of the header image.
 * @return {string} The Google Doc rendered as an HTML string.
 */
function docToHtml(docId, fullName, headerImageUrl) {
  var url = 'https://docs.google.com/feeds/download/documents/export/Export?id=' +
            docId + '&exportFormat=html';
  var param = {
    method: 'get',
    headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true,
  };
  var content = UrlFetchApp.fetch(url, param).getContentText();

  // Replace a placeholder (e.g., "{{FullName}}") in your Google Doc content with the actual Full Name
  content = content.replace('{{FullName}}', fullName);

  // Replace a placeholder (e.g., "{{HeaderImage}}") with the HTML code for the header image
  content = content.replace('{{HeaderImage}}', '<img src="' + headerImageUrl + '">');

  return content;
}

