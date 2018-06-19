/* Sends email to user. 
 * NOTE: if function is called on a callback (like onFormSubmit)
 * will send email from every account that has given this app
 * permission. Solution is to run this app on a single account
 * only (no shared Spreadsheets!)
 */
function SendEmailToUser(user, templatefile, attachmentID, subject){
  Logger.log('sending email to ' + user.email)
  var attachment = DriveApp.getFileById(attachmentID);
  var template = HtmlService.createTemplateFromFile(templatefile)
  template.name = user.name;
  htmlBody = template.evaluate().getContent();
  MailApp.sendEmail({
    to: user.email,
    subject: subject,
    htmlBody: htmlBody,
    attachments:  [attachment.getAs(MimeType.PDF)]
      })
}