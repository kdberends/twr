function SendEmailToUser(user, templatefile, attachmentID){
  Logger.log('sending email to ' + user.email)
    var attachment = DriveApp.getFileById(attachmentID);
    
    var template = HtmlService.createTemplateFromFile(templatefile)
    template.name = user.name;
    htmlBody = template.evaluate().getContent();
    MailApp.sendEmail({
      to: user.email,
      subject: "Jouw voorlopige Triviumweekrooster",
      htmlBody: htmlBody,
      attachments:  [attachment.getAs(MimeType.PDF)]
      })
}