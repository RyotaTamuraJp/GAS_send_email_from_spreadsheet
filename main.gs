const subject = "Subject"
const options = {
  name: "SenderName",
  from: "hoge@gmail.cam"
}
const emailBodyTemplate = `
If you don't see the HTML, you'll see this text.
`

function sendEmail() {
  var mySheet = SpreadsheetApp.getActiveSheet();
  var rows = mySheet.getDataRange().getLastRow();

  for (var i = 2; i <= rows; i++) {
    const company = mySheet.getRange(i,1).getValue();
    const name = mySheet.getRange(i,2).getValue();
    const to = mySheet.getRange(i,3).getValue();
    const emailTextBody = `${company}\n${name}\n${emailBodyTemplate}`;

    var templateEmail = HtmlService.createTemplateFromFile('contents');
    templateEmail.params = [company, name];
    html = templateEmail.evaluate().getContent();
    options['htmlBody'] = html

    GmailApp.sendEmail(to, subject, emailTextBody, options);

    const remaining = MailApp.getRemainingDailyQuota();
    console.log(`i=${i}: Sent to "${toEmail}". The remaining ${remaining} can be sent today.`)
  }
}
