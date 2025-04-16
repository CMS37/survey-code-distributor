function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("InputSidebar")
      .setTitle("섭외 메일 입력");
  SpreadsheetApp.getUi().showSidebar(html);
}

function appendEmailsToSheet(rawText) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("섭외메일시트");
  if (!sheet) throw new Error("섭외메일시트 시트를 찾을 수 없습니다.");

  const emailList = rawText
    .split(/\r?\n/)
    .map(email => email.trim())
    .filter(email => email.length > 0);

  if (emailList.length === 0) throw new Error("붙여넣은 이메일 목록이 비어 있습니다.");

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, emailList.length, 1).setValues(emailList.map(email => [email]));
}
