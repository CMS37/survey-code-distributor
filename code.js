const showSidebar = () => {
  const html = HtmlService.createHtmlOutputFromFile("InputSidebar")
    .setTitle("섭외 메일 입력");
  SpreadsheetApp.getUi().showSidebar(html);
};

const appendEmailsToSheet = (rawText) => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailSheet = ss.getSheetByName("섭외메일시트");
  const viewSheet = ss.getSheetByName("필터링");

  if (!emailSheet) throw new Error("섭외메일시트 시트를 찾을 수 없습니다.");
  if (!viewSheet) throw new Error("필터링 시트를 찾을 수 없습니다.");

  const emailList = rawText
    .split(/\r?\n/)
    .map(email => email.trim())
    .filter(email => email.length > 0);

  if (emailList.length === 0) throw new Error("붙여넣은 이메일 목록이 비어 있습니다.");

  const existingCodes = new Set(emailSheet.getRange("B2:B" + emailSheet.getLastRow()).getValues().flat().filter(Boolean));
  const now = new Date();

  const newRows = [];
  const newCodes = [];

  for (const email of emailList) {
    let newCode;
    do {
      newCode = makeRandomCode();
    } while (existingCodes.has(newCode));
    existingCodes.add(newCode);

    newRows.push([email, newCode, now]);
    newCodes.push(newCode);
  }

  if (newRows.length > 0) {
    const startRow = emailSheet.getLastRow() + 1;
    emailSheet.getRange(startRow, 1, newRows.length, 3).setValues(newRows);
  }

  // 정규표현식 생성 및 G1에 삽입
  const allCodes = emailSheet.getRange("B2:B" + emailSheet.getLastRow()).getValues().flat().filter(Boolean);
  const regex = `^(${allCodes.join("|")})$`;
  const regexCell = viewSheet.getRange("G2");

  regexCell.setValue(regex);
  regexCell.setWrap(false);
};

const makeRandomCode = (length = 6) => {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  return Array.from({ length }, () => chars[Math.floor(Math.random() * chars.length)]).join("");
};
