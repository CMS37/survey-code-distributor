const updateCodeUsageStatus = () => {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const emailSheet = ss.getSheetByName("섭외메일시트");
	const responseSheet = ss.getSheetByName("설문지 응답 시트 1");
  
	if (!emailSheet || !responseSheet) throw new Error("시트를 찾을 수 없습니다.");
  
	const responseCodes = responseSheet
	  .getRange(2, 2, responseSheet.getLastRow() - 1) // B열
	  .getValues()
	  .flat()
	  .filter(Boolean);
  
	const responseMap = new Map();
	responseCodes.forEach(code => {
	  responseMap.set(code, (responseMap.get(code) || 0) + 1);
	});
  
	const codeRange = emailSheet.getRange("B2:B" + emailSheet.getLastRow());
	const statusRange = emailSheet.getRange("D2:D" + emailSheet.getLastRow());
	const codes = codeRange.getValues().flat();
  
	const statuses = codes.map(code => {
	  const count = responseMap.get(code);
	  if (!count) return "미사용";
	  return count === 1 ? "사용" : "중복";
	});
  
	statusRange.setValues(statuses.map(s => [s]));
  };
  