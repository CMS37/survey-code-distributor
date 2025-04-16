const createFormSubmitTrigger = () => {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	ScriptApp.newTrigger("updateCodeUsageStatus")
	  .forSpreadsheet(ss)
	  .onFormSubmit()
	  .create();
};
  