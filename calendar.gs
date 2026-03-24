function getCollaborators() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEMPLATE_SHEET);
  const data = sheet.getDataRange().getValues();

  const collaborators = [];
  for (let i = 1; i < data.length; i++) {
    const isPermanent = data[i][0] === "P";
    const name = data[i][1];
    const email = data[i][2];
    if (isPermanent && name) {
      collaborators.push({ name, email });
    }
  }
  return collaborators;
}

function updateCalendarEmail(name, newEmailRaw) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEMPLATE_SHEET);
  const data = sheet.getDataRange().getValues();
  const emails = JSON.parse(newEmailRaw || '[]');

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === name && data[i][0] === "P") {
      if (emails.length > 0) {
        sheet.getRange(i + 1, 3).setValue(JSON.stringify(emails));
      } else {
        sheet.getRange(i + 1, 3).setValue('');
      }
      return true;
    }
  }
  return false;
}