function doPost(e) {
  try {
    const raw = e.postData.contents;
    Logger.log('Raw received: ' + raw);

    const data = JSON.parse(raw);
    Logger.log('Parsed data: ' + JSON.stringify(data));
    Logger.log('formType: ' + data.formType);
    Logger.log('city: ' + data.city);

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.formType === 'GE') {
      const sheet = ss.getSheetByName('GE_Log');
      sheet.appendRow([
        data.city,
        data.dateEntered,
        data.trial,
        data.levelCompleted,
        data.dateCompleted,
        data.contributedPoints || '',
        data.guildRank || '',
        data.encCompleted || '',
        data.encAvailable || '',
        data.notes
      ]);
    } else if (data.formType === 'GBG') {
      const sheet = ss.getSheetByName('GBG_Log');
      sheet.appendRow([
        data.city,
        data.date,
        data.trialFought,
        data.battlesFought,
        data.peakAttrition,
        data.notes
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('ERROR: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
