// EVOLUTION TIMER — Google Apps Script
// Pegá este código en Extensions > Apps Script de tu Google Sheets
// Publicá como Web App: Execute as Me, Access Anyone

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet(data.gym || 'Evolution');
    
    const sessions = data.sessions || [];
    
    sessions.forEach(session => {
      const row = [
        session.date ? new Date(session.date).toLocaleString('es-AR') : new Date().toLocaleString('es-AR'),
        session.athlete || '',
        (session.mode || '').toUpperCase(),
        session.result || '',
        session.detail || '',
        session.mode === 'laps' && session.laps ? session.laps.length : '',
        session.mode === 'laps' && session.laps ? 
          session.laps.map((l, i) => `V${i+1}:${formatTime(l.split)}`).join(' | ') : ''
      ];
      sheet.appendRow(row);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, added: sessions.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, status: 'Evolution Timer API activa' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // Headers
    const headers = ['FECHA / HORA', 'ATLETA', 'MODO', 'RESULTADO', 'DETALLE', 'VUELTAS', 'DETALLE VUELTAS'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Style headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#0a0a0a');
    headerRange.setFontColor('#e8001d');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(10);
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Column widths
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 90);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(6, 80);
    sheet.setColumnWidth(7, 300);
  }
  
  return sheet;
}

function formatTime(seconds) {
  const m = Math.floor(seconds / 60);
  const s = seconds % 60;
  return `${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`;
}
