/**
 * Paste this into: Extensions → Apps Script (from your orders spreadsheet)
 * Deploy → New deployment → Web app
 * Execute as: Me
 * Who has access: Anyone
 * Copy the Web app URL into the website checkout "Connect Google Sheet" box.
 */
var SHEET_ID = '1OIcfFUSImMuMg-iUxe_sXR9xygVeRkF7RKrSfJntCKA';
var SHEET_URL = 'https://docs.google.com/spreadsheets/d/1OIcfFUSImMuMg-iUxe_sXR9xygVeRkF7RKrSfJntCKA/edit';
var NOTIFY_EMAIL = 'browniesbrookiess@gmail.com';

function doGet(e) { return handle_(e); }
function doPost(e) { return handle_(e); }

function handle_(e) {
  try {
    var d = (e && e.parameter) ? e.parameter : {};
    if (e && e.postData && e.postData.contents) {
      try {
        var j = JSON.parse(e.postData.contents);
        for (var k in j) if (!d[k]) d[k] = j[k];
      } catch (ignore) {}
    }
    if (!d.name || !d.phone) {
      return out_({ ok: false, error: 'Missing name or phone' });
    }
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp','Name','Village / City','Phone','Email','Items','Total','Status']);
      sheet.getRange(1,1,1,8).setFontWeight('bold');
    }
    var when = new Date();
    sheet.appendRow([when, d.name, d.city || '', d.phone, d.email || '', d.items || '', d.total || '', 'New']);
    try {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        replyTo: d.email || NOTIFY_EMAIL,
        subject: 'New order — ' + d.name + ' — ' + (d.total || ''),
        body:
          'New order from the website\n\n' +
          'Name: ' + d.name + '\n' +
          'Village / city: ' + (d.city || '') + '\n' +
          'Phone: ' + d.phone + '\n' +
          'Email: ' + (d.email || '') + '\n' +
          'Items: ' + (d.items || '') + '\n' +
          'Total: ' + (d.total || '') + '\n' +
          'Time: ' + when + '\n\n' +
          'Spreadsheet:\n' + SHEET_URL
      });
    } catch (mailErr) {}
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
