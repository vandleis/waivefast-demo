// Apps Script: Receives POSTs from the static site and runs an email sequence.
// Deploy as Web App -> New deployment -> Web app.
// Authorize Gmail send when prompted.

const SHEET_NAME = 'Leads';
const FOLLOWUP_DAYS = [0,2,4,7]; // Day offsets for sends

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['timestamp','name','phone','email','state','invoice_number','total','job_address','raw_text','stripe_link','status','last_sent','next_send_index']);
    }
    const row = [
      payload.timestamp || new Date().toISOString(),
      payload.submitter.name || '',
      payload.submitter.phone || '',
      payload.submitter.email || '',
      payload.state || '',
      payload.invoiceData.invoice_number || '',
      payload.invoiceData.total || '',
      payload.invoiceData.job_address || '',
      payload.invoiceData.raw || '',
      payload.stripe_link || '',
      'new',
      '',
      0
    ];
    sheet.appendRow(row);
    // send immediate email (Day 0)
    sendEmailForRow(sheet, sheet.getLastRow());
    return ContentService.createTextOutput(JSON.stringify({ok:true, row: sheet.getLastRow()})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function sendEmailForRow(sheet, rowNumber) {
  const row = sheet.getRange(rowNumber,1,1,sheet.getLastColumn()).getValues()[0];
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const obj = {};
  headers.forEach((h,i)=> obj[h] = row[i]);
  const to = obj['email'];
  const firstName = (obj['name'] || '').split(' ')[0] || '';
  const stripeLink = obj['stripe_link'] || '';
  const personalizedLine = `Hi ${firstName}, I can turn an invoice photo into a state-compliant lien waiver and a payment link so you get paid faster. First 20 waivers free.`;
  const subject = 'Get paid faster — first 20 waivers free';
  const body = `${personalizedLine}\n\nI generated a waiver for your invoice. Download the PDF you received and use this link for faster payments: ${stripeLink}\n\nReply to this email or call/text ${obj['phone']} to start the pilot.\n\n— WaiveFast`;
  if (to && to.includes('@')) {
    GmailApp.sendEmail(to, subject, body, {name: "WaiveFast"});
    sheet.getRange(rowNumber, headers.indexOf('status')+1).setValue('emailed');
    sheet.getRange(rowNumber, headers.indexOf('last_sent')+1).setValue(new Date().toISOString());
    sheet.getRange(rowNumber, headers.indexOf('next_send_index')+1).setValue(1);
  }
}

function processFollowups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  rows.forEach((row, idx) => {
    const rnum = idx+2;
    const obj = {};
    headers.forEach((h,i)=> obj[h] = row[i]);
    let nextIndex = Number(obj['next_send_index'] || 0);
    if (nextIndex >= FOLLOWUP_DAYS.length) return;
    const created = new Date(obj['timestamp']);
    const sendDate = new Date(created);
    sendDate.setDate(sendDate.getDate() + FOLLOWUP_DAYS[nextIndex]);
    if (new Date() >= sendDate) {
      const to = obj['email'];
      if (to && to.includes('@')) {
        const firstName = (obj['name'] || '').split(' ')[0] || '';
        const subject = nextIndex === 0 ? 'Get paid faster — first 20 waivers free' : `Quick follow-up — ${firstName}, can we run your first waiver?`;
        const body = `Hi ${firstName || ''},\n\nJust following up on our offer to process your first 20 waivers free and show how much faster you can get paid. Use this link to pay for fast processing (optional): ${obj['stripe_link'] || ''}\n\nReply to this message or call/text ${obj['phone']}.\n\n— WaiveFast`;
        GmailApp.sendEmail(to, subject, body, {name: "WaiveFast"});
        sheet.getRange(rnum, headers.indexOf('last_sent')+1).setValue(new Date().toISOString());
        sheet.getRange(rnum, headers.indexOf('next_send_index')+1).setValue(nextIndex+1);
      }
    }
  });
}