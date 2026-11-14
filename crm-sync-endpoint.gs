/**
 * ثبت همهٔ کلیک‌های پنل فروش (sales-crm-panel.html) در یک Google Sheet مشترک تیم
 * هر کلیک = یک ردیف در برگهٔ «رویدادها»؛ برگهٔ «مشتریان» آخرین وضعیت هر مشتری را نگه می‌دارد.
 *
 * راه‌اندازی: Google Sheet تازه › Extensions › Apps Script › این کد را بگذارید ›
 * Deploy › New deployment › Web app (Execute as: Me · Who has access: Anyone) ›
 * آدرس را در پنل جای CRM_SYNC_ENDPOINT = '' بگذارید.
 */
/** دکمهٔ «دریافت پرونده‌های تیم» در پنل: خلاصهٔ همهٔ مشتریان را به‌صورت JSON برمی‌گرداند. */
function doGet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('مشتریان');
  const rows = sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues() : [];
  const out = rows.map(r => ({id: r[0], name: r[1], phone: String(r[2]), agent: r[3], stage: r[4], type: r[5], last: r[6],
    next: r[7] instanceof Date ? r[7].toISOString() : '', updated: r[8] instanceof Date ? r[8].toISOString() : ''}));
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents), ss = SpreadsheetApp.getActiveSpreadsheet();
    const ev = ss.getSheetByName('رویدادها') || ss.insertSheet('رویدادها');
    if (ev.getLastRow() === 0) ev.appendRow(['زمان', 'کارشناس', 'شناسه', 'نام', 'موبایل', 'مرحله', 'تیپ', 'رویداد', 'توضیح', 'پیگیری بعدی']);
    ev.appendRow([new Date(d.at), d.agent, d.leadId, d.name, "'" + d.phone, d.stage, d.type, d.kind, d.text, d.next ? new Date(d.next) : '']);
    const cu = ss.getSheetByName('مشتریان') || ss.insertSheet('مشتریان');
    if (cu.getLastRow() === 0) cu.appendRow(['شناسه', 'نام', 'موبایل', 'کارشناس', 'مرحله', 'تیپ', 'آخرین رویداد', 'پیگیری بعدی', 'به‌روزرسانی']);
    const row = [d.leadId, d.name, "'" + d.phone, d.agent, d.stage, d.type, d.kind + ' · ' + d.text, d.next ? new Date(d.next) : '', new Date()];
    const ids = cu.getLastRow() > 1 ? cu.getRange(2, 1, cu.getLastRow() - 1, 1).getValues().flat() : [];
    const i = ids.indexOf(d.leadId);
    if (i < 0) cu.appendRow(row); else cu.getRange(i + 2, 1, 1, row.length).setValues([row]);
    return ContentService.createTextOutput('ok');
  } finally { lock.releaseLock(); }
}
