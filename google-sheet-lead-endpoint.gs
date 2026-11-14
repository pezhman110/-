/**
 * دریافت لید صفحهٔ سرمایه‌گذاری در Google Sheet
 * هر پرونده (leadId) یک ردیف است و با هر قدم به‌روز می‌شود؛
 * پس کسی که بعد از قدم ۲ صفحه را می‌بندد هم با نام و شماره در فهرست می‌ماند.
 *
 * راه‌اندازی (یک بار، حدود ۵ دقیقه):
 * 1. یک Google Sheet تازه بسازید.
 * 2. Extensions > Apps Script را باز کنید، همهٔ کد را با این فایل جایگزین و ذخیره کنید.
 * 3. Deploy > New deployment > نوع Web app
 *    Execute as: Me   ·   Who has access: Anyone
 * 4. آدرس Web app را کپی کنید و در صفحه جای FORM_ENDPOINT = '' بگذارید.
 */
const HEADERS = ['leadId', 'وضعیت', 'قدم', 'نام', 'موبایل', 'ایمیل', 'ساعت تماس',
  'طرح‌های پیشنهادی', 'جواب‌ها', 'utm_source', 'utm_medium', 'utm_campaign', 'utm_content',
  'ارجاع', 'صفحه', 'اولین ثبت', 'آخرین به‌روزرسانی'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
    const status = d.stage === 'completed' || d.stage === 'call-time' ? 'تکمیل' :
                   d.stage === 'skipped' ? 'تکمیل (ادامهٔ فرم را نخواست؛ تماس بگیرید)' :
                   d.stage === 'left-page' ? 'نیمه‌کاره (صفحه بسته شد)' : 'در حال تکمیل';
    const answers = (d.answers || []).map(x => x.q + ' ← ' + x.a).join('\n');
    const now = new Date();
    const row = [d.leadId, status, (d.step || '') + ' از ' + (d.of || ''), d.name, "'" + d.phone, d.email,
      d.callTime, (d.plans || []).join('، '), answers, d.utm_source || '', d.utm_medium || '',
      d.utm_campaign || '', d.utm_content || '', d.referrer || '', d.page || '', now, now];

    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat() : [];
    const i = ids.indexOf(d.leadId);
    if (i < 0) {
      sh.appendRow(row);
    } else {
      const r = i + 2;
      const first = sh.getRange(r, 16).getValue();
      const prevStatus = sh.getRange(r, 2).getValue();
      row[15] = first || now;
      if (String(prevStatus).startsWith('تکمیل') && !status.startsWith('تکمیل')) row[1] = prevStatus;  // وضعیت تکمیل برنمی‌گردد
      sh.getRange(r, 1, 1, row.length).setValues([row]);
    }
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
