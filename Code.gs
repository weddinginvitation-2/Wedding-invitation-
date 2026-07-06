/*******************************************************
 * نظام دعوة زفاف — محمود & اسراء | 27 يوليو 2026
 * قاعة فندق بوليفارد
 *
 * خطوات التشغيل:
 * 1) أنشئ Google Sheet جديد وافتح Extensions ▸ Apps Script
 * 2) الصق هذا الكود واحفظ
 * 3) شغّل دالة setupSheet مرة واحدة (سيطلب الصلاحيات) — ستُنشأ ورقة "الضيوف" منسّقة
 * 4) Deploy ▸ New deployment ▸ Web app
 *    Execute as: Me  |  Who has access: Anyone
 * 5) انسخ رابط /exec وضعه في SCRIPT_URL داخل index.html و admin.html
 *******************************************************/

const SHEET_NAME   = 'الضيوف';
const ADMIN_KEY    = 'wedding2026';   // نفس كلمة مرور لوحة الإدارة (لأمر فتح/إغلاق التسجيل)
const CODE_PREFIX  = 'MI';            // بادئة الأكواد (محمود & اسراء)
const CODE_DIGITS  = 5;
const TIMEZONE     = 'Asia/Muscat';

/* ═══════════ نقطة الدخول (JSONP) ═══════════ */
function doGet(e) {
  const p  = (e && e.parameter) || {};
  const cb = p.callback || 'callback';
  let result;
  try {
    switch (String(p.action || '').toLowerCase()) {
      case 'rsvp':      result = handleRsvp(p);      break;
      case 'checkin':   result = handleCheckin(p);   break;
      case 'getguests': result = handleGetGuests();  break;
      case 'status':    result = { status: 'success', registrationOpen: isRegOpen() }; break;
      case 'setreg':    result = handleSetReg(p);    break;
      default:          result = { status: 'error', message: 'إجراء غير معروف' };
    }
  } catch (err) {
    result = { status: 'error', message: 'خطأ في الخادم: ' + err };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(result) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/* ═══════════ إعداد الشيت المنسّق (شغّلها يدوياً مرة واحدة) ═══════════ */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  ss.setSpreadsheetTimeZone(TIMEZONE);
  sh.setRightToLeft(true);

  const headers = ['#', 'اسم الضيف', 'الكود', 'رقم الجوال', 'تاريخ التسجيل', 'الحضور', 'وقت الحضور'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  // تنسيق الترويسة: أخضر ليلي + ذهبي
  sh.getRange(1, 1, 1, headers.length)
    .setBackground('#12332B')
    .setFontColor('#E9D8A6')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);

  // عرض الأعمدة
  const widths = [50, 220, 110, 140, 170, 90, 170];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // محاذاة وتنسيق عام
  sh.getRange('A2:A').setHorizontalAlignment('center');
  sh.getRange('C2:D').setHorizontalAlignment('center').setFontFamily('Courier New');
  sh.getRange('E2:G').setHorizontalAlignment('center');

  // تلوين شرطي: صف أخضر فاتح عند الحضور
  const rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="✅ حضر"')
    .setBackground('#E7F3E7')
    .setRanges([sh.getRange('A2:G1000')])
    .build());
  sh.setConditionalFormatRules(rules);

  // فتح التسجيل افتراضياً
  PropertiesService.getScriptProperties().setProperty('REG_OPEN', 'true');
  Logger.log('✅ تم تجهيز الشيت وفتح التسجيل');
}

/* ═══════════ التسجيل (RSVP) ═══════════ */
function handleRsvp(p) {
  if (!isRegOpen()) return { status: 'closed', message: 'التسجيل مغلق حالياً' };

  const name  = String(p.name || '').trim();
  const phone = normPhone(p.phone);
  if (!name) return { status: 'error', message: 'الاسم مطلوب' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh   = getSheet();
    const data = sh.getDataRange().getValues();

    // قاعدة: جوال واحد = كود واحد
    if (phone) {
      for (let i = 1; i < data.length; i++) {
        if (normPhone(data[i][3]) === phone) {
          return { status: 'exists', name: data[i][1], code: String(data[i][2]) };
        }
      }
    }

    const code = generateCode(data);
    const row  = data.length + 1;
    const now  = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd HH:mm');
    sh.appendRow([row - 1, name, code, phone ? "'" + phone : '', now, '⏳ لم يحضر', '']);
    return { status: 'success', name: name, code: code };
  } finally {
    lock.releaseLock();
  }
}

/* ═══════════ تسجيل الحضور ═══════════ */
function handleCheckin(p) {
  const code = String(p.code || '').trim().toUpperCase();
  if (!code) return { status: 'error', message: 'الكود مطلوب' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh   = getSheet();
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim().toUpperCase() === code) {
        if (String(data[i][5]).indexOf('حضر') === 0 || String(data[i][5]).indexOf('✅') === 0) {
          return { status: 'duplicate', name: data[i][1], attendTime: String(data[i][6]) };
        }
        const now = Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm:ss — yyyy/MM/dd');
        sh.getRange(i + 1, 6).setValue('✅ حضر');
        sh.getRange(i + 1, 7).setValue(now);
        return { status: 'success', name: data[i][1], time: now };
      }
    }
    return { status: 'notfound', message: 'الكود غير موجود' };
  } finally {
    lock.releaseLock();
  }
}

/* ═══════════ جلب الضيوف (للوحة الإدارة) ═══════════ */
function handleGetGuests() {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  const guests = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    guests.push({
      name:       String(data[i][1]),
      code:       String(data[i][2]),
      phone:      String(data[i][3]).replace(/^'/, ''),
      regTime:    String(data[i][4]),
      attended:   String(data[i][5]).indexOf('✅') === 0,
      attendTime: String(data[i][6])
    });
  }
  return { status: 'success', guests: guests, registrationOpen: isRegOpen() };
}

/* ═══════════ فتح/إغلاق التسجيل ═══════════ */
function handleSetReg(p) {
  if (String(p.key || '') !== ADMIN_KEY) return { status: 'error', message: 'غير مصرّح' };
  const open = String(p.open) === 'true';
  PropertiesService.getScriptProperties().setProperty('REG_OPEN', open ? 'true' : 'false');
  return { status: 'success', registrationOpen: open };
}

function isRegOpen() {
  // مفتوح افتراضياً
  return PropertiesService.getScriptProperties().getProperty('REG_OPEN') !== 'false';
}

/* ═══════════ أدوات ═══════════ */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { setupSheet(); sh = ss.getSheetByName(SHEET_NAME); }
  return sh;
}

function normPhone(v) {
  let s = String(v || '').replace(/^'/, '').replace(/\D/g, '');
  if (!s) return '';
  if (s.startsWith('00968')) s = s.slice(5);
  else if (s.startsWith('968') && s.length > 8) s = s.slice(3);
  return s;
}

function generateCode(data) {
  const existing = {};
  for (let i = 1; i < data.length; i++) existing[String(data[i][2])] = true;
  let code;
  do {
    let digits = '';
    for (let i = 0; i < CODE_DIGITS; i++) digits += Math.floor(Math.random() * 10);
    code = CODE_PREFIX + digits;
  } while (existing[code]);
  return code;
}
