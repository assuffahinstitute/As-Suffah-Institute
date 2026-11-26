/**
 * As-Suffah Institute — Online Madrasa Management System
 * Google Apps Script backend. Paste into Code.gs (with Logo.gs and Index.html).
 */
const SHEETS = {
  Banners:    ['BannerID','Image','Title','Link','Order','Show'],
  Board:      ['MemberID','Name','Role','Photo','Order','Show'],
  Reviews:    ['ReviewID','Name','Country','Rating','Text','Status','Date'],
  Students:   ['ID','Name','Guardian','Phone','Course','Version','Country','TeacherID','MonthlyFee','Status','AdmissionDate','Notes','MeetLink'],
  Teachers:   ['Code','Name','Phone','Email','Version','Country','FreeTime','ClassesPerWeek','SharePercent','Status'],
  Payments:   ['PayID','Date','StudentID','Month','Amount','TeacherID','TeacherShare','InstituteShare','InvoiceNo','Method'],
  TeacherPay: ['PayID','Date','TeacherID','Amount','Note'],
  Expenses:   ['ExpID','Date','Category','Amount','Note'],
  Users:      ['Username','FullName','Email','Role','ReadOnly','Pages','PassHash','TeacherCode','Phone'],
  Classes:    ['ClassID','Date','Time','StudentID','TeacherID','StudentStatus','StudentReason','TeacherStatus','TeacherReason','MarkedAt','MarkedBy','StartedAt','EndedAt','Minutes'],
  Fields:     ['FieldID','Entity','Label','Type','Options','Required','Placeholder','ShowInTable','Filter','Chart','Order'],
  Themes:     ['ThemeID','Word','BgColor','TextColor','Description']
};
const ENT = {Students:1, Teachers:1, Expenses:1};
const FTYPES = ['text','email','phone','number','date','textarea','select','multiselect','url','file','autoid'];
const NUM = {Minutes:1,Rating:1,Order:1,Amount:1,MonthlyFee:1,ClassesPerWeek:1,SharePercent:1,TeacherShare:1,InstituteShare:1};

/* The website calls these functions over HTTPS (see portal.html). Only this list is reachable. */
const PUBLIC_FNS = {getPublic, submitReview, login, logout, api, requestReset, confirmReset, registerTeacher, getRegForm};
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = Object.prototype.hasOwnProperty.call(PUBLIC_FNS, req.fn) ? PUBLIC_FNS[req.fn] : null;
    if (!fn) throw new Error('Unknown function');
    out = {result: fn.apply(null, Array.isArray(req.args) ? req.args : [])};
  } catch (err) {
    out = {error: String((err && err.message) || err)};
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.ping)   // quick check that the latest version is deployed
    return ContentService.createTextOutput(JSON.stringify({ok: true, api: 'v2'})).setMimeType(ContentService.MimeType.JSON);
  const t = HtmlService.createTemplateFromFile('Index');
  const m = e && e.parameter ? e.parameter.m : '';
  t.MODE = ['teacher', 'register'].indexOf(m) > -1 ? m : 'admin';   // whitelisted: lets the website open the right login tab
  t.LOGO = 'data:image/png;base64,' + LOGO_B64;
  return t.evaluate().setTitle('As-Suffah Institute')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);   // lets the website show this app inside itself
}

/* ---------- helpers ---------- */
const CHK_ = {};
const FB_DEFAULT = 'https://www.facebook.com/profile.php?id=61592474635111';
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) {
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    const h = SHEETS[name];
    sh.getRange(1,1,sh.getMaxRows(),h.length).setNumberFormat('@');
    sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold').setBackground('#D9902A').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  if (!CHK_[name]) { // add any newly introduced base columns to sheets created by an older version
    CHK_[name] = 1;
    const want = SHEETS[name];
    if (want) {
      const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      want.forEach(k => {
        if (h.indexOf(k) < 0) {
          const col = sh.getLastColumn() + 1;
          if (col > sh.getMaxColumns()) sh.insertColumnAfter(sh.getMaxColumns());
          sh.getRange(1, col, sh.getMaxRows(), 1).setNumberFormat('@');
          sh.getRange(1, col).setValue(k).setFontWeight('bold').setBackground('#D9902A').setFontColor('#fff');
          h.push(k);
          drop_(name);
        }
      });
    }
  }
  return sh;
}
const MEM_ = {};
function drop_(name) { delete MEM_[name]; try { CacheService.getScriptCache().remove('sh_' + name); } catch (e) {} }
function all_(name) {
  const cp = rows => rows.map(o => Object.assign({}, o));
  if (MEM_[name]) return cp(MEM_[name]);               // same request: no re-read
  const cache = CacheService.getScriptCache();
  const hit = cache.get('sh_' + name);                  // other requests within 90 s
  if (hit) { MEM_[name] = JSON.parse(hit); return cp(MEM_[name]); }
  const sh = sheet_(name), v = sh.getDataRange().getValues(), h = v[0];
  const rows = v.slice(1).filter(r => r[0] !== '').map(r => {
    const o = {};
    h.forEach((k,i) => o[k] = NUM[k] ? Number(r[i]) || 0 : (r[i] instanceof Date ? Utilities.formatDate(r[i], 'Asia/Dhaka', 'yyyy-MM-dd') : String(r[i])));
    return o;
  });
  MEM_[name] = rows;
  try { const js = JSON.stringify(rows); if (js.length < 90000) cache.put('sh_' + name, js, 90); } catch (e) {}
  return cp(rows);
}
function headers_(sh) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]; }
function add_(name, obj) {
  const sh = sheet_(name), h = headers_(sh);
  sh.getRange(sh.getLastRow() + 1, 1, 1, h.length).setNumberFormat('@').setValues([h.map(k => obj[k] === undefined ? '' : obj[k])]);
  drop_(name);
}
function addMany_(name, objs) {
  if (!objs.length) return;
  const sh = sheet_(name), h = headers_(sh);
  sh.getRange(sh.getLastRow() + 1, 1, objs.length, h.length).setNumberFormat('@').setValues(objs.map(o => h.map(k => o[k] === undefined ? '' : o[k])));
  drop_(name);
}
function update_(name, id, patch) {
  const sh = sheet_(name), v = sh.getDataRange().getValues(), h = v[0];
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(id)) {
    const row = v[i].slice();
    Object.keys(patch).forEach(k => { const c = h.indexOf(k); if (c > -1) row[c] = patch[k]; });
    sh.getRange(i + 1, 1, 1, h.length).setNumberFormat('@').setValues([row]);
    drop_(name);
    return true;
  }
  throw new Error('Record not found');
}
function del_(name, id) {
  const sh = sheet_(name), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(id)) { sh.deleteRow(i + 1); drop_(name); return true; }
  throw new Error('Record not found');
}
function nextIds_(prefix, name, pad, count) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const p = PropertiesService.getScriptProperties(), k = 'seq_' + name;
    const n = Number(p.getProperty(k) || 0);
    p.setProperty(k, String(n + count));
    const out = []; for (let i = 1; i <= count; i++) out.push(prefix + String(n + i).padStart(pad, '0'));
    return out;
  } finally { lock.releaseLock(); }
}
function nextId_(prefix, name, pad) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const p = PropertiesService.getScriptProperties();
    const k = 'seq_' + name;
    const n = Number(p.getProperty(k) || 0) + 1;
    p.setProperty(k, String(n));
    return prefix + String(n).padStart(pad, '0');
  } finally { lock.releaseLock(); }
}
const today_ = () => Utilities.formatDate(new Date(), 'Asia/Dhaka', 'yyyy-MM-dd');
function hash_(s) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'assuffah|' + s));
}

/* ---------- auth ---------- */
function ensureAdmin_() {
  if (all_('Users').length === 0)
    add_('Users', {Username:'admin', FullName:'Administrator', Email:'', Role:'Admin', ReadOnly:'No', Pages:'*', PassHash:hash_('admin123')});
}
function canPage_(me, p) { return me.role === 'Admin' || me.pages === '*' || String(me.pages).split(',').indexOf(p) > -1; }
function init_(me) {
  ensureThemes_();
  return {settings: settings_(), teachers: me.role === 'Teacher' ? [] : all_('Teachers'), me,
          fields: all_('Fields').sort((a,b) => a.Order - b.Order), themes: all_('Themes')};
}
function normPhone_(s) { return String(s || '').replace(/\D/g, '').slice(-10); }
function findUser_(ident) {
  const id = String(ident || '').trim().toLowerCase();
  if (!id) return null;
  const users = all_('Users');
  let u = users.find(x => x.Username.toLowerCase() === id || (x.Email && x.Email.toLowerCase() === id));
  if (u) return u;
  const p = normPhone_(id);
  if (p.length >= 10 && id.indexOf('@') < 0) {
    u = users.find(x => normPhone_(x.Phone) === p);
    if (u) return u;
    const t = all_('Teachers').find(x => normPhone_(x.Phone) === p);
    if (t) return users.find(x => x.TeacherCode === t.Code) || null;
  }
  return null;
}
function login(username, password, mode) {
  ensureAdmin_();
  Utilities.sleep(150);
  const u = findUser_(username);
  if (!u || u.PassHash !== hash_(password)) return {ok:false};
  if (mode === 'teacher' && u.Role !== 'Teacher') return {ok:false, msg:'wrongtab_admin'};
  if (mode === 'admin' && u.Role === 'Teacher') return {ok:false, msg:'wrongtab_teacher'};
  if (u.Role === 'Teacher') {
    const t = all_('Teachers').find(x => x.Code === u.TeacherCode);
    if (!t || t.Status === 'Inactive') return {ok:false, msg:'inactive'};
    if (t.Status === 'Pending') return {ok:false, msg:'pending'};
  }
  const token = Utilities.getUuid();
  const me = {username:u.Username, name:u.FullName || u.Username, role:u.Role, ro:u.ReadOnly === 'Yes', pages:u.Pages || '*', email:u.Email, teacherCode:u.TeacherCode || ''};
  CacheService.getScriptCache().put('tok_' + token, JSON.stringify(me), 21600);
  let first = null;   // first screen's data is sent with the login to save two round trips
  try { first = me.role === 'Teacher' ? tdash_(me) : (canPage_(me, 'dashboard') ? dashboard_() : null); } catch (e) {}
  return {ok:true, token, me, defaultPw: u.PassHash === hash_('admin123'), boot:{init: init_(me), first}};
}
function auth_(token) {
  const v = token && CacheService.getScriptCache().get('tok_' + token);
  if (!v) throw new Error('SESSION_EXPIRED');
  return JSON.parse(v);
}
function logout(token) { CacheService.getScriptCache().remove('tok_' + token); }

/* ---------- settings ---------- */
function settings_() {
  const p = PropertiesService.getScriptProperties();
  return {
    name: p.getProperty('s_name') || 'As-Suffah Institute',
    address: p.getProperty('s_address') || '',
    phone: p.getProperty('s_phone') || '',
    emergency: p.getProperty('s_emergency') || '',
    facebook: p.getProperty('s_facebook') || FB_DEFAULT
  };
}

/* ---------- single entry point ---------- */
const WRITE = ['addClasses','updClass','delClass','attSave','addStudent','updStudent','delStudent','addTeacher','updTeacher','delTeacher','addPayment','delPayment',
  'addTeacherPay','addExpense','delExpense','saveSettings'];
const ADMIN = ['cmsList','cmsAdd','cmsUpd','cmsDel','users','addUser','updUser','delUser','saveSettings','addField','updField','delField','addTheme','updTheme','delTheme'];
WRITE.push('cmsAdd','cmsUpd','cmsDel','clsStart','clsEnd','addField','updField','delField','addTheme','updTheme','delTheme','upload');
const ACT_PAGES = {
  dashboard:['dashboard'],
  students:['students','admission','finance','invoices','attendance'],
  classes:['attendance'], classHours:['attendance'], addClasses:['attendance'], updClass:['attendance'], delClass:['attendance'], addStudent:['admission','students'], updStudent:['students'], delStudent:['students'],
  teachers:['teachers','finance'], addTeacher:['teachers'], updTeacher:['teachers'], delTeacher:['teachers'],
  payments:['invoices','finance'], addPayment:['finance'], delPayment:['invoices'], invoice:['invoices','finance'], invoicePdf:['invoices','finance'],
  teacherPays:['finance'], addTeacherPay:['finance'], earnings:['finance'], expenses:['finance'], addExpense:['finance'], delExpense:['finance'],
  upload:['admission','students','teachers','finance']
};

function api(token, action, d) {
  const me = auth_(token);
  d = d || {};
  if (me.role === 'Teacher' && ['init','tdash','attSave','clsStart','clsEnd','changePassword','myProfile','saveMyProfile','upload'].indexOf(action) < 0) throw new Error('Not allowed');
  if (me.ro && WRITE.indexOf(action) > -1) throw new Error('Read-only account / শুধু দেখার অনুমতি আছে');
  if (me.role !== 'Admin' && ADMIN.indexOf(action) > -1) throw new Error('Admin only');
  if (me.role !== 'Admin' && me.role !== 'Teacher' && me.pages !== '*' && ACT_PAGES[action]) {
    const mine = String(me.pages).split(',');
    if (!ACT_PAGES[action].some(p => mine.indexOf(p) > -1)) throw new Error('No permission for this page / এই পেজের অনুমতি নেই');
  }
  switch (action) {
    case 'init':       return init_(me);
    case 'tdash':      return tdash_(me);
    case 'clsStart':   return clsStart_(me, d);
    case 'clsEnd':     return clsEnd_(me, d);
    case 'classHours': return classHours_();
    case 'cmsList':    return {Banners: all_('Banners').sort(byOrder_), Board: all_('Board').sort(byOrder_), Reviews: all_('Reviews').reverse()};
    case 'cmsAdd':     return cmsAdd_(d);
    case 'cmsUpd':     cmsMeta_(d.sheet); update_(d.sheet, d.id, d.patch); pubDrop_(); return true;
    case 'cmsDel':     cmsMeta_(d.sheet); del_(d.sheet, d.id); pubDrop_(); return true;
    case 'myProfile':  return myProfile_(me);
    case 'saveMyProfile': return saveMyProfile_(me, d);
    case 'attSave':    return attSave_(me, d);
    case 'classes':    return all_('Classes').sort((a, b) => (b.Date + b.Time) < (a.Date + a.Time) ? -1 : 1);
    case 'addClasses': return addClasses_(d);
    case 'updClass':   return updClass_(me, d);
    case 'delClass':   del_('Classes', d.ClassID); return true;
    case 'dashboard':  return dashboard_();
    case 'students':   return all_('Students');
    case 'addStudent': return addStudent_(d);
    case 'updStudent': if (d.patch.MonthlyFee !== undefined) d.patch.MonthlyFee = Number(d.patch.MonthlyFee) || 0; update_('Students', d.ID, d.patch); return true;
    case 'delStudent': del_('Students', d.ID); return true;
    case 'teachers':   return teachersFull_();
    case 'addTeacher': return addTeacher_(d);
    case 'updTeacher': return updTeacher_(d);
    case 'delTeacher':
      if (all_('Students').some(s => s.TeacherID === d.Code)) throw new Error('This teacher still has students — reassign them first / আগে শিক্ষার্থীদের অন্য শিক্ষকে সরান');
      del_('Teachers', d.Code);
      const tu = all_('Users').find(u => u.TeacherCode === d.Code); if (tu) del_('Users', tu.Username);
      return true;
    case 'payments':   return all_('Payments').reverse();
    case 'addPayment': return addPayment_(d);
    case 'delPayment': del_('Payments', d.PayID); return true;
    case 'teacherPays':return all_('TeacherPay').reverse();
    case 'addTeacherPay': d.PayID = nextId_('TP-', 'tpay', 4); d.Date = d.Date || today_(); add_('TeacherPay', d); return true;
    case 'earnings':   return teachersFull_();
    case 'expenses':   return all_('Expenses').reverse();
    case 'addExpense': fillAuto_('Expenses', d); d.ExpID = nextId_('EX-', 'exp', 4); d.Date = d.Date || today_(); add_('Expenses', d); return true;
    case 'delExpense': del_('Expenses', d.ExpID); return true;
    case 'invoice':    return invoice_(d.PayID);
    case 'invoicePdf': return invoicePdf_(d.PayID);
    case 'users':      return all_('Users').filter(u => u.Role !== 'Teacher').map(u => { delete u.PassHash; return u; });
    case 'addUser':
      if (!/^[A-Za-z0-9_.@-]{3,40}$/.test(d.Username || '')) throw new Error('Username: 3-40 letters, digits, . _ - @ only');
      if (!d.pw || d.pw.length < 6) throw new Error('Username and a 6+ character password are required');
      if (all_('Users').some(u => u.Username.toLowerCase() === d.Username.toLowerCase())) throw new Error('Username already exists');
      add_('Users', {Username:d.Username, FullName:d.FullName, Email:d.Email, Phone:d.Phone || '', Role:d.Role || 'Staff', ReadOnly:d.ReadOnly || 'No', Pages:d.Pages || '*', PassHash:hash_(d.pw)});
      return true;
    case 'updUser':
      const patch = d.patch; if (patch.pw) { patch.PassHash = hash_(patch.pw); delete patch.pw; }
      update_('Users', d.Username, patch); return true;
    case 'delUser':
      if (d.Username === me.username) throw new Error('You cannot delete your own account');
      del_('Users', d.Username); return true;
    case 'upload':     return upload_(d);
    case 'addField':   return addField_(d);
    case 'updField':   return updField_(d);
    case 'delField':   del_('Fields', d.FieldID); return true;
    case 'addTheme':   d.ThemeID = nextId_('TH-', 'theme', 3); add_('Themes', d); return true;
    case 'updTheme':   update_('Themes', d.ThemeID, d.patch); return true;
    case 'delTheme':   del_('Themes', d.ThemeID); return true;
    case 'saveProfile':update_('Users', me.username, {FullName: d.FullName || me.name, Email: d.Email || ''}); return true;
    case 'saveSettings':
      const p = PropertiesService.getScriptProperties();
      ['name','address','phone','emergency'].forEach(k => p.setProperty('s_' + k, d[k] || ''));
      p.setProperty('s_facebook', /^https:\/\//i.test(d.facebook || '') ? d.facebook : '');
      return settings_();
    case 'changePassword':
      if (!d.pw || d.pw.length < 6) throw new Error('Password must be at least 6 characters');
      update_('Users', me.username, {PassHash: hash_(d.pw)}); return true;
  }
  throw new Error('Unknown action');
}

/* ---------- business logic ---------- */
function addStudent_(d) {
  d.ID = nextId_('ASI-', 'student', 4);
  fillAuto_('Students', d);
  d.Status = 'Active'; d.AdmissionDate = d.AdmissionDate || today_();
  d.MonthlyFee = Number(d.MonthlyFee) || 0;
  add_('Students', d);
  return d.ID;
}
function addPayment_(d) {
  const st = all_('Students').find(s => s.ID === d.StudentID);
  if (!st) throw new Error('Student not found');
  const amt = Number(d.Amount);
  if (!(amt > 0)) throw new Error('Invalid amount');
  const tc = all_('Teachers').find(t => t.Code === st.TeacherID);
  const pct = tc ? tc.SharePercent : 0;
  const share = Math.round(amt * pct) / 100;
  const rec = {
    PayID: nextId_('PAY-', 'pay', 5), Date: d.Date || today_(), StudentID: st.ID,
    Month: d.Month, Amount: amt, TeacherID: st.TeacherID || '',
    TeacherShare: share, InstituteShare: amt - share,
    InvoiceNo: nextId_('INV-', 'inv', 5), Method: d.Method || 'Cash'
  };
  add_('Payments', rec);
  return rec;
}
function teachersFull_() {
  const st = all_('Students'), pay = all_('Payments'), tp = all_('TeacherPay');
  const rev = {}, earn = {}, paid = {}, mine = {};
  pay.forEach(p => { rev[p.TeacherID] = (rev[p.TeacherID] || 0) + p.Amount; earn[p.TeacherID] = (earn[p.TeacherID] || 0) + p.TeacherShare; });
  tp.forEach(p => { paid[p.TeacherID] = (paid[p.TeacherID] || 0) + p.Amount; });
  st.forEach(s => { (mine[s.TeacherID] = mine[s.TeacherID] || []).push(s); });
  return all_('Teachers').map(t => {
    const m = mine[t.Code] || [], revenue = rev[t.Code] || 0, earned = earn[t.Code] || 0, pd = paid[t.Code] || 0;
    return Object.assign({}, t, {
      studentCount: m.length, activeCount: m.filter(s => s.Status === 'Active').length,
      students: m.map(s => ({ID:s.ID, Name:s.Name, Status:s.Status, Country:s.Country, Version:s.Version})),
      revenue, earned, paid: pd, due: earned - pd, instituteIncome: revenue - earned
    });
  });
}
function dashboard_() {
  const st = all_('Students'), pay = all_('Payments'), tp = all_('TeacherPay'), ex = all_('Expenses');
  const sum = (a, f) => a.reduce((x, y) => x + f(y), 0);
  const collected = sum(pay, p => p.Amount), salaries = sum(tp, p => p.Amount), other = sum(ex, e => e.Amount);
  const m = today_().slice(0, 7);
  const monthly = {};
  pay.forEach(p => { const k = p.Date.slice(0,7); (monthly[k] = monthly[k] || {inc:0, exp:0}).inc += p.Amount; });
  tp.forEach(p => { const k = p.Date.slice(0,7); (monthly[k] = monthly[k] || {inc:0, exp:0}).exp += p.Amount; });
  ex.forEach(p => { const k = p.Date.slice(0,7); (monthly[k] = monthly[k] || {inc:0, exp:0}).exp += p.Amount; });
  const paidThisMonth = new Set(pay.filter(p => p.Month === m).map(p => p.StudentID));
  return {
    totalStudents: st.length, activeStudents: st.filter(s => s.Status === 'Active').length,
    newThisMonth: st.filter(s => s.AdmissionDate.slice(0,7) === m).length,
    teachers: all_('Teachers').length, pendingTeachers: all_('Teachers').filter(t => t.Status === 'Pending').length,
    income: collected, salaries, otherExpense: other, expense: salaries + other,
    net: collected - salaries - other,
    incomeThisMonth: (monthly[m] || {inc: 0}).inc,
    charts: builtinCharts_(st).concat(charts_()),
    att: attToday_(),
    unpaid: st.filter(s => s.Status === 'Active' && !paidThisMonth.has(s.ID)).length,
    monthly: Object.keys(monthly).sort().slice(-12).map(k => Object.assign({month:k}, monthly[k]))
  };
}

/* ---------- invoice ---------- */
function invoiceData_(payId) {
  const p = all_('Payments').find(x => x.PayID === payId);
  if (!p) throw new Error('Payment not found');
  const s = all_('Students').find(x => x.ID === p.StudentID) || {};
  const t = all_('Teachers').find(x => x.Code === p.TeacherID) || {};
  return {p, s, t, set: settings_()};
}
function invoiceHtml_(payId) {
  const {p, s, t, set} = invoiceData_(payId);
  const esc = x => String(x || '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
  return '<html><head><meta charset="utf-8"><style>' +
   'body{font-family:"Noto Serif Bengali","Noto Serif",serif;color:#0B2D36;margin:0;padding:28px}' +
   '.top{background:#0B2D36;color:#fff;padding:16px 22px;border-bottom:5px solid #D9902A}' +
   '.top img{height:70px;vertical-align:middle;margin-right:14px}.top b{font-size:22px;color:#D9902A}' +
   'table{width:100%;border-collapse:collapse;margin-top:18px}td,th{padding:9px 10px;border-bottom:1px solid #e4d9c4;text-align:left;font-size:14px}' +
   'th{background:#faf3e4}.tot{font-size:18px;font-weight:bold;color:#B5701A}.f{margin-top:30px;font-size:12px;color:#666;text-align:center}' +
   '</style></head><body><div class="top"><img src="data:image/png;base64,' + LOGO_B64 + '"><b>' + esc(set.name) + '</b><br>' +
   '<span style="font-size:12px">' + esc(set.address) + ' ' + esc(set.phone) + '</span></div>' +
   '<h2 style="margin:18px 0 0">ফি রশিদ / Fee Invoice</h2>' +
   '<table><tr><th>Invoice No / ইনভয়েস নং</th><td>' + esc(p.InvoiceNo) + '</td><th>Date / তারিখ</th><td>' + esc(p.Date) + '</td></tr>' +
   '<tr><th>Student / শিক্ষার্থী</th><td>' + esc(s.Name) + '</td><th>ID / আইডি</th><td>' + esc(s.ID) + '</td></tr>' +
   '<tr><th>Guardian / অভিভাবক</th><td>' + esc(s.Guardian) + '</td><th>Course / কোর্স</th><td>' + esc(s.Course) + '</td></tr>' +
   '<tr><th>Teacher / শিক্ষক</th><td>' + esc(t.Name) + '</td><th>Month / মাস</th><td>' + esc(p.Month) + '</td></tr>' +
   '<tr><th>Method / পদ্ধতি</th><td>' + esc(p.Method) + '</td><th>Amount / পরিমাণ</th><td class="tot">৳ ' + p.Amount.toLocaleString('en-US') + '</td></tr></table>' +
   '<div class="f">জাযাকাল্লাহু খাইরান — Thank you. This is a computer-generated invoice.' + (set.facebook ? '<br>Facebook: ' + esc(set.facebook) : '') + '</div></body></html>';
}
function invoice_(payId) { return {html: invoiceHtml_(payId)}; }
function invoicePdf_(payId) {
  const blob = Utilities.newBlob(invoiceHtml_(payId), 'text/html', 'invoice.html').getAs('application/pdf');
  return {name: 'Invoice-' + payId + '.pdf', b64: Utilities.base64Encode(blob.getBytes())};
}


/* ---------- custom fields (dynamic columns) ---------- */
function fillAuto_(entity, d) {
  all_('Fields').filter(f => f.Entity === entity && f.Type === 'autoid').forEach(f => { d[f.Label] = nextId_('', 'f_' + f.FieldID, 1); });
}
function addField_(d) {
  if (!ENT[d.Entity]) throw new Error('Invalid entity');
  if (FTYPES.indexOf(d.Type) < 0) throw new Error('Invalid type');
  const label = String(d.Label || '').trim();
  if (!label || label.length > 40) throw new Error('Field name required (max 40 chars)');
  const sh = sheet_(d.Entity), h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (h.some(x => String(x).toLowerCase() === label.toLowerCase())) throw new Error('A column with this name already exists');
  const c = h.length + 1;
  if (c > sh.getMaxColumns()) sh.insertColumnAfter(sh.getMaxColumns());
  sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@');
  sh.getRange(1, c).setValue(label).setFontWeight('bold').setBackground('#12454B').setFontColor('#fff');
  drop_(d.Entity);
  add_('Fields', {FieldID: nextId_('F-', 'field', 3), Entity: d.Entity, Label: label, Type: d.Type, Options: d.Options || '',
    Required: d.Required || 'No', Placeholder: d.Placeholder || '', ShowInTable: d.ShowInTable || 'Yes', Filter: d.Filter || 'No',
    Chart: d.Chart || 'None', Order: all_('Fields').length + 1});
  return true;
}
function updField_(d) {
  const f = all_('Fields').find(x => x.FieldID === d.FieldID);
  if (!f) throw new Error('Field not found');
  const p = d.patch; delete p.Entity;
  if (p.Type && FTYPES.indexOf(p.Type) < 0) throw new Error('Invalid type');
  if (p.Label !== undefined && p.Label.trim() !== f.Label) {
    const label = p.Label.trim(); if (!label) throw new Error('Field name required');
    const sh = sheet_(f.Entity), h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (h.some(x => String(x).toLowerCase() === label.toLowerCase())) throw new Error('A column with this name already exists');
    const c = h.indexOf(f.Label); if (c > -1) sh.getRange(1, c + 1).setValue(label);
    drop_(f.Entity);
    p.Label = label;
  } else delete p.Label;
  update_('Fields', d.FieldID, p);
  return true;
}
function upload_(d) {
  if (!d.b64 || d.b64.length > 7000000) throw new Error('File too large (max ~5 MB) / ফাইল ৫ MB-এর বেশি');
  const props = PropertiesService.getScriptProperties();
  let folder;
  try { folder = DriveApp.getFolderById(props.getProperty('upload_folder')); }
  catch (e) { folder = DriveApp.createFolder('As-Suffah Uploads'); props.setProperty('upload_folder', folder.getId()); }
  const name = String(d.name || 'file').replace(/[\\\/:*?"<>|]/g, '_').slice(0, 80);
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(d.b64), d.mime || 'application/octet-stream', name));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return {url: file.getUrl(), name: name, id: file.getId()};
}
function builtinCharts_(st) {
  const count = k => { const m = {}; st.forEach(s => { const v = String(s[k] || '').trim(); if (v) m[v] = (m[v] || 0) + 1; });
    return Object.keys(m).map(x => ({v: x, n: m[x]})).sort((a, b) => b.n - a.n); };
  return [{key: 'country', entity: 'Students', chart: 'Bar', data: count('Country')},
          {key: 'version', entity: 'Students', chart: 'Pie', data: count('Version')}];
}
function charts_() {
  return all_('Fields').filter(f => ENT[f.Entity] && f.Chart && f.Chart !== 'None').map(f => {
    const cnt = {};
    all_(f.Entity).forEach(r => String(r[f.Label] || '').split(',').map(x => x.trim()).filter(Boolean).forEach(v => cnt[v] = (cnt[v] || 0) + 1));
    return {label: f.Label, entity: f.Entity, chart: f.Chart, data: Object.keys(cnt).map(k => ({v: k, n: cnt[k]}))};
  });
}

/* ---------- themes (badge colours) ---------- */
function ensureThemes_() {
  const have = {}; all_('Themes').forEach(t => have[t.Word.toLowerCase()] = 1);
  const want = [['Active','#e3f3ea','#1f6b49','Active / সক্রিয়'],['Inactive','#f7e4e0','#a3301f','Inactive / নিষ্ক্রিয়'],
    ['Pending','#fbefd3','#8a5d00','Pending / অপেক্ষমাণ'],['Paid','#dff0f2','#12454B','Paid / পরিশোধিত'],
    ['Unpaid','#f7e4e0','#a3301f','Unpaid / বাকি'],['Done','#e3f3ea','#1f6b49','Done / সম্পন্ন'],
    ['Present','#e3f3ea','#1f6b49','Present / উপস্থিত'],['Absent','#f7e4e0','#a3301f','Absent / অনুপস্থিত'],['Approved','#e3f3ea','#1f6b49','Approved / অনুমোদিত'],['Hidden','#f7e4e0','#a3301f','Hidden / লুকানো']]
    .filter(w => !have[w[0].toLowerCase()]);
  if (!want.length) return;
  const ids = nextIds_('TH-', 'theme', 3, want.length);
  addMany_('Themes', want.map((w, i) => ({ThemeID: ids[i], Word: w[0], BgColor: w[1], TextColor: w[2], Description: w[3]})));
}

/* ---------- password reset by email OTP ---------- */
function requestReset(ident) {
  ident = String(ident || '').trim().toLowerCase();
  if (!ident) return {ok: true};
  ensureAdmin_();
  const u = findUser_(ident);
  const cache = CacheService.getScriptCache();
  if (u && u.Email && !cache.get('rl_' + u.Username)) {
    const code = String(100000 + (parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 7), 16) % 900000));
    cache.put('otp_' + u.Username, JSON.stringify({h: hash_(code), a: 0}), 600);
    cache.put('rl_' + u.Username, '1', 60);
    MailApp.sendEmail({to: u.Email, subject: 'As-Suffah Institute — password reset code / পাসওয়ার্ড রিসেট কোড',
      htmlBody: '<div style="font-family:serif;max-width:420px"><h3 style="color:#0B2D36">As-Suffah Institute</h3>' +
        '<p>আপনার পাসওয়ার্ড রিসেট কোড / Your reset code:</p><p style="font-size:28px;letter-spacing:6px;color:#B5701A"><b>' + code + '</b></p>' +
        '<p style="color:#666">১০ মিনিট পর্যন্ত বৈধ / Valid for 10 minutes. আপনি না চাইলে এটি উপেক্ষা করুন / Ignore if not requested.</p></div>'});
  }
  return {ok: true}; // same answer whether or not the account exists
}
function confirmReset(ident, code, newPw) {
  ident = String(ident || '').trim().toLowerCase();
  if (!newPw || newPw.length < 6) return {ok: false, msg: 'Password must be at least 6 characters'};
  const u = findUser_(ident);
  const cache = CacheService.getScriptCache();
  const raw = u && cache.get('otp_' + u.Username);
  if (!raw) return {ok: false, msg: 'Code expired or invalid'};
  const o = JSON.parse(raw);
  if (o.a >= 5) { cache.remove('otp_' + u.Username); return {ok: false, msg: 'Too many attempts'}; }
  if (o.h !== hash_(String(code).trim())) { o.a++; cache.put('otp_' + u.Username, JSON.stringify(o), 600); return {ok: false, msg: 'Wrong code'}; }
  update_('Users', u.Username, {PassHash: hash_(newPw)});
  cache.remove('otp_' + u.Username);
  return {ok: true};
}


/* ---------- teachers: each gets their own login (email + password) ---------- */
const EMAIL_RE = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
function addTeacher_(d) {
  const email = String(d.Email || '').trim().toLowerCase();
  if (!EMAIL_RE.test(email)) throw new Error('Valid email required / সঠিক ইমেইল দিন');
  if (!d.pw || d.pw.length < 6) throw new Error('Password must be at least 6 characters / পাসওয়ার্ড কমপক্ষে ৬ অক্ষর');
  if (all_('Users').some(u => u.Username.toLowerCase() === email || String(u.Email).toLowerCase() === email)) throw new Error('This email is already used / এই ইমেইল আগে ব্যবহৃত');
  const pw = d.pw; delete d.pw;
  fillAuto_('Teachers', d);
  d.Code = nextId_('T-', 'teacher', 3); d.Status = 'Active'; d.Email = email;
  add_('Teachers', d);
  add_('Users', {Username: email, FullName: d.Name, Email: email, Role: 'Teacher', ReadOnly: 'No', Pages: 'teacher', TeacherCode: d.Code, Phone: d.Phone || '', PassHash: hash_(pw)});
  return d.Code;
}
function updTeacher_(d) {
  const p = d.patch || {}, pw = p.pw; delete p.pw;
  if (pw && pw.length < 6) throw new Error('Password must be at least 6 characters / পাসওয়ার্ড কমপক্ষে ৬ অক্ষর');
  const users = all_('Users'), u = users.find(x => x.TeacherCode === d.Code);
  if (p.Email !== undefined) {
    p.Email = String(p.Email).trim().toLowerCase();
    if (p.Email) {
      if (!EMAIL_RE.test(p.Email)) throw new Error('Valid email required / সঠিক ইমেইল দিন');
      if (users.some(x => x.TeacherCode !== d.Code && (x.Username.toLowerCase() === p.Email || String(x.Email).toLowerCase() === p.Email))) throw new Error('This email is already used / এই ইমেইল আগে ব্যবহৃত');
    }
  }
  update_('Teachers', d.Code, p);
  if (u) {
    const up = {};
    if (p.Email && p.Email !== u.Username) { up.Username = p.Email; up.Email = p.Email; }
    if (p.Name) up.FullName = p.Name;
    if (p.Phone !== undefined) up.Phone = p.Phone;
    if (pw) up.PassHash = hash_(pw);
    if (Object.keys(up).length) update_('Users', u.Username, up);
  } else if (p.Email && pw) {
    add_('Users', {Username: p.Email, FullName: p.Name || '', Email: p.Email, Role: 'Teacher', ReadOnly: 'No', Pages: 'teacher', TeacherCode: d.Code, Phone: p.Phone || '', PassHash: hash_(pw)});
  }
  return true;
}

/* ---------- class schedule + attendance ---------- */
function nowDhaka_() {
  const d = new Date();
  return {date: Utilities.formatDate(d, 'Asia/Dhaka', 'yyyy-MM-dd'), time: Utilities.formatDate(d, 'Asia/Dhaka', 'HH:mm')};
}
function tdash_(me) {
  const code = me.teacherCode, now = nowDhaka_();
  const st = all_('Students').filter(s => s.TeacherID === code);
  const byId = {}; st.forEach(s => byId[s.ID] = s);
  const cls = all_('Classes').filter(c => c.TeacherID === code && byId[c.StudentID]);
  const slim = c => ({ClassID:c.ClassID, Date:c.Date, Time:c.Time, StudentID:c.StudentID, Name:byId[c.StudentID].Name,
    StudentStatus:c.StudentStatus, StudentReason:c.StudentReason, TeacherStatus:c.TeacherStatus, TeacherReason:c.TeacherReason,
    MarkedAt:c.MarkedAt, StartedAt:c.StartedAt, EndedAt:c.EndedAt, Minutes:c.Minutes, MeetLink:byId[c.StudentID].MeetLink, open: c.Date === now.date && now.time >= c.Time});
  const asc = (a, b) => (a.Date + a.Time) < (b.Date + b.Time) ? -1 : 1;
  const next = {};
  cls.filter(c => c.Date >= now.date).sort(asc).forEach(c => { if (!next[c.StudentID]) next[c.StudentID] = {Date:c.Date, Time:c.Time}; });
  return {
    now,
    counts: {total: st.length, active: st.filter(s => s.Status === 'Active').length, inactive: st.filter(s => s.Status !== 'Active').length},
    students: st.map(s => ({ID:s.ID, Name:s.Name, Status:s.Status, next: next[s.ID] || null})),
    today: cls.filter(c => c.Date === now.date && byId[c.StudentID].Status === 'Active').sort(asc).map(slim),
    upcoming: cls.filter(c => c.Date > now.date).sort(asc).slice(0, 40).map(slim),
    past: cls.filter(c => c.Date < now.date).sort((a, b) => asc(b, a)).slice(0, 40).map(slim)
  };
}
function attPatch_(d) {
  const ok = v => v === '' || v === 'Present' || v === 'Absent';
  const p = {StudentStatus: d.StudentStatus || '', TeacherStatus: d.TeacherStatus || '',
             StudentReason: String(d.StudentReason || '').slice(0, 200), TeacherReason: String(d.TeacherReason || '').slice(0, 200)};
  if (!ok(p.StudentStatus) || !ok(p.TeacherStatus)) throw new Error('Invalid status');
  if (p.StudentStatus !== 'Absent') p.StudentReason = '';
  if (p.TeacherStatus !== 'Absent') p.TeacherReason = '';
  return p;
}
function attSave_(me, d) {
  const c = all_('Classes').find(x => x.ClassID === d.ClassID);
  if (!c || c.TeacherID !== me.teacherCode) throw new Error('Class not found');
  const now = nowDhaka_();
  if (c.Date < now.date) throw new Error('এই দিনের এন্ট্রির সময় শেষ (রাত ১২টার পর আর হয় না) / Entry closed after midnight');
  if (c.Date > now.date || now.time < c.Time) throw new Error('এখনো ক্লাসের সময় হয়নি / Class has not started yet');
  const p = attPatch_(d);
  if (!p.TeacherStatus) throw new Error('Select teacher present/absent / শিক্ষক উপস্থিত না অনুপস্থিত বাছুন');
  if (p.TeacherStatus === 'Present' && !p.StudentStatus) throw new Error('Select student present/absent / শিক্ষার্থী উপস্থিত না অনুপস্থিত বাছুন');
  p.MarkedAt = now.date + ' ' + now.time; p.MarkedBy = me.username;
  update_('Classes', c.ClassID, p);
  return {MarkedAt: p.MarkedAt};
}
function addClasses_(d) {
  const s = all_('Students').find(x => x.ID === d.StudentID);
  if (!s) throw new Error('Student not found');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.Date || '') || !/^\d{2}:\d{2}$/.test(d.Time || '')) throw new Error('Date and time required / তারিখ ও সময় দিন');
  const days = String(d.Days || '').split(',').filter(x => x !== '').map(Number);
  const weeks = Math.min(Math.max(Number(d.Weeks) || 4, 1), 26);
  const dates = [];
  if (!days.length) dates.push(d.Date);
  else {
    const start = new Date(d.Date + 'T00:00:00Z').getTime();
    for (let i = 0; i < weeks * 7; i++) { const x = new Date(start + i * 86400000); if (days.indexOf(x.getUTCDay()) > -1) dates.push(x.toISOString().slice(0, 10)); }
  }
  const have = {}; all_('Classes').forEach(c => have[c.StudentID + c.Date + c.Time] = 1);
  const fresh = dates.filter(x => !have[s.ID + x + d.Time]).slice(0, 200);
  if (!fresh.length) return {added: 0, skipped: dates.length};
  const ids = nextIds_('CL-', 'class', 5, fresh.length);
  addMany_('Classes', fresh.map((x, i) => ({ClassID: ids[i], Date: x, Time: d.Time, StudentID: s.ID, TeacherID: s.TeacherID, StudentStatus:'', StudentReason:'', TeacherStatus:'', TeacherReason:'', MarkedAt:'', MarkedBy:''})));
  return {added: fresh.length, skipped: dates.length - fresh.length};
}
function updClass_(me, d) {
  const p = d.patch || {}, out = attPatch_(p);
  if (p.Date !== undefined) { if (!/^\d{4}-\d{2}-\d{2}$/.test(p.Date)) throw new Error('Invalid date'); out.Date = p.Date; }
  if (p.Time !== undefined) { if (!/^\d{2}:\d{2}$/.test(p.Time)) throw new Error('Invalid time'); out.Time = p.Time; }
  if (p.Minutes !== undefined) out.Minutes = Math.max(0, Number(p.Minutes) || 0);
  if (out.StudentStatus || out.TeacherStatus) { const n = nowDhaka_(); out.MarkedAt = n.date + ' ' + n.time; out.MarkedBy = me.username; }
  update_('Classes', d.ClassID, out);
  return true;
}

function attToday_() {
  const today = nowDhaka_().date, cls = all_('Classes').filter(c => c.Date === today);
  const sn = {}, tn = {}; all_('Students').forEach(s => sn[s.ID] = s.Name); all_('Teachers').forEach(t => tn[t.Code] = t.Name);
  const absences = [];
  cls.sort((a, b) => a.Time < b.Time ? -1 : 1).forEach(c => {
    if (c.TeacherStatus === 'Absent') absences.push({Time: c.Time, who: 'T', Name: tn[c.TeacherID] || c.TeacherID, Reason: c.TeacherReason});
    if (c.StudentStatus === 'Absent') absences.push({Time: c.Time, who: 'S', Name: sn[c.StudentID] || c.StudentID, Reason: c.StudentReason});
  });
  return {today: cls.length, sPresent: cls.filter(c => c.StudentStatus === 'Present').length, sAbsent: cls.filter(c => c.StudentStatus === 'Absent').length,
          tAbsent: cls.filter(c => c.TeacherStatus === 'Absent').length, pending: cls.filter(c => !c.TeacherStatus).length, absences: absences.slice(0, 50)};
}


/* ---------- new teacher self-registration (needs admin approval) ---------- */
function registerTeacher(d) {
  d = d || {};
  const cache = CacheService.getScriptCache(), k = 'reg_' + Utilities.formatDate(new Date(), 'Asia/Dhaka', 'yyyyMMddHH');
  const n = Number(cache.get(k) || 0);
  if (n >= 30) return {ok: false, msg: 'Too many requests, try again later / পরে চেষ্টা করুন'};
  cache.put(k, String(n + 1), 3600);
  const name = String(d.Name || '').trim().slice(0, 80), email = String(d.Email || '').trim().toLowerCase(),
        phone = String(d.Phone || '').trim().slice(0, 20), pw = String(d.pw || '');
  if (!name) return {ok: false, msg: 'Name required / নাম দিন'};
  if (!EMAIL_RE.test(email)) return {ok: false, msg: 'Valid email required / সঠিক ইমেইল দিন'};
  if (normPhone_(phone).length < 10) return {ok: false, msg: 'Valid phone number required / সঠিক ফোন নম্বর দিন'};
  if (pw.length < 6) return {ok: false, msg: 'Password must be at least 6 characters / পাসওয়ার্ড কমপক্ষে ৬ অক্ষর'};
  ensureAdmin_();
  const users = all_('Users'), teachers = all_('Teachers');
  if (users.some(u => u.Username.toLowerCase() === email || String(u.Email).toLowerCase() === email) || teachers.some(t => String(t.Email).toLowerCase() === email))
    return {ok: false, msg: 'This email is already registered / এই ইমেইল আগে নিবন্ধিত'};
  const np = normPhone_(phone);
  if (teachers.some(t => normPhone_(t.Phone) === np) || users.some(u => normPhone_(u.Phone) === np))
    return {ok: false, msg: 'This phone number is already registered / এই ফোন নম্বর আগে নিবন্ধিত'};
  const code = nextId_('T-', 'teacher', 3);
  const rec = {Code: code, Name: name, Phone: phone, Email: email, ClassesPerWeek: 0, SharePercent: 50, Status: 'Pending',
               Version: String(d.Version || '').trim().slice(0, 80), Country: String(d.Country || '').trim().slice(0, 80), FreeTime: String(d.FreeTime || '').trim().slice(0, 120)};
  const custom = regFields_();
  for (let i = 0; i < custom.length; i++) {            // admin-defined extra teacher fields
    const f = custom[i], v = String(d[f.Label] || '').trim().slice(0, 300);
    if (f.Required === 'Yes' && !v) return {ok: false, msg: f.Label + ' required / আবশ্যক'};
    rec[f.Label] = v;
  }
  fillAuto_('Teachers', rec);
  add_('Teachers', rec);
  add_('Users', {Username: email, FullName: name, Email: email, Phone: phone, Role: 'Teacher', ReadOnly: 'No', Pages: 'teacher', TeacherCode: code, PassHash: hash_(pw)});
  try {
    const esc = x => String(x).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
    users.filter(u => u.Role === 'Admin' && u.Email).forEach(u => MailApp.sendEmail({to: u.Email,
      subject: 'New teacher registration / নতুন শিক্ষকের রেজিস্ট্রেশন',
      htmlBody: '<p>' + esc(name) + ' (' + esc(email) + ', ' + esc(phone) + ') রেজিস্ট্রেশন করেছেন। শিক্ষক পেজ থেকে অনুমোদন করুন। / registered as a teacher; please approve in the Teachers page.</p>'}));
  } catch (e) {}
  return {ok: true};
}

/* ---------- teacher registration form + teacher's own profile ---------- */
function regFields_() {   // extra teacher fields shown on the public form (no file upload: not logged in yet)
  return all_('Fields').filter(f => f.Entity === 'Teachers' && f.Type !== 'file' && f.Type !== 'autoid').sort((a, b) => a.Order - b.Order);
}
function getRegForm() { return {fields: regFields_()}; }
const OWN_KEYS_ = ['Name','Phone','Version','Country','FreeTime'];
function ownKeys_() { return OWN_KEYS_.concat(all_('Fields').filter(f => f.Entity === 'Teachers' && f.Type !== 'autoid').map(f => f.Label)); }
function myProfile_(me) {
  const t = all_('Teachers').find(x => x.Code === me.teacherCode);
  if (!t) throw new Error('Profile not found');
  const o = {Email: t.Email}; ownKeys_().forEach(k => o[k] = t[k] === undefined ? '' : t[k]);
  return o;
}
function saveMyProfile_(me, d) {
  const patch = {};
  ownKeys_().forEach(k => { if (d[k] !== undefined) patch[k] = String(d[k]).trim().slice(0, 300); });
  if (patch.Name !== undefined && !patch.Name) throw new Error('Name required / নাম দিন');
  if (patch.Phone !== undefined) {
    const np = normPhone_(patch.Phone);
    if (np.length < 10) throw new Error('Valid phone number required / সঠিক ফোন নম্বর দিন');
    if (all_('Teachers').some(t => t.Code !== me.teacherCode && normPhone_(t.Phone) === np) ||
        all_('Users').some(u => u.TeacherCode !== me.teacherCode && normPhone_(u.Phone) === np)) throw new Error('This phone number is already registered / এই ফোন নম্বর আগে নিবন্ধিত');
  }
  update_('Teachers', me.teacherCode, patch);
  const up = {}; if (patch.Name) up.FullName = patch.Name; if (patch.Phone !== undefined) up.Phone = patch.Phone;
  if (Object.keys(up).length) update_('Users', me.username, up);
  return true;
}


/* ---------- website content (banners, board, reviews) + public data for the website ---------- */
const CMS = {Banners: ['BN-', 3], Board: ['BM-', 3], Reviews: ['RV-', 4]};
const byOrder_ = (a, b) => a.Order - b.Order;
function cmsMeta_(sheet) { if (!CMS[sheet]) throw new Error('Unknown list'); return CMS[sheet]; }
function pubDrop_() { try { CacheService.getScriptCache().remove('pub'); } catch (e) {} }
function cmsAdd_(d) {
  const m = cmsMeta_(d.sheet), r = d.row || {};
  r[SHEETS[d.sheet][0]] = nextId_(m[0], 'cms_' + d.sheet, m[1]);
  if (d.sheet === 'Reviews') { r.Date = today_(); r.Status = r.Status || 'Approved'; }
  add_(d.sheet, r); pubDrop_(); return true;
}
function getPublic() {   // no login: only what the public website shows; cached 5 minutes
  const cache = CacheService.getScriptCache(), hit = cache.get('pub');
  if (hit) return JSON.parse(hit);
  const st = all_('Students').filter(s => s.Status === 'Active'), countries = {};
  st.forEach(s => { const k = String(s.Country || '').split('/')[0].trim().toLowerCase(); if (k) countries[k] = 1; });
  const rev = all_('Reviews').filter(r => r.Status === 'Approved'), rated = rev.filter(r => r.Rating > 0);
  const out = {
    stats: {students: st.length, countries: Object.keys(countries).length, reviews: rev.length,
            rating: rated.length ? Math.round(rated.reduce((a, r) => a + r.Rating, 0) / rated.length * 10) / 10 : 0},
    banners: all_('Banners').filter(b => b.Show !== 'No' && b.Image).sort(byOrder_).map(b => ({Image: b.Image, Title: b.Title, Link: b.Link})),
    board: all_('Board').filter(b => b.Show !== 'No').sort(byOrder_).map(b => ({Name: b.Name, Role: b.Role, Photo: b.Photo})),
    reviews: rev.slice(-12).reverse().map(r => ({Name: r.Name, Country: r.Country, Rating: r.Rating, Text: r.Text}))
  };
  try { cache.put('pub', JSON.stringify(out), 300); } catch (e) {}
  return out;
}
function submitReview(d) {   // parents write a review on the website; it waits for admin approval
  d = d || {};
  const cache = CacheService.getScriptCache(), n = Number(cache.get('rv_n') || 0);
  if (n >= 20) return {ok: false};
  const name = String(d.Name || '').trim().slice(0, 60), text = String(d.Text || '').trim().slice(0, 600), rating = Math.round(Number(d.Rating) || 0);
  if (name.length < 2 || text.length < 10 || rating < 1 || rating > 5) return {ok: false};
  cache.put('rv_n', String(n + 1), 3600);
  add_('Reviews', {ReviewID: nextId_('RV-', 'cms_Reviews', 4), Name: name, Country: String(d.Country || '').trim().slice(0, 60), Rating: rating, Text: text, Status: 'Pending', Date: today_()});
  return {ok: true};
}

/* ---------- class timing: teacher starts / ends a class, minutes are calculated ---------- */
function stamp_() { return Utilities.formatDate(new Date(), 'Asia/Dhaka', "yyyy-MM-dd'T'HH:mm:ss'+06:00'"); }
function clsOwn_(me, d) {
  const c = all_('Classes').find(x => x.ClassID === d.ClassID);
  if (!c || c.TeacherID !== me.teacherCode) throw new Error('Class not found');
  const now = nowDhaka_();
  if (c.Date < now.date) throw new Error('এই দিনের এন্ট্রির সময় শেষ / Entry closed');
  if (c.Date > now.date || now.time < c.Time) throw new Error('এখনো ক্লাসের সময় হয়নি / Class has not started yet');
  return c;
}
function clsStart_(me, d) {
  const c = clsOwn_(me, d);
  if (c.EndedAt) throw new Error('ক্লাস আগেই শেষ হয়েছে / Class already ended');
  if (c.StartedAt) return {StartedAt: c.StartedAt};
  const s = stamp_(); update_('Classes', c.ClassID, {StartedAt: s}); return {StartedAt: s};
}
function clsEnd_(me, d) {
  const c = clsOwn_(me, d);
  if (!c.StartedAt) throw new Error('আগে ক্লাস শুরু করুন / Start the class first');
  if (c.EndedAt) return {EndedAt: c.EndedAt, Minutes: c.Minutes};
  const e = stamp_(), min = Math.max(1, Math.round((new Date(e) - new Date(c.StartedAt)) / 60000));
  update_('Classes', c.ClassID, {EndedAt: e, Minutes: min}); return {EndedAt: e, Minutes: min};
}
function classHours_() {
  const tn = {}; all_('Teachers').forEach(t => tn[t.Code] = t.Name);
  const m = {};
  all_('Classes').forEach(c => {
    const k = c.Date.slice(0, 7) + '|' + c.TeacherID;
    const o = m[k] = m[k] || {Month: c.Date.slice(0, 7), Teacher: tn[c.TeacherID] || c.TeacherID, Scheduled: 0, Held: 0, Minutes: 0, TeacherAbsent: 0};
    o.Scheduled++; if (c.Minutes > 0) { o.Held++; o.Minutes += c.Minutes; } if (c.TeacherStatus === 'Absent') o.TeacherAbsent++;
  });
  return Object.keys(m).map(k => m[k]).sort((a, b) => a.Month < b.Month ? 1 : a.Month > b.Month ? -1 : 0);
}
