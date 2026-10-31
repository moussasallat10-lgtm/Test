// عيادة سام – حفظ المرضى في Google Sheet مع مستخدمين وأدمن
// المستخدمون وكلمات سرهم (مشفّرة) تُحفظ داخل السكربت، وتُدار من الإعدادات في الصفحة.
// أول مرة يُنشأ أدمن افتراضي: الاسم «سام» وكلمة السر sam2026

function props_() { return PropertiesService.getScriptProperties(); }
function hash_(salt, pw) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function mk_(name, pw, role) {
  var salt = Utilities.getUuid();
  return { name: name, salt: salt, hash: hash_(salt, pw), role: role };
}
function users_() {
  var s = props_().getProperty('users');
  if (s) return JSON.parse(s);
  var u = [mk_('سام', 'sam2026', 'admin')];
  props_().setProperty('users', JSON.stringify(u));
  return u;
}
function auth_(n, pw) {
  var l = users_();
  for (var i = 0; i < l.length; i++) {
    if (l[i].name === n && l[i].hash === hash_(l[i].salt, String(pw || ''))) return { name: l[i].name, role: l[i].role };
  }
  return null;
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('Patients');
  if (!s) {
    s = ss.insertSheet('Patients');
    s.appendRow(['id','first','last','father','mother','age','phone','total','paid','remaining','json']);
    s.getRange(2, 7, s.getMaxRows(), 1).setNumberFormat('@');
  }
  if (s.getLastColumn() < 13) s.getRange(1, 12, 1, 2).setValues([['updated_by', 'updated_at']]);
  return s;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// حماية من التخمين: 10 محاولات خاطئة تقفل الدخول 10 دقائق
function blocked_() { return Number(CacheService.getScriptCache().get('fails') || 0) >= 10; }
function fail_() {
  var c = CacheService.getScriptCache();
  c.put('fails', String(Number(c.get('fails') || 0) + 1), 600);
  Utilities.sleep(1500);
  return out_({ error: 'auth' });
}

function doGet(e) {
  var q = e.parameter;
  if (q.action === 'users') return out_({ users: users_().map(function (u) { return u.name; }) });
  if (blocked_()) return out_({ error: 'locked' });
  var u = auth_(q.name, q.key);
  if (!u) return fail_();
  CacheService.getScriptCache().remove('fails');
  if (q.action === 'list') {
    if (u.role !== 'admin') return out_({ error: 'forbidden' });
    return out_({ users: users_().map(function (x) { return { name: x.name, role: x.role }; }) });
  }
  var v = sheet_().getDataRange().getValues(), list = [];
  for (var i = 1; i < v.length; i++) {
    try { list.push(JSON.parse(v[i][10])); } catch (x) {}
  }
  return out_({ patients: list, user: u });
}

function doPost(e) {
  var b = JSON.parse(e.postData.contents);
  if (blocked_()) return out_({ error: 'locked' });
  var u = auth_(b.name, b.key);
  if (!u) return fail_();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (b.action === 'addUser' || b.action === 'delUser') {
      if (u.role !== 'admin') return out_({ error: 'forbidden' });
      var l = users_();
      if (b.action === 'addUser') {
        var nm = String((b.user && b.user.name) || '').trim(), pw = String((b.user && b.user.password) || '');
        if (!nm || pw.length < 4) return out_({ error: 'invalid' });
        if (l.some(function (x) { return x.name === nm; })) return out_({ error: 'exists' });
        l.push(mk_(nm, pw, b.user.role === 'staff' ? 'staff' : 'admin'));
      } else {
        var t = l.filter(function (x) { return x.name === b.target; })[0];
        if (!t) return out_({ error: 'invalid' });
        if (t.role === 'admin' && l.filter(function (x) { return x.role === 'admin'; }).length < 2) return out_({ error: 'last_admin' });
        l = l.filter(function (x) { return x.name !== b.target; });
      }
      props_().setProperty('users', JSON.stringify(l));
      return out_({ ok: true });
    }
    var s = sheet_(), p = b.patient, row = 0;
    var ids = s.getRange(1, 1, s.getLastRow(), 1).getValues();
    for (var i = 1; i < ids.length; i++) {
      if (String(ids[i][0]) === String(p.id)) { row = i + 1; break; }
    }
    // الموظف (staff) لا يستطيع حذف زيارة أو دفعة
    if (row && u.role !== 'admin') {
      try {
        var old = JSON.parse(s.getRange(row, 11).getValue());
        if ((p.visits || []).length < (old.visits || []).length || (p.payments || []).length < (old.payments || []).length) {
          return out_({ error: 'forbidden' });
        }
      } catch (x) {}
    }
    p.by = u.name;
    p.at = new Date().toISOString();
    var tt = 0, d = 0;
    (p.visits || []).forEach(function (v) { tt += v.total; d += v.paid; });
    (p.payments || []).forEach(function (x) { d += x.amount; });
    var line = [String(p.id), p.first, p.last, p.father || '', p.mother || '', p.age || '', p.phone || '', tt, d, Math.max(0, tt - d), JSON.stringify(p), u.name, p.at];
    if (row) s.getRange(row, 1, 1, line.length).setValues([line]);
    else s.appendRow(line);
    return out_({ ok: true });
  } finally {
    lock.release();
  }
}
