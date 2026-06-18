/**
 * ГОЛОСУВАННЯ ЖБК «МЕТРОБУДІВНИК-3» — бекенд (Google Apps Script)
 * Зберігає голоси + систему адміністраторів (вхід, пароль, відновлення по email,
 * підтвердження/бан додаткових адмінів) у Google Таблиці.
 *
 * НАЛАШТУВАННЯ (один раз):
 *   1) Google Таблиця -> Розширення -> Apps Script -> встав цей код у Code.gs.
 *   2) Зміни SETUP_KEY нижче на свій секретний код (потрібен ЛИШЕ для створення
 *      найпершого адміна). Решта адмінів реєструються самі й чекають підтвердження.
 *   3) Deploy -> New deployment -> Web app: Execute as = Me, Who has access = Anyone.
 *   4) Скопіюй URL у index.html і admin.html (SCRIPT_URL).
 *   5) При зміні коду -> Deploy -> Manage deployments -> Edit -> New version.
 *
 * ПРИВАТНІСТЬ: публічно (GET) віддаються ЛИШЕ підсумки. Усі персональні дані
 * та керування — тільки через POST з токеном активного адміна.
 */

// ====== НАЛАШТУВАННЯ ======
var SETUP_KEY = 'demo';                       // секрет для створення ПЕРШОГО адміна
var SHEET_NAME = 'Голоси';
var ADMIN_SHEET = 'Адміни';
var QUESTION = 'Чи потрібне нам резервне опалення?';
var BUILDING = { paradni: 6, poverhy: 10, perFloor: 4 };
// ===========================

var MAX_APT = BUILDING.paradni * BUILDING.poverhy * BUILDING.perFloor; // 240
var TOKEN_DAYS = 7;
var REC_MIN = 15; // термін дії коду відновлення, хвилин

var HEADERS = [
  'Час', 'Квартира №', 'Розташування', 'ПІБ власника', 'Телефон', 'Відповідь',
  'IP', 'Місто (IP)', 'Регіон (IP)', 'Країна (IP)',
  'GPS широта', 'GPS довгота', 'GPS точність (м)', 'Пристрій', 'User-Agent'
];
var A_HEAD = ['email','hash','salt','status','created','token','token_exp','rec_code','rec_exp'];
var FIO_RE = /^[А-Яа-яЁёЇїІіЄєҐґ'’\-]{2,}(\s+[А-Яа-яЁёЇїІіЄєҐґ'’\-]{2,})+$/;

// ---------- утиліти ----------
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet_(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(SHEET_NAME);
  if(!sh){ sh=ss.insertSheet(SHEET_NAME); sh.appendRow(HEADERS); sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function aSheet_(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(ADMIN_SHEET);
  if(!sh){ sh=ss.insertSheet(ADMIN_SHEET); sh.appendRow(A_HEAD); sh.getRange(1,1,1,A_HEAD.length).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function tally_(sh){
  var r={yes:0,no:0,abstain:0,total:0}, last=sh.getLastRow();
  if(last<2) return r;
  var v=sh.getRange(2,6,last-1,1).getValues();
  for(var i=0;i<v.length;i++){var a=String(v[i][0]);if(a==='yes')r.yes++;else if(a==='no')r.no++;else if(a==='abstain')r.abstain++;}
  r.total=r.yes+r.no+r.abstain; return r;
}
function hex_(bytes){ var s=''; for(var i=0;i<bytes.length;i++) s+=('0'+(bytes[i]&255).toString(16)).slice(-2); return s; }
function hash_(pw,salt){ return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt+':'+pw, Utilities.Charset.UTF_8)); }
function rnd_(n,alpha){ var c=alpha||'ABCDEFGHJKLMNPQRSTUVWXYZ23456789abcdefghijkmnpqrstuvwxyz', s=''; for(var i=0;i<n;i++) s+=c.charAt(Math.floor(Math.random()*c.length)); return s; }
function now_(){ return new Date().getTime(); }
function normEmail_(e){ return String(e||'').trim().toLowerCase(); }

// ---------- доступ до адмінів ----------
function adminRow_(email){ // -> {idx, obj} або null
  var sh=aSheet_(), last=sh.getLastRow(); if(last<2) return null;
  var data=sh.getRange(2,1,last-1,A_HEAD.length).getValues();
  email=normEmail_(email);
  for(var i=0;i<data.length;i++){
    if(normEmail_(data[i][0])===email){
      var o={}; for(var k=0;k<A_HEAD.length;k++) o[A_HEAD[k]]=data[i][k];
      return {idx:i+2, obj:o};
    }
  }
  return null;
}
function adminCount_(){ var sh=aSheet_(); return Math.max(0, sh.getLastRow()-1); }
function setCell_(idx,col,val){ aSheet_().getRange(idx, A_HEAD.indexOf(col)+1, 1, 1).setValue(val); }
function issueToken_(idx){ var t=rnd_(28); setCell_(idx,'token',t); setCell_(idx,'token_exp', now_()+TOKEN_DAYS*864e5); return t; }
function authToken_(token){ // -> {idx,obj} активного адміна або null
  if(!token) return null;
  var sh=aSheet_(), last=sh.getLastRow(); if(last<2) return null;
  var data=sh.getRange(2,1,last-1,A_HEAD.length).getValues();
  for(var i=0;i<data.length;i++){
    var o={}; for(var k=0;k<A_HEAD.length;k++) o[A_HEAD[k]]=data[i][k];
    if(o.token && String(o.token)===String(token)){
      if(o.status!=='active') return null;
      if(Number(o.token_exp) < now_()) return null;
      return {idx:i+2, obj:o};
    }
  }
  return null;
}

// ====== GET: лише публічні підсумки ======
function doGet(e){
  return json_({ ok:true, question:QUESTION, building:BUILDING, results:tally_(sheet_()) });
}

// ====== POST: голос або адмін-дії ======
function doPost(e){
  var d; try{ d=JSON.parse(e.postData.contents); }catch(err){ return json_({ok:false,error:'Невірний запит'}); }
  if(d && d.action) return handleAdmin_(d);
  return handleVote_(d);
}

// ---------- голосування ----------
function handleVote_(d){
  var lock=LockService.getScriptLock();
  try{ lock.waitLock(15000); }catch(err){ return json_({ok:false,error:'Сервер зайнятий, спробуйте ще раз'}); }
  try{
    var apt=parseInt(d.apt,10), fio=String(d.fio||'').trim().replace(/\s+/g,' '),
        phone=String(d.phone||'').trim(), choice=String(d.choice||'').trim();
    if(!(apt>=1&&apt<=MAX_APT)) return json_({ok:false,error:'Невірний номер квартири'});
    if(!FIO_RE.test(fio)) return json_({ok:false,error:'ПІБ — лише українські/російські літери, мінімум прізвище та ім\'я'});
    var digits=phone.replace(/\D/g,'');
    if(!/^(?:38)?0\d{9}$/.test(digits)) return json_({ok:false,error:'Невірний номер телефону'});
    if(['yes','no','abstain'].indexOf(choice)<0) return json_({ok:false,error:'Оберіть варіант відповіді'});

    var sh=sheet_(), last=sh.getLastRow();
    if(last>=2){
      var aptCol=sh.getRange(2,2,last-1,1).getValues();
      for(var i=0;i<aptCol.length;i++) if(parseInt(aptCol[i][0],10)===apt)
        return json_({ok:false,error:'Квартира № '+apt+' вже проголосувала. Змінити голос не можна.',results:tally_(sh)});
    }
    sh.appendRow([new Date(),apt,String(d.place||'кв. '+apt),fio,"'"+digits,choice,
      d.ip||'',d.ipCity||'',d.ipRegion||'',d.ipCountry||'',d.lat||'',d.lon||'',d.acc||'',d.device||'',d.ua||'']);
    return json_({ok:true,results:tally_(sh)});
  }catch(err){ return json_({ok:false,error:'Помилка: '+err}); }
  finally{ lock.releaseLock(); }
}

// ---------- адмін-система ----------
function handleAdmin_(d){
  var a=d.action;

  // публічно: чи існує вже хоч один адмін (для першого налаштування)
  if(a==='admin_status') return json_({ok:true, exists: adminCount_()>0});

  var lock=LockService.getScriptLock();
  try{ lock.waitLock(15000); }catch(err){ return json_({ok:false,error:'Сервер зайнятий'}); }
  try{
    var email=normEmail_(d.email), pw=String(d.password||'');

    if(a==='admin_setup'){
      if(String(d.setupKey||'')!==SETUP_KEY) return json_({ok:false,error:'Невірний код налаштування'});
      if(adminCount_()>0) return json_({ok:false,error:'Адміністратор вже створений. Скористайтесь входом.'});
      if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return json_({ok:false,error:'Невірний email'});
      if(pw.length<6) return json_({ok:false,error:'Пароль мінімум 6 символів'});
      var salt=rnd_(12);
      aSheet_().appendRow([email,hash_(pw,salt),salt,'active',new Date(),'',0,'',0]);
      var r=adminRow_(email); var t=issueToken_(r.idx);
      return json_({ok:true, token:t, me:{email:email,status:'active'}});
    }

    if(a==='admin_register'){
      if(adminCount_()===0) return json_({ok:false,error:'Спочатку має бути створений перший адміністратор'});
      if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return json_({ok:false,error:'Невірний email'});
      if(pw.length<6) return json_({ok:false,error:'Пароль мінімум 6 символів'});
      if(adminRow_(email)) return json_({ok:false,error:'Такий email вже зареєстрований'});
      var s2=rnd_(12);
      aSheet_().appendRow([email,hash_(pw,s2),s2,'pending',new Date(),'',0,'',0]);
      notifyActive_('Новий адміністратор очікує підтвердження: '+email);
      return json_({ok:true, pending:true});
    }

    if(a==='admin_login'){
      var r=adminRow_(email);
      if(!r) return json_({ok:false,error:'Невірний email або пароль'});
      if(r.obj.status==='banned') return json_({ok:false,error:'Доступ заборонено'});
      if(r.obj.status==='pending') return json_({ok:false,error:'Акаунт очікує підтвердження першим адміністратором'});
      if(hash_(pw,r.obj.salt)!==String(r.obj.hash)) return json_({ok:false,error:'Невірний email або пароль'});
      var t2=issueToken_(r.idx);
      return json_({ok:true, token:t2, me:{email:email,status:'active'}});
    }

    if(a==='admin_recover_request'){
      var r3=adminRow_(email);
      if(!r3 || r3.obj.status==='banned') return json_({ok:true}); // не розкриваємо існування
      var code=rnd_(6,'0123456789');
      setCell_(r3.idx,'rec_code',code); setCell_(r3.idx,'rec_exp', now_()+REC_MIN*60000);
      try{ MailApp.sendEmail(email,'Код відновлення доступу — голосування ЖБК',
        'Ваш код відновлення: '+code+'\nДійсний '+REC_MIN+' хв.\nЯкщо ви не запитували — проігноруйте цей лист.'); }catch(e2){}
      return json_({ok:true});
    }

    if(a==='admin_recover_confirm'){
      var r4=adminRow_(email);
      if(!r4) return json_({ok:false,error:'Невірний код'});
      if(!r4.obj.rec_code || String(r4.obj.rec_code)!==String(d.code||'')) return json_({ok:false,error:'Невірний код'});
      if(Number(r4.obj.rec_exp) < now_()) return json_({ok:false,error:'Код прострочено'});
      if(String(d.password||'').length<6) return json_({ok:false,error:'Пароль мінімум 6 символів'});
      var salt4=rnd_(12);
      setCell_(r4.idx,'salt',salt4); setCell_(r4.idx,'hash',hash_(d.password,salt4));
      setCell_(r4.idx,'rec_code',''); setCell_(r4.idx,'rec_exp',0);
      if(r4.obj.status==='pending') {} // лишається pending до підтвердження
      var t4 = r4.obj.status==='active' ? issueToken_(r4.idx) : '';
      return json_({ok:true, token:t4, status:r4.obj.status});
    }

    // --- далі тільки для активного адміна (потрібен токен) ---
    var me=authToken_(d.token);
    if(!me) return json_({ok:false,error:'Потрібен вхід',needLogin:true});

    if(a==='admin_data'){
      var sh=sheet_(), last=sh.getLastRow(), rows=[];
      if(last>=2){
        var data=sh.getRange(2,1,last-1,HEADERS.length).getValues();
        for(var i=0;i<data.length;i++){var x=data[i];rows.push({time:x[0],apt:x[1],place:x[2],fio:x[3],phone:x[4],choice:x[5],
          ip:x[6],ipCity:x[7],ipRegion:x[8],ipCountry:x[9],lat:x[10],lon:x[11],acc:x[12],device:x[13],ua:x[14]});}
      }
      return json_({ok:true, question:QUESTION, results:tally_(sh), rows:rows, me:{email:me.obj.email}});
    }

    if(a==='admin_list'){
      var ash=aSheet_(), al=ash.getLastRow(), list=[];
      if(al>=2){ var ad=ash.getRange(2,1,al-1,A_HEAD.length).getValues();
        for(var j=0;j<ad.length;j++) list.push({email:ad[j][0],status:ad[j][3],created:ad[j][4],me:normEmail_(ad[j][0])===normEmail_(me.obj.email)}); }
      return json_({ok:true, admins:list});
    }

    if(a==='admin_approve'||a==='admin_ban'||a==='admin_unban'){
      var target=adminRow_(d.target);
      if(!target) return json_({ok:false,error:'Адміна не знайдено'});
      if(normEmail_(target.obj.email)===normEmail_(me.obj.email)) return json_({ok:false,error:'Не можна змінювати власний статус'});
      var ns = a==='admin_ban' ? 'banned' : 'active';
      setCell_(target.idx,'status',ns);
      if(ns==='banned'){ setCell_(target.idx,'token',''); } // вибиваємо сесію
      return json_({ok:true});
    }

    return json_({ok:false,error:'Невідома дія'});
  }catch(err){ return json_({ok:false,error:'Помилка: '+err}); }
  finally{ lock.releaseLock(); }
}

function notifyActive_(msg){
  try{
    var sh=aSheet_(), last=sh.getLastRow(); if(last<2) return;
    var d=sh.getRange(2,1,last-1,A_HEAD.length).getValues();
    for(var i=0;i<d.length;i++) if(d[i][3]==='active')
      MailApp.sendEmail(d[i][0],'Голосування ЖБК — сповіщення', msg);
  }catch(e){}
}
