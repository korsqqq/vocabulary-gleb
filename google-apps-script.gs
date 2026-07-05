/****************************************************************
 *  Word Quiz — приём результатов в Google Таблицу
 *  ------------------------------------------------------------
 *  Куда вставлять: открой Google Таблицу → меню «Расширения» →
 *  «Apps Script», удали весь пример и вставь этот код целиком.
 *  Затем «Развернуть» → «Новое развёртывание» → тип «Веб-приложение»,
 *  «У кого есть доступ: Все» → скопируй ссылку и отдай её мне
 *  (или сам вставь в index.html в строку  const SHEET_URL = "…").
 ****************************************************************/

var VISITS = "Входы";       // вкладка с логинами
var DICTS  = "Диктанты";    // вкладка с диктантами

function getSheet_(name, headers){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if(!sh){ sh = ss.insertSheet(name); sh.appendRow(headers); }
  else if(sh.getLastRow() === 0){ sh.appendRow(headers); }
  return sh;
}

/* Приём данных из тренажёра (вход или диктант) */
function doPost(e){
  try{
    var d = JSON.parse(e.postData.contents);
    if(d.type === "visit"){
      getSheet_(VISITS, ["Имя","Время входа"])
        .appendRow([ d.name, new Date(d.ts || Date.now()) ]);
    } else if(d.type === "dictation"){
      var sh = getSheet_(DICTS, ["Имя","Дата","Время (сек)","Направление","Всего","Верно",
                        "Ошибок","Процент","Работа над ошибками","Темы","Слова с ошибками","Режим"]);
      if(!sh.getRange(1,12).getValue()) sh.getRange(1,12).setValue("Режим");   // дополнить старую шапку
      var actLabel = {quiz:"Диктант", cards:"Карточки", match:"Игра"}[d.act] || "Диктант";
      sh.appendRow([
          d.name,
          new Date(d.finished || Date.now()),
          Math.round((d.durationMs || 0) / 1000),
          (d.act && d.act !== "quiz") ? "—" : (d.mode === "ruen" ? "RU->EN" : "EN->RU"),
          d.total, d.correct, d.wrong, d.percent,
          d.errorRound ? "да" : "нет",
          (d.themes || []).join(" | "),
          (d.errors || []).map(function(x){
            return x.en + " (" + String(x.ru || "").split(/[,;\/]/)[0] + ")";
          }).join(" | "),
          actLabel
        ]);
    }
    return ContentService.createTextOutput("ok");
  }catch(err){
    return ContentService.createTextOutput("error: " + err);
  }
}

/* Выдача всего журнала обратно в панель «Koba» (через JSONP) */
function doGet(e){
  var students = {};
  function rec(name){
    var key = String(name || "").trim().toLowerCase();
    if(!key) return null;
    if(!students[key]) students[key] = {name:String(name).trim(), firstSeen:0, visits:[], dictations:[]};
    return students[key];
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var vs = ss.getSheetByName(VISITS);
  if(vs && vs.getLastRow() > 1){
    vs.getRange(2,1,vs.getLastRow()-1,2).getValues().forEach(function(r){
      var s = rec(r[0]); if(!s) return;
      var ts = (r[1] instanceof Date) ? r[1].getTime() : Date.parse(r[1]);
      if(ts){ s.visits.push(ts); if(!s.firstSeen || ts < s.firstSeen) s.firstSeen = ts; }
    });
  }

  var ds = ss.getSheetByName(DICTS);
  if(ds && ds.getLastRow() > 1){
    ds.getRange(2,1,ds.getLastRow()-1,12).getValues().forEach(function(r){
      var s = rec(r[0]); if(!s) return;
      var ts = (r[1] instanceof Date) ? r[1].getTime() : Date.parse(r[1]);
      var errStr = String(r[10] || "");
      var errors = errStr ? errStr.split(" | ").map(function(t){
        var m = t.match(/^(.*)\s+\((.*)\)$/);
        return m ? {en:m[1], ru:m[2]} : {en:t, ru:""};
      }) : [];
      var actMap = {"Карточки":"cards", "Игра":"match"};
      s.dictations.push({
        act: actMap[String(r[11] || "")] || "quiz",
        finished: ts || 0,
        durationMs: (Number(r[2]) || 0) * 1000,
        mode: (r[3] === "RU->EN") ? "ruen" : "enru",
        themes: String(r[9] || "").split(" | ").filter(String),
        total: Number(r[4]) || 0,
        correct: Number(r[5]) || 0,
        wrong: Number(r[6]) || 0,
        percent: parseInt(r[7], 10) || 0,
        errorRound: (r[8] === "да"),
        errors: errors
      });
      if(ts && (!s.firstSeen || ts < s.firstSeen)) s.firstSeen = ts;
    });
  }

  var out = JSON.stringify({students:students});
  var cb  = e && e.parameter && e.parameter.callback;
  if(cb){
    return ContentService.createTextOutput(cb + "(" + out + ")")
             .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(out)
           .setMimeType(ContentService.MimeType.JSON);
}
