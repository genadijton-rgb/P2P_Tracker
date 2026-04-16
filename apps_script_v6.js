// P2P Tracker v6 — полный скрипт
// НОВОЕ в v6:
//  + calcSalary()       — Расчёт зарплаты (% от оборота по каждому исполнителю)
//  + checkAnomalies()   — Алерты: 3+ убыточных подряд у одного сотрудника + спред ниже порога
//  + checkCompanyLimits() — Проверка дневных лимитов по компаниям
//  + saveRateHistory()  — Автозапись курса в RateHistory при каждой сделке
//  + sendWeeklyReport() — Авто-отчёт на email перед закрытием недели
//  + buildAnalytics()   — БЛОК 9: сравнение с прошлой неделей из Архива

// ══════════════════════════════════════════════════════
// 1. ОСНОВНОЙ ТРИГГЕР — автодата + защита + курс + лимиты + алерты
// ══════════════════════════════════════════════════════
function myonEdit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  if (sheetName !== "Journal") return;

  var row = e.range.getRow();
  var col = e.range.getColumn();

  if (row < 3) return;

  // ── Колонка B (2): автодата в колонку C (3) ──
  if (col === 2) {
    var value = e.range.getValue();
    if (value !== "" && value !== null) {
      var dateCell = sheet.getRange(row, 3);
      if (dateCell.getValue() === "" || dateCell.getValue() === null) {
        var now = new Date();
        var formatted = Utilities.formatDate(now, "Asia/Novosibirsk", "dd.MM.yyyy HH:mm");
        dateCell.setValue(formatted);
      }
    }
  }

  // ── Колонка H (8): защита исполнителя ──
  if (col === 8) {
    var props = PropertiesService.getScriptProperties();
    var key = "exec_" + row;
    var saved = props.getProperty(key);
    var current = sheet.getRange(row, 8).getValue();

    if (saved !== null && saved !== "") {
      if (String(current) !== String(saved)) {
        sheet.getRange(row, 8).setValue(saved);
      }
    } else {
      if (current !== "" && current !== null) {
        props.setProperty(key, String(current));
      }
    }
  }

  // ── Колонка F (6): автокурс Rapira для Kaya и WinWin ──
  if (col === 6) {
    var compVal = e.range.getValue();
    var compName = String(compVal).trim().toLowerCase();
    if (compName === "kaya" || compName === "winwin") {
      var rateCell = sheet.getRange(row, 14);
      var existingRate = rateCell.getValue();
      if (existingRate === "" || existingRate === null || existingRate === 0) {
        try {
          var response = UrlFetchApp.fetch("https://api.rapira.net/market/symbol-thumb", {muteHttpExceptions: true});
          var data = JSON.parse(response.getContentText());
          var usdtRub = null;
          for (var k = 0; k < data.length; k++) {
            if (data[k].symbol === "USDT/RUB") { usdtRub = data[k].close; break; }
          }
          if (usdtRub === null && data.data) {
            for (var k2 = 0; k2 < data.data.length; k2++) {
              if (data.data[k2].symbol === "USDT/RUB") { usdtRub = data.data[k2].close; break; }
            }
          }
          if (usdtRub && usdtRub > 0) {
            rateCell.setValue(usdtRub);
            saveRateHistory(usdtRub, String(compVal).trim());
            var dV = sheet.getRange(row, 4).getValue();
            var jV = sheet.getRange(row, 10).getValue();
            if (dV && dV !== 0 && (jV === "" || jV === null || jV === 0)) {
              sheet.getRange(row, 10).setValue(Math.round((dV / usdtRub) * 100) / 100);
            }
          }
        } catch(err) {}
      }
      // Проверяем лимиты по компании
      checkCompanyLimits(String(compVal).trim());
    }
  }

  // ── Авторасчёт J(10) = D(4) / N(14) для Kaya и WinWin ──
  if (col === 4 || col === 14) {
    var fComp = sheet.getRange(row, 6).getValue();
    var compLow = String(fComp).trim().toLowerCase();
    if (compLow === "kaya" || compLow === "winwin") {
      var dVal = sheet.getRange(row, 4).getValue();
      var nVal = sheet.getRange(row, 14).getValue();
      if (dVal && dVal !== 0 && nVal && nVal !== 0) {
        var jEmpty = sheet.getRange(row, 10).getValue();
        if (jEmpty === "" || jEmpty === null || jEmpty === 0) {
          sheet.getRange(row, 10).setValue(Math.round((dVal / nVal) * 100) / 100);
        }
      }
    }
  }

  // ── При заполнении J (10): проверяем аномалии у исполнителя ──
  if (col === 10) {
    var jVal = e.range.getValue();
    if (jVal !== "" && jVal !== null && jVal !== 0) {
      var executor = sheet.getRange(row, 8).getValue();
      if (executor) {
        checkAnomalies(executor);
      }
    }
  }

  // ── Лог изменений ──
  if (col !== 3) {
    writeLog(sheetName, e.range.getA1Notation(), e.oldValue, e.range.getValue());
  }
}

// ══════════════════════════════════════════════════════
// 2. ЛОГ ИЗМЕНЕНИЙ
// ══════════════════════════════════════════════════════
function writeLog(sheetName, cell, oldVal, newVal) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var log = ss.getSheetByName("ChangeLog");
    if (!log) return;
    var user = Session.getActiveUser().getEmail() || "неизвестно";
    var now = new Date();
    var formatted = Utilities.formatDate(now, "Asia/Novosibirsk", "dd.MM.yyyy HH:mm:ss");
    var lastRow = log.getLastRow();
    var nextRow = Math.max(lastRow + 1, 6);
    log.getRange(nextRow, 2).setValue(formatted);
    log.getRange(nextRow, 3).setValue(sheetName);
    log.getRange(nextRow, 4).setValue(cell);
    log.getRange(nextRow, 5).setValue(oldVal !== undefined ? oldVal : "");
    log.getRange(nextRow, 6).setValue(newVal !== undefined ? newVal : "");
    log.getRange(nextRow, 7).setValue(user);
  } catch(err) {}
}

// ══════════════════════════════════════════════════════
// 3. ЗАКРЫТЬ НЕДЕЛЮ (с авто-отчётом)
// ══════════════════════════════════════════════════════
function closeWeek() {
  var OWNER = "genadijton@gmail.com";
  if (Session.getActiveUser().getEmail() !== OWNER) {
    SpreadsheetApp.getUi().alert("Доступ запрещён. Обратитесь к администратору.");
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var result = ui.alert(
    "Закрыть неделю?",
    "Будет выполнено:\n" +
    "1. Авто-отчёт недели отправится на почту\n" +
    "2. Все сделки скопируются на отдельный лист (скрытый)\n" +
    "3. Итоги недели скопируются в Архив\n" +
    "4. Журнал очистится (строки 3–2002)\n" +
    "5. Выплачено ₽ обнулится\n" +
    "6. Исполнители разблокируются\n\n" +
    "Продолжить?",
    ui.ButtonSet.YES_NO
  );
  if (result !== ui.Button.YES) return;

  // Сначала отправляем отчёт
  sendWeeklyReport();

  var journal = ss.getSheetByName("Journal");
  var settings = ss.getSheetByName("Settings");
  var balances = ss.getSheetByName("Balances");
  var archive = ss.getSheetByName("Archive");
  var dashboard = ss.getSheetByName("Dashboard");

  if (!journal || !settings || !balances || !archive) {
    ui.alert("Ошибка: не найден один из листов.");
    return;
  }

  var week = settings.getRange("D5").getValue();
  var totalDeals = dashboard.getRange("B5").getValue();
  var totalTurnover = balances.getRange("E21").getValue();
  var totalProfitUSDT = balances.getRange("F21").getValue();
  var totalProfitRUB = balances.getRange("G21").getValue();
  var totalPaid = balances.getRange("J21").getValue();
  var compShare = balances.getRange("G21").getValue() - balances.getRange("H21").getValue();
  var avgProfit = totalDeals > 0 ? totalProfitUSDT / totalDeals : 0;

  var bestEmp = "", maxProfit = 0;
  for (var i = 0; i < 15; i++) {
    var empProfit = balances.getRange(6 + i, 6).getValue();
    var empName   = balances.getRange(6 + i, 3).getValue();
    if (empProfit > maxProfit) { maxProfit = empProfit; bestEmp = empName; }
  }

  var archiveLastRow = archive.getLastRow();
  var archiveNextRow = Math.max(archiveLastRow + 1, 6);
  for (var ar = 6; ar <= 57; ar++) {
    if (archive.getRange(ar, 3).getValue() === "" || archive.getRange(ar, 3).getValue() === null) {
      archiveNextRow = ar; break;
    }
  }

  archive.getRange(archiveNextRow, 2).setValue(archiveNextRow - 5);
  archive.getRange(archiveNextRow, 3).setValue(week);
  archive.getRange(archiveNextRow, 4).setValue(totalDeals);
  archive.getRange(archiveNextRow, 5).setValue(totalTurnover);
  archive.getRange(archiveNextRow, 6).setValue(totalProfitUSDT);
  archive.getRange(archiveNextRow, 7).setValue(totalProfitRUB);
  archive.getRange(archiveNextRow, 8).setValue(totalPaid);
  archive.getRange(archiveNextRow, 9).setValue(compShare);
  archive.getRange(archiveNextRow, 10).setValue(avgProfit);
  archive.getRange(archiveNextRow, 11).setValue(bestEmp);

  var lastDataRow = journal.getLastRow();
  if (lastDataRow >= 3) {
    var dealData = journal.getRange(3, 1, lastDataRow - 2, 15).getValues();
    var filteredDeals = dealData.filter(function(row) { return row[1] !== ""; });

    if (filteredDeals.length > 0) {
      var sheetName = String(week).replace(/[\/\\\?\*\[\]:]/g, "-").substring(0, 31);
      var existingSheet = ss.getSheetByName(sheetName);
      if (existingSheet) { sheetName = sheetName.substring(0, 27) + "(2)"; }

      var dealsSheet = ss.insertSheet(sheetName);
      var headers = ["ID площадки","ID сервиса","Дата/время","₽ получили","USDT отдали",
        "Токен(комп.рекв)","Заморозка","Исполнитель","Комп.продажи","USDT получили",
        "₽ отдали","Спред%","Прибыль USDT","Курс","Комментарий"];
      dealsSheet.getRange(1, 1, 1, 15).setValues([headers]);
      dealsSheet.getRange(1, 1, 1, 15).setFontWeight("bold").setBackground("#1c3a4a").setFontColor("#ffffff");
      dealsSheet.getRange(2, 1, filteredDeals.length, 15).setValues(filteredDeals);
      var summaryRow = filteredDeals.length + 3;
      dealsSheet.getRange(summaryRow, 1).setValue("ИТОГО").setFontWeight("bold");
      dealsSheet.getRange(summaryRow, 2).setValue("Сделок: " + filteredDeals.length).setFontWeight("bold");
      dealsSheet.getRange(summaryRow, 6).setValue("Прибыль USDT:").setFontWeight("bold");
      dealsSheet.getRange(summaryRow, 7).setFormula("=SUM(M2:M" + (filteredDeals.length + 1) + ")").setFontWeight("bold");
      dealsSheet.hideSheet();
    }
  }

  journal.getRange("A3:O2002").clearContent();
  balances.getRange("J6:J20").setValue(0);

  var props = PropertiesService.getScriptProperties();
  var allKeys = props.getKeys();
  for (var k = 0; k < allKeys.length; k++) {
    if (allKeys[k].indexOf("exec_") === 0) { props.deleteProperty(allKeys[k]); }
  }

  var askWeek = ui.prompt(
    "Новый период",
    "Введите название следующей недели (например: 21.04–27.04.2026):",
    ui.ButtonSet.OK_CANCEL
  );
  if (askWeek.getSelectedButton() === ui.Button.OK && askWeek.getResponseText() !== "") {
    settings.getRange("D5").setValue(askWeek.getResponseText());
  }

  ui.alert("✓ Неделя закрыта!\n\n" +
    "Отчёт отправлен на genadijton@gmail.com.\n" +
    "Сделки сохранены на листе (скрыт).\n" +
    "Итоги в Архиве. Журнал очищен.");
}

// ══════════════════════════════════════════════════════
// NEW: АВТО-ОТЧЁТ НА EMAIL ПЕРЕД ЗАКРЫТИЕМ НЕДЕЛИ
// ══════════════════════════════════════════════════════
function sendWeeklyReport() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var journal = ss.getSheetByName("Journal");
    var settings = ss.getSheetByName("Settings");
    var balances = ss.getSheetByName("Balances");

    var week = settings ? settings.getRange("D5").getValue() : "—";
    var lastRow = journal ? journal.getLastRow() : 2;

    // Собираем данные из журнала
    var deals = [];
    if (lastRow >= 3) {
      var data = journal.getRange(3, 1, lastRow - 2, 15).getValues();
      for (var i = 0; i < data.length; i++) {
        var r = data[i];
        if (!r[1] || r[1] === "") continue;
        var isDone = r[9] !== "" && r[9] !== null && r[9] !== 0;
        if (!isDone) continue;
        var spreadRaw = r[11];
        var spreadPct = (typeof spreadRaw === "number") ? spreadRaw * 100 : parseFloat(String(spreadRaw).replace("%","")) || 0;
        deals.push({
          dt: r[2], rub: r[3]||0, executor: r[7]||"—",
          company: r[5]||"—", profit: parseFloat(r[12])||0,
          spread: spreadPct, rate: parseFloat(r[13])||0
        });
      }
    }

    var totalDeals = deals.length;
    var totalVol = deals.reduce(function(s,d){ return s + d.rub; }, 0);
    var totalProfit = deals.reduce(function(s,d){ return s + d.profit; }, 0);
    var profDeals = deals.filter(function(d){ return d.profit > 0; }).length;
    var avgProfit = totalDeals > 0 ? totalProfit / totalDeals : 0;
    var avgRate = deals.length > 0 ? deals.reduce(function(s,d){ return s+d.rate; },0) / deals.length : 0;

    // По исполнителям
    var empMap = {};
    deals.forEach(function(d) {
      if (!empMap[d.executor]) empMap[d.executor] = {deals:0, profit:0, vol:0};
      empMap[d.executor].deals++;
      empMap[d.executor].profit += d.profit;
      empMap[d.executor].vol += d.rub;
    });
    var empRows = Object.keys(empMap).map(function(n) {
      return {name:n, deals:empMap[n].deals, profit:empMap[n].profit, vol:empMap[n].vol};
    }).sort(function(a,b){ return b.profit - a.profit; });

    // По компаниям
    var compMap = {};
    deals.forEach(function(d) {
      if (!compMap[d.company]) compMap[d.company] = {deals:0, profit:0};
      compMap[d.company].deals++;
      compMap[d.company].profit += d.profit;
    });
    var compRows = Object.keys(compMap).map(function(n) {
      return {name:n, deals:compMap[n].deals, profit:compMap[n].profit};
    }).sort(function(a,b){ return b.profit - a.profit; });

    // Строим HTML письмо
    var html = '<div style="font-family:Arial,sans-serif;max-width:640px;margin:0 auto;">';
    html += '<div style="background:#1c3a4a;color:#fff;padding:20px 24px;border-radius:8px 8px 0 0;">';
    html += '<h2 style="margin:0;font-size:20px;">📊 Отчёт P2P Tracker</h2>';
    html += '<p style="margin:6px 0 0;opacity:0.7;font-size:14px;">Неделя: ' + week + '</p></div>';

    html += '<div style="background:#f5f5f5;padding:20px 24px;">';

    // KPI блок
    html += '<table style="width:100%;border-collapse:collapse;margin-bottom:16px;">';
    html += '<tr>';
    html += kpiCell("Сделок завершено", totalDeals);
    html += kpiCell("Прибыльных", profDeals + " (" + (totalDeals>0?(profDeals/totalDeals*100).toFixed(0):0) + "%)");
    html += kpiCell("Оборот ₽", fmtNum(totalVol,0) + " ₽");
    html += kpiCell("Прибыль USDT", (totalProfit>=0?"+":"") + fmtNum(totalProfit,2) + " ₮");
    html += '</tr><tr>';
    html += kpiCell("Средняя прибыль", (avgProfit>=0?"+":"") + fmtNum(avgProfit,2) + " ₮/сд.");
    html += kpiCell("Средний курс", fmtNum(avgRate,2) + " ₽/$");
    html += kpiCell("% прибыльных", (totalDeals>0?(profDeals/totalDeals*100).toFixed(1):0) + "%");
    html += kpiCell("Убыточных", (totalDeals - profDeals));
    html += '</tr></table>';

    // Таблица по сотрудникам
    html += '<h3 style="margin:16px 0 8px;font-size:15px;color:#1c3a4a;">По исполнителям</h3>';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<tr style="background:#1c3a4a;color:#fff;"><th style="padding:8px;text-align:left;">Исполнитель</th><th style="padding:8px;text-align:right;">Сделок</th><th style="padding:8px;text-align:right;">Оборот ₽</th><th style="padding:8px;text-align:right;">Прибыль ₮</th></tr>';
    empRows.forEach(function(e, idx) {
      var bg = idx % 2 === 0 ? "#fff" : "#f0f4f8";
      var color = e.profit >= 0 ? "#2d7a2d" : "#a12c2c";
      html += '<tr style="background:' + bg + ';">';
      html += '<td style="padding:7px 8px;font-weight:600;">' + e.name + '</td>';
      html += '<td style="padding:7px 8px;text-align:right;">' + e.deals + '</td>';
      html += '<td style="padding:7px 8px;text-align:right;">' + fmtNum(e.vol,0) + ' ₽</td>';
      html += '<td style="padding:7px 8px;text-align:right;color:' + color + ';font-weight:600;">' + (e.profit>=0?"+":"") + fmtNum(e.profit,2) + ' ₮</td>';
      html += '</tr>';
    });
    html += '</table>';

    // Таблица по компаниям
    html += '<h3 style="margin:16px 0 8px;font-size:15px;color:#1c3a4a;">По компаниям-реквизитчикам</h3>';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<tr style="background:#1c3a4a;color:#fff;"><th style="padding:8px;text-align:left;">Компания</th><th style="padding:8px;text-align:right;">Сделок</th><th style="padding:8px;text-align:right;">Прибыль ₮</th><th style="padding:8px;text-align:right;">Ср. прибыль</th></tr>';
    compRows.forEach(function(c, idx) {
      var bg = idx % 2 === 0 ? "#fff" : "#f0f4f8";
      var color = c.profit >= 0 ? "#2d7a2d" : "#a12c2c";
      html += '<tr style="background:' + bg + ';">';
      html += '<td style="padding:7px 8px;font-weight:600;">' + c.name + '</td>';
      html += '<td style="padding:7px 8px;text-align:right;">' + c.deals + '</td>';
      html += '<td style="padding:7px 8px;text-align:right;color:' + color + ';font-weight:600;">' + (c.profit>=0?"+":"") + fmtNum(c.profit,2) + ' ₮</td>';
      html += '<td style="padding:7px 8px;text-align:right;">' + (c.deals>0?fmtNum(c.profit/c.deals,2):"—") + ' ₮</td>';
      html += '</tr>';
    });
    html += '</table>';

    html += '<p style="margin:20px 0 0;font-size:12px;color:#888;">Отчёт сформирован автоматически системой P2P Tracker</p>';
    html += '</div></div>';

    MailApp.sendEmail({
      to: "genadijton@gmail.com",
      subject: "📊 P2P Отчёт недели: " + week + " | Прибыль: " + (totalProfit>=0?"+":"") + fmtNum(totalProfit,2) + " ₮",
      htmlBody: html
    });
  } catch(err) {
    // Не критично — продолжаем закрытие недели
    Logger.log("sendWeeklyReport error: " + err.message);
  }
}

function kpiCell(label, value) {
  return '<td style="background:#fff;border-radius:6px;padding:10px 12px;text-align:center;margin:4px;">' +
    '<div style="font-size:11px;color:#888;text-transform:uppercase;margin-bottom:4px;">' + label + '</div>' +
    '<div style="font-size:16px;font-weight:700;color:#1c3a4a;">' + value + '</div></td>';
}

function fmtNum(n, dec) {
  if (n === null || n === undefined) return "—";
  return n.toLocaleString("ru-RU", {minimumFractionDigits:dec, maximumFractionDigits:dec});
}

// ══════════════════════════════════════════════════════
// NEW: РАСЧЁТ ЗАРПЛАТЫ (% от оборота)
// ══════════════════════════════════════════════════════
function calcSalary() {
  var OWNER = "genadijton@gmail.com";
  if (Session.getActiveUser().getEmail() !== OWNER) {
    SpreadsheetApp.getUi().alert("Доступ запрещён.");
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Читаем % из Settings (D10 — ячейка с % зарплаты, если нет — спрашиваем)
  var settings = ss.getSheetByName("Settings");
  var salaryPct = 0;
  if (settings) {
    var settingVal = settings.getRange("D10").getValue();
    if (settingVal && settingVal > 0) { salaryPct = parseFloat(settingVal); }
  }

  if (!salaryPct || salaryPct <= 0) {
    var r = ui.prompt(
      "Расчёт зарплаты",
      "Введите % от оборота ₽ для всех сотрудников (например: 2.5 = 2,5%):\n\n" +
      "💡 Совет: можно задать постоянный % в ячейке Settings!D10 — тогда не нужно вводить каждый раз.",
      ui.ButtonSet.OK_CANCEL
    );
    if (r.getSelectedButton() !== ui.Button.OK) return;
    salaryPct = parseFloat(r.getResponseText().replace(",", "."));
    if (isNaN(salaryPct) || salaryPct <= 0) { ui.alert("Некорректный процент."); return; }
  }

  // Собираем данные из журнала (только завершённые сделки — заполнен J)
  var journal = ss.getSheetByName("Journal");
  var lastRow = journal.getLastRow();
  if (lastRow < 3) { ui.alert("Журнал пуст."); return; }

  var data = journal.getRange(3, 1, lastRow - 2, 15).getValues();
  var empMap = {};

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (!r[1] || r[1] === "") continue;             // нет ID сервиса
    var isDone = r[9] !== "" && r[9] !== null && r[9] !== 0;
    if (!isDone) continue;                           // незавершённая
    var executor = String(r[7] || "Неизвестно").trim();
    var rub = parseFloat(r[3]) || 0;                // D — ₽ получили (оборот)
    var profit = parseFloat(r[12]) || 0;            // M — Прибыль USDT

    if (!empMap[executor]) empMap[executor] = {deals:0, vol:0, profit:0};
    empMap[executor].deals++;
    empMap[executor].vol += rub;
    empMap[executor].profit += profit;
  }

  if (Object.keys(empMap).length === 0) {
    ui.alert("Нет завершённых сделок для расчёта.");
    return;
  }

  // Создаём / пересоздаём лист «Зарплата»
  var salSheet = ss.getSheetByName("Зарплата");
  if (salSheet) { ss.deleteSheet(salSheet); }
  salSheet = ss.insertSheet("Зарплата");

  var ACCENT = "#1c3a4a";
  var FG     = "#ffffff";
  var ALT    = "#f0f4f8";

  // Заголовок
  salSheet.getRange("A1:G1").merge().setValue("💰 РАСЧЁТ ЗАРПЛАТЫ")
    .setBackground(ACCENT).setFontColor(FG).setFontSize(14).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  salSheet.setRowHeight(1, 36);

  salSheet.getRange("A2").setValue("% от оборота:").setFontWeight("bold");
  salSheet.getRange("B2").setValue(salaryPct + "%").setFontWeight("bold").setFontColor("#e67e22");

  var settingsRef = settings ? settings.getRange("D5").getValue() : "—";
  salSheet.getRange("C2").setValue("Неделя:");
  salSheet.getRange("D2:E2").merge().setValue(settingsRef);

  // Заголовки таблицы
  var headers = ["Исполнитель", "Сделок", "Оборот ₽", "Прибыль USDT", "Зарплата ₽", "Статус выплаты", "Комментарий"];
  salSheet.getRange(4, 1, 1, 7).setValues([headers])
    .setBackground(ACCENT).setFontColor(FG).setFontWeight("bold");

  var empArr = Object.keys(empMap).map(function(n) {
    return {name:n, deals:empMap[n].deals, vol:empMap[n].vol, profit:empMap[n].profit};
  }).sort(function(a,b){ return b.vol - a.vol; });

  var totalVol = 0, totalProfit = 0, totalSalary = 0;

  for (var ei = 0; ei < empArr.length; ei++) {
    var e = empArr[ei];
    var salary = Math.round(e.vol * salaryPct / 100);
    var rowIdx = 5 + ei;
    var isAlt = ei % 2 === 1;
    var bg = isAlt ? ALT : "#ffffff";

    salSheet.getRange(rowIdx, 1).setValue(e.name);
    salSheet.getRange(rowIdx, 2).setValue(e.deals);
    salSheet.getRange(rowIdx, 3).setValue(e.vol).setNumberFormat("#,##0");
    salSheet.getRange(rowIdx, 4).setValue(e.profit.toFixed(2));
    salSheet.getRange(rowIdx, 5).setValue(salary).setNumberFormat("#,##0").setFontWeight("bold").setFontColor("#1a6b2a");
    salSheet.getRange(rowIdx, 6).setValue("⏳ Ожидает");
    salSheet.getRange(rowIdx, 7).setValue("");
    salSheet.getRange(rowIdx, 1, 1, 7).setBackground(bg);

    totalVol += e.vol;
    totalProfit += e.profit;
    totalSalary += salary;
  }

  // Итого
  var totRow = 5 + empArr.length + 1;
  salSheet.getRange(totRow, 1, 1, 7).setBackground("#dde8f0").setFontWeight("bold");
  salSheet.getRange(totRow, 1).setValue("ИТОГО");
  salSheet.getRange(totRow, 2).setValue(empArr.reduce(function(s,e){ return s+e.deals; },0));
  salSheet.getRange(totRow, 3).setValue(totalVol).setNumberFormat("#,##0");
  salSheet.getRange(totRow, 4).setValue(totalProfit.toFixed(2));
  salSheet.getRange(totRow, 5).setValue(totalSalary).setNumberFormat("#,##0").setFontColor("#1a6b2a");

  // Ширина колонок
  salSheet.setColumnWidth(1, 150);
  salSheet.setColumnWidth(2, 70);
  salSheet.setColumnWidth(3, 120);
  salSheet.setColumnWidth(4, 120);
  salSheet.setColumnWidth(5, 120);
  salSheet.setColumnWidth(6, 130);
  salSheet.setColumnWidth(7, 180);

  ss.setActiveSheet(salSheet);

  ui.alert("✓ Зарплата рассчитана!\n\n" +
    "% от оборота: " + salaryPct + "%\n" +
    "Сотрудников: " + empArr.length + "\n" +
    "Итого к выплате: " + totalSalary.toLocaleString("ru-RU") + " ₽\n\n" +
    "Лист «Зарплата» открыт.");
}

// ══════════════════════════════════════════════════════
// NEW: АЛЕРТЫ НА АНОМАЛИИ (вызывается из myonEdit при заполнении J)
// ══════════════════════════════════════════════════════
function checkAnomalies(executorName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var journal = ss.getSheetByName("Journal");
    var settings = ss.getSheetByName("Settings");
    var lastRow = journal.getLastRow();
    if (lastRow < 3) return;

    // Порог спреда из Settings!D11 (по умолчанию 1%)
    var spreadThreshold = 1.0;
    if (settings) {
      var stVal = settings.getRange("D11").getValue();
      if (stVal && stVal > 0) { spreadThreshold = parseFloat(stVal); }
    }

    var data = journal.getRange(3, 1, lastRow - 2, 15).getValues();

    // Фильтруем завершённые сделки данного исполнителя, сортируем по порядку строк (они уже в порядке)
    var empDeals = [];
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (!r[1] || r[1] === "") continue;
      var isDone = r[9] !== "" && r[9] !== null && r[9] !== 0;
      if (!isDone) continue;
      if (String(r[7]).trim() !== String(executorName).trim()) continue;

      var profit = parseFloat(r[12]) || 0;
      var spreadRaw = r[11];
      var spreadPct = (typeof spreadRaw === "number") ? spreadRaw * 100 : parseFloat(String(spreadRaw).replace("%","")) || 0;
      empDeals.push({profit: profit, spread: spreadPct});
    }

    if (empDeals.length === 0) return;

    var alerts = [];

    // 1. Проверяем 3+ убыточных подряд (последние сделки)
    if (empDeals.length >= 3) {
      var last3 = empDeals.slice(-3);
      var allNeg = last3.every(function(d){ return d.profit <= 0; });
      if (allNeg) {
        alerts.push("⚠️ УБЫТОЧНЫЕ СДЕЛКИ ПОДРЯД\n\n" +
          "Исполнитель: " + executorName + "\n" +
          "Последние 3 сделки подряд убыточны!\n" +
          "Прибыли: " + last3.map(function(d){ return d.profit.toFixed(2)+"₮"; }).join(", ") + "\n\n" +
          "Рекомендуем проверить условия работы.");
      }
    }

    // 2. Проверяем спред последней сделки
    var lastDeal = empDeals[empDeals.length - 1];
    if (lastDeal.spread > 0 && lastDeal.spread < spreadThreshold) {
      alerts.push("⚠️ НИЗКИЙ СПРЕД\n\n" +
        "Исполнитель: " + executorName + "\n" +
        "Спред последней сделки: " + lastDeal.spread.toFixed(2) + "% (порог: " + spreadThreshold + "%)\n\n" +
        "Возможно невыгодные условия по этой компании.");
    }

    if (alerts.length === 0) return;

    // Отправляем email
    var alertText = alerts.join("\n\n---\n\n");
    MailApp.sendEmail({
      to: "genadijton@gmail.com",
      subject: "🚨 P2P Алерт: " + executorName + " — аномалия обнаружена",
      body: "P2P Tracker обнаружил аномалию:\n\n" + alertText +
            "\n\n---\nАвтоматическое уведомление системы P2P Tracker"
    });
  } catch(err) {
    Logger.log("checkAnomalies error: " + err.message);
  }
}

// ══════════════════════════════════════════════════════
// NEW: ЛИМИТЫ ПО КОМПАНИЯМ (вызывается при заполнении F)
// ══════════════════════════════════════════════════════
function checkCompanyLimits(companyName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settings = ss.getSheetByName("Settings");
    if (!settings) return;

    // Лимиты хранятся в Settings начиная с D14:E28
    // D = Компания, E = Лимит сделок в день
    var limitsData = settings.getRange("D14:E28").getValues();
    var limit = 0;
    for (var i = 0; i < limitsData.length; i++) {
      if (String(limitsData[i][0]).trim().toLowerCase() === companyName.toLowerCase()) {
        limit = parseInt(limitsData[i][1]) || 0;
        break;
      }
    }
    if (!limit || limit <= 0) return; // Лимит не задан — не проверяем

    // Считаем сделки с этой компанией за сегодня
    var journal = ss.getSheetByName("Journal");
    var lastRow = journal.getLastRow();
    if (lastRow < 3) return;

    var tz = "Asia/Novosibirsk";
    var today = Utilities.formatDate(new Date(), tz, "dd.MM.yyyy");
    var data = journal.getRange(3, 1, lastRow - 2, 15).getDisplayValues();

    var todayCount = 0;
    for (var j = 0; j < data.length; j++) {
      var r = data[j];
      if (!r[1] || r[1] === "") continue;
      if (String(r[5]).trim().toLowerCase() !== companyName.toLowerCase()) continue;
      var dateStr = String(r[2]).trim();
      if (dateStr.indexOf(today) === 0) { todayCount++; }
    }

    if (todayCount >= limit) {
      // Предупреждение в таблице (toast)
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "⚠️ Компания " + companyName + ": достигнут дневной лимит (" + limit + " сделок). " +
        "Сегодня уже " + todayCount + " сделок.",
        "🚨 Лимит достигнут",
        15
      );

      // И email
      MailApp.sendEmail({
        to: "genadijton@gmail.com",
        subject: "🚨 P2P Лимит: " + companyName + " — " + todayCount + "/" + limit + " сделок сегодня",
        body: "Достигнут дневной лимит по компании " + companyName + ".\n\n" +
              "Лимит: " + limit + " сделок/день\n" +
              "Фактически сегодня: " + todayCount + " сделок\n\n" +
              "Рекомендуем остановить приём заявок от этой компании на сегодня."
      });
    } else if (todayCount >= limit * 0.8) {
      // 80% лимита — предупреждение без email
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "ℹ️ " + companyName + ": использовано " + todayCount + "/" + limit + " сделок сегодня (80% лимита).",
        "Лимит близок",
        8
      );
    }
  } catch(err) {
    Logger.log("checkCompanyLimits error: " + err.message);
  }
}

// ══════════════════════════════════════════════════════
// NEW: ИСТОРИЯ КУРСОВ (записывается автоматически)
// ══════════════════════════════════════════════════════
function saveRateHistory(rate, company) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hist = ss.getSheetByName("RateHistory");

    // Создаём лист если нет
    if (!hist) {
      hist = ss.insertSheet("RateHistory");
      hist.getRange("A1:D1").setValues([["Дата/время","Курс USDT/RUB","Компания","Источник"]])
        .setBackground("#1c3a4a").setFontColor("#fff").setFontWeight("bold");
    }

    var tz = "Asia/Novosibirsk";
    var now = Utilities.formatDate(new Date(), tz, "dd.MM.yyyy HH:mm");
    var nextRow = Math.max(hist.getLastRow() + 1, 2);
    hist.getRange(nextRow, 1).setValue(now);
    hist.getRange(nextRow, 2).setValue(rate);
    hist.getRange(nextRow, 3).setValue(company || "");
    hist.getRange(nextRow, 4).setValue("Rapira API");
  } catch(err) {
    Logger.log("saveRateHistory error: " + err.message);
  }
}

// ══════════════════════════════════════════════════════
// 4. МЕНЮ В ТАБЛИЦЕ
// ══════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("P2P Tracker")
    .addItem("📅 Закрыть неделю", "closeWeek")
    .addSeparator()
    .addItem("💰 Расчёт зарплаты", "calcSalary")
    .addSeparator()
    .addItem("💱 Проставить курс по периоду", "applyRateByPeriod")
    .addSeparator()
    .addItem("🔓 Разблокировать строку H (ввести номер)", "unlockExecutorRow")
    .addSeparator()
    .addItem("🧮 Открыть калькулятор сделок", "openCalculator")
    .addSeparator()
    .addItem("📊 Обновить аналитику сделок", "buildAnalytics")
    .addToUi();
}

// ══════════════════════════════════════════════════════
// 7. ОТКРЫТЬ КАЛЬКУЛЯТОР СДЕЛОК
// ══════════════════════════════════════════════════════
function openCalculator() {
  var html = HtmlService.createHtmlOutputFromFile("Calc")
    .setWidth(520)
    .setHeight(700)
    .setTitle("Калькулятор сделок")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ══════════════════════════════════════════════════════
// 6. ПРОСТАВИТЬ КУРС ПО ПЕРИОДУ
// ══════════════════════════════════════════════════════
function applyRateByPeriod() {
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt("Шаг 1 из 4", "Введите название компании точно как в журнале (например: ARGOS):", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK || r1.getResponseText() === "") return;
  var company = r1.getResponseText().trim();

  var r2 = ui.prompt("Шаг 2 из 4", "Дата НАЧАЛА периода (DD.MM.YYYY, например: 13.04.2026):", ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK || r2.getResponseText() === "") return;

  var r3 = ui.prompt("Шаг 3 из 4", "Дата КОНЦА и время (DD.MM.YYYY HH:MM, например: 13.04.2026 16:00):", ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK || r3.getResponseText() === "") return;

  var r4 = ui.prompt("Шаг 4 из 4", "Введите курс USDT/RUB (например: 94.50):", ui.ButtonSet.OK_CANCEL);
  if (r4.getSelectedButton() !== ui.Button.OK || r4.getResponseText() === "") return;

  var rate = parseFloat(r4.getResponseText().replace(",", "."));
  if (isNaN(rate) || rate <= 0) { ui.alert("Некорректный курс."); return; }

  var fromStr = displayToSortable(r2.getResponseText().trim());
  var toStr   = displayToSortable(r3.getResponseText().trim());
  if (!fromStr || !toStr) { ui.alert("Не удалось распознать даты. Проверьте формат DD.MM.YYYY или DD.MM.YYYY HH:MM."); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var journal = ss.getSheetByName("Journal");
  var lastRow = journal.getLastRow();
  var updated = 0;
  if (lastRow < 3) { ui.alert("Журнал пуст."); return; }

  var dataRange = lastRow - 2;
  var colC  = journal.getRange(3, 3,  dataRange, 1).getDisplayValues();
  var colD  = journal.getRange(3, 4,  dataRange, 1).getValues();
  var colF  = journal.getRange(3, 6,  dataRange, 1).getValues();
  var colJ  = journal.getRange(3, 10, dataRange, 1).getValues();
  var colN  = journal.getRange(3, 14, dataRange, 1).getValues();
  var colNF = journal.getRange(3, 14, dataRange, 1).getFormulas();
  var defaultFormula = "=Settings!$D$7";

  var newRates = colN.map(function(r) { return [r[0]]; });
  var newUsdt  = colJ.map(function(r) { return [r[0]]; });
  var updatedUsdt = 0;

  for (var i = 0; i < dataRange; i++) {
    var dateVal = colC[i][0], compVal = colF[i][0], rateVal = colN[i][0];
    var rateFormula = colNF[i][0];
    if (!dateVal || dateVal === "") continue;
    var rowStr = displayToSortable(String(dateVal).trim());
    if (!rowStr) continue;
    var compMatch = String(compVal).trim().toLowerCase() === company.toLowerCase();
    var dateMatch = rowStr >= fromStr && rowStr <= toStr;
    var rateStr = String(rateVal).trim();
    var isDefault = (rateStr === "" || rateStr === "0" || rateVal === 0 || rateVal === null)
                 || (rateFormula !== "" && rateFormula.replace(/\s/g,"").toUpperCase() === defaultFormula.replace(/\s/g,"").toUpperCase());
    if (compMatch && dateMatch && isDefault) {
      newRates[i][0] = rate;
      updated++;
      var jVal = colJ[i][0], dVal = colD[i][0];
      var jEmpty = (jVal === "" || jVal === null || jVal === 0);
      if (jEmpty && dVal && dVal !== "" && dVal !== 0 && rate > 0) {
        newUsdt[i][0] = Math.round((dVal / rate) * 100) / 100;
        updatedUsdt++;
      }
    }
  }

  if (updated > 0) { journal.getRange(3, 14, dataRange, 1).setValues(newRates); }
  if (updatedUsdt > 0) { journal.getRange(3, 10, dataRange, 1).setValues(newUsdt); }

  // Сохраняем в историю курсов
  if (updated > 0) { saveRateHistory(rate, company); }

  ui.alert("Готово!", "Курс " + rate + " проставлен в " + updated + " строках\nUSDT получили заполнено в " + updatedUsdt + " строках\nКомпания: " + company + "\nПериод: " + r2.getResponseText() + " — " + r3.getResponseText(), ui.ButtonSet.OK);
}

// ══════════════════════════════════════════════════════
// 5. РАЗБЛОКИРОВАТЬ СТРОКУ H
// ══════════════════════════════════════════════════════
function unlockExecutorRow() {
  var OWNER = "genadijton@gmail.com";
  if (Session.getActiveUser().getEmail() !== OWNER) {
    SpreadsheetApp.getUi().alert("Доступ запрещён. Обратитесь к администратору.");
    return;
  }
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Разблокировать исполнителя", "Введите номер строки для разблокировки (например: 5):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var rowNum = parseInt(response.getResponseText());
  if (isNaN(rowNum) || rowNum < 3) { ui.alert("Некорректный номер строки."); return; }
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty("exec_" + rowNum);
  ui.alert("Строка " + rowNum + " разблокирована. Теперь можно изменить исполнителя.");
}

// ══════════════════════════════════════════════════════
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДАТ
// ══════════════════════════════════════════════════════
function displayToSortable(str) {
  try {
    str = String(str).trim();
    var hours = 0, minutes = 0, parts;
    if (str.indexOf(" ") > -1) {
      var dt = str.split(" "); parts = dt[0].split(".");
      var tp = dt[1].split(":"); hours = parseInt(tp[0]); minutes = parseInt(tp[1]);
    } else { parts = str.split("."); }
    var day = parseInt(parts[0]), month = parseInt(parts[1]) - 1, year = parseInt(parts[2]);
    if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
    return year + "-" + (month+1<10?"0":"")+(month+1) + "-" + (day<10?"0":"")+day + " " + (hours<10?"0":"")+hours + ":" + (minutes<10?"0":"")+minutes;
  } catch(e) { return null; }
}

function parseDateUTC(str, tz) {
  try {
    str = String(str).trim(); var hours = 0, minutes = 0, parts;
    if (str.indexOf(" ") > -1) {
      var dt = str.split(" "); parts = dt[0].split("."); var tp = dt[1].split(":");
      hours = parseInt(tp[0]); minutes = parseInt(tp[1]);
    } else { parts = str.split("."); }
    var day = parseInt(parts[0]), month = parseInt(parts[1]) - 1, year = parseInt(parts[2]);
    if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
    var formatted = Utilities.formatString("%04d/%02d/%02d %02d:%02d", year, month+1, day, hours, minutes);
    return Utilities.parseDate(formatted, tz, "yyyy/MM/dd HH:mm");
  } catch(e) { return null; }
}

function normalizeToLocal(val, tz) {
  try {
    if (val instanceof Date) { return Utilities.formatDate(val, tz, "yyyy-MM-dd HH:mm"); }
    var s = String(val).trim(); var hours = 0, minutes = 0, parts;
    if (s.indexOf(" ") > -1) {
      var dt = s.split(" "); parts = dt[0].split("."); var tp = dt[1].split(":");
      hours = parseInt(tp[0]); minutes = parseInt(tp[1]);
    } else { parts = s.split("."); }
    var day = parseInt(parts[0]), month = parseInt(parts[1]) - 1, year = parseInt(parts[2]);
    if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
    return year + "-" + (month+1<10?"0":"")+(month+1) + "-" + (day<10?"0":"")+day + " " + (hours<10?"0":"")+hours + ":" + (minutes<10?"0":"")+minutes;
  } catch(e) { return null; }
}

function parseDate(str, tz) {
  try {
    str = str.trim(); var parts, hours = 0, minutes = 0;
    if (str.indexOf(" ") > -1) {
      var dt = str.split(" "); parts = dt[0].split("."); var timeParts = dt[1].split(":");
      hours = parseInt(timeParts[0]); minutes = parseInt(timeParts[1]);
    } else { parts = str.split("."); }
    var day = parseInt(parts[0]), month = parseInt(parts[1]) - 1, year = parseInt(parts[2]);
    return new Date(year, month, day, hours, minutes);
  } catch(e) { return null; }
}

// ══════════════════════════════════════════════════════
// 8. АНАЛИТИКА ПО СДЕЛКАМ (с БЛОКОМ 9 — сравнение недель)
// ══════════════════════════════════════════════════════
function buildAnalytics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var journal = ss.getSheetByName("Journal");
  var archive = ss.getSheetByName("Archive");
  var settings = ss.getSheetByName("Settings");

  if (!journal) { SpreadsheetApp.getUi().alert("Лист «Journal» не найден."); return; }

  var lastRow = journal.getLastRow();
  if (lastRow < 3) { SpreadsheetApp.getUi().alert("В журнале нет данных."); return; }

  var TZ = "Asia/Novosibirsk";
  var rawData = journal.getRange(3, 1, lastRow - 2, 15).getValues();

  // Берём только завершённые (J заполнена)
  var deals = [];
  for (var i = 0; i < rawData.length; i++) {
    var r = rawData[i];
    if (!r[1] || r[1] === "") continue;
    var isDone = r[9] !== "" && r[9] !== null && r[9] !== 0;
    if (!isDone) continue;

    var dtVal = r[2];
    var dtStr = "";
    if (dtVal instanceof Date) {
      dtStr = Utilities.formatDate(dtVal, TZ, "dd.MM.yyyy HH:mm");
    } else {
      dtStr = String(dtVal).trim();
    }

    var spreadRaw = r[11];
    var spreadPct;
    if (typeof spreadRaw === "number") {
      spreadPct = spreadRaw * 100;
    } else {
      spreadPct = parseFloat(String(spreadRaw).replace("%","").replace(",",".")) || 0;
    }

    var hour = 0;
    try {
      var timePart = dtStr.indexOf(" ") > -1 ? dtStr.split(" ")[1] : "";
      if (timePart) hour = parseInt(timePart.split(":")[0]);
    } catch(ex) {}

    deals.push({
      dt: dtStr, rub: parseFloat(r[3])||0, executor: String(r[7]||"—").trim(),
      company: String(r[5]||"—").trim(), profit: parseFloat(r[12])||0,
      spread: spreadPct, freeze: String(r[6]||"").trim().toLowerCase(),
      rate: parseFloat(r[13])||0, hour: hour
    });
  }

  var totalDeals = deals.length;
  if (totalDeals === 0) { SpreadsheetApp.getUi().alert("Нет завершённых сделок для анализа."); return; }

  var totalProfit = deals.reduce(function(s,d){ return s+d.profit; }, 0);
  var profDeals   = deals.filter(function(d){ return d.profit > 0; }).length;
  var lossDeals   = totalDeals - profDeals;
  var avgProfit   = totalProfit / totalDeals;
  var totalVol    = deals.reduce(function(s,d){ return s+d.rub; }, 0);
  var avgSpread   = deals.reduce(function(s,d){ return s+d.spread; }, 0) / totalDeals;

  // Создаём лист Аналитика
  var existSheet = ss.getSheetByName("Аналитика");
  if (existSheet) { ss.deleteSheet(existSheet); }
  var sheet = ss.insertSheet("Аналитика");

  var row = 1;
  var ACCENT   = "#1c3a4a";
  var DARK_BG  = "#2c4a5a";
  var HEADER_FG = "#ffffff";
  var ALT_BG   = "#f0f4f8";

  function writeTitle(text, bgColor, fgColor, fontSize) {
    var r = sheet.getRange(row, 1, 1, 8);
    r.merge().setValue(text).setBackground(bgColor).setFontColor(fgColor)
     .setFontSize(fontSize||12).setFontWeight("bold").setVerticalAlignment("middle");
    sheet.setRowHeight(row, 32); row++;
  }
  function writeHeaders(headers, bg) {
    var r = sheet.getRange(row, 1, 1, headers.length);
    r.setValues([headers]).setBackground(bg||DARK_BG).setFontColor(HEADER_FG)
     .setFontWeight("bold").setFontSize(10);
    sheet.setRowHeight(row, 22); row++;
  }
  function writeRow(values, isAlt, isTotal) {
    var r = sheet.getRange(row, 1, 1, values.length);
    r.setValues([values]);
    if (isTotal) { r.setBackground("#dde8f0").setFontWeight("bold"); }
    else if (isAlt) { r.setBackground(ALT_BG); }
    else { r.setBackground("#ffffff"); }
    sheet.setRowHeight(row, 20); row++;
  }
  function blankRow() { row++; }

  // ═══ БЛОК 1: ОБЩАЯ СВОДКА ═══
  writeTitle("📊 ОБЩАЯ СВОДКА", ACCENT, HEADER_FG, 13);
  writeHeaders(["Показатель","Значение","","Показатель","Значение","","",""], DARK_BG);
  writeRow(["Всего завершённых сделок", totalDeals, "", "Прибыльных сделок", profDeals + " (" + (profDeals/totalDeals*100).toFixed(1) + "%)", "", "", ""], false);
  writeRow(["Общая прибыль USDT", totalProfit.toFixed(2)+" ₮", "", "Убыточных сделок", lossDeals + " (" + (lossDeals/totalDeals*100).toFixed(1) + "%)", "", "", ""], true);
  writeRow(["Средняя прибыль/сделку", avgProfit.toFixed(3)+" ₮", "", "Общий оборот ₽", totalVol.toLocaleString("ru-RU")+" ₽", "", "", ""], false);
  writeRow(["Средний спред", avgSpread.toFixed(2)+"%", "", "Сделок в журнале", rawData.filter(function(r){return r[1]!=="";}).length, "", "", ""], true);
  blankRow();

  // ═══ БЛОК 2: ПО КОМПАНИЯМ ═══
  writeTitle("🏢 ПО КОМПАНИЯМ-РЕКВИЗИТЧИКАМ", ACCENT, HEADER_FG, 12);
  writeHeaders(["Компания","Сделок","Оборот ₽","Прибыль USDT","Ср. прибыль","Ср. спред %","Доля от прибыли",""], DARK_BG);
  var compMap = {};
  deals.forEach(function(d) {
    if (!compMap[d.company]) compMap[d.company] = {cnt:0, profit:0, spread:0, vol:0};
    compMap[d.company].cnt++; compMap[d.company].profit+=d.profit;
    compMap[d.company].spread+=d.spread; compMap[d.company].vol+=d.rub;
  });
  var compArr = Object.keys(compMap).map(function(n){ return {name:n, d:compMap[n]}; })
    .sort(function(a,b){ return b.d.profit - a.d.profit; });
  compArr.forEach(function(c, i) {
    var share = totalProfit !== 0 ? (c.d.profit/totalProfit*100).toFixed(1)+"%" : "—";
    writeRow([c.name, c.d.cnt, c.d.vol.toLocaleString("ru-RU")+" ₽",
      c.d.profit.toFixed(2)+" ₮", (c.d.profit/c.d.cnt).toFixed(3)+" ₮",
      (c.d.spread/c.d.cnt).toFixed(2)+"%", share, ""], i%2===1);
  });
  writeRow(["ИТОГО", totalDeals, totalVol.toLocaleString("ru-RU")+" ₽",
    totalProfit.toFixed(2)+" ₮", avgProfit.toFixed(3)+" ₮",
    avgSpread.toFixed(2)+"%", "100%", ""], false, true);
  blankRow();

  // ═══ БЛОК 3: ПО РАЗМЕРУ ═══
  writeTitle("📏 ПО РАЗМЕРУ СДЕЛКИ", ACCENT, HEADER_FG, 12);
  writeHeaders(["Диапазон ₽","Сделок","Оборот ₽","Прибыль USDT","Ср. прибыль","% от всех","",""], DARK_BG);
  var buckets = [
    {name:"< 1 000 ₽",    min:0,      max:1000},
    {name:"1 000–5 000 ₽", min:1000,   max:5000},
    {name:"5 000–10 000 ₽",min:5000,   max:10000},
    {name:"10 000–30 000 ₽",min:10000, max:30000},
    {name:"30 000–100 000 ₽",min:30000,max:100000},
    {name:"> 100 000 ₽",  min:100000, max:Infinity}
  ];
  buckets.forEach(function(b) { b.cnt=0; b.profit=0; b.vol=0; });
  deals.forEach(function(d) {
    for (var bi = 0; bi < buckets.length; bi++) {
      if (d.rub >= buckets[bi].min && d.rub < buckets[bi].max) {
        buckets[bi].cnt++; buckets[bi].profit+=d.profit; buckets[bi].vol+=d.rub; break;
      }
    }
  });
  buckets.forEach(function(b, i) {
    if (b.cnt === 0) return;
    writeRow([b.name, b.cnt, b.vol.toLocaleString("ru-RU")+" ₽",
      b.profit.toFixed(2)+" ₮", b.cnt>0?(b.profit/b.cnt).toFixed(3)+" ₮":"—",
      (b.cnt/totalDeals*100).toFixed(1)+"%","",""], i%2===1);
  });
  blankRow();

  // ═══ БЛОК 4: ПО ВРЕМЕНИ ═══
  writeTitle("⏰ ПО ВРЕМЕНИ СУТОК", ACCENT, HEADER_FG, 12);
  writeHeaders(["Период","Часы","Сделок","Прибыль USDT","Ср. прибыль","% от всех","",""], DARK_BG);
  var periods = [
    {name:"Ночь",   range:"00–06", h0:0,  h1:6},
    {name:"Утро",   range:"06–12", h0:6,  h1:12},
    {name:"День",   range:"12–18", h0:12, h1:18},
    {name:"Вечер",  range:"18–24", h0:18, h1:24}
  ];
  periods.forEach(function(p) { p.cnt=0; p.profit=0; });
  deals.forEach(function(d) {
    for (var pi=0; pi<periods.length; pi++) {
      if (d.hour >= periods[pi].h0 && d.hour < periods[pi].h1) {
        periods[pi].cnt++; periods[pi].profit+=d.profit; break;
      }
    }
  });
  periods.forEach(function(p, i) {
    writeRow([p.name, p.range, p.cnt, p.cnt>0?p.profit.toFixed(2)+" ₮":"—",
      p.cnt>0?(p.profit/p.cnt).toFixed(3)+" ₮":"—",
      (p.cnt/totalDeals*100).toFixed(1)+"%","",""], i%2===1);
  });
  blankRow();

  // ═══ БЛОК 5: ПО ИСПОЛНИТЕЛЯМ ═══
  writeTitle("👤 ПО ИСПОЛНИТЕЛЯМ", ACCENT, HEADER_FG, 12);
  writeHeaders(["Исполнитель","Сделок","Оборот ₽","Прибыль USDT","Ср. прибыль","% прибыльных","Заморозок",""], DARK_BG);
  var empMap2 = {};
  deals.forEach(function(d) {
    if (!empMap2[d.executor]) empMap2[d.executor] = {cnt:0, profit:0, profCnt:0, spread:0, freeze:0, vol:0};
    empMap2[d.executor].cnt++; empMap2[d.executor].profit+=d.profit;
    empMap2[d.executor].vol+=d.rub;
    if (d.profit>0) empMap2[d.executor].profCnt++;
    empMap2[d.executor].spread+=d.spread;
    if (d.freeze==="да"||d.freeze==="yes"||d.freeze==="1") empMap2[d.executor].freeze++;
  });
  var empArr2 = Object.keys(empMap2).map(function(n){ return {name:n, d:empMap2[n]}; })
    .sort(function(a,b){ return b.d.profit - a.d.profit; });
  empArr2.forEach(function(e, i) {
    writeRow([e.name, e.d.cnt, e.d.vol.toLocaleString("ru-RU")+" ₽",
      e.d.profit.toFixed(2)+" ₮", (e.d.profit/e.d.cnt).toFixed(3)+" ₮",
      (e.d.profCnt/e.d.cnt*100).toFixed(0)+"%", e.d.freeze, ""], i%2===1);
  });
  blankRow();

  // ═══ БЛОК 6: ЗАМОРОЗКИ ═══
  writeTitle("🧊 АНАЛИЗ ЗАМОРОЗОК", ACCENT, HEADER_FG, 12);
  writeHeaders(["","Сделок","Прибыль USDT","Ср. прибыль","Ср. спред %","Доля от всех","",""], DARK_BG);
  var frzY = {cnt:0,profit:0,spread:0}, frzN = {cnt:0,profit:0,spread:0};
  deals.forEach(function(d) {
    var f = (d.freeze==="да"||d.freeze==="yes"||d.freeze==="1") ? frzY : frzN;
    f.cnt++; f.profit+=d.profit; f.spread+=d.spread;
  });
  writeRow(["С заморозкой",frzY.cnt,frzY.profit.toFixed(2)+" ₮",
    frzY.cnt>0?(frzY.profit/frzY.cnt).toFixed(3)+" ₮":"—",
    frzY.cnt>0?(frzY.spread/frzY.cnt).toFixed(2)+"%":"—",
    (frzY.cnt/totalDeals*100).toFixed(1)+"%","",""], false);
  writeRow(["Без заморозки",frzN.cnt,frzN.profit.toFixed(2)+" ₮",
    frzN.cnt>0?(frzN.profit/frzN.cnt).toFixed(3)+" ₮":"—",
    frzN.cnt>0?(frzN.spread/frzN.cnt).toFixed(2)+"%":"—",
    (frzN.cnt/totalDeals*100).toFixed(1)+"%","",""], true);
  blankRow();

  // ═══ БЛОК 7: ТОП СДЕЛОК ═══
  writeTitle("🏆 ТОП-10 САМЫХ ПРИБЫЛЬНЫХ СДЕЛОК", ACCENT, HEADER_FG, 12);
  writeHeaders(["Дата/время","Компания","₽ получили","Прибыль USDT","Спред %","Исполнитель","Заморозка",""], DARK_BG);
  var sorted = deals.slice().sort(function(a,b){ return b.profit - a.profit; });
  for (var i = 0; i < Math.min(10, sorted.length); i++) {
    var d = sorted[i];
    writeRow([d.dt, d.company, d.rub.toLocaleString("ru-RU")+" ₽",
      d.profit.toFixed(3)+" ₮", d.spread.toFixed(2)+"%", d.executor, d.freeze, ""], i%2===1);
  }
  blankRow();
  writeTitle("📉 ТОП-10 УБЫТОЧНЫХ / НУЛЕВЫХ СДЕЛОК", "#5a1a1a", HEADER_FG, 12);
  writeHeaders(["Дата/время","Компания","₽ получили","Прибыль USDT","Спред %","Исполнитель","Заморозка",""], DARK_BG);
  var lossSorted = deals.slice().sort(function(a,b){ return a.profit - b.profit; });
  for (var i = 0; i < Math.min(10, lossSorted.length); i++) {
    var d = lossSorted[i];
    writeRow([d.dt, d.company, d.rub.toLocaleString("ru-RU")+" ₽",
      d.profit.toFixed(3)+" ₮", d.spread.toFixed(2)+"%", d.executor, d.freeze, ""], i%2===1);
  }
  blankRow();

  // ═══ БЛОК 8: РЕКОМЕНДАЦИИ ═══
  writeTitle("💡 РЕКОМЕНДАЦИИ НА ОСНОВЕ ДАННЫХ", "#1a3a2a", HEADER_FG, 12);
  var recs = [];
  if (compArr.length > 0) {
    var best = compArr[0], worst = compArr[compArr.length-1];
    var bestAvg = best.d.profit/best.d.cnt, worstAvg = worst.d.profit/worst.d.cnt;
    recs.push(["✅ Лучшая компания-реквизитчик", best.name+" — средняя прибыль "+bestAvg.toFixed(2)+" ₮/сделку, спред "+(best.d.spread/best.d.cnt).toFixed(1)+"%"]);
    if (worst.name !== best.name && worstAvg < bestAvg*0.5) {
      recs.push(["⚠️ Наименее выгодная компания", worst.name+" — средняя прибыль всего "+worstAvg.toFixed(2)+" ₮/сделку. Рекомендуем ограничить объём."]);
    }
  }
  var bestBucket = null, bestBucketAvg = -999;
  for (var b=0; b<buckets.length; b++) {
    if (buckets[b].cnt >= 5) {
      var avg = buckets[b].profit/buckets[b].cnt;
      if (avg > bestBucketAvg) { bestBucketAvg=avg; bestBucket=buckets[b]; }
    }
  }
  if (bestBucket) {
    recs.push(["✅ Оптимальный размер платежа", bestBucket.name+" — средняя прибыль "+bestBucketAvg.toFixed(2)+" ₮/сделку. Таких сделок "+bestBucket.cnt+" шт."]);
  }
  var bestPeriod=null, bestPeriodAvg=-999, worstPeriod=null, worstPeriodAvg=999;
  for (var p=0; p<periods.length; p++) {
    if (periods[p].cnt >= 5) {
      var pavg = periods[p].profit/periods[p].cnt;
      if (pavg > bestPeriodAvg) { bestPeriodAvg=pavg; bestPeriod=periods[p]; }
      if (pavg < worstPeriodAvg) { worstPeriodAvg=pavg; worstPeriod=periods[p]; }
    }
  }
  if (bestPeriod) {
    recs.push(["⏰ Лучшее время для сделок", bestPeriod.name+" — средняя прибыль "+bestPeriodAvg.toFixed(2)+" ₮/сделку ("+bestPeriod.cnt+" сделок)."]);
  }
  if (frzY.cnt>0&&frzN.cnt>0) {
    var fyAvg=frzY.profit/frzY.cnt, fnAvg=frzN.profit/frzN.cnt;
    var diff = fnAvg !== 0 ? ((fnAvg-fyAvg)/Math.abs(fnAvg)*100) : 0;
    if (diff > 10) {
      recs.push(["🧊 Избегайте заморозок", "Сделки без заморозки приносят на "+diff.toFixed(0)+"% больше ("+fnAvg.toFixed(2)+" vs "+fyAvg.toFixed(2)+" ₮). Заморозок: "+frzY.cnt+"."]);
    }
  }
  if (empArr2.length > 0) {
    var topEmp = empArr2[0];
    recs.push(["🏆 Лучший исполнитель", topEmp.name+" — "+topEmp.d.cnt+" сделок, прибыль "+topEmp.d.profit.toFixed(2)+" ₮, средняя "+(topEmp.d.profit/topEmp.d.cnt).toFixed(2)+" ₮/сделку."]);
  }
  var profPct = profDeals/totalDeals*100;
  recs.push(["📊 Общий вывод",
    profPct>=90 ? "Отличная неделя: "+profPct.toFixed(0)+"% сделок прибыльны. Средняя прибыль "+avgProfit.toFixed(2)+" ₮/сделку."
    : profPct>=70 ? profPct.toFixed(0)+"% сделок прибыльны. Есть потенциал роста — фокус на лучших компаниях."
    : "Только "+profPct.toFixed(0)+"% сделок прибыльны. Рекомендуем пересмотреть выбор компаний и временные слоты."
  ]);

  writeHeaders(["Рекомендация","Детали","","","","","",""], "#1a3a2a");
  for (var ri=0; ri<recs.length; ri++) {
    var recRow = sheet.getRange(row, 1, 1, 8);
    recRow.setValues([[recs[ri][0], recs[ri][1], "", "", "", "", "", ""]]);
    recRow.setBackground(ri%2===0 ? "#f0faf4" : "#e0f5e9");
    recRow.setFontSize(11);
    sheet.getRange(row, 1).setFontWeight("bold");
    sheet.setRowHeight(row, 28);
    sheet.getRange(row, 2, 1, 7).merge();
    row++;
  }
  blankRow();

  // ═══ БЛОК 9: СРАВНЕНИЕ С ПРОШЛОЙ НЕДЕЛЕЙ ═══
  writeTitle("📅 СРАВНЕНИЕ С ПРОШЛОЙ НЕДЕЛЕЙ", "#2a1a4a", HEADER_FG, 12);

  var prevWeekData = null;
  if (archive) {
    // Читаем архив — ищем две последние записи
    var archLastRow = archive.getLastRow();
    if (archLastRow >= 7) {
      // Ищем последнюю заполненную строку в архиве (столбец C — неделя)
      var lastArchRow = -1;
      for (var ar = 6; ar <= archLastRow; ar++) {
        if (archive.getRange(ar, 3).getValue() !== "") { lastArchRow = ar; }
      }
      if (lastArchRow > 0) {
        var pWeek = archive.getRange(lastArchRow, 3).getValue();
        var pDeals = parseFloat(archive.getRange(lastArchRow, 4).getValue()) || 0;
        var pTurnover = parseFloat(archive.getRange(lastArchRow, 5).getValue()) || 0;
        var pProfitU = parseFloat(archive.getRange(lastArchRow, 6).getValue()) || 0;
        var pPaid = parseFloat(archive.getRange(lastArchRow, 8).getValue()) || 0;
        var pAvg = parseFloat(archive.getRange(lastArchRow, 10).getValue()) || 0;
        var pBest = archive.getRange(lastArchRow, 11).getValue() || "—";
        prevWeekData = {week:pWeek, deals:pDeals, vol:pTurnover, profitU:pProfitU, paid:pPaid, avg:pAvg, best:pBest};
      }
    }
  }

  var curWeek = settings ? settings.getRange("D5").getValue() : "Текущая";

  if (!prevWeekData) {
    writeHeaders(["Показатель","Текущая неделя","Прошлая неделя","Изменение","","","",""], DARK_BG);
    writeRow(["Данные прошлой недели", "—", "Нет данных в Архиве", "После первого closeWeek() здесь появится сравнение", "", "", "", ""], false);
  } else {
    function delta(cur, prev) {
      if (!prev || prev === 0) return "+0%";
      var pct = ((cur - prev) / Math.abs(prev) * 100);
      return (pct >= 0 ? "+" : "") + pct.toFixed(1) + "%";
    }
    writeHeaders(["Показатель", ""+curWeek, ""+prevWeekData.week, "Изменение","","","",""], DARK_BG);
    writeRow(["Сделок завершено", totalDeals, prevWeekData.deals, delta(totalDeals, prevWeekData.deals),"","","",""], false);
    writeRow(["Оборот ₽", totalVol.toLocaleString("ru-RU")+" ₽", (prevWeekData.vol||0).toLocaleString("ru-RU")+" ₽", delta(totalVol, prevWeekData.vol),"","","",""], true);
    writeRow(["Прибыль USDT", totalProfit.toFixed(2)+" ₮", (prevWeekData.profitU||0).toFixed(2)+" ₮", delta(totalProfit, prevWeekData.profitU),"","","",""], false);
    writeRow(["Средняя прибыль/сд", avgProfit.toFixed(3)+" ₮", (prevWeekData.avg||0).toFixed(3)+" ₮", delta(avgProfit, prevWeekData.avg),"","","",""], true);
    writeRow(["% прибыльных сделок", (profDeals/totalDeals*100).toFixed(1)+"%", "—", "—","","","",""], false);
    writeRow(["Лучший сотрудник", empArr2.length>0?empArr2[0].name:"—", prevWeekData.best, "—","","","",""], true);
  }

  // Форматирование
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 140);
  sheet.setFrozenRows(0);

  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(ss.getNumSheets());

  SpreadsheetApp.getUi().alert("✓ Аналитика обновлена!\n\nСделок проанализировано: " + totalDeals + "\n\n" +
    "БЛОК 9 — сравнение с прошлой неделей:\n" +
    (prevWeekData ? "Найдена неделя «" + prevWeekData.week + "» в Архиве." : "Данные прошлой недели появятся после первого «Закрыть неделю»."));
}
