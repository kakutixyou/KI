function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📅 予定表自動作成')
    .addSubMenu(ui.createMenu('🏢 【就労移行】メニュー')
      .addItem('🚀 翌月分を全自動作成 (データ＆描画)', 'auto_Shuro')
      .addItem('🎯 指定月を全自動作成 (データ＆描画)', 'specify_Shuro')
      .addItem('🗓️ 1年分一括データ生成＆カレンダー全描画', 'year_Shuro')
      .addSeparator()
      .addItem('📝 (手動) 予定データを生成', 'manualData_Shuro')
      .addItem('🖨️ (手動) 印刷シートを描画', 'manualPrint_Shuro')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('🏠 【生活訓練】メニュー')
      .addItem('🚀 翌月分を全自動作成 (データ＆描画)', 'auto_Seikatsu')
      .addItem('🎯 指定月を全自動作成 (データ＆描画)', 'specify_Seikatsu')
      .addItem('🗓 1年分一括データ生成＆カレンダー全描画', 'year_Seikatsu')
      .addSeparator()
      .addItem('📝 (手動) 予定データを生成', 'manualData_Seikatsu')
      .addItem('🖨️ (手動) 印刷シートを描画', 'manualPrint_Seikatsu')
    )
    .addToUi();
}

// ====================================================
// メニューから呼び出される仲介関数
// ====================================================
function auto_Shuro() { autoCreateNextMonth('就労'); }
function auto_Seikatsu() { autoCreateNextMonth('生活'); }

function specify_Shuro() { createSpecifiedMonth('就労'); }
function specify_Seikatsu() { createSpecifiedMonth('生活'); }

function year_Shuro() { createOneYearData('就労'); }
function year_Seikatsu() { createOneYearData('生活'); }

function manualData_Shuro() { promptForMonthData('就労'); }
function manualData_Seikatsu() { promptForMonthData('生活'); }

function manualPrint_Shuro() { promptForPrintLayout('就労'); }
function manualPrint_Seikatsu() { promptForPrintLayout('生活'); }


// ====================================================
// 🎌 祝日データを一度だけ取得して使い回す
// ====================================================
let globalHolidays = null;
function getHolidays() {
  if (globalHolidays !== null) return globalHolidays;
  globalHolidays = [];
  try {
    const response = UrlFetchApp.fetch("https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv");
    const csvData = Utilities.parseCsv(response.getContentText("Shift_JIS"));
    for (let i = 1; i < csvData.length; i++) {
      if (csvData[i][0]) globalHolidays.push(Utilities.formatDate(new Date(csvData[i][0]), "JST", "yyyy/MM/dd"));
    }
  } catch(e) { }
  return globalHolidays;
}

// ====================================================
// 🌟 翌月分を全自動作成
// ====================================================
function autoCreateNextMonth(type) {
  const ui = SpreadsheetApp.getUi();
  const today = new Date();
  
  let year = today.getFullYear();
  let month = today.getMonth() + 2; 
  if (month > 12) {
    year += 1;
    month = 1;
  }

  const response = ui.alert(
    `【${type}】翌月分の全自動作成`,
    `【 ${year}年${month}月 】の「${type}」予定データ生成と、印刷用シートの作成を一気に実行しますか？`,
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    generateSchedule(year, month, type, true);
    drawPrintLayout(year, month, type, true, 1);
    ui.alert('✨ 完了しました！', `${year}年${month}月の「${type}」予定生成とカレンダー描画が完了しました。`, ui.ButtonSet.OK);
  }
}

// ====================================================
// 🎯 指定月を全自動作成
// ====================================================
function createSpecifiedMonth(type) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    `【${type}】指定月の全自動作成`,
    `作成したい年月を「YYYY/MM」の形式で入力してください\n（例: 2026/12）\n\n※予定データ作成と印刷用シート描画を同時に行います。`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    const input = response.getResponseText();
    const parts = input.split('/');
    if (parts.length === 2) {
      let y = parseInt(parts[0], 10);
      let m = parseInt(parts[1], 10);
      generateSchedule(y, m, type, true);
      drawPrintLayout(y, m, type, true, 1);
      ui.alert('✨ 完了しました！', `${y}年${m}月の「${type}」予定生成とカレンダー描画が完了しました。`, ui.ButtonSet.OK);
    } else {
      ui.alert('エラー', 'フォーマットが正しくありません。「2026/12」のように入力してください。', ui.ButtonSet.OK);
    }
  }
}

// ====================================================
// 🗓 1年分一括データ生成 ＆ カレンダー全描画
// ====================================================
function createOneYearData(type) {
  const ui = SpreadsheetApp.getUi();
  const today = new Date();
  let startYear = today.getFullYear();
  let startMonth = today.getMonth() + 1; 
  
  let endMonth = startMonth + 11;
  let endYear = startYear + Math.floor((endMonth - 1) / 12);
  let endMonthDisp = ((endMonth - 1) % 12) + 1;

  const response = ui.alert(
    `【${type}】1年分の全自動作成`,
    `現在の月から12ヶ月分（${startYear}年${startMonth}月 〜 ${endYear}年${endMonthDisp}月）の\n「${type}」予定データと印刷用カレンダーを一気に作成しますか？`,
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    getHolidays(); 
    
    let nextRow = 1;
    for (let i = 0; i < 12; i++) {
      let targetMonth = startMonth + i;
      let y = startYear + Math.floor((targetMonth - 1) / 12);
      let m = ((targetMonth - 1) % 12) + 1;
      
      generateSchedule(y, m, type, true); 
      nextRow = drawPrintLayout(y, m, type, true, nextRow);
    }
    ui.alert(
      '✨ 1年分の作成完了！', 
      `「${type}_予定データ」に12ヶ月分の予定を作成し、\n「${type}_印刷用」に1年分のカレンダーを全て描画しました！`, 
      ui.ButtonSet.OK
    );
  }
}

// ====================================================
// ① 予定データを生成する機能 (ベース処理)
// ====================================================
function promptForMonthData(type) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(`【${type}】(手動)予定データ生成`, '予定データを作成したい年月を「YYYY/MM」の形式で入力してください\n（例: 2026/11）', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const input = response.getResponseText();
    const parts = input.split('/');
    if (parts.length === 2) {
      generateSchedule(parseInt(parts[0], 10), parseInt(parts[1], 10), type, false);
    } else {
      ui.alert('フォーマットが正しくありません。「2026/11」のように入力してください。');
    }
  }
}

function generateSchedule(year, month, type, isAuto = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ruleSheetName = `${type}_ルール設定`;
  const dataSheetName = `${type}_予定データ`;
  
  const ruleSheet = ss.getSheetByName(ruleSheetName);
  const dataSheet = ss.getSheetByName(dataSheetName);

  if (!ruleSheet || !dataSheet) {
    SpreadsheetApp.getUi().alert(`エラー：「${ruleSheetName}」または「${dataSheetName}」シートが見つかりません。`);
    return;
  }

  let holidays = getHolidays(); 

  const ruleData = ruleSheet.getDataRange().getValues();
  const rules = [];
  
  for (let i = 1; i < ruleData.length; i++) {
    if (!ruleData[i][0]) continue;
    
    let mVal = ruleData[i][0];
    let wVal = ruleData[i][1];

    // 【超強化】スプレッドシートが勝手に日付形式(Date)に変換していた場合への対策
    if (mVal instanceof Date) mVal = mVal.getMonth() + 1;
    if (wVal instanceof Date) wVal = wVal.getDate();

    // 12/26 や 12月26日 のように入力されていた場合、不要な文字を切り落として抽出
    let mStr = String(mVal).trim();
    if (mStr !== '毎月') {
      mStr = mStr.split(',').map(m => {
        let t = m.trim();
        if (t.indexOf('/') !== -1) t = t.split('/')[0];
        if (t.indexOf('月') !== -1) t = t.split('月')[0];
        return parseInt(t.replace(/[^0-9]/g, ''), 10);
      }).join(',');
    }

    rules.push({
      months: mStr,
      weeks: String(wVal).trim(), 
      dayName: String(ruleData[i][2]).trim(),
      timeSlot: String(ruleData[i][3]).trim(),
      eventName: String(ruleData[i][4]).trim()
    });
  }

  const timeMapping = {
    'AM': { start: '10:00', end: '12:00' },
    'PM': { start: '13:00', end: '15:00' },
    'PM2': { start: '15:00', end: '17:00' },
    'ALL_PM': { start: '13:00', end: '17:00' }
  };

  const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
  const daysInMonth = new Date(year, month, 0).getDate();
  const newRows = [];
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');

  const headers = dataSheet.getDataRange().getValues()[0] || [];
  const colId = headers.indexOf('企画ID');
  const colName = headers.indexOf('企画名');
  const colDate = headers.indexOf('開催日');
  const colStart = headers.indexOf('開始時間');
  const colEnd = headers.indexOf('終了時間');
  const colStatus = headers.indexOf('承認ステータス');
  const colCreated = headers.indexOf('登録日時');

  const existingSlots = new Set();
  let maxId = 0;
  const existingData = dataSheet.getDataRange().getValues();
  for (let i = 1; i < existingData.length; i++) {
    if(colId !== -1) {
      const idStr = String(existingData[i][colId]);
      if (idStr && idStr.startsWith('T')) {
        const num = parseInt(idStr.replace('T', ''), 10);
        if (num > maxId) maxId = num;
      }
    }
    
    let d = (colDate !== -1) ? existingData[i][colDate] : "";
    let t = (colStart !== -1) ? existingData[i][colStart] : "";
    if (d && t) {
      let dStr = (d instanceof Date) ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd') : String(d);
      let tStr = (t instanceof Date) ? Utilities.formatDate(t, Session.getScriptTimeZone(), 'HH:mm') : String(t).trim();
      existingSlots.add(dStr + '_' + tStr); // 既に予定がある枠を記録
    }
  }

  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(year, month - 1, d);
    const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');

    if (holidays.indexOf(dateStr) !== -1) continue;

    const dayName = dayNames[date.getDay()];
    const weekNum = Math.ceil(d / 7);

    let dailySlots = { 'AM': null, 'PM': null, 'PM2': null, 'ALL_PM': null };
    
    rules.filter(r => {
      const monthMatch = (r.months === '毎月' || String(r.months).split(',').map(s=>s.trim()).includes(String(month)));
      
      let isMatch = false;
      if (r.weeks === '毎週') {
        isMatch = (r.dayName === dayName);
      } else {
        const weeksOrDays = String(r.weeks).split(',').map(s=>s.trim());
        for (let wd of weeksOrDays) {
          // 【超強化】「12/26」や「26日」から数字「26」だけを正確に抜き出す
          let text = wd;
          if (text.indexOf('/') !== -1) text = text.split('/').pop();
          if (text.indexOf('月') !== -1) text = text.split('月').pop();
          let num = parseInt(text.replace(/[^0-9]/g, ''), 10); 
          
          if (!isNaN(num) && num >= 6) { 
            // 6以上の数字は「日にち指定」
            if (num === d) {
              isMatch = true; 
              break;
            }
          } else if (!isNaN(num)) {
            // 5以下の数字は「第○週指定」
            if (num === weekNum && r.dayName === dayName) {
              isMatch = true;
              break;
            }
          }
        }
      }
      return monthMatch && isMatch;
    }).forEach(r => { 
      dailySlots[r.timeSlot] = r.eventName; 
    });

    let hasOutdoor = false;
    for (let k in dailySlots) {
      if (dailySlots[k] && String(dailySlots[k]).indexOf('野外') !== -1) hasOutdoor = true;
    }
    if (hasOutdoor || dailySlots['ALL_PM']) dailySlots['AM'] = null; 
    if (dailySlots['ALL_PM']) {
      dailySlots['PM'] = null;
      dailySlots['PM2'] = null;
    }

    ['AM', 'PM', 'PM2', 'ALL_PM'].forEach(slot => {
      if (dailySlots[slot]) {
        let startTime = timeMapping[slot].start;
        // すでに予定が存在する場合は上書きしない（スキップする）
        if (existingSlots.has(dateStr + '_' + startTime)) return; 

        maxId++;
        let rowData = new Array(headers.length).fill(''); 
        
        if(colId !== -1) rowData[colId] = 'T' + ('000' + maxId).slice(-3);
        if(colName !== -1) rowData[colName] = dailySlots[slot];
        if(colDate !== -1) rowData[colDate] = dateStr;
        if(colStart !== -1) rowData[colStart] = startTime;
        if(colEnd !== -1) rowData[colEnd] = timeMapping[slot].end;
        if(colStatus !== -1) rowData[colStatus] = '予定';
        if(colCreated !== -1) rowData[colCreated] = nowStr;

        newRows.push(rowData);
      }
    });
  }

  if (newRows.length > 0) {
    dataSheet.getRange(dataSheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
    SpreadsheetApp.flush(); 
    if (!isAuto) SpreadsheetApp.getUi().alert(`✅ ${year}年${month}月の「${type}」予定データを作成しました！`);
  } else if (!isAuto) {
    SpreadsheetApp.getUi().alert(`ℹ️ ${year}年${month}月の「${type}」予定は既に作成されているか、追加する予定がありません。`);
  }
}

// ====================================================
// ② 印刷用シートを描画する機能 (ベース処理・縦積み対応)
// ====================================================
function promptForPrintLayout(type) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(`【${type}】(手動)印刷用シート作成`, '印刷したい年月を「YYYY/MM」の形式で入力してください\n（例: 2026/11）', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const input = response.getResponseText();
    const parts = input.split('/');
    if (parts.length === 2) {
      drawPrintLayout(parseInt(parts[0], 10), parseInt(parts[1], 10), type, false, 1);
    } else {
      ui.alert('フォーマットが正しくありません。「2026/11」のように入力してください。');
    }
  }
}

function drawPrintLayout(year, month, type, isAuto = false, startRow = 1) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const printSheetName = `${type}_印刷用`;
  const dataSheetName = `${type}_予定データ`;
  const ruleSheetName = `${type}_ルール設定`;

  let sheet = ss.getSheetByName(printSheetName);
  if (!sheet) sheet = ss.insertSheet(printSheetName);
  
  if (startRow === 1) {
    sheet.clear(); 
    sheet.setColumnWidth(1, 40); 
    for(let i=2; i<=7; i++) sheet.setColumnWidth(i, 130); 
  }
  
  sheet.getRange(startRow, 1, 1, 7).merge();
  sheet.getRange(startRow, 1).setValue(`${year}年 ${month}月 【${type}】 カレンダー`).setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");

  sheet.getRange(startRow + 1, 1, 1, 7).setValues([['日', '月', '火', '水', '木', '金', '土']]);
  sheet.getRange(startRow + 1, 1, 1, 7).setBackground('#efefef').setHorizontalAlignment('center');
  
  let holidays = getHolidays(); 

  const dataSheet = ss.getSheetByName(dataSheetName);
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert(`エラー：「${dataSheetName}」シートが見つかりません。`);
    return startRow;
  }
  const dataValues = dataSheet.getDataRange().getValues();
  const headers = dataValues[0];
  
  const col = {
    name: headers.indexOf('企画名'),
    date: headers.indexOf('開催日'),
    time: headers.indexOf('開始時間')
  };

  if (col.name === -1 || col.date === -1 || col.time === -1) {
    SpreadsheetApp.getUi().alert(`エラー：${dataSheetName}シートに「企画名」「開催日」「開始時間」の列が必要です。`);
    return startRow;
  }

  let scheduleMap = {}; 
  let uniqueDrawSet = new Set(); 

  for(let i = 1; i < dataValues.length; i++) {
    let rowData = dataValues[i];
    let eventName = rowData[col.name];
    let d = rowData[col.date];
    let time = rowData[col.time]; 
    
    if(!d || !time || !eventName) continue;
    
    let dateStr = (d instanceof Date) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd") : String(d);
    let timeStr = (time instanceof Date) ? Utilities.formatDate(time, Session.getScriptTimeZone(), "HH:mm") : String(time).trim();
    
    let uniqueKey = dateStr + '_' + timeStr + '_' + eventName;
    if (uniqueDrawSet.has(uniqueKey)) continue;
    uniqueDrawSet.add(uniqueKey);

    let eventObj = { name: eventName };

    if (!scheduleMap[dateStr]) scheduleMap[dateStr] = { 'AM': [], 'PM1': [], 'PM2': [] };
    if (timeStr.indexOf('10:') !== -1) scheduleMap[dateStr]['AM'].push(eventObj);
    if (timeStr.indexOf('13:') !== -1) scheduleMap[dateStr]['PM1'].push(eventObj);
    if (timeStr.indexOf('15:') !== -1) scheduleMap[dateStr]['PM2'].push(eventObj);
  }

  let startDate = new Date(year, month - 1, 1);
  let startDayOfWeek = startDate.getDay();
  let currentDate = new Date(startDate);
  currentDate.setDate(currentDate.getDate() - startDayOfWeek); 
  
  let row = startRow + 2; 
  for (let w = 1; w <= 6; w++) { 
    if (w > 1 && currentDate.getMonth() !== (month - 1) && currentDate.getFullYear() >= year) break;
    
    sheet.setRowHeight(row, 20);      
    sheet.setRowHeight(row + 1, 35);  
    sheet.setRowHeight(row + 2, 35);  
    sheet.setRowHeight(row + 3, 35);  
    
    sheet.getRange(row + 1, 1).setValue(''); 
    sheet.getRange(row + 2, 1).setValue('午後①');
    sheet.getRange(row + 3, 1).setValue('午後②');
    sheet.getRange(row + 1, 1, 3, 1).setHorizontalAlignment('center').setVerticalAlignment('middle');
    
    sheet.getRange(row, 1, 4, 7).setBorder(true, true, true, true, true, false); 

    for (let c = 1; c <= 7; c++) {
      let dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
      
      sheet.getRange(row, c).setValue(currentDate.getDate());
      sheet.getRange(row, c).setHorizontalAlignment('left');
      
      let dayPlan = scheduleMap[dateStr] || { 'AM': [], 'PM1': [], 'PM2': [] };
      let allEvents = [...dayPlan['AM'], ...dayPlan['PM1'], ...dayPlan['PM2']];
      
      let isGray = false;
      let isHolidayOrSunday = (c === 1 || holidays.indexOf(dateStr) !== -1);
      
      if (isHolidayOrSunday) {
        isGray = true; 
      } else if (c === 7 && allEvents.length === 0) {
        isGray = true; 
      }

      let amCell = sheet.getRange(row + 1, c);
      let pm1Cell = sheet.getRange(row + 2, c);
      let pm2Cell = sheet.getRange(row + 3, c);

      if (isGray) {
        sheet.getRange(row, c, 4, 1).setBackground('#f3f3f3');
        amCell.setValue('');
        pm1Cell.setValue('');
        pm2Cell.setValue('');
      } else {
        sheet.getRange(row, c, 4, 1).setBackground('#ffffff');
        
        let amNames = dayPlan['AM'].map(e => e.name).join('\n');
        let pm1Names = dayPlan['PM1'].map(e => e.name).join('\n');
        let pm2Names = dayPlan['PM2'].map(e => e.name).join('\n');
        
        amCell.setValue(amNames);
        pm1Cell.setValue(pm1Names);
        pm2Cell.setValue(pm2Names);
        
        [amCell, pm1Cell, pm2Cell].forEach(cell => {
          cell.setHorizontalAlignment('left').setVerticalAlignment('top'); 
        });
      }
      
      amCell.setBorder(null, null, true, null, null, null); 
      currentDate.setDate(currentDate.getDate() + 1); 
    }
    
    sheet.getRange(row + 1, 2, 3, 6).setWrap(true);
    row += 4; 
  }
  
  let notesRow = row;
  sheet.getRange(notesRow, 1, 1, 7).merge(); 
  sheet.getRange(notesRow, 1).setValue(`【${type} お知らせ・運用ルール】`);
  sheet.getRange(notesRow, 1).setBackground('#444444').setFontColor('#ffffff').setHorizontalAlignment('center');
  
  let contentRow = notesRow + 1;
  sheet.getRange(contentRow, 1, 4, 7).merge(); 
  sheet.getRange(contentRow, 1).setBorder(true, true, true, true, false, false);
  sheet.getRange(contentRow, 1).setVerticalAlignment('top').setWrap(true);

  const ruleSheet = ss.getSheetByName(ruleSheetName);
  if (ruleSheet) {
    let noticeText = ruleSheet.getRange('H1').getValue();
    if(noticeText) {
      sheet.getRange(contentRow, 1).setValue(noticeText);
    } else {
      sheet.getRange(contentRow, 1).setValue(`（${ruleSheetName}シートのH1セルにお知らせを書くと、ここに自動で表示されます）`);
    }
  }
  
  return contentRow + 4; 
}
