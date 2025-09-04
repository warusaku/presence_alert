/**
 * 出勤簿シート用GASエンドポイント
 * 2025-07-10
 * 2025-09-04 施設情報連携対応版
 *
 * このスクリプトは出勤簿用の別スプレッドシートで動作します
 * Omada Webhook受信用GASから送信されたデータを受け取り、記録します
 */

function doPost(e) {
  // 並行実行の衝突回避（最大30秒待機）
  const lock = LockService.getDocumentLock();
  try { lock.waitLock(30 * 1000); } catch (err) { console.error('Lock acquisition failed:', err); }

  try {
    /* ---------- 1. 受信データ解析 ---------- */
    const raw = (e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const data = _safeJson_(raw);

    if (!data.username || (!data.timestamp && !data.timestampMs) || !data.facilityName) {
      return _error_('必須パラメータ（username, timestamp/timestampMs, facilityName）が不足しています');
    }

    /* ---------- 2. スプレッドシート準備 ---------- */
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 出勤簿シート（メイン）
    const attendanceSheet = ss.getSheetByName('出勤簿') || _createAttendanceSheet_(ss);

    // 生データ保存用シート
    const rawDataSheet = ss.getSheetByName('raw_data') || _createRawDataSheet_(ss);

    // 月次集計シート
    const monthlySummarySheet = ss.getSheetByName('月次集計') || _createMonthlySummarySheet_(ss);

    /* ---------- 3. 生データ記録 ---------- */
    rawDataSheet.appendRow([
      new Date(),
      data.username,
      data.timestampMs || data.timestamp,
      data.state,
      data.name,
      data.devicename,
      data.ipaddr,
      data.MAC,
      data.description,
      data.siteName, // サイト名を追加
      data.facilityName, // 施設表示名を追加
      JSON.stringify(data)
    ]);

    /* ---------- 4. 出勤簿更新 ---------- */
    _updateAttendanceSheet_(attendanceSheet, data);

    /* ---------- 5. 月次集計更新 ---------- */
    _updateMonthlySummary_(monthlySummarySheet, attendanceSheet);

    return _ok_();
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ================================================================= */
/* 出勤簿更新処理 --------------------------------------------------- */
/* ================================================================= */

function _updateAttendanceSheet_(sheet, data) {
  const timestamp = data.timestampMs ? new Date(Number(data.timestampMs)) : _parseTimestamp_(data.timestamp);
  const dateStr = Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyy/MM/dd');
  const timeStr = Utilities.formatDate(timestamp, 'Asia/Tokyo', 'HH:mm:ss');
  const facilityName = data.facilityName;
  // 退勤はgas1の判定文言（"退勤" を含む）に依存させる
  const isDeparture = (data.description && data.description.indexOf('退勤') !== -1);
  const isOnline = data.state === 'ONLINE';
  const OUTING_WINDOW_MS = 8 * 60 * 60 * 1000; // 8時間以内は外出扱い

  // 1. スプレッドシートから全データを一度に読み込む
  const lastRow = sheet.getLastRow();
  const allValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 8).getValues() : [];
  let targetRow = -1;

  // 2. まず同日の記録を後ろから探す
  for (let i = allValues.length - 1; i >= 0; i--) {
    const rowDateStr = Utilities.formatDate(new Date(allValues[i][0]), 'Asia/Tokyo', 'yyyy/MM/dd');
    if (rowDateStr === dateStr && allValues[i][1] === data.username && allValues[i][2] === facilityName) {
      targetRow = i + 2;
      break;
    }
  }

  // 3. 退勤イベントで、かつ同日に記録がない場合、前日の記録を探す
  if (targetRow === -1 && isDeparture) {
    const yesterday = new Date(timestamp.getTime());
    yesterday.setDate(timestamp.getDate() - 1);
    const yesterdayStr = Utilities.formatDate(yesterday, 'Asia/Tokyo', 'yyyy/MM/dd');
    for (let i = allValues.length - 1; i >= 0; i--) {
      const rowData = allValues[i];
      const rowDateStr = Utilities.formatDate(new Date(rowData[0]), 'Asia/Tokyo', 'yyyy/MM/dd');
      const rowDepartureTime = rowData[4]; // 退勤時刻
      if (rowDateStr === yesterdayStr && rowData[1] === data.username && rowData[2] === facilityName && !rowDepartureTime) {
        targetRow = i + 2; // 前日の行をターゲットにする
        break;
      }
    }
  }

  // 4. データの更新または新規追加
  if (targetRow === -1) {
    // 新規行を追加
    const newRow = [
      timestamp, data.username, facilityName,
      isOnline ? timeStr : '',
      isDeparture ? timeStr : '',
      '', '', data.description
    ];
    sheet.appendRow(newRow);
  } else {
    // 既存行を更新
    const arrivalTime = sheet.getRange(targetRow, 4).getValue();
    const departureTime = sheet.getRange(targetRow, 5).getValue();

    // 外出復帰判定: 既に退勤時刻があり、8時間以内にONLINEで戻った
    if (isOnline && departureTime) {
      const dep = departureTime instanceof Date ? departureTime : new Date(sheet.getRange(targetRow, 1).getValue().toDateString() + ' ' + departureTime);
      const diffMs = timestamp.getTime() - dep.getTime();
      if (diffMs >= 0 && diffMs <= OUTING_WINDOW_MS) {
        // 退勤を取り消し、外出として備考に記録
        sheet.getRange(targetRow, 5).clearContent();
        const outingMinutes = Math.round(diffMs / 60000);
        const outingHours = Math.floor(outingMinutes / 60);
        const outingMinsR = outingMinutes % 60;
        const currentNote = sheet.getRange(targetRow, 8).getValue();
        const extra = `[${timeStr}] 外出から復帰（外出${outingHours}時間${outingMinsR}分） [OUTING_MINUTES=${outingMinutes}]`;
        sheet.getRange(targetRow, 8).setValue(currentNote + '\n' + extra);
        // 退勤取消に伴い実働時間は未確定に戻るため、計算はしない
      } else if (isOnline && !arrivalTime) {
        // 8時間超の復帰、かつ出社時刻未設定なら通常通り出社時刻を設定
        sheet.getRange(targetRow, 4).setValue(timeStr);
      }
    } else if (isOnline && !arrivalTime) {
      // 通常の初回出社記録
      sheet.getRange(targetRow, 4).setValue(timeStr);
    } else if (isDeparture) {
      // 退勤記録
      sheet.getRange(targetRow, 5).setValue(timeStr);
      _calculateWorkTime_(sheet, targetRow);
    }

    // 備考追記（常に記録）
    const currentNote2 = sheet.getRange(targetRow, 8).getValue();
    sheet.getRange(targetRow, 8).setValue(currentNote2 + '\n' + `[${timeStr}] ${data.description}`);
  }
}

/* ================================================================= */
/* 実働時間計算 ----------------------------------------------------- */
/* ================================================================= */

function _calculateWorkTime_(sheet, row) {
  const arrivalTime = sheet.getRange(row, 4).getValue(); // D列
  const departureTime = sheet.getRange(row, 5).getValue(); // E列

  if (!arrivalTime || !departureTime) return;

  // 時刻文字列をDateオブジェクトに変換
  const dateCell = sheet.getRange(row, 1).getValue();
  
  // arrivalTimeとdepartureTimeがDateオブジェクトかチェックし、そうでなければ変換
  const arrival = arrivalTime instanceof Date ? arrivalTime : new Date(dateCell.toDateString() + ' ' + arrivalTime);
  const departure = departureTime instanceof Date ? departureTime : new Date(dateCell.toDateString() + ' ' + departureTime);


  // 総滞在時間（ミリ秒）
  const workMillis = departure.getTime() - arrival.getTime();

  // 備考から外出分（分）を抽出
  const noteText = String(sheet.getRange(row, 8).getValue() || '');
  let outingMinutesSum = 0;
  const re = /\[OUTING_MINUTES=(\d+)\]/g;
  let m;
  while ((m = re.exec(noteText)) !== null) {
    outingMinutesSum += parseInt(m[1], 10) || 0;
  }

  // 外出分控除後の実働見込み
  let netMillis = workMillis - (outingMinutesSum * 60 * 1000);
  if (netMillis < 0) netMillis = 0;

  // 休憩時間の自動計算は「外出控除後の実働」に対して行う
  let breakMinutes = 0;
  const netHours = netMillis / (1000 * 60 * 60);
  if (netHours >= 8) {
    breakMinutes = 60;
  } else if (netHours >= 6) {
    breakMinutes = 45;
  }

  // 実働時間 = 外出控除後 - 休憩
  let actualWorkMillis = netMillis - (breakMinutes * 60 * 1000);
  if (actualWorkMillis < 0) actualWorkMillis = 0;

  // 時間フォーマット
  const breakTime = breakMinutes > 0 ? `${breakMinutes}分` : '';
  const actualWorkHours = Math.floor(actualWorkMillis / (1000 * 60 * 60));
  const actualWorkMinutes = Math.floor((actualWorkMillis % (1000 * 60 * 60)) / (1000 * 60));
  const actualWorkTime = `${actualWorkHours}時間${actualWorkMinutes}分`;

  // 表示用休憩時間（外出分を注記）
  let displayBreak = breakTime || '0分';
  if (outingMinutesSum > 0) {
    displayBreak += ` + 外出${outingMinutesSum}分`;
  }

  // セルに設定
  sheet.getRange(row, 6).setValue(displayBreak); // F列
  sheet.getRange(row, 7).setValue(actualWorkTime); // G列
}

/* ================================================================= */
/* 月次集計更新 ----------------------------------------------------- */
/* ================================================================= */

function _updateMonthlySummary_(summarySheet, attendanceSheet) {
  const now = new Date();
  const currentMonth = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM');

  // 出勤簿から当月のデータを取得
  const lastRow = attendanceSheet.getLastRow();
  if (lastRow <= 1) return;

  const allData = attendanceSheet.getRange(2, 1, lastRow - 1, 7).getValues(); // 7列に

  // ユーザーと施設ごとの集計
  const summaryMap = new Map();

  allData.forEach(row => {
    const date = row[0];
    const user = row[1];
    const facility = row[2];
    const arrivalTime = row[3];
    const workTime = row[6];

    if (!date || !user || !facility) return;

    const monthStr = Utilities.formatDate(new Date(date), 'Asia/Tokyo', 'yyyy/MM');
    if (monthStr !== currentMonth) return;

    const summaryKey = `${user}_${facility}`;
    if (!summaryMap.has(summaryKey)) {
      summaryMap.set(summaryKey, {
        user: user,
        facility: facility,
        days: 0,
        totalMinutes: 0
      });
    }

    const summary = summaryMap.get(summaryKey);

    // 出勤日数カウント（出社時刻がある場合のみ）
    if (arrivalTime) {
      summary.days++;
    }

    // 実働時間の集計
    if (workTime) {
      const match = workTime.match(/(\d+)時間(\d+)分/);
      if (match) {
        summary.totalMinutes += parseInt(match[1], 10) * 60 + parseInt(match[2], 10);
      }
    }
  });

  // 既存の集計を読み込み、当月以外は保持
  const existingLastRow = summarySheet.getLastRow();
  const keepRows = [];
  if (existingLastRow > 1) {
    const existing = summarySheet.getRange(2, 1, existingLastRow - 1, 6).getValues();
    for (const row of existing) {
      if (row[0] && row[0] !== currentMonth) {
        keepRows.push(row);
      }
    }
  }

  // シート内容部をクリア（フォーマット維持）
  if (existingLastRow > 1) {
    summarySheet.getRange(2, 1, existingLastRow - 1, 6).clearContent();
  }

  // まず当月以外の既存行を書き戻す
  let rowIndex = 2;
  if (keepRows.length > 0) {
    summarySheet.getRange(rowIndex, 1, keepRows.length, 6).setValues(keepRows);
    rowIndex += keepRows.length;
  }

  // 次に当月の集計を書き込む
  summaryMap.forEach((summary) => {
    const totalHours = Math.floor(summary.totalMinutes / 60);
    const totalMins = summary.totalMinutes % 60;
    summarySheet.getRange(rowIndex, 1, 1, 6).setValues([
      [
        currentMonth,
        summary.user,
        summary.facility,
        summary.days,
        `${totalHours}時間${totalMins}分`,
        new Date()
      ]
    ]);
    rowIndex++;
  });
}


/* ================================================================= */
/* シート作成関数 --------------------------------------------------- */
/* ================================================================= */

function _createAttendanceSheet_(ss) {
  const sheet = ss.insertSheet('出勤簿');
  sheet.getRange(1, 1, 1, 8).setValues([
    ['日付', 'ユーザー名', '施設', '出社時刻', '退社時刻', '休憩時間', '実働時間', '備考']
  ]);

  // 列幅調整
  sheet.setColumnWidth(1, 100); // 日付
  sheet.setColumnWidth(2, 120); // ユーザー名
  sheet.setColumnWidth(3, 120); // 施設
  sheet.setColumnWidth(4, 80);  // 出社時刻
  sheet.setColumnWidth(5, 80);  // 退社時刻
  sheet.setColumnWidth(6, 80);  // 休憩時間
  sheet.setColumnWidth(7, 100); // 実働時間
  sheet.setColumnWidth(8, 300); // 備考

  // ヘッダー装飾
  const header = sheet.getRange(1, 1, 1, 8);
  header.setBackground('#4a86e8');
  header.setFontColor('#ffffff');
  header.setFontWeight('bold');

  return sheet;
}

function _createRawDataSheet_(ss) {
  const sheet = ss.insertSheet('raw_data');
  sheet.getRange(1, 1, 1, 12).setValues([
    [
      '受信時刻', 'username', 'timestamp', 'state', 'name', 
      'devicename', 'ipaddr', 'MAC', 'description', 
      'siteName', 'facilityName', 'raw_json'
    ]
  ]);

  // ヘッダー装飾
  const header = sheet.getRange(1, 1, 1, 12);
  header.setBackground('#666666');
  header.setFontColor('#ffffff');
  header.setFontWeight('bold');

  return sheet;
}

function _createMonthlySummarySheet_(ss) {
  const sheet = ss.insertSheet('月次集計');
  sheet.getRange(1, 1, 1, 6).setValues([
    ['年月', 'ユーザー名', '施設', '出勤日数', '総実働時間', '更新日時']
  ]);

  // ヘッダー装飾
  const header = sheet.getRange(1, 1, 1, 6);
  header.setBackground('#0f9d58');
  header.setFontColor('#ffffff');
  header.setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 150);

  return sheet;
}

/* ================================================================= */
/* ユーティリティ関数 ----------------------------------------------- */
/* ================================================================= */

function _ok_() {
  return ContentService.createTextOutput(JSON.stringify({status: 'ok'}))
         .setMimeType(ContentService.MimeType.JSON);
}

function _error_(message) {
  return ContentService.createTextOutput(JSON.stringify({status: 'error', message: message}))
         .setMimeType(ContentService.MimeType.JSON);
}

function _safeJson_(txt) {
  try { 
    return JSON.parse(txt); 
  } catch(e) { 
    return { parseError: e.toString(), raw: txt }; 
  } 
}

/* ================================================================= */
/* 定期実行用関数（必要に応じて設定） --------------------------------- */
/* ================================================================= */

function dailyAttendanceCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheetByName('出勤簿');
  if (!attendanceSheet) return;

  const lastRow = attendanceSheet.getLastRow();
  if (lastRow <= 1) return;

  const now = new Date();
  const gracePeriodHours = 24; // 24時間以上の勤務を異常とみなす

  const dataRange = attendanceSheet.getRange(2, 1, lastRow - 1, 8);
  const values = dataRange.getValues();

  for (let i = 0; i < values.length; i++) {
    const rowData = values[i];
    const arrivalDate = rowData[0];
    const departureTime = rowData[4];

    if (arrivalDate && !departureTime) {
      // 出勤時刻から24時間以上経過しているかチェック
      const arrivalTimestamp = new Date(arrivalDate).getTime();
      if ((now.getTime() - arrivalTimestamp) > (gracePeriodHours * 60 * 60 * 1000)) {
        const rowIndex = i + 2;
        
        // 退勤時刻を出勤日の23:59:59に設定
        const endOfDay = new Date(arrivalTimestamp);
        endOfDay.setHours(23, 59, 59, 999);
        
        attendanceSheet.getRange(rowIndex, 5).setValue(endOfDay);
        
        const currentNote = values[i][7];
        attendanceSheet.getRange(rowIndex, 8).setValue(
          currentNote + `\n[システム] ${gracePeriodHours}時間以上退勤未記録のため、同日23:59で仮設定`
        );
        
        _calculateWorkTime_(attendanceSheet, rowIndex);
      }
    }
  }

  const monthlySummarySheet = ss.getSheetByName('月次集計');
  if (monthlySummarySheet) {
    _updateMonthlySummary_(monthlySummarySheet, attendanceSheet);
  }
}

/**
 * 初期セットアップ用関数
 */
function setupAttendanceSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName('出勤簿')) {
    _createAttendanceSheet_(ss);
  }

  if (!ss.getSheetByName('raw_data')) {
    _createRawDataSheet_(ss);
  }

  if (!ss.getSheetByName('月次集計')) {
    _createMonthlySummarySheet_(ss);
  }

  SpreadsheetApp.getUi().alert('シートのセットアップが完了しました。');
}

/**
 * raw_dataシートの記録を元に出勤簿シートを再構築する関数。
 * メニューから手動で一度だけ実行することを想定しています。
 */
function reprocessRawData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('確認', 'raw_dataシートを元に出勤簿シートを再構築します。出勤簿の既存のデータはクリアされます。よろしいですか？', ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    ui.alert('処理を中断しました。');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawDataSheet = ss.getSheetByName('raw_data');
  const attendanceSheet = ss.getSheetByName('出勤簿');
  const monthlySummarySheet = ss.getSheetByName('月次集計');

  if (!rawDataSheet || !attendanceSheet) {
    ui.alert('エラー: raw_dataシートまたは出勤簿シートが見つかりません。');
    return;
  }

  // 1. 出勤簿をクリア（ヘッダーは残す）
  if (attendanceSheet.getLastRow() > 1) {
    attendanceSheet.getRange(2, 1, attendanceSheet.getLastRow() - 1, 8).clearContent();
  }

  // 2. raw_dataから全データを読み込む
  const rawDataLastRow = rawDataSheet.getLastRow();
  if (rawDataLastRow <= 1) {
    ui.alert('raw_dataシートに処理するデータがありません。');
    return;
  }
  const rawValues = rawDataSheet.getRange(2, 1, rawDataLastRow - 1, 11).getValues();

  // 3. イベント発生時刻（timestamp列）でソートする
  rawValues.sort(function(a, b) {
    return new Date(a[2]) - new Date(b[2]); // Column C (index 2) is timestamp
  });

  // 4. 1行ずつ処理して出勤簿を更新
  SpreadsheetApp.getActiveSpreadsheet().toast('出勤簿の再構築を開始しました... データ量に応じて時間がかかります。');
  for (const row of rawValues) {
    const data = {
      username:     row[1],  // username
      timestamp:    row[2],  // timestamp
      state:        row[3],  // state
      description:  row[8],  // description
      facilityName: row[10] // facilityName
    };

    // 必須データがない場合はスキップ
    if (!data.username || !data.timestamp || !data.facilityName) {
      continue;
    }

    // 既存のロジックを再利用して出勤簿を更新
    _updateAttendanceSheet_(attendanceSheet, data);
  }

  // 5. 月次集計を更新
  if (monthlySummarySheet) {
    _updateMonthlySummary_(monthlySummarySheet, attendanceSheet);
  }

  ui.alert('出勤簿の再構築が完了しました。');
}

/**
 * スプレッドシートを開いた時にカスタムメニューを追加する
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('カスタムメニュー')
      .addItem('出勤簿をraw_dataから再構築', 'reprocessRawData')
      .addToUi();
}
/**
 * さまざまな形式のtimestampを安全にDateへ変換
 * - 数値(文字列含む): 1e12未満→秒、以上→ミリ秒
 * - ISO8601文字列: Date.parseに委譲
 * - それ以外: new Date(v) にフォールバック
 */
function _parseTimestamp_(v) {
  if (typeof v === 'number') {
    const ms = v < 1e12 ? v * 1000 : v;
    return new Date(ms);
  }
  if (typeof v === 'string') {
    if (/^\d+$/.test(v)) {
      const n = Number(v);
      const ms = n < 1e12 ? n * 1000 : n;
      return new Date(ms);
    }
    // ISOや一般的な日時表現
    const t = Date.parse(v);
    if (!isNaN(t)) return new Date(t);
  }
  // フォールバック
  return new Date(v);
}
