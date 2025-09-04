/**
 * Omada Webhook 受信用 (Google Apps Script) - 修正版
 * 2025-07-10 出勤簿連携・Discord通知対応版
 * 2025-09-04 施設情報連携対応版
 */

function doPost(e) {
  // 並行実行の衝突回避（最大30秒待機）
  const lock = LockService.getDocumentLock();
  try { lock.waitLock(30 * 1000); } catch (err) { console.error('Lock acquisition failed:', err); }

  try {
    /* ---------- 1. 受信 ---------- */
    const raw     = (e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const payload = _safeJson_(raw);

    /* ---------- 2. ログ保存（全件） ---------- */
    _writeLog_(payload);

    /* ---------- 3. MAC / ステータス / IP / 時刻 / 施設情報 抽出 ---------- */
    const info = _extractInfo_(payload);
    if (!info.macRaw || !info.status) return _ok_();

    const macNorm = _normalizeMac_(info.macRaw);

    /* ---------- 4. mac / facility シート照合 ---------- */
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const macSheet  = ss.getSheetByName('mac') || ss.insertSheet('mac');
    const macData   = _getMacData_(macSheet);
    
    if (!macData.has(macNorm)) return _ok_();
    
    const macInfo   = macData.get(macNorm);
    const dispName  = macInfo.displayName;
    const username  = macInfo.username;
    const endpoint  = macInfo.gasEndpoint;
    const webhook   = macInfo.discordWebhook;

    const facilitySheet = ss.getSheetByName('facility') || ss.insertSheet('facility');
    const facilityData  = _getFacilityData_(facilitySheet);
    
    // 施設未登録でも処理継続（フォールバックはsiteNameそのまま）
    let facilityName = '';
    if (info.siteName && facilityData.has(info.siteName)) {
      facilityName = facilityData.get(info.siteName);
    } else {
      facilityName = info.siteName || '';
      console.warn('facility未登録: ', info.siteName);
    }

  /* ---------- 5. data シートへ書き込み ---------- */
  const dataSheet = ss.getSheetByName('data') || ss.insertSheet('data');
  const newRow = [
    info.eventDate,
    info.macRaw.toUpperCase(),
    dispName,
    info.status,
    info.ip || '',
    facilityName, // 施設名を追加
    JSON.stringify(payload)
  ];
  dataSheet.appendRow(newRow);

  /* ---------- 6. 出勤判定ロジック ---------- */
  const judgmentResult = _judgeAttendance_(dataSheet, macSheet, username, info.status, info.eventDate);
  // 直近に追加した行の「判定内容」列（8列目）へ書き込み
  try {
    const appendedRowIndex = dataSheet.getLastRow();
    dataSheet.getRange(appendedRowIndex, 8).setValue(judgmentResult.description);
  } catch (e) {
    console.warn('判定内容の書き込みに失敗:', e);
  }
    
    /* ---------- 7. 出勤簿エンドポイントへ送信 ---------- */
  // 施設名が未確定（空）の場合は出勤簿への送信をスキップ
  if (endpoint && facilityName) {
    const postData = {
      username: username,
      // 互換性のため従来の文字列と、機械可読なEpoch(ms)の両方を送付
      timestamp: _formatDateTime_(info.eventDate),
      timestampMs: info.eventDate.getTime(),
      state: info.status,
      name: username,
      devicename: dispName,
      ipaddr: info.ip || '',
      MAC: info.macRaw.toUpperCase(),
      description: judgmentResult.description,
      siteName: info.siteName, // サイト名を追加
      facilityName: facilityName // 施設表示名を追加
    };
      
      try {
        UrlFetchApp.fetch(endpoint, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify(postData)
        });
      } catch(e) {
        console.error('出勤簿エンドポイントへの送信エラー:', e);
      }
    }

    /* ---------- 8. Discord通知 ---------- */
    if (webhook) {
      _sendDiscordNotification_(webhook, username, dispName, info.status, judgmentResult.description, info.eventDate, facilityName);
    }

    return _ok_();
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  } else if (endpoint && !facilityName) {
    console.warn('facilityNameが空のためgas2への送信をスキップ');
  }
}

/* ================================================================= */
/* 出勤判定ロジック ------------------------------------------------- */
/* ================================================================= */

function _judgeAttendance_(dataSheet, macSheet, username, currentStatus, eventDate) {
  const lastRow = dataSheet.getLastRow();

  // ヘッダー行と直近の行（現在イベント）を除外して履歴を探索
  const historyCount = Math.max(0, lastRow - 2);
  const allData = historyCount > 0
    ? dataSheet.getRange(2, 1, historyCount, 7).getValues().reverse()
    : [];
  const macData = _getMacData_(macSheet);
  const currentDateStr = _formatDate_(eventDate);

  // Find previous entries for the same user
  let sameDayEntry = null;
  let previousEntry = null;
  let otherDeviceOnline = false;

  const currentDeviceMacNorm = _normalizeMac_(lastRow >= 1 ? (dataSheet.getRange(lastRow, 2).getValue() || '') : '');

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const rowMac = _normalizeMac_(row[1]);
    const rowStatus = row[3];
    const rowUserInfo = macData.get(rowMac);
    if (!rowUserInfo || rowUserInfo.username !== username) continue;

    const rowDateStr = _formatDate_(row[0]);
    
    // Check for other online devices on the same day (excluding the current event's device)
    if (rowMac !== currentDeviceMacNorm && rowStatus === 'ONLINE' && rowDateStr === currentDateStr) {
      otherDeviceOnline = true;
    }
    
    if (rowDateStr === currentDateStr && !sameDayEntry) {
      sameDayEntry = {row: row, index: i};
    }
    
    if (!previousEntry) {
      previousEntry = {row: row, index: i, dateStr: rowDateStr};
    }
  }

  let description = '';

  if (currentStatus === 'ONLINE') {
    if (!sameDayEntry) {
      description = '出勤(最初のエントリ)';
      if (previousEntry && previousEntry.row[3] === 'OFFLINE' && previousEntry.dateStr !== currentDateStr) {
        const rowIndex = lastRow - previousEntry.index - 1; // index=0が直近行のため-1補正
        // Ensure rowIndex is valid before setting value
        if(rowIndex > 1) {
            dataSheet.getRange(rowIndex, 8).setValue('前日最後のOFFLINE=退勤と判定');
        }
      }
    } else {
      const prevStatus = sameDayEntry.row[3];
      if (prevStatus === 'OFFLINE') {
        const interval = _calculateInterval_(sameDayEntry.row[0], eventDate);
        description = `ONLINEに復帰しました interval=${interval}`;
      } else {
        description = 'ONLINEに復帰しました前回OFFLINEが記録されてません';
      }
    }
    if (otherDeviceOnline) {
      description += ' //別の端末ではすでにONLINEへ復帰済み';
    }
  } else { // currentStatus === 'OFFLINE'
    const userDevices = _getUserDevices_(macData, username);
    const currentDeviceMac = currentDeviceMacNorm;
    const otherDevices = userDevices.filter(mac => mac !== currentDeviceMac);
    let isAnotherDeviceOnline = false;

    for (const deviceMac of otherDevices) {
      const lastStatusOfDevice = allData.find(row => {
        const rowMac = _normalizeMac_(row[1]);
        const rowDateStr = _formatDate_(row[0]);
        return rowMac === deviceMac && rowDateStr === currentDateStr;
      });

      if (lastStatusOfDevice && lastStatusOfDevice[3] === 'ONLINE') {
        isAnotherDeviceOnline = true;
        break;
      }
    }

    if (isAnotherDeviceOnline) {
      description = 'OFFLINE=デバイスがwifiから切断されました (他のデバイスはONLINE)';
    } else {
      description = '退勤(最終デバイスのOFFLINE)';
    }
  }
  
  return {description: description};
}

/* ================================================================= */
/* Discord通知 ----------------------------------------------------- */
/* ================================================================= */

function _sendDiscordNotification_(webhook, username, devicename, status, description, eventDate, facilityName) {
  const color = status === 'ONLINE' ? 0x00ff00 : 0xff0000;
  const statusEmoji = status === 'ONLINE' ? '🟢' : '🔴';
  const placeLabel = status === 'ONLINE' ? '出勤先' : '退勤先';
  
  const embed = {
    embeds: [{
      title: `${statusEmoji} ${username} - ${status} (${facilityName})`, // 施設名を追加
      description: description,
      color: color,
      fields: [
        {
          name: placeLabel,
          value: facilityName || '未登録',
          inline: true
        },
        {
          name: 'デバイス名',
          value: devicename,
          inline: true
        },
        {
          name: '時刻',
          value: _formatDateTime_(eventDate),
          inline: true
        }
      ],
      timestamp: eventDate.toISOString()
    }]
  };
  
  try {
    UrlFetchApp.fetch(webhook, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(embed)
    });
  } catch(e) {
    console.error('Discord通知エラー:', e);
  }
}

/* ================================================================= */
/* ユーティリティ関数 ----------------------------------------------- */
/* ================================================================= */

function _ok_(){
  // Omadaが期待する形式でレスポンスを返す
  const response = ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'Webhook received successfully'
  }));
  
  response.setMimeType(ContentService.MimeType.JSON);
  
  // HTTPステータスコードを明示的に200に設定
  return response;
}

function _safeJson_(txt){
  try { return JSON.parse(txt); }
  catch(e){ return { parseError: e.toString(), raw: txt }; }
}

function _normalizeMac_(m){
  return m.replace(/[^0-9A-Fa-f]/g,'').toUpperCase();
}

function _eventStatus_(msg=''){
  const m = msg.toLowerCase();
  if (m.includes('went online') || m.includes('online')  || m.includes('オンライン'))  return 'ONLINE';
  if (m.includes('went offline')|| m.includes('offline') || m.includes('オフライン')) return 'OFFLINE';
  return '';
}

function _getMacData_(sheet){
  const rows = sheet.getLastRow();
  if (rows <= 1) return new Map();

  // ヘッダー行を除外して読み込み
  const vals = sheet.getRange(2, 1, rows - 1, 5).getValues();
  const map  = new Map();
  vals.forEach(r => {
    if (r[0]) {
      const macNorm = _normalizeMac_(r[0]);
      if (!/^[0-9A-F]{12}$/.test(macNorm)) return; // 不正なMACを除外
      map.set(macNorm, {
        displayName: r[1] || '',
        gasEndpoint: r[2] || '',
        username: r[3] || '',
        discordWebhook: r[4] || ''
      });
    }
  });
  return map;
}

// facilityシートからデータを取得する関数を追加
function _getFacilityData_(sheet) {
  const rows = sheet.getLastRow();
  if (rows <= 1) return new Map();

  // ヘッダー行を除外して読み込み
  const vals = sheet.getRange(2, 1, rows - 1, 2).getValues();
  const map = new Map();
  vals.forEach(r => {
    if (r[0]) {
      map.set(r[0], r[1] || '');
    }
  });
  return map;
}

function _getUserDevices_(macData, username) {
  const devices = [];
  macData.forEach((value, key) => {
    if (value.username === username) {
      devices.push(key);
    }
  });
  return devices;
}

function _writeLog_(payload){
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const logSh = ss.getSheetByName('log') || ss.insertSheet('log');
  logSh.appendRow([new Date(), JSON.stringify(payload)]);
}

function _extractInfo_(obj){
  let macRaw = '', status = '', ip = '', siteName = '';

  siteName = obj.Site || ''; // Site情報を取得

  if (obj.data && obj.data.clientMac){
    macRaw = obj.data.clientMac;
    ip     = obj.data.clientIp || '';
    status = _eventStatus_(obj.msg || '');
  } else if (Array.isArray(obj.text) && obj.text[0]){
    const line = obj.text[0];
    macRaw = (line.match(/client:[^:\]]+:([0-9A-Fa-f:-]+)/) || [])[1] || '';
    ip     = (line.match(/IP:\s*([0-9.]+)/)               || [])[1] || '';
    status = _eventStatus_(line);
  }

  // タイムスタンプは秒/ミリ秒の両方に対応
  const tsRaw = Number(obj.timestamp);
  let ms;
  if (isFinite(tsRaw)) {
    ms = tsRaw < 1e12 ? tsRaw * 1000 : tsRaw;
  } else {
    ms = Date.now();
  }
  let   eventDate = new Date(ms);

  return { macRaw, status, ip, eventDate, siteName }; // siteNameを返す
}

function _formatDate_(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
}

function _formatDateTime_(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

function _calculateInterval_(startDate, endDate) {
  const diff = endDate.getTime() - startDate.getTime();
  const hours = Math.floor(diff / (1000 * 60 * 60));
  const minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
  const seconds = Math.floor((diff % (1000 * 60)) / 1000);
  
  return `${hours}時間${minutes}分${seconds}秒`;
}

/* ================================================================= */
/* テスト用関数（開発時のみ使用） ------------------------------------ */
/* ================================================================= */

function testSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // macシートのセットアップ
  const macSheet = ss.getSheetByName('mac') || ss.insertSheet('mac');
  if (macSheet.getLastRow() === 0) {
    macSheet.getRange(1, 1, 1, 5).setValues([
      ['MACアドレス', 'デバイス名', 'GASエンドポイント', 'ユーザー名', 'Discord Webhook']
    ]);
  }
  
  // facilityシートのセットアップ
  const facilitySheet = ss.getSheetByName('facility') || ss.insertSheet('facility');
  if (facilitySheet.getLastRow() === 0) {
    facilitySheet.getRange(1, 1, 1, 2).setValues([
      ['サイト名', '表示名']
    ]);
    facilitySheet.getRange(2, 1, 1, 2).setValues([
      ['Akihabara_office', '秋葉原事務所']
    ]);
  }

  // dataシートのセットアップ  
  const dataSheet = ss.getSheetByName('data') || ss.insertSheet('data');
  if (dataSheet.getLastRow() === 0) {
    dataSheet.getRange(1, 1, 1, 8).setValues([ // 8列に
      ['タイムスタンプ', 'MAC', '表示名', 'ステータス', 'IP', '施設名', '元JSON', '判定内容']
    ]);
  }
  
  // logシートのセットアップ
  const logSheet = ss.getSheetByName('log') || ss.insertSheet('log');
  if (logSheet.getLastRow() === 0) {
    logSheet.getRange(1, 1, 1, 2).setValues([
      ['受信時刻', 'ペイロード']
    ]);
  }
}

/**
 * Webhook動作確認用関数
 * Omadaからのテストリクエストをシミュレート
 */
function testWebhook() {
  const testPayload = {
    postData: {
      contents: JSON.stringify({
        Site: "Akihabara_office", // テスト用にSiteを追加
        timestamp: Date.now(),
        data: {
          clientMac: "AA:BB:CC:DD:EE:FF",
          clientIp: "192.168.1.100"
        },
        msg: "Client AA:BB:CC:DD:EE:FF went online"
      })
    }
  };
  
  const result = doPost(testPayload);
  console.log('Test Result:', result.getContent());
}
