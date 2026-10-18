/**
 * 虹揚廣告 預售實價登錄查詢 — Google Apps Script 線上版
 * 資料來源：雲端硬碟中名為「實價登錄」的 Google 試算表（若有多份同名檔，取最近更新的一份）
 * 工作表：預售屋買賣
 */
const FILE_NAME = '實價登錄';
const SHEET_NAME = '預售屋買賣';
const BASE_VERSION = 4;   // 目前版本起算號碼（V4）；之後每次發布含程式變更的新版本，系統會自動 +1（見 getVersion_）
const LOG_BOOK = '實價登錄使用記錄';   // 使用記錄另存成獨立試算表（首次有人使用時自動建立），不會被每次重傳「實價登錄」蓋掉
const SHEET_SET = '網頁設定';   // 標題處「最新資料」說明文字：B2 儲存格（可自行修改；沒有此工作表時用下方預設值）
const LATEST_DEFAULT = '交易日期在115/8/21~115/8/31之預售屋';
const SHEET_591 = '591檢索';   // 社區 ↔ 591 建案對照表（由 Excel「實價登錄」的 591檢索 工作表轉入）

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 網址帶 ?api=data → 回傳 JSON 資料（給 PWA／App 使用）；?api=log → 記錄一次使用；其餘照舊回傳網頁 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.api === 'data') {
    try { return jsonOut_(getData()); }
    catch (err) { return jsonOut_({ error: String(err && err.message || err) }); }
  }
  if (p.api === 'log') {
    try { logVisit(String(p.cid || ''), String(p.device || '')); } catch (err) {}
    return jsonOut_({ ok: true });
  }
  // 資料直接由伺服器放進網頁（用文字取代，不使用樣板語法），不依賴瀏覽器端的 google.script.run
  let json;
  try { json = JSON.stringify(getData()); }
  catch (e) { json = JSON.stringify({ error: String(e && e.message || e) }); }
  json = json.replace(/\//g, '\\/').replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  const raw = HtmlService.createHtmlOutputFromFile('Index').getContent();
  const html = raw.replace('/*DATA_PLACEHOLDER*/null', function () { return json; });   // 用函式避免 $ 被當成特殊符號
  return HtmlService.createHtmlOutput(html)
    .setTitle('虹揚廣告預售實價登錄查詢')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 找出最新的「實價登錄」Google 試算表 */
function findFile_() {
  const it = DriveApp.searchFiles(
    'title = "' + FILE_NAME + '" and mimeType = "application/vnd.google-apps.spreadsheet" and trashed = false');
  let best = null;
  while (it.hasNext()) {
    const f = it.next();
    if (!best || f.getLastUpdated() > best.getLastUpdated()) best = f;
  }
  if (!best) throw new Error('雲端硬碟找不到名為「' + FILE_NAME + '」的 Google 試算表');
  return best;
}

/**
 * 解析交易年月日（欄 H）→ {y: 民國年, m: 月}
 * 支援：文字/數字 1141113、991231（民國 2 碼年）、114/11/13、2025/11/13（西元自動換算）
 */
function parseYM_(v) {
  const s = String(v).trim();
  if (!s) return null;
  let y, m;
  const parts = s.split(/[\/\-\.年月日\s]+/).filter(Boolean);
  if (parts.length >= 2) {
    y = +parts[0]; m = +parts[1];
  } else {
    const d = s.replace(/\D/g, '');
    if (d.length < 5) return null;
    y = +d.slice(0, d.length - 4);   // 年 = 去掉最後 4 碼（月日）
    m = +d.slice(d.length - 4, d.length - 2);
  }
  if (y > 1911) y -= 1911;
  if (!y || m < 1 || m > 12) return null;
  return { y: y, m: m };
}

/** 讀取試算表並整理成網頁用資料（較耗時；結果會放進快取，見 getData） */
function buildData_(file) {
  const t0 = Date.now(), tm = {};
  const ss = SpreadsheetApp.openById(file.getId());
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('找不到工作表「' + SHEET_NAME + '」');
  tm.open = Date.now() - t0;
  const n = sh.getLastRow() - 1;
  // 只讀需要的欄位（A、H、L、P～AC、AH），不整張讀 → 比讀 34 欄快很多
  const colA = n > 0 ? sh.getRange(2, 1, n, 1).getValues() : [];
  const colH = n > 0 ? sh.getRange(2, 8, n, 1).getDisplayValues() : [];   // H 欄用顯示文字，避免「通用格式」被轉成數字或日期
  const colHv = n > 0 ? sh.getRange(2, 8, n, 1).getValues() : [];
  const colL = n > 0 ? sh.getRange(2, 12, n, 1).getValues() : [];
  const blk = n > 0 ? sh.getRange(2, 16, n, 14).getValues() : [];          // P(16)…AC(29)：P=0, V=6, Y=9, Z=10, AC=13
  const colAH = (n > 0 && sh.getLastColumn() >= 34) ? sh.getRange(2, 34, n, 1).getValues() : [];
  tm.read = Date.now() - t0;
  const num = v => (typeof v === 'number' ? v : (parseFloat(String(v).replace(/,/g, '')) || 0));
  const rows = [];
  let skipped = 0;
  for (let i = 0; i < n; i++) {
    const town = String(colA[i][0]).trim();
    if (!town || town === '累積統計') continue;          // 排除空白列與小計列
    let ym = parseYM_(colH[i][0]);
    if (!ym && colHv[i][0] instanceof Date) ym = { y: colHv[i][0].getFullYear() - 1911, m: colHv[i][0].getMonth() + 1 };
    if (!ym) { skipped++; continue; }
    const b = blk[i];
    rows.push([
      town,                       // 0 鄉鎮市區 (A)
      ym.y * 100 + ym.m,          // 1 成交年月 民國 yyymm (H 前五碼)
      String(b[13]).trim(),       // 2 建案名稱 (AC)
      num(b[0]),                  // 3 建物移轉總面積㎡ (P)
      num(b[9]),                  // 4 車位移轉總面積㎡ (Y)
      num(b[6]),                  // 5 總價元 (V)
      num(b[10]),                 // 6 車位總價元 (Z)
      String(colL[i][0]).trim(),  // 7 建物型態 (L)
      colAH.length ? num(colAH[i][0]) : 0   // 8 裝潢費金額 (AH)；單價計算時剔除
    ]);
  }
  tm.parse = Date.now() - t0;
  // 591 對照：key = 鄉鎮|社區名稱 → [591編號, 591建案名稱, 判定]
  // 樂居銷控表：key 同上 → 銷控表網址（取自 591檢索 O 欄 HYPERLINK；查無者不放入，網頁改用 Google 搜尋）
  const l591 = {}, lj = {};
  const s591 = ss.getSheetByName(SHEET_591);
  if (s591 && s591.getLastRow() > 1) {
    const m = s591.getLastRow() - 1;
    const lc = s591.getLastColumn();
    // 依「標題名稱」找欄位，之後在 591檢索 插入/刪除欄位都不會再錯位
    const hdr = s591.getRange(1, 1, 1, lc).getDisplayValues()[0].map(x => String(x).trim());
    const col = (names) => { for (const n of names) { const k = hdr.indexOf(n); if (k >= 0) return k; } return -1; };
    const cTown = col(['鄉鎮市區']), cName = col(['社區(建案名稱)', '社區']),
          c591n = col(['591建案名稱']), cId = col(['591編號']), cVer = col(['判定']),
          cLj = col(['樂居銷控表']);
    const vals = s591.getRange(2, 1, m, lc).getValues();
    const fo = cLj >= 0 ? s591.getRange(2, cLj + 1, m, 1).getFormulas() : [];
    vals.forEach((r, i) => {
      const town = String(r[cTown]).trim(), name = String(r[cName]).trim(),
            id = String(r[cId]).trim(), verdict = String(r[cVer]).trim();
      if (town && name && id && verdict !== '未找到') l591[town + '|' + name] = [id, String(r[c591n]).trim(), verdict];
      if (cLj >= 0 && town && name) {
        const mm = /HYPERLINK\("([^"]+)"/i.exec(String(fo[i][0] || ''));
        if (mm && /leju\.com\.tw\/community\//.test(mm[1])) lj[town + '|' + name] = mm[1];
      }
    });
  }
  let latest = '';
  try {
    const st = ss.getSheetByName(SHEET_SET);
    if (st && st.getLastRow() >= 2) latest = String(st.getRange(2, 2).getDisplayValue() || '').trim();
  } catch (e) {}
  if (!latest) latest = LATEST_DEFAULT;
  tm.total = Date.now() - t0;
  console.log('buildData_ 耗時(ms)：' + JSON.stringify(tm) + '，筆數 ' + rows.length);
  return {
    rows: rows, l591: l591, lj: lj, latest: latest, fileId: file.getId(), skipped: skipped,
    updated: Utilities.formatDate(file.getLastUpdated(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm')
  };
}

/** 把資料壓縮後分段放進快取（每段 < 100KB）；試算表沒更新時，之後開啟只需約 1～2 秒 */
function cachePut_(key, obj) {
  try {
    const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(obj), 'application/json', 'd.json')).getBytes());
    const size = 90000, parts = Math.ceil(b64.length / size), cache = CacheService.getScriptCache(), put = {};
    for (let i = 0; i < parts; i++) put[key + '_' + i] = b64.substr(i * size, size);
    put[key + '_n'] = String(parts);
    cache.putAll(put, 21600);
  } catch (e) { console.warn('快取寫入失敗（不影響功能）：' + e); }
}
function cacheGet_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const n = parseInt(cache.get(key + '_n'), 10);
    if (!n) return null;
    const keys = []; for (let i = 0; i < n; i++) keys.push(key + '_' + i);
    const got = cache.getAll(keys);
    let b64 = ''; for (let i = 0; i < n; i++) { if (!got[keys[i]]) return null; b64 += got[keys[i]]; }
    const bytes = Utilities.base64Decode(b64);
    return JSON.parse(Utilities.ungzip(Utilities.newBlob(bytes, 'application/x-gzip', 'd.gz')).getDataAsString());
  } catch (e) { return null; }
}

/** 回傳全部交易資料（精簡欄位），由網頁端篩選計算 */
function getData() {
  const file = findFile_();
  const key = 'D4_' + file.getId().slice(0, 8) + '_' + file.getLastUpdated().getTime();
  let d = cacheGet_(key);
  if (!d) { d = buildData_(file); cachePut_(key, d); }
  let version = '';
  try { version = getVersion_(); } catch (e) { version = 'V' + BASE_VERSION; }
  d.version = version;
  return d;
}

/** 診斷用：在編輯器選 testGetData 按「執行」，到「執行記錄」看各步驟耗時，確認資料庫連結是否正常 */
function testGetData() {
  const t0 = Date.now();
  const file = findFile_();
  console.log('找到檔案：' + file.getName() + '｜更新 ' + file.getLastUpdated() + '｜' + (Date.now() - t0) + 'ms');
  const d = buildData_(file);
  cachePut_('D4_' + file.getId().slice(0, 8) + '_' + file.getLastUpdated().getTime(), d);   // 順便暖快取
  const t1 = Date.now(); let v = ''; try { v = getVersion_(); } catch (e) { v = '版本取得失敗：' + e; }
  console.log('版本 ' + v + '｜取得版本耗時 ' + (Date.now() - t1) + 'ms');
  console.log('筆數 ' + d.rows.length + '｜591 對照 ' + Object.keys(d.l591).length + '｜樂居連結 ' + Object.keys(d.lj).length + '｜略過 ' + d.skipped + '｜總耗時 ' + (Date.now() - t0) + 'ms');
}


/** 版本號：以「全部程式碼＋網頁內容」的雜湊判斷是否發布了新版本；有變更就自動 +1，首次為 BASE_VERSION（V4）。 */
function getVersion_() {
  const props = PropertiesService.getScriptProperties();
  const src = [doGet, findFile_, parseYM_, getData, buildData_, cachePut_, cacheGet_, getLogBook_, setupLogBook_, logVisit, hdr_, noteFormula_]
    .map(f => f.toString()).join('\n') + HtmlService.createHtmlOutputFromFile('Index').getContent();
  const cur = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, src, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    let num = parseInt(props.getProperty('VER_NUM'), 10) || 0;
    const h = props.getProperty('VER_HASH');
    if (!h || num < BASE_VERSION) num = Math.max(num, BASE_VERSION);
    else if (h !== cur) num += 1;
    props.setProperties({ VER_NUM: String(num), VER_HASH: cur });
    return 'V' + num;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/* ================= 使用記錄 ================= */
const LOG_LAYOUT = '3';   // 使用記錄版面版本；變更版面時加 1，程式會自動升級既有試算表

/** 取得（必要時建立／升級）「實價登錄使用記錄」試算表。
 *  不同版本、不同專案的程式只要在同一個 Google 帳號下，就會用名稱找到同一份記錄檔，記到一起。 */
function getLogBook_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const savedId = props.getProperty('LOG_BOOK_ID');
  if (savedId) { try { ss = SpreadsheetApp.openById(savedId); } catch (e) { ss = null; } }
  if (!ss) {
    const it = DriveApp.searchFiles('title = "' + LOG_BOOK + '" and mimeType = "application/vnd.google-apps.spreadsheet" and trashed = false');
    ss = it.hasNext() ? SpreadsheetApp.openById(it.next().getId()) : SpreadsheetApp.create(LOG_BOOK);
    props.setProperty('LOG_BOOK_ID', ss.getId());
    props.deleteProperty('LOG_LAYOUT_DONE');
  }
  if (props.getProperty('LOG_LAYOUT_DONE') !== LOG_LAYOUT + ':' + ss.getId()) {
    setupLogBook_(ss);
    props.setProperty('LOG_LAYOUT_DONE', LOG_LAYOUT + ':' + ss.getId());
  }
  return ss;
}

/** 「備註」欄公式：依裝置代碼（去空白）到「裝置備註」工作表比對 */
function noteFormula_(row) {
  return '=IFERROR(INDEX(FILTER(\'裝置備註\'!$B$2:$B,TRIM(\'裝置備註\'!$A$2:$A)=TRIM($D' + row + ')),1),"")';
}

function hdr_(range) { return range.setFontWeight('bold').setBackground('#b0191e').setFontColor('#ffffff'); }

/** 建立或升級版面（可重複執行）：使用記錄 / 每月統計 / 裝置備註；並把月份、日期修正為純文字 */
function setupLogBook_(ss) {
  ss.setSpreadsheetTimeZone('Asia/Taipei');
  let log = ss.getSheetByName('使用記錄') || ss.getSheets()[0].setName('使用記錄');
  hdr_(log.getRange(1, 1, 1, 7).setValues([['時間', '月份', '日期', '裝置代碼', '裝置類型', '備註（來自「裝置備註」）', '版本']]));
  log.setFrozenRows(1);
  log.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  log.getRange('B:E').setNumberFormat('@');          // 純文字，避免「2026-10」被轉成日期數字
  log.setColumnWidths(1, 1, 160); log.setColumnWidths(2, 3, 100); log.setColumnWidth(4, 190); log.setColumnWidth(6, 200); log.setColumnWidth(7, 130);
  // 修正既有資料列的月份／日期（由時間欄重新產生純文字），並把「備註」改成即時對照「裝置備註」的公式（之後補填備註，過去記錄也會自動更新）
  const n = log.getLastRow() - 1;
  if (n > 0) {
    log.getRange(2, 6, n, 1).setFormulas(Array.from({ length: n }, (_, i) => [noteFormula_(i + 2)]));
    const times = log.getRange(2, 1, n, 1).getValues();
    const md = times.map(r => r[0] instanceof Date
      ? [Utilities.formatDate(r[0], 'Asia/Taipei', 'yyyy-MM'), Utilities.formatDate(r[0], 'Asia/Taipei', 'yyyy-MM-dd')] : ['', '']);
    log.getRange(2, 2, n, 2).setNumberFormat('@').setValues(md);
  }

  let note = ss.getSheetByName('裝置備註');
  if (!note) {
    note = ss.insertSheet('裝置備註');
    hdr_(note.getRange('A1:B1').setValues([['裝置代碼', '備註（例：Tommy 辦公室電腦）']]));
    note.getRange('A:A').setNumberFormat('@');
    note.setColumnWidth(1, 190); note.setColumnWidth(2, 260); note.setFrozenRows(1);
    note.getRange('D1').setValue('說明：把自己或同事的「裝置代碼」貼在 A 欄、名稱寫在 B 欄，之後新的記錄會自動帶入備註，並從「外部裝置數」中排除。');
  }

  let st = ss.getSheetByName('每月統計');
  if (!st) st = ss.insertSheet('每月統計');
  st.clear();
  const L = "'使用記錄'!";
  hdr_(st.getRange('A1:G1').setValues([['月份', '開啟次數', '不重複裝置數', '手機次數', '電腦次數', '有使用的天數', '外部裝置數（排除有備註者）']]));
  st.getRange('A2').setFormula('=IFERROR(SORT(UNIQUE(FILTER(' + L + 'B2:B,' + L + 'B2:B<>"")),1,FALSE),"")');
  st.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIF(' + L + 'B:B,A2:A)))');
  st.getRange('C2').setFormula('=ARRAYFORMULA(IF(A2:A="","",MAP(A2:A,LAMBDA(m,COUNTUNIQUE(FILTER(' + L + 'D2:D,' + L + 'B2:B=m))))))');
  st.getRange('D2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(' + L + 'B:B,A2:A,' + L + 'E:E,"手機")))');
  st.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(' + L + 'B:B,A2:A,' + L + 'E:E,"電腦")))');
  st.getRange('F2').setFormula('=ARRAYFORMULA(IF(A2:A="","",MAP(A2:A,LAMBDA(m,COUNTUNIQUE(FILTER(' + L + 'C2:C,' + L + 'B2:B=m))))))');
  st.getRange('G2').setFormula('=ARRAYFORMULA(IF(A2:A="","",MAP(A2:A,LAMBDA(m,IFERROR(COUNTUNIQUE(FILTER(' + L + 'D2:D,' + L + 'B2:B=m,' + L + 'F2:F="")),0)))))');
  hdr_(st.getRange('I1:J1').setValues([['累計', '']]));
  st.getRange('I2:I4').setValues([['總開啟次數'], ['不重複裝置數'], ['統計起始日']]);
  st.getRange('J2').setFormula('=COUNTA(' + L + 'A2:A)');
  st.getRange('J3').setFormula('=IFERROR(COUNTUNIQUE(' + L + 'D2:D),0)');
  st.getRange('J4').setFormula('=IFERROR(MIN(' + L + 'A2:A),"")').setNumberFormat('yyyy-mm-dd');
  hdr_(st.getRange('L1:N1').setValues([['版本', '開啟次數', '不重複裝置數']]));
  st.getRange('L2').setFormula('=IFERROR(SORT(UNIQUE(FILTER(' + L + 'G2:G,' + L + 'G2:G<>"")),1,FALSE),"")');
  st.getRange('M2').setFormula('=ARRAYFORMULA(IF(L2:L="","",COUNTIF(' + L + 'G:G,L2:L)))');
  st.getRange('N2').setFormula('=ARRAYFORMULA(IF(L2:L="","",MAP(L2:L,LAMBDA(v,COUNTUNIQUE(FILTER(' + L + 'D2:D,' + L + 'G2:G=v))))))');
  st.setColumnWidths(1, 6, 105); st.setColumnWidth(7, 210); st.setColumnWidths(9, 2, 110); st.setColumnWidths(12, 3, 130);
  st.setFrozenRows(1);
  ss.setActiveSheet(st);
  ss.moveActiveSheet(1);
}

/** 由網頁載入完成後呼叫：記錄一次使用。clientId 為瀏覽器產生的隨機代碼（不含個人資料）。 */
function logVisit(clientId, device) {
  const cid = String(clientId || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 24) || 'unknown';
  const dev = device === '手機' ? '手機' : '電腦';
  let ver = '';
  try { ver = getVersion_(); } catch (e) { ver = 'V' + BASE_VERSION; }   // 先取版本（內含鎖定，不可放在下方鎖定區內）
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getLogBook_();
    const now = new Date();
    const log = ss.getSheetByName('使用記錄');
    const row = log.getLastRow() + 1;
    log.getRange(row, 1, 1, 7).setValues([[
      now,
      Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM'),
      Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd'),
      cid, dev, '', ver
    ]]);
    log.getRange(row, 6).setFormula(noteFormula_(row));   // 備註由公式即時對照「裝置備註」
    SpreadsheetApp.flush();
  } catch (e) {
    console.error('logVisit 失敗：' + e);   // 失敗不影響查詢功能；原因可在「執行作業」查看
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** 手動測試／手動升級版面：在編輯器選這個函式按「執行」 */
function testLog() {
  const ss = getLogBook_();
  Logger.log('使用記錄網址：' + ss.getUrl());
}
