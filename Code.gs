/**
 * 社会基盤ユニットQMS管理システム v3.4
 */
const CONFIG = {
  PROGRESS_SS_ID: '1fhMHbLWHeSIF4HRd9d44Aqmw0aS7j4WYKI3FsGQyKKE',
  PHONEBOOK_SS_ID: '1x6Uy711HFPwdLPFNxyCvMk0Fo77XC4MJNEMh29_n0Lo',
  ORDER_FOLDER_ID: '10azgUkgwEKMxfmv5O9GVmwuos9KiAB-y',
  KENMEI_FOLDER_ID: '1kcAYs0mXtCc2qsei9NNiI8fY_GwZzLi3',
  HEADER_ROW:8, DATA_START_ROW:9,
  COL_IMPORT_START:2, COL_IMPORT_COUNT:8,
  COL_DETAIL_START:10, COL_TOTAL:47,
  COL_SS_ID:51, // AY列
  COL_COMMENT_STATUS:52, // AZ列: コメント状況
  DEPARTMENTS: {
    '土木本部':['耐震技術部','技術開発部','土木設計部','風力技術部'],
    'ジオフロント本部':['バックエンド技術部','地下開発技術部']
  },
  EXCEL_COLS:{BU_NAME:3,KENMEI_CODE:6,KEIYAKU_KENMEI:7,KOKI_START:8,KOKI_END:9,TOUNENDO_JUCHU:11,KYAKUSAKI_KUBUN:17,KOKYAKU_NAME:20}
};

const DETAIL_FIELDS = [
  '_J','_K','group','tanto','jisshi','shinsa','approver','tokki',
  'keiyaku_sakusei','keiyaku_risk','keiyaku_henkou','keiyaku_risk2',
  'hinshitsu_henkou','hinshitsu_risk','sekkei_umu','sekkei_tetsuzuki',
  'irai_kaisu','irai_jisshi','itaku_umu','itaku_tetsuzuki',
  'itaku_kaisu','itaku_jisshi','gijiroku_no','gijiroku_date',
  'shinchoku_kaisu','shinchoku_jisshi','kentou_kaisu','kentou_jisshi',
  'kensa_kaisu','kensa_jisshi','datousei_kaisu','datousei_jisshi',
  'risk_sheet','risk_taiou','risk_yuukou','kenshu','denshika','qc_bu','qc_unit'
];

// ===== Web App =====
function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var tpl = HtmlService.createTemplateFromFile('Index');
  tpl.approveMode = (p.mode === 'approve') ? true : false;
  tpl.contractApproveMode = (p.mode === 'contract_approve') ? true : false;
  tpl.contractApproveToken = p.ctoken || '';
  tpl.genericApproveMode = (p.genericApprove === 'true') ? true : false;
  tpl.genericApproveToken = p.token || '';
  tpl.commentMode = (p.mode === 'comments') ? true : false;
  tpl.commentDept = p.dept || '';
  tpl.commentRow = p.row || '';
  tpl.approveToken = p.token || '';
  tpl.approveMrow = p.mrow || '';
  return tpl.evaluate()
    .setTitle('社会基盤ユニットQMS管理システム')
    .addMetaTag('viewport','width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(f){return HtmlService.createHtmlOutputFromFile(f).getContent();}
function getOrganizationStructure(){return CONFIG.DEPARTMENTS;}
function getWebAppUrl(){return ScriptApp.getService().getUrl();}
function getCurrentUserEmail(){return Session.getActiveUser().getEmail();}

// ===== 管理者判定 =====
function isAdmin() {
  var email = getCurrentUserEmail().toLowerCase();
  var admins = (PropertiesService.getScriptProperties().getProperty('QC_MANAGERS') || '').toLowerCase().split(',');
  for (var i = 0; i < admins.length; i++) { if (admins[i].trim() === email) return true; }
  return false;
}

// ===== フロー設定CRUD =====
function getFlowSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID);
  var sh = ss.getSheetByName('フロー設定');
  if (!sh) { sh = ss.insertSheet('フロー設定'); sh.getRange(1,1,1,4).setValues([['id','name','json','updated']]); sh.setFrozenRows(1); }
  return sh;
}

function getFlowList() {
  try {
    var sh = getFlowSheet_();
    var lr = sh.getLastRow();
    if (lr < 2) return [];
    var data = sh.getRange(2,1,lr-1,4).getValues();
    return data.map(function(r){ return { id:sv(r[0]), name:sv(r[1]), updated:sv(r[3]) }; });
  } catch(e) { return []; }
}

function getFlowData(flowId) {
  try {
    var sh = getFlowSheet_();
    var lr = sh.getLastRow();
    if (lr < 2) return null;
    var data = sh.getRange(2,1,lr-1,3).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === flowId) return { id:sv(data[i][0]), name:sv(data[i][1]), json:sv(data[i][2]) };
    }
    return null;
  } catch(e) { return null; }
}

function saveFlowData(flowId, name, json) {
  try {
    if (!isAdmin()) return { success:false, message:'管理者権限がありません。' };
    var sh = getFlowSheet_();
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
    var lr = sh.getLastRow();
    if (lr >= 2) {
      var ids = sh.getRange(2,1,lr-1,1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === flowId) {
          sh.getRange(i+2,2,1,3).setValues([[name, json, now]]);
          return { success:true, message:'フロー「'+name+'」を更新しました。' };
        }
      }
    }
    sh.appendRow([flowId, name, json, now]);
    return { success:true, message:'フロー「'+name+'」を作成しました。' };
  } catch(e) { return { success:false, message:e.message }; }
}

function deleteFlow(flowId) {
  try {
    if (!isAdmin()) return { success:false, message:'管理者権限がありません。' };
    var sh = getFlowSheet_();
    var lr = sh.getLastRow();
    if (lr < 2) return { success:false, message:'フローが見つかりません。' };
    var ids = sh.getRange(2,1,lr-1,1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === flowId) { sh.deleteRow(i+2); return { success:true, message:'削除しました。' }; }
    }
    return { success:false, message:'フローが見つかりません。' };
  } catch(e) { return { success:false, message:e.message }; }
}
// メールから氏名を検索
function getNameByEmail(email) {
  if (!email) return '';
  try {
    var rows = SpreadsheetApp.openById(CONFIG.PHONEBOOK_SS_ID).getSheets()[0].getDataRange().getValues().slice(1);
    for (var i = 0; i < rows.length; i++) {
      var mail = String(rows[i][8] || '').trim(); // I列=メール
      if (mail && mail.toLowerCase() === email.toLowerCase()) return String(rows[i][5]).trim(); // F列=氏名
    }
    // @の前を返す
    return email.split('@')[0];
  } catch(e) { return email.split('@')[0]; }
}

// ===== 部門データ取得 =====
function getDepartmentData(dept) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID);
    var sh = ss.getSheetByName(dept);
    if (!sh) return { error: 'シート「' + dept + '」が見つかりません。' };
    var lr = sh.getLastRow();
    if (lr < CONFIG.DATA_START_ROW) return { data: [], department: dept };
    var vals = sh.getRange(CONFIG.DATA_START_ROW, 2, lr - CONFIG.DATA_START_ROW + 1, CONFIG.COL_TOTAL).getValues();
    var ayVals = sh.getRange(CONFIG.DATA_START_ROW, CONFIG.COL_SS_ID, lr - CONFIG.DATA_START_ROW + 1, 2).getValues();
    var data = [];
    for (var i = 0; i < vals.length; i++) {
      var r = vals[i]; if (!r[0] && !r[3]) continue;
      var item = { no:r[0], buName:r[1], code:r[2], kenmeiName:r[3],
        kikiStart:r[4]?fd(r[4]):'', kikiEnd:r[5]?fd(r[5]):'',
        juchuAmount:r[6], contractor:r[7],
        rowIndex:i+CONFIG.DATA_START_ROW, projectSSId:ayVals[i][0]||'',
        commentStatus:sv(ayVals[i][1]||''), detail:{} };
      for (var d = 0; d < DETAIL_FIELDS.length; d++) {
        var k = DETAIL_FIELDS[d]; if (k.charAt(0)==='_') continue;
        var v = r[8+d]; item.detail[k] = (v instanceof Date) ? fd(v) : (v!=null?String(v):'');
      }
      data.push(item);
    }
    return { data:data, department:dept };
  } catch(e) { return { error:e.message }; }
}

function fd(d) { if (d instanceof Date) return d.getFullYear()+'/'+('0'+(d.getMonth()+1)).slice(-2)+'/'+('0'+d.getDate()).slice(-2); return String(d); }
// 安全文字列変換（Date型をgoogle.script.runで返すとnullになるため全てString化）
function sv(v) { if (v == null || v === '') return ''; if (v instanceof Date) return fd(v); return String(v); }
// 時刻セル用フォーマッタ（Date型の時刻 → "HH:MM" 10分単位に丸め）
function ft(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date) {
    var h = v.getHours(), m = Math.round(v.getMinutes() / 10) * 10;
    if (m >= 60) { m = 0; h = (h + 1) % 24; }
    return ('0'+h).slice(-2) + ':' + ('0'+m).slice(-2);
  }
  return String(v);
}
function getHonbuByDept(dept) { for (var h in CONFIG.DEPARTMENTS) { if (CONFIG.DEPARTMENTS[h].indexOf(dept)>=0) return h; } return ''; }

// ===== 件名フォルダ・スプレッドシート管理 =====
function ensureProjectSS(dept, rowIndex, code, kenmeiName) {
  try {
    var sh = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID).getSheetByName(dept);
    var existing = sh.getRange(rowIndex, CONFIG.COL_SS_ID).getValue();
    if (existing) {
      try { SpreadsheetApp.openById(existing); return { ssId: existing }; } catch(e) { /* 壊れていたら再作成 */ }
    }
    // テンプレートを探す
    var rootFolder = DriveApp.getFolderById(CONFIG.KENMEI_FOLDER_ID);
    var tmpl = null;
    var files = rootFolder.getFiles();
    while (files.hasNext()) { var f = files.next(); if (f.getName() === '件名_Default') { tmpl = f; break; } }
    if (!tmpl) return { error: 'テンプレート「件名_Default」が見つかりません。' };

    // 部署フォルダを探す/作成
    var deptFolder = findOrCreateFolder_(rootFolder, dept);

    // 件名フォルダを作成: 「件名コード_件名」
    var kenmeiFolder = findOrCreateFolder_(deptFolder, code + '_' + kenmeiName);

    // テンプレートをコピーして件名フォルダに保存
    var copy = tmpl.makeCopy(code + '_' + kenmeiName, kenmeiFolder);
    var newId = copy.getId();

    // シート準備
    var newSs = SpreadsheetApp.openById(newId);
    if (!newSs.getSheetByName('グループ・実施体制・特記事項')) {
      var s = newSs.insertSheet('グループ・実施体制・特記事項');
      s.getRange(1,1,1,7).setValues([['グループ','担当者','実施責任者','審査責任者','承認者','特記事項','設定日']]);
    }
    if (!newSs.getSheetByName('議事録')) {
      var s2 = newSs.insertSheet('議事録');
      s2.getRange(1,1,1,13).setValues([['議事録No','年月日','開始時間','終了時間','場所','出席者（相手）','出席者（当社）','資料','協議内容','作成者','実施責任者','作成日','承認日']]);
    }
    // AY列に保存
    sh.getRange(rowIndex, CONFIG.COL_SS_ID).setValue(newId);
    return { ssId: newId };
  } catch(e) { return { error: e.message }; }
}

function findOrCreateFolder_(parent, name) {
  var subs = parent.getFoldersByName(name);
  if (subs.hasNext()) return subs.next();
  return parent.createFolder(name);
}

// 件名フォルダを取得（SS親フォルダ）
function getKenmeiFolderId_(ssId) {
  try {
    var f = DriveApp.getFileById(ssId);
    var parents = f.getParents();
    return parents.hasNext() ? parents.next().getId() : null;
  } catch(e) { return null; }
}

// ===== グループ・実施体制・特記事項 =====
function saveGroupTaisei(ssId, data) {
  try {
    var ss = SpreadsheetApp.openById(ssId);
    var sh = ss.getSheetByName('グループ・実施体制・特記事項');
    if (!sh) { sh = ss.insertSheet('グループ・実施体制・特記事項'); sh.getRange(1,1,1,7).setValues([['グループ','担当者','実施責任者','審査責任者','承認者','特記事項','設定日']]); }
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
    sh.getRange(2,1,1,7).setValues([[data.group||'',data.tanto||'',data.jisshi||'',data.shinsa||'',data.approver||'',data.tokki||'',now]]);
    return { success:true, message:'保存しました。' };
  } catch(e) { return { success:false, message:e.message }; }
}

function getGroupTaisei(ssId) {
  try {
    var sh = SpreadsheetApp.openById(ssId).getSheetByName('グループ・実施体制・特記事項');
    if (!sh || sh.getLastRow() < 2) return { group:'',tanto:'',jisshi:'',shinsa:'',approver:'',tokki:'',date:'' };
    var r = sh.getRange(2,1,1,7).getValues()[0];
    return { group:sv(r[0]),tanto:sv(r[1]),jisshi:sv(r[2]),shinsa:sv(r[3]),approver:sv(r[4]),tokki:sv(r[5]),date:sv(r[6]) };
  } catch(e) { return { group:'',tanto:'',jisshi:'',shinsa:'',approver:'',tokki:'',date:'' }; }
}

function syncGroupToProgress(dept, rowIndex, data) {
  try {
    var sh = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID).getSheetByName(dept);
    var startCol = CONFIG.COL_DETAIL_START + 2; // L列
    sh.getRange(rowIndex, startCol, 1, 6).setValues([[data.group||'',data.tanto||'',data.jisshi||'',data.shinsa||'',data.approver||'',data.tokki||'']]);
  } catch(e) {}
}

// ===== 詳細セクション保存 =====
function saveDetailSection(dept, rowIndex, updates) {
  try {
    var sh = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID).getSheetByName(dept);
    if (!sh) return { success:false, message:'シートが見つかりません。' };
    for (var key in updates) {
      var idx = DETAIL_FIELDS.indexOf(key);
      if (idx >= 0) sh.getRange(rowIndex, CONFIG.COL_DETAIL_START + idx).setValue(updates[key]);
    }
    return { success:true, message:'保存しました。' };
  } catch(e) { return { success:false, message:e.message }; }
}

// ===== 汎用フォームエンジン CRUD =====
function ensureGenericSheet_(ssId, sheetName) {
  var ss = SpreadsheetApp.openById(ssId);
  var sh = ss.getSheetByName(sheetName);
  if (sh) return sh;
  sh = ss.insertSheet(sheetName);
  sh.setFrozenRows(1);
  return sh;
}

function getGenericFormData(ssId, sheetName) {
  try {
    var sh = ensureGenericSheet_(ssId, sheetName);
    var lc = sh.getLastColumn();
    if (lc < 1) return null;
    var lr = sh.getLastRow();
    if (lr < 2) return null;
    var headers = sh.getRange(1, 1, 1, lc).getValues()[0];
    var vals = sh.getRange(2, 1, 1, lc).getValues()[0];
    var result = {};
    for (var i = 0; i < headers.length; i++) {
      var k = sv(headers[i]);
      if (k && k.indexOf('_approval_') !== 0) result[k] = sv(vals[i]);
    }
    return Object.keys(result).length ? result : null;
  } catch(e) { return null; }
}

function saveGenericFormData(ssId, sheetName, data) {
  try {
    var sh = ensureGenericSheet_(ssId, sheetName);
    // 既存の承認状態を保持
    var existingApproval = {};
    var lc = sh.getLastColumn();
    if (lc > 0 && sh.getLastRow() >= 2) {
      var headers = sh.getRange(1, 1, 1, lc).getValues()[0];
      var vals = sh.getRange(2, 1, 1, lc).getValues()[0];
      for (var i = 0; i < headers.length; i++) {
        var h = sv(headers[i]);
        if (h.indexOf('_approval_') === 0) existingApproval[h] = sv(vals[i]);
      }
    }
    // フォームデータ（_approval_で始まるキーは除外）
    var keys = [], vals2 = [];
    for (var key in data) {
      if (key.indexOf('_approval_') === 0) continue;
      keys.push(key);
      vals2.push(data[key] !== undefined && data[key] !== null ? String(data[key]) : '');
    }
    // 承認列を追加
    for (var ak in existingApproval) {
      keys.push(ak);
      vals2.push(existingApproval[ak]);
    }
    if (!keys.length) return { success:false, message:'保存するデータがありません。' };
    if (lc > 0) {
      sh.getRange(1, 1, 1, lc).clearContent();
      if (sh.getLastRow() >= 2) sh.getRange(2, 1, 1, lc).clearContent();
    }
    sh.getRange(1, 1, 1, keys.length).setValues([keys]);
    sh.getRange(2, 1, 1, vals2.length).setValues([vals2]);
    return { success:true, message: sheetName + 'を保存しました。' };
  } catch(e) { return { success:false, message: e.message }; }
}

// 帳票No自動採番（同じシート名の承認記録から最大Noを取得して+1）
function getNextFormNo(ssId, sheetName) {
  try {
    var sh = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return '1';
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var noIdx = -1;
    for (var i = 0; i < headers.length; i++) { if (sv(headers[i]) === 'no') { noIdx = i; break; } }
    if (noIdx < 0) return '1';
    var val = sh.getRange(2, noIdx+1).getValue();
    return val ? String(Number(val) + 1) : '1';
  } catch(e) { return '1'; }
}

// ===== 契約内容確認 =====
function ensureContractSheet_(ssId) {
  var ss = SpreadsheetApp.openById(ssId);
  var sh = ss.getSheetByName('契約内容確認');
  if (sh) return sh;
  sh = ss.insertSheet('契約内容確認');
  var COLS = 35;
  sh.getRange(1,1,1,COLS).merge().setValue('契約内容確認').setBackground('#1a5276').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(12);
  sh.getRange(2,1,1,2).merge().setValue('基本情報').setBackground('#2471a3').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2,3,1,6).merge().setValue('受注関連情報顧客要求事項').setBackground('#2471a3').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2,9,1,7).merge().setValue('顧客不満足発生リスクの確認').setBackground('#922b21').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2,16,1,10).merge().setValue('契約前段階リスク確認').setBackground('#7d3c98').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2,26,1,4).merge().setValue('成果品チェック体制の選定').setBackground('#1a7a5c').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2,30,1,6).merge().setValue('承認ルート（確認者→部長→本部長）').setBackground('#d4ac0d').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(3,1,1,15).setValues([[
    '受付日','確認者','営業からの受領情報',
    '発注内示書又は確認記録','積算依頼書又は確認記録',
    '見積依頼書（仕様書、設計書）又は確認記録',
    '社会的要求事項の確認記録','その他の確認記録',
    '顧客要求事項に対し必要な技術を保有しているか',
    '必要な人材・体制を自部署で確保できるか',
    '必要な物的資源を自部署で確保できるか',
    '他部署あるいは委託先の協力により解決できるか',
    '協力先',
    '原子力関連要領を適用するか',
    '仕様書不利益内容確認'
  ]]).setBackground('#e3edf5').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  sh.getRange(3,16,1,10).setValues([[
    '1.新規顧客','2.新規商材','3.業務内容不明確','4.与信問題',
    '5.契約変更','6.工期変更','7.確認記録不備',
    '8.金額不足','9.損害賠償','10.その他'
  ]]).setBackground('#e8daef').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  sh.getRange(3,26,1,4).setValues([['①受注額5000万\n②重要度難易度\n③定型設計(YES/NO)','選定体制\n(A-1～3/B-1～3\n/C/特別)','特別体制\n採用理由','特別体制\n承認者/審査者'
  ]]).setBackground('#d5f5e3').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  sh.getRange(3,30,1,6).setValues([['部長\n(自動設定)','部長\n承認日','部長判定\n(approved/rejected)','本部長\n(自動設定)','本部長\n承認日','本部長判定\n(approved/rejected)'
  ]]).setBackground('#fdebd0').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  sh.setFrozenRows(3);
  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 100);
  for (var i = 3; i <= 8; i++) sh.setColumnWidth(i, 160);
  for (var j = 9; j <= 15; j++) sh.setColumnWidth(j, 140);
  for (var k = 16; k <= 25; k++) sh.setColumnWidth(k, 140);
  for (var m = 26; m <= COLS; m++) sh.setColumnWidth(m, 130);
  sh.setRowHeight(3, 80);
  sh.getRange(3,1,1,COLS).setBorder(true,true,true,true,true,true);
  return sh;
}

function getContractData(ssId) {
  try {
    var sh = ensureContractSheet_(ssId);
    if (sh.getLastRow() < 4) return { exists: false };
    var lc = Math.max(sh.getLastColumn(), 35);
    var r = sh.getRange(4, 1, 1, lc).getValues()[0];
    var hasData = false;
    for (var i = 0; i < r.length; i++) { if (r[i] !== '') { hasData = true; break; } }
    if (!hasData) return { exists: false };
    var result = {
      exists: true,
      date: sv(r[0]), confirmer: sv(r[1]), eigyo: sv(r[2]),
      hacchu: sv(r[3]), sekisan: sv(r[4]), mitsumori: sv(r[5]),
      shakai: sv(r[6]), sonota: sv(r[7]),
      risk1: sv(r[8]||''), risk2: sv(r[9]||''), risk3: sv(r[10]||''),
      risk4: sv(r[11]||''), risk5: sv(r[12]||''),
      risk6: sv(r[13]||''), risk7: sv(r[14]||''),
      riskTable: [],
      tqAnswers: sv(r[25]||''), taisei: sv(r[26]||''),
      spReason: sv(r[27]||''), spDetail: sv(r[28]||''),
      bucho: sv(r[29]||''), buchoDate: sv(r[30]||''), buchoStatus: sv(r[31]||''),
      honbucho: sv(r[32]||''), honbuchoDate: sv(r[33]||''), honbuchoStatus: sv(r[34]||'')
    };
    for (var t = 0; t < 10; t++) result.riskTable.push(sv(r[15 + t] || ''));
    return result;
  } catch(e) { return { error: e.message }; }
}

function saveContractData(ssId, data) {
  try {
    var sh = ensureContractSheet_(ssId);
    var row = [data.date, data.confirmer, data.eigyo,
      data.hacchu, data.sekisan, data.mitsumori, data.shakai, data.sonota,
      data.risk1||'', data.risk2||'', data.risk3||'',
      data.risk4||'', data.risk5||'', data.risk6||'', data.risk7||''];
    var rt = data.riskTable || [];
    for (var t = 0; t < 10; t++) row.push(rt[t] || '');
    // Z-AC: 体制 (4列)
    var tqAnswers = (data.tq1||'') + '|' + (data.tq2||'') + '|' + (data.tq3||'');
    var spDetail = (data.spApprover||'') + '|' + (data.spReviewer||'');
    row.push(tqAnswers, data.taisei||'', data.spReason||'', spDetail);
    // AD-AI: 承認欄 (6列) - 既存値を維持
    var lr = sh.getLastRow();
    if (lr >= 4) {
      var existing = sh.getRange(4, 30, 1, 6).getValues()[0];
      for (var a = 0; a < 6; a++) row.push(sv(existing[a]));
    } else {
      for (var b = 0; b < 6; b++) row.push('');
    }
    sh.getRange(4, 1, 1, 35).setValues([row]);
    return { success:true, message:'契約内容確認を保存しました。' };
  } catch(e) { return { success:false, message:e.message }; }
}

// ===== 品質管理担当判定 =====
function isQCManager() {
  try {
    var email = Session.getActiveUser().getEmail().toLowerCase();
    var prop = PropertiesService.getScriptProperties().getProperty('QC_MANAGERS') || '';
    var list = prop.split(',').map(function(e){ return e.trim().toLowerCase(); });
    return list.indexOf(email) >= 0;
  } catch(e) { return false; }
}

// ===== コメントシステム =====
function ensureCommentSheet_(ssId) {
  var ss = SpreadsheetApp.openById(ssId);
  var sh = ss.getSheetByName('コメント');
  if (sh) return sh;
  sh = ss.insertSheet('コメント');
  sh.getRange(1,1,1,8).setValues([['No','日時','投稿者','投稿者メール','帳票名','内容','確認済み','確認日']]);
  sh.setFrozenRows(1);
  sh.setColumnWidth(6, 400);
  return sh;
}

function getComments(ssId) {
  try {
    var sh = ensureCommentSheet_(ssId);
    var lr = sh.getLastRow();
    if (lr < 2) return [];
    var data = sh.getRange(2, 1, lr-1, 8).getValues();
    var result = [];
    for (var i = 0; i < data.length; i++) {
      result.push({
        no: sv(data[i][0]), date: sv(data[i][1]), author: sv(data[i][2]),
        email: sv(data[i][3]), section: sv(data[i][4]), content: sv(data[i][5]),
        resolved: sv(data[i][6]), resolvedDate: sv(data[i][7]), sheetRow: i + 2
      });
    }
    return result;
  } catch(e) { return []; }
}

function addComment(ssId, dept, rowIndex, data) {
  try {
    if (!isQCManager()) return { success:false, message:'品質管理担当のみコメント可能です。' };
    var sh = ensureCommentSheet_(ssId);
    var lr = sh.getLastRow();
    var nextNo = lr < 2 ? 1 : lr;
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
    var email = Session.getActiveUser().getEmail();
    var authorName = getNameByEmail(email);
    sh.getRange(lr+1, 1, 1, 8).setValues([[nextNo, now, authorName, email, data.section||'', data.content||'', '', '']]);

    // 進捗管理表にコメント状況を設定
    var psh = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID).getSheetByName(dept);
    if (psh) psh.getRange(Number(rowIndex), CONFIG.COL_COMMENT_STATUS).setValue('未確認');

    // メール送信
    sendCommentEmail_(ssId, dept, rowIndex, data.section, data.content, authorName);

    return { success:true, message:'コメントを投稿しました。' };
  } catch(e) { return { success:false, message:e.message }; }
}

function markCommentResolved(ssId, commentRow, dept, rowIndex) {
  try {
    if (!isQCManager()) return { success:false, message:'品質管理担当のみ操作可能です。' };
    var sh = ensureCommentSheet_(ssId);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
    sh.getRange(Number(commentRow), 7).setValue('確認済み');
    sh.getRange(Number(commentRow), 8).setValue(now);

    // 全コメントが確認済みか確認
    var lr = sh.getLastRow();
    if (lr >= 2) {
      var statuses = sh.getRange(2, 7, lr-1, 1).getValues();
      var allResolved = true;
      for (var i = 0; i < statuses.length; i++) {
        if (String(statuses[i][0]) !== '確認済み') { allResolved = false; break; }
      }
      if (allResolved) {
        var psh = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID).getSheetByName(dept);
        if (psh) psh.getRange(Number(rowIndex), CONFIG.COL_COMMENT_STATUS).setValue('');
      }
    }
    return { success:true, message:'確認済みにしました。' };
  } catch(e) { return { success:false, message:e.message }; }
}

function sendCommentEmail_(ssId, dept, rowIndex, section, content, authorName) {
  try {
    // 担当者・実施責任者の名前を取得
    var psh = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID).getSheetByName(dept);
    var tantoIdx = CONFIG.COL_DETAIL_START + DETAIL_FIELDS.indexOf('tanto');
    var jisshiIdx = CONFIG.COL_DETAIL_START + DETAIL_FIELDS.indexOf('jisshi');
    var tantoName = sv(psh.getRange(Number(rowIndex), tantoIdx).getValue());
    var jisshiName = sv(psh.getRange(Number(rowIndex), jisshiIdx).getValue());
    var kenmei = sv(psh.getRange(Number(rowIndex), 5).getValue());
    var code = sv(psh.getRange(Number(rowIndex), 4).getValue());

    // 名前→メール変換
    var tantoEmail = getEmailByName_(tantoName);
    var jisshiEmail = getEmailByName_(jisshiName);
    if (!tantoEmail) return;

    var appUrl = ScriptApp.getService().getUrl();
    var link = appUrl + '?mode=comments&dept=' + encodeURIComponent(dept) + '&row=' + rowIndex;

    var subject = '【QMS】品質管理コメント - ' + code + ' ' + kenmei;
    var body = tantoName + ' 様\n\n'
      + '品質管理担当（' + authorName + '）より、以下の帳票にコメントがありました。\n\n'
      + '件名: ' + code + ' ' + kenmei + '\n'
      + '帳票: ' + section + '\n'
      + 'コメント: ' + content + '\n\n'
      + '以下のリンクからコメントを確認してください:\n' + link + '\n';

    var options = {};
    if (jisshiEmail && jisshiEmail !== tantoEmail) options.cc = jisshiEmail;
    options.name = 'QMS管理システム';
    GmailApp.sendEmail(tantoEmail, subject, body, options);
  } catch(e) { Logger.log('sendCommentEmail error: ' + e.message); }
}

function getEmailByName_(name) {
  if (!name) return '';
  try {
    var rows = SpreadsheetApp.openById(CONFIG.PHONEBOOK_SS_ID).getSheets()[0].getDataRange().getValues().slice(1);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][5]).trim() === name && String(rows[i][11]).trim()) return String(rows[i][11]).trim(); // L列=送信用メール
    }
  } catch(e) {}
  return '';
}

// ログインGmailから所属部を特定し、部長・本部長を自動取得
function getApproversByLoginEmail_(email) {
  try {
    var rows = SpreadsheetApp.openById(CONFIG.PHONEBOOK_SS_ID).getSheets()[0].getDataRange().getValues().slice(1);
    var myDept = '', myHonbu = '';
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][8]).trim().toLowerCase() === email.toLowerCase()) {
        myDept = String(rows[i][2]).trim();
        myHonbu = String(rows[i][1]).trim();
        break;
      }
    }
    if (!myDept) return { error: '電話帳に登録されていません。' };
    var bucho = null, honbucho = null;
    for (var j = 0; j < rows.length; j++) {
      var pos = String(rows[j][4]).trim(), dept = String(rows[j][2]).trim(), hb = String(rows[j][1]).trim();
      var nm = String(rows[j][5]).trim(), sendEmail = String(rows[j][11]).trim();
      if (dept === myDept && (pos === '部長' || /^部長\s*●$/.test(pos)) && !bucho) {
        bucho = { name: nm, email: sendEmail };
      }
      if (hb === myHonbu && (pos === '本部長' || /^本部長\s*●$/.test(pos)) && !honbucho) {
        honbucho = { name: nm, email: sendEmail };
      }
    }
    return { bucho: bucho, honbucho: honbucho, dept: myDept, honbu: myHonbu };
  } catch(e) { return { error: e.message }; }
}

// クライアントから呼び出し可能な公開版
function getApproversByEmail(email) { return getApproversByLoginEmail_(email); }

// 部署・本部指定で部長・本部長を検索
function getApproversByDept(dept, honbu) {
  try {
    var rows = SpreadsheetApp.openById(CONFIG.PHONEBOOK_SS_ID).getSheets()[0].getDataRange().getValues().slice(1);
    var bucho = null, honbucho = null;
    for (var j = 0; j < rows.length; j++) {
      var pos = String(rows[j][4]).trim(), d = String(rows[j][2]).trim(), hb = String(rows[j][1]).trim();
      var nm = String(rows[j][5]).trim(), sendEmail = String(rows[j][11]).trim();
      if (d === dept && (pos === '部長' || /^部長\s*●$/.test(pos)) && !bucho) {
        bucho = { name: nm, email: sendEmail };
      }
      if (hb === honbu && (pos === '本部長' || /^本部長\s*●$/.test(pos)) && !honbucho) {
        honbucho = { name: nm, email: sendEmail };
      }
    }
    return { bucho: bucho, honbucho: honbucho, dept: dept, honbu: honbu };
  } catch(e) { return { error: e.message }; }
}

// ===== 契約内容確認 承認ワークフロー =====
function getContractApprovalSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID);
  var sh = ss.getSheetByName('契約承認依頼');
  if (!sh) {
    sh = ss.insertSheet('契約承認依頼');
    sh.getRange(1,1,1,14).setValues([['token','ssId','dept','rowIndex','code','kenmei','customer','confirmer','confirmerEmail','approverName','approverEmail','step','requestDate','status']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function findContractApprovalRecord_(token) {
  var sh = getContractApprovalSheet_();
  var lr = sh.getLastRow();
  if (lr < 2) return null;
  var data = sh.getRange(2, 1, lr - 1, 14).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(token)) {
      return {
        sheetRow: i + 2,
        token: data[i][0], ssId: String(data[i][1]),
        dept: String(data[i][2]), rowIndex: Number(data[i][3]),
        code: String(data[i][4]), kenmei: String(data[i][5]),
        customer: String(data[i][6]), confirmer: String(data[i][7]),
        confirmerEmail: String(data[i][8]), approverName: String(data[i][9]),
        approverEmail: String(data[i][10]), step: String(data[i][11]),
        requestDate: sv(data[i][12]), status: String(data[i][13]) || 'pending'
      };
    }
  }
  return null;
}

function sendContractApproval(ssId, dept, rowIndex) {
  try {
    var loginEmail = getCurrentUserEmail();
    var confirmerName = getNameByEmail(loginEmail);
    var approvers = getApproversByLoginEmail_(loginEmail);
    if (approvers.error) return { success:false, message: approvers.error };
    if (!approvers.bucho || !approvers.bucho.email) return { success:false, message: '部長のメールアドレスが見つかりません。' };

    var cd = getContractData(ssId);
    if (!cd || !cd.exists) return { success:false, message: '契約内容確認データがありません。' };

    var psh = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID).getSheetByName(dept);
    var code = sv(psh.getRange(Number(rowIndex), 4).getValue());
    var kenmei = sv(psh.getRange(Number(rowIndex), 5).getValue());
    var customer = cd.date ? (sv(psh.getRange(Number(rowIndex), 9).getValue())) : '';

    // 部長名・本部長名をSSに記録
    var csh = ensureContractSheet_(ssId);
    csh.getRange(4, 30).setValue(approvers.bucho.name);
    csh.getRange(4, 32).setValue('pending');
    if (approvers.honbucho) csh.getRange(4, 33).setValue(approvers.honbucho.name);

    // トークン生成
    var token = Utilities.getUuid();
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
    var ash = getContractApprovalSheet_();
    ash.appendRow([token, ssId, dept, rowIndex, code, kenmei, customer||'', confirmerName, loginEmail, approvers.bucho.name, approvers.bucho.email, 'bucho', now, 'pending']);

    // 承認依頼メール送信
    sendContractApprovalEmail_(token, approvers.bucho.name, approvers.bucho.email, code, kenmei, cd.date, confirmerName, customer, 'bucho');
    writeApprovalRecord_(ssId, '契約内容確認', 'keiyaku', '承認依頼送信', confirmerName, 'bucho', '', code, kenmei, '契約内容確認');
    return { success:true, message: '承認依頼を ' + approvers.bucho.name + ' に送信しました。' };
  } catch(e) { return { success:false, message: e.message }; }
}

function sendContractApprovalEmail_(token, approverName, approverEmail, code, kenmei, date, confirmer, customer, step) {
  var url = getWebAppUrl() + '?mode=contract_approve&ctoken=' + token;
  var stepLabel = step === 'bucho' ? '部長' : '本部長';
  var subject = '【承認依頼】契約内容確認 - ' + code + ' ' + kenmei;
  var htmlBody = '<div style="font-family:sans-serif;max-width:600px;margin:0 auto;color:#2c3e50;">'
    + '<div style="background:linear-gradient(135deg,#0e3a56,#1a5276);color:#fff;padding:20px 24px;border-radius:8px 8px 0 0;">'
    + '<h2 style="margin:0;font-size:18px;">社会基盤ユニットQMS管理システム</h2>'
    + '<p style="margin:4px 0 0;font-size:14px;opacity:.8;">契約内容確認 承認依頼（' + stepLabel + '）</p></div>'
    + '<div style="background:#fff;border:1px solid #dce1e6;border-top:none;padding:24px;border-radius:0 0 8px 8px;">'
    + '<p style="font-size:16px;margin:0 0 16px;">' + stepLabel + ' <strong>' + approverName + '</strong> 様</p>'
    + '<p style="font-size:16px;margin:0 0 20px;">契約内容確認の承認申請がありました。</p>'
    + '<table style="width:100%;border-collapse:collapse;margin:0 0 24px;">'
    + '<tr><th style="text-align:left;padding:10px 14px;background:#e8eaed;border:1px solid #bbb;width:120px;">顧客名</th><td style="padding:10px 14px;border:1px solid #bbb;">' + (customer||'') + '</td></tr>'
    + '<tr><th style="text-align:left;padding:10px 14px;background:#e8eaed;border:1px solid #bbb;">件名コード</th><td style="padding:10px 14px;border:1px solid #bbb;">' + code + '</td></tr>'
    + '<tr><th style="text-align:left;padding:10px 14px;background:#e8eaed;border:1px solid #bbb;">件名</th><td style="padding:10px 14px;border:1px solid #bbb;">' + kenmei + '</td></tr>'
    + '<tr><th style="text-align:left;padding:10px 14px;background:#e8eaed;border:1px solid #bbb;">受付日</th><td style="padding:10px 14px;border:1px solid #bbb;">' + (date||'') + '</td></tr>'
    + '<tr><th style="text-align:left;padding:10px 14px;background:#e8eaed;border:1px solid #bbb;">確認者</th><td style="padding:10px 14px;border:1px solid #bbb;">' + confirmer + '</td></tr>'
    + '</table>'
    + '<div style="text-align:center;margin:24px 0;">'
    + '<a href="' + url + '" style="display:inline-block;background:#27ae60;color:#fff;text-decoration:none;padding:14px 40px;border-radius:8px;font-size:16px;font-weight:700;">内容を確認する</a></div>'
    + '</div></div>';
  MailApp.sendEmail({ to: approverEmail, subject: subject, body: '契約内容確認の承認依頼です。\n' + url, htmlBody: htmlBody });
}

function getContractApprovalData(token) {
  try {
    if (!token) return { error: 'トークンが空です。' };
    var rec = findContractApprovalRecord_(token);
    if (!rec) return { error: '承認依頼が見つかりません。トークン: ' + String(token).substring(0,8) + '...' };
    if (rec.status !== 'pending') return { error: 'この承認依頼は既に処理済みです。(ステータス: ' + rec.status + ')' };
    var cd = getContractData(rec.ssId);
    if (!cd) cd = { exists: false };
    return { rec: rec, data: cd };
  } catch(e) { return { error: e.message }; }
}

function approveContract(token) {
  try {
    var rec = findContractApprovalRecord_(token);
    if (!rec) return { success:false, message: '承認依頼が見つかりません。' };
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
    var csh = ensureContractSheet_(rec.ssId);

    // ステータス更新
    var ash = getContractApprovalSheet_();
    ash.getRange(rec.sheetRow, 14).setValue('approved');

    if (rec.step === 'bucho') {
      csh.getRange(4, 31).setValue(now); // 部長承認日
      csh.getRange(4, 32).setValue('approved');
      // 本部長へ承認依頼
      var approvers = getApproversByLoginEmail_(rec.confirmerEmail);
      if (approvers.honbucho && approvers.honbucho.email) {
        var cd = getContractData(rec.ssId);
        var contractDate = (cd && cd.date) ? cd.date : '';
        var newToken = Utilities.getUuid();
        var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
        ash.appendRow([newToken, rec.ssId, rec.dept, rec.rowIndex, rec.code, rec.kenmei, rec.customer, rec.confirmer, rec.confirmerEmail, approvers.honbucho.name, approvers.honbucho.email, 'honbucho', nowStr, 'pending']);
        sendContractApprovalEmail_(newToken, approvers.honbucho.name, approvers.honbucho.email, rec.code, rec.kenmei, contractDate, rec.confirmer, rec.customer, 'honbucho');
        writeApprovalRecord_(rec.ssId, '契約内容確認', 'keiyaku', '部長承認', approvers.bucho?approvers.bucho.name:'', 'bucho', '', rec.code, rec.kenmei, '契約内容確認');
        return { success:true, message: '部長承認完了。本部長へ承認依頼を送信しました。' };
      }
      writeApprovalRecord_(rec.ssId, '契約内容確認', 'keiyaku', '部長承認', '', 'bucho', '', rec.code, rec.kenmei, '契約内容確認');
      return { success:true, message: '部長承認完了。（本部長情報なし）' };
    } else {
      // 本部長承認 → 完了 + PDF生成 + 完了通知
      csh.getRange(4, 34).setValue(now); // 本部長承認日
      csh.getRange(4, 35).setValue('approved');
      try { generateContractPDF_(rec.ssId, rec.code, rec.kenmei); } catch(pe) { Logger.log('Contract PDF error: ' + pe); }
      try { sendApprovalCompleteEmail_(rec.confirmer, rec.confirmerEmail, rec.code, rec.kenmei, '契約内容確認', rec.dept, rec.rowIndex); } catch(ne) { Logger.log('Notify error: ' + ne); }
      writeApprovalRecord_(rec.ssId, '契約内容確認', 'keiyaku', '本部長承認（完了）', '', 'honbucho', '', rec.code, rec.kenmei, '契約内容確認');
      return { success:true, message: '本部長承認完了。PDFを生成しました。' };
    }
  } catch(e) { return { success:false, message: e.message }; }
}

function rejectContract(token, comment) {
  try {
    var rec = findContractApprovalRecord_(token);
    if (!rec) return { success:false, message: '承認依頼が見つかりません。' };
    var ash = getContractApprovalSheet_();
    ash.getRange(rec.sheetRow, 14).setValue('rejected');
    var csh = ensureContractSheet_(rec.ssId);
    var col = rec.step === 'bucho' ? 32 : 35;
    csh.getRange(4, col).setValue('rejected');

    // 確認者に差戻メール
    var confirmerSendEmail = getEmailByName_(rec.confirmer);
    if (confirmerSendEmail) {
      var stepLabel = rec.step === 'bucho' ? '部長' : '本部長';
      MailApp.sendEmail({
        to: confirmerSendEmail,
        subject: '【差戻】契約内容確認 - ' + rec.code + ' ' + rec.kenmei,
        body: rec.confirmer + ' 様\n\n契約内容確認が' + stepLabel + 'により否決されました。\n\n理由: ' + (comment||'(未記入)') + '\n\n修正後、再度承認依頼を送信してください。'
      });
    }
    writeApprovalRecord_(rec.ssId, '契約内容確認', 'keiyaku', '否決', '', rec.step, comment||'', rec.code, rec.kenmei, '契約内容確認');
    return { success:true, message: '否決しました。確認者に差戻メールを送信しました。' };
  } catch(e) { return { success:false, message: e.message }; }
}

// ===== 承認完了通知メール =====
function sendApprovalCompleteEmail_(confirmerName, confirmerEmail, code, kenmei, docType, dept, rowIndex) {
  try {
    var sendTo = getEmailByName_(confirmerName) || confirmerEmail;
    if (!sendTo) return;
    var appUrl = getWebAppUrl();
    var link = (dept && rowIndex) ? appUrl + '?mode=comments&dept=' + encodeURIComponent(dept) + '&row=' + rowIndex : appUrl;
    var subject = '【承認完了】' + docType + ' - ' + code + ' ' + kenmei;
    var htmlBody = '<div style="font-family:\'Segoe UI\',sans-serif;max-width:600px;margin:0 auto;color:#2c3e50;">'
      + '<div style="background:linear-gradient(135deg,#1e8449,#27ae60);color:#fff;padding:24px;border-radius:12px 12px 0 0;">'
      + '<h2 style="margin:0;font-size:18px;">✓ 承認完了のお知らせ</h2></div>'
      + '<div style="background:#fff;border:1px solid #e0e0e0;border-top:none;padding:28px;border-radius:0 0 12px 12px;">'
      + '<p style="font-size:16px;margin:0 0 20px;">' + escHtml_(confirmerName) + ' 様</p>'
      + '<p style="font-size:15px;margin:0 0 20px;line-height:1.7;">以下の申請について、すべての承認が完了しました。PDFが件名フォルダに保存されています。</p>'
      + '<table style="width:100%;border-collapse:collapse;margin:0 0 24px;">'
      + '<tr><td style="padding:12px 16px;background:#f8f9fa;border:1px solid #e0e0e0;font-weight:600;width:100px;">種別</td><td style="padding:12px 16px;border:1px solid #e0e0e0;">' + escHtml_(docType) + '</td></tr>'
      + '<tr><td style="padding:12px 16px;background:#f8f9fa;border:1px solid #e0e0e0;font-weight:600;">件名コード</td><td style="padding:12px 16px;border:1px solid #e0e0e0;">' + escHtml_(code) + '</td></tr>'
      + '<tr><td style="padding:12px 16px;background:#f8f9fa;border:1px solid #e0e0e0;font-weight:600;">件名</td><td style="padding:12px 16px;border:1px solid #e0e0e0;">' + escHtml_(kenmei) + '</td></tr>'
      + '</table>'
      + '<div style="text-align:center;margin:24px 0;">'
      + '<a href="' + link + '" style="display:inline-block;background:#1a5276;color:#fff;text-decoration:none;'
      + 'padding:14px 40px;border-radius:8px;font-size:16px;font-weight:700;letter-spacing:1px;">'
      + '件名の詳細を確認する</a></div>'
      + '<p style="font-size:12px;color:#7f8c8d;margin:16px 0 0;text-align:center;">ボタンが表示されない場合:<br>'
      + '<a href="' + link + '" style="color:#2471a3;word-break:break-all;">' + link + '</a></p>'
      + '</div></div>';
    MailApp.sendEmail({ to: sendTo, subject: subject, body: confirmerName + ' 様\n\n' + docType + '（' + code + ' ' + kenmei + '）の承認が完了しました。\n\n確認: ' + link, htmlBody: htmlBody, name: 'QMS管理システム' });
  } catch(e) { Logger.log('sendApprovalCompleteEmail_ error: ' + e.message); }
}

// ===== 契約内容確認 PDF（全インラインスタイル）=====
function generateContractPDF_(ssId, code, kenmei) {
  var cd = getContractData(ssId);
  if (!cd || !cd.exists) return;
  var e_ = escHtml_;

  var html = pdfHead_('契約内容確認', 'Contract Review', code, kenmei);

  // 基本情報
  html += '<table class="t"><tr><th>受付日</th><td>' + e_(cd.date) + '</td><th>確認者</th><td>' + e_(cd.confirmer) + '</td></tr></table>';

  // 受注関連情報
  html += '<div class="sec">受注関連情報・顧客要求事項</div>';
  html += '<table class="t">';
  if (cd.eigyo) html += '<tr><th>営業受領情報</th><td colspan="3">' + e_(cd.eigyo).replace(/\|/g, '、') + '</td></tr>';
  var ynPairs = [['hacchu','発注内示書又は確認記録'],['sekisan','積算依頼書又は確認記録'],['mitsumori','見積依頼書（仕様書、設計書）又は確認記録'],['shakai','社会的要求事項の確認記録'],['sonota','その他の確認記録']];
  ynPairs.forEach(function(p) { var v = cd[p[0]]; if (v) html += '<tr><th>' + p[1] + '</th><td class="yn" colspan="3">' + e_(v) + '</td></tr>'; });
  html += '</table>';

  // 顧客不満足リスク
  html += '<div class="sec">顧客不満足発生リスクの確認</div>';
  html += '<table class="t">';
  var riskQs = ['顧客要求事項に対し必要な技術を保有しているか','必要な人材・体制を自部署で確保できるか','必要な物的資源を自部署で確保できるか','他部署あるいは委託先の協力により解決できるか','協力先','原子力関連要領を適用するか','仕様書の不利益内容確認'];
  var riskKeys = ['risk1','risk2','risk3','risk4','risk5','risk6','risk7'];
  for (var ri = 0; ri < riskKeys.length; ri++) {
    var rv = cd[riskKeys[ri]]; if (rv) html += '<tr><th>' + riskQs[ri] + '</th><td class="yn" colspan="3">' + e_(rv).replace(/\|/g, '、') + '</td></tr>';
  }
  html += '</table>';

  // リスク表
  var hasRT = cd.riskTable && cd.riskTable.some(function(v) { return v; });
  if (hasRT) {
    var rtNames = ['1.新規顧客','2.新規商材','3.業務内容不明確','4.与信問題','5.契約内容の変更','6.工期の変更','7.確認記録の不備','8.金額不足','9.損害賠償の上限','10.その他'];
    html += '<div class="sec">契約前段階で想定されるリスクの確認・対応結果</div>';
    html += '<table class="t"><tr><th>リスク項目</th><th style="width:60px;">審議</th><th style="width:60px;">有無程度</th><th>対策 / 具体策</th></tr>';
    for (var ti = 0; ti < 10; ti++) {
      var tv = cd.riskTable[ti]; if (!tv) continue;
      var tp = tv.split('|');
      html += '<tr><td class="bold">' + rtNames[ti] + '</td><td class="center">' + e_(tp[0]||'') + '</td><td class="center">' + e_(tp[1]||'') + ' / ' + e_(tp[2]||'') + '</td><td>' + e_(tp[3]||'') + '</td></tr>';
    }
    html += '</table>';
  }

  // 体制選定
  if (cd.tqAnswers || cd.taisei) {
    html += '<div class="sec">成果品チェック体制の選定</div><table class="t">';
    if (cd.tqAnswers) { var tqa = cd.tqAnswers.split('|'); html += '<tr><th>①受注額 / ②重要度 / ③定型</th><td class="yn" colspan="3">' + e_(tqa.join(' / ')) + '</td></tr>'; }
    if (cd.taisei) html += '<tr><th>選定体制</th><td class="yn" colspan="3">' + e_(cd.taisei) + '</td></tr>';
    if (cd.spReason) html += '<tr><th>特別体制理由</th><td colspan="3">' + e_(cd.spReason) + '</td></tr>';
    html += '</table>';
  }

  // 承認欄
  html += '<div class="sec">承認</div><table class="t">';
  html += '<tr><th>部長</th><td class="yn">' + e_(cd.bucho) + '</td><th>承認日</th><td class="yn">' + e_(cd.buchoDate) + '</td></tr>';
  html += '<tr><th>本部長</th><td class="yn">' + e_(cd.honbucho) + '</td><th>承認日</th><td class="yn">' + e_(cd.honbuchoDate) + '</td></tr>';
  html += '</table></body></html>';

  var pdfFile = createPdfFromHtml_(html, '契約内容確認_' + code, ssId);
  lockPdfFile_(pdfFile);
}

function escHtml_(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

// 契約承認ステータスチェック（編集ロック判定用）
function checkContractApprovalStatus(ssId) {
  try {
    var sh = getContractApprovalSheet_();
    var lr = sh.getLastRow();
    if (lr < 2) return { status: 'none' };
    var data = sh.getRange(2, 1, lr-1, 14).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][1]) === ssId && String(data[i][13]) === 'pending') {
        return { status: 'pending', step: String(data[i][11]) };
      }
    }
    return { status: 'none' };
  } catch(e) { return { status: 'none' }; }
}

// 議事録承認ステータスチェック（編集ロック判定用）
function checkMinutesApprovalStatus(ssId, mrow) {
  try {
    var sh = getApprovalSheet_();
    var lr = sh.getLastRow();
    if (lr < 2) return { status: 'none' };
    var data = sh.getRange(2, 1, lr-1, 12).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][1]) === ssId && Number(data[i][2]) === Number(mrow) && String(data[i][11]) === 'pending') {
        return { status: 'pending' };
      }
    }
    return { status: 'none' };
  } catch(e) { return { status: 'none' }; }
}

// ===== 汎用帳票 承認ワークフロー =====
function getGenericApprovalSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID);
  var sh = ss.getSheetByName('汎用承認依頼');
  if (!sh) {
    sh = ss.insertSheet('汎用承認依頼');
    sh.getRange(1,1,1,16).setValues([['token','ssId','dept','rowIndex','code','kenmei','formId','formTitle','confirmer','confirmerEmail','approverName','approverEmail','step','requestDate','status','sheetName']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function sendGenericApproval(ssId, dept, honbu, rowIndex, formId, formTitle, sheetName) {
  try {
    // 承認ルート検証（実施責任者・部長・本部長の全員チェック）
    var route = validateApprovalRoute(dept, honbu, ssId);
    if (!route.valid) return { success:false, message:route.message };

    var email = Session.getActiveUser().getEmail();
    var confirmer = getNameByEmail(email) || email;

    var ss2 = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID);
    var psh = ss2.getSheetByName(dept);
    var row = psh ? psh.getRange(rowIndex, 1, 1, 10).getValues()[0] : [];
    var code = sv(row[3]) || '', kenmei = sv(row[4]) || '';

    var token = Utilities.getUuid();
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');

    // 汎用承認依頼シートにtoken保存
    var ash = getGenericApprovalSheet_();
    ash.appendRow([token, ssId, dept, rowIndex, code, kenmei, formId, formTitle, confirmer, email, route.jisshi.name, route.jisshi.email, 'jisshi', now, 'pending', sheetName]);

    // 実施責任者（最下位）にメール送信
    sendGenericApprovalMail_(token, route.jisshi.name, route.jisshi.email, formTitle, code, kenmei, confirmer, '実施責任者');
    writeApprovalRecord_(ssId, formTitle, formId, '承認依頼送信（実施責任者）', confirmer, 'jisshi', '', code, kenmei, sheetName);
    return { success:true, message:'実施責任者（'+route.jisshi.name+'）に承認依頼を送信しました。' };
  } catch(e) { return { success:false, message:e.message }; }
}

// 承認メール共通
function sendGenericApprovalMail_(token, toName, toEmail, formTitle, code, kenmei, confirmer, stepLabel) {
  var appUrl = getWebAppUrl();
  var link = appUrl + '?genericApprove=true&token=' + token;
  var e_ = escHtml_;
  var htmlBody = '<div style="font-family:sans-serif;max-width:600px;margin:0 auto;">'
    + '<div style="background:linear-gradient(135deg,#0f3460,#16213e);color:#fff;padding:20px 24px;border-radius:8px 8px 0 0;">'
    + '<h2 style="margin:0;font-size:18px;">承認依頼（'+e_(stepLabel)+'）</h2></div>'
    + '<div style="padding:20px 24px;border:1px solid #e0e0e0;border-top:none;">'
    + '<table style="width:100%;border-collapse:collapse;margin-bottom:16px;">'
    + '<tr><td style="padding:8px 12px;background:#f8f9fa;font-weight:bold;width:100px;border:1px solid #e0e0e0;">種別</td>'
    + '<td style="padding:8px 12px;border:1px solid #e0e0e0;">'+e_(formTitle)+'</td></tr>'
    + '<tr><td style="padding:8px 12px;background:#f8f9fa;font-weight:bold;border:1px solid #e0e0e0;">件名コード</td>'
    + '<td style="padding:8px 12px;border:1px solid #e0e0e0;">'+e_(code)+'</td></tr>'
    + '<tr><td style="padding:8px 12px;background:#f8f9fa;font-weight:bold;border:1px solid #e0e0e0;">件名</td>'
    + '<td style="padding:8px 12px;border:1px solid #e0e0e0;">'+e_(kenmei)+'</td></tr>'
    + '<tr><td style="padding:8px 12px;background:#f8f9fa;font-weight:bold;border:1px solid #e0e0e0;">確認者</td>'
    + '<td style="padding:8px 12px;border:1px solid #e0e0e0;">'+e_(confirmer)+'</td></tr></table>'
    + '<p style="text-align:center;"><a href="'+link+'" style="display:inline-block;padding:14px 40px;background:#0f3460;color:#fff;text-decoration:none;border-radius:6px;font-weight:bold;font-size:15px;">承認画面を開く</a></p>'
    + '</div></div>';
  MailApp.sendEmail({ to:toEmail, subject:'【承認依頼】'+formTitle+' - '+code+' '+kenmei, htmlBody:htmlBody });
}

function getGenericApprovalData(token) {
  try {
    var ash = getGenericApprovalSheet_();
    var lr = ash.getLastRow();
    if (lr < 2) return { error:'承認依頼が見つかりません。' };
    var data = ash.getRange(2,1,lr-1,16).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === token) {
        var rec = { sheetRow:i+2, token:sv(data[i][0]), ssId:sv(data[i][1]), dept:sv(data[i][2]), rowIndex:Number(data[i][3]),
          code:sv(data[i][4]), kenmei:sv(data[i][5]), formId:sv(data[i][6]), formTitle:sv(data[i][7]),
          confirmer:sv(data[i][8]), confirmerEmail:sv(data[i][9]), step:sv(data[i][12]), status:sv(data[i][14]), sheetName:sv(data[i][15]) };
        if (rec.status !== 'pending') return { error:'この承認依頼は既に処理済みです。' };
        var formData = getGenericFormData(rec.ssId, rec.sheetName);
        return { rec:rec, formData:formData };
      }
    }
    return { error:'承認依頼が見つかりません。' };
  } catch(e) { return { error:e.message }; }
}

function approveGenericForm(token) {
  try {
    var ash = getGenericApprovalSheet_();
    var lr = ash.getLastRow();
    if (lr < 2) return { success:false, message:'承認依頼が見つかりません。' };
    var data = ash.getRange(2,1,lr-1,16).getValues();
    var rec = null, recRow = -1;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === token) { rec = data[i]; recRow = i+2; break; }
    }
    if (!rec) return { success:false, message:'承認依頼が見つかりません。' };
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
    var step = sv(rec[12]);
    ash.getRange(recRow, 15).setValue('approved');

    var ssId = sv(rec[1]), dept = sv(rec[2]), formId = sv(rec[6]), formTitle = sv(rec[7]);
    var confirmer = sv(rec[8]), confirmerEmail = sv(rec[9]), sheetName = sv(rec[15]);
    var code = sv(rec[4]), kenmei = sv(rec[5]);
    var honbu = getHonbuByDept(dept);

    if (step === 'jisshi') {
      // 実施責任者承認 → 部長へ
      var ap = getApproversByDept(dept, honbu);
      if (ap.bucho && ap.bucho.email) {
        var newToken = Utilities.getUuid();
        var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
        ash.appendRow([newToken, ssId, dept, Number(rec[3]), code, kenmei, formId, formTitle, confirmer, confirmerEmail, ap.bucho.name, ap.bucho.email, 'bucho', nowStr, 'pending', sheetName]);
        sendGenericApprovalMail_(newToken, ap.bucho.name, ap.bucho.email, formTitle, code, kenmei, confirmer, '部長');
        writeApprovalRecord_(ssId, formTitle, formId, '実施責任者承認', sv(rec[10]), 'jisshi', '', code, kenmei, sheetName);
        return { success:true, message:'実施責任者承認完了。部長（'+ap.bucho.name+'）へ承認依頼を送信しました。' };
      }
      writeApprovalRecord_(ssId, formTitle, formId, '実施責任者承認', sv(rec[10]), 'jisshi', '', code, kenmei, sheetName);
      return { success:true, message:'実施責任者承認完了。（部長情報なし）' };

    } else if (step === 'bucho') {
      // 部長承認 → 本部長へ
      var ap2 = getApproversByDept(dept, honbu);
      if (ap2.honbucho && ap2.honbucho.email) {
        var newToken2 = Utilities.getUuid();
        var nowStr2 = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
        ash.appendRow([newToken2, ssId, dept, Number(rec[3]), code, kenmei, formId, formTitle, confirmer, confirmerEmail, ap2.honbucho.name, ap2.honbucho.email, 'honbucho', nowStr2, 'pending', sheetName]);
        sendGenericApprovalMail_(newToken2, ap2.honbucho.name, ap2.honbucho.email, formTitle, code, kenmei, confirmer, '本部長');
        writeApprovalRecord_(ssId, formTitle, formId, '部長承認', sv(rec[10]), 'bucho', '', code, kenmei, sheetName);
        return { success:true, message:'部長承認完了。本部長（'+ap2.honbucho.name+'）へ承認依頼を送信しました。' };
      }
      writeApprovalRecord_(ssId, formTitle, formId, '部長承認', sv(rec[10]), 'bucho', '', code, kenmei, sheetName);
      return { success:true, message:'部長承認完了。（本部長情報なし）' };

    } else {
      // 本部長承認 → 完了 + PDF + 通知
      try {
        var formData = getGenericFormData(ssId, sheetName);
        generateGenericPDF_(ssId, code, kenmei, formTitle, formData);
      } catch(pe) { Logger.log('Generic PDF error: ' + pe); }
      try {
        sendApprovalCompleteEmail_(confirmer, confirmerEmail, code, kenmei, formTitle, dept, Number(rec[3]));
      } catch(ne) { Logger.log('Notify error: ' + ne); }
      writeApprovalRecord_(ssId, formTitle, formId, '本部長承認（完了）', sv(rec[10]), 'honbucho', '', code, kenmei, sheetName);
      return { success:true, message:'本部長承認完了。PDFを生成しました。' };
    }
  } catch(e) { return { success:false, message:e.message }; }
}

function rejectGenericForm(token, comment) {
  try {
    var ash = getGenericApprovalSheet_();
    var lr = ash.getLastRow();
    if (lr < 2) return { success:false, message:'承認依頼が見つかりません。' };
    var data = ash.getRange(2,1,lr-1,16).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === token) {
        ash.getRange(i+2, 15).setValue('rejected');
        var sendTo = getEmailByName_(sv(data[i][8])) || sv(data[i][9]);
        if (sendTo) {
          MailApp.sendEmail({ to:sendTo, subject:'【差戻】'+sv(data[i][7])+' - '+sv(data[i][4])+' '+sv(data[i][5]),
            body:sv(data[i][8])+' 様\n\n'+sv(data[i][7])+'が'+getStepLabel_(sv(data[i][12]))+'により否決されました。\n\n理由: '+(comment||'(未記入)')+'\n\n修正後、再度承認依頼を送信してください。' });
        }
        writeApprovalRecord_(sv(data[i][1]), sv(data[i][7]), sv(data[i][6]), '否決（'+getStepLabel_(sv(data[i][12]))+'）', sv(data[i][10]), sv(data[i][12]), comment||'', sv(data[i][4]), sv(data[i][5]), sv(data[i][15]));
        return { success:true, message:'否決しました。確認者に差戻メールを送信しました。' };
      }
    }
    return { success:false, message:'承認依頼が見つかりません。' };
  } catch(e) { return { success:false, message:e.message }; }
}

function getStepLabel_(step) {
  if (step === 'jisshi') return '実施責任者';
  if (step === 'bucho') return '部長';
  if (step === 'honbucho') return '本部長';
  return step;
}

// 承認状態チェック（汎用承認依頼シートから該当帳票のpendingを検索）
function checkGenericApprovalStatus(ssId, formId) {
  try {
    var ash = getGenericApprovalSheet_();
    var lr = ash.getLastRow();
    if (lr < 2) return { status:'none' };
    var data = ash.getRange(2,1,lr-1,16).getValues();
    for (var i = data.length-1; i >= 0; i--) {
      if (String(data[i][1]) === ssId && String(data[i][6]) === formId && String(data[i][14]) === 'pending') {
        return { status:'pending', step:sv(data[i][12]) };
      }
    }
    return { status:'none' };
  } catch(e) { return { status:'none' }; }
}

// 承認ルート検証（実施責任者・部長・本部長の全員チェック）
function validateApprovalRoute(dept, honbu, ssId) {
  try {
    var missing = [];
    // 実施責任者（グループ体制設定から）
    var jisshi = null;
    if (ssId) {
      var gt = getGroupTaisei(ssId);
      if (gt.jisshi) {
        var jEmail = getStaffEmail(gt.jisshi);
        if (jEmail) jisshi = { name:gt.jisshi, email:jEmail };
      }
    }
    if (!jisshi) missing.push('実施責任者');
    // 部長・本部長
    var ap = getApproversByDept(dept, honbu);
    if (!ap.bucho || !ap.bucho.email) missing.push('部長');
    if (!ap.honbucho || !ap.honbucho.email) missing.push('本部長');
    if (missing.length) return { valid:false, message:missing.join('・')+'の情報が見つかりません。グループ体制設定と電話帳を確認してください。' };
    return { valid:true, jisshi:jisshi, bucho:ap.bucho, honbucho:ap.honbucho };
  } catch(e) { return { valid:false, message:e.message }; }
}

// ===== 承認記録（各件名SSに記録）=====
function getApprovalRecordSheet_(ssId) {
  var ss = SpreadsheetApp.openById(ssId);
  var sh = ss.getSheetByName('承認記録');
  if (!sh) {
    sh = ss.insertSheet('承認記録');
    sh.getRange(1,1,1,10).setValues([['日時','帳票種別','帳票ID','アクション','実行者','承認ステップ','コメント','件名コード','件名','シート名']]);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 140);
    sh.setColumnWidth(2, 160);
    sh.setColumnWidth(4, 100);
    sh.setColumnWidth(5, 100);
  }
  return sh;
}

function writeApprovalRecord_(ssId, formTitle, formId, action, actor, step, comment, code, kenmei, sheetName) {
  try {
    var sh = getApprovalRecordSheet_(ssId);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    sh.appendRow([now, formTitle||'', formId||'', action||'', actor||'', step||'', comment||'', code||'', kenmei||'', sheetName||'']);
  } catch(e) { Logger.log('writeApprovalRecord_ error: ' + e.message); }
}

function getApprovalRecords(ssId) {
  try {
    var sh = getApprovalRecordSheet_(ssId);
    var lr = sh.getLastRow();
    if (lr < 2) return [];
    var data = sh.getRange(2, 1, lr-1, 10).getValues();
    return data.map(function(r) {
      return { date:sv(r[0]), formTitle:sv(r[1]), formId:sv(r[2]), action:sv(r[3]), actor:sv(r[4]), step:sv(r[5]), comment:sv(r[6]), code:sv(r[7]), kenmei:sv(r[8]), sheetName:sv(r[9]) };
    }).reverse();
  } catch(e) { return []; }
}

// ===== 報告書アップロード =====
function uploadReport(ssId, fileName, base64Data) {
  try {
    var folderId = getKenmeiFolderId_(ssId);
    if (!folderId) return { success:false, message:'件名フォルダが見つかりません。' };
    var folder = DriveApp.getFolderById(folderId);
    var reportFolder = findOrCreateFolder_(folder, '報告書');
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, 'application/octet-stream', fileName);
    var file = reportFolder.createFile(blob);
    return { success:true, message:'アップロード完了: ' + fileName, url: file.getUrl() };
  } catch(e) { return { success:false, message:e.message }; }
}

// ===== 議事録 =====
function getMinutesList(ssId) {
  try {
    var sh = SpreadsheetApp.openById(ssId).getSheetByName('議事録');
    if (!sh || sh.getLastRow() < 2) return [];
    var data = sh.getRange(2, 1, sh.getLastRow()-1, 13).getValues();
    // PDF検索用: 件名フォルダ
    var folderId = getKenmeiFolderId_(ssId);
    var pdfMap = {};
    if (folderId) {
      try {
        var folder = DriveApp.getFolderById(folderId);
        var pdfs = folder.getFilesByType('application/pdf');
        while (pdfs.hasNext()) {
          var pf = pdfs.next();
          var nm = pf.getName();
          var match = nm.match(/議事録_No(\d+)\.pdf/);
          if (match) pdfMap[match[1]] = pf.getUrl();
        }
      } catch(pe) {}
    }
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (!r[0] && !r[1]) continue; // 空行スキップ
      var noStr = String(r[0]);
      var sheetRow = i + 2; // ヘッダー1行なのでデータ行=i+2
      result.push({
        sheetRow: sheetRow, no: sv(r[0]), date: sv(r[1]),
        startTime: ft(r[2]), endTime: ft(r[3]), place: sv(r[4]),
        attendeesOther: sv(r[5]), attendeesUs: sv(r[6]),
        creator: sv(r[9]), responsible: sv(r[10]),
        createDate: sv(r[11]), approvalDate: sv(r[12]),
        pdfUrl: pdfMap[noStr] || ''
      });
    }
    return result;
  } catch(e) { return []; }
}

function getMinutesDetail(ssId, mrow) {
  try {
    var sh = SpreadsheetApp.openById(ssId).getSheetByName('議事録');
    if (!sh) return null;
    var rowNum = Number(mrow);
    if (rowNum < 2 || rowNum > sh.getLastRow()) return null;
    var r = sh.getRange(rowNum, 1, 1, 13).getValues()[0];
    return { sheetRow: rowNum, no:sv(r[0]), date:sv(r[1]), startTime:ft(r[2]), endTime:ft(r[3]),
      place:sv(r[4]), attendeesOther:sv(r[5]), attendeesUs:sv(r[6]),
      materials:sv(r[7]), content:sv(r[8]), creator:sv(r[9]),
      responsible:sv(r[10]), createDate:sv(r[11]), approvalDate:sv(r[12]) };
  } catch(e) { Logger.log('getMinutesDetail error: ' + e.message); return null; }
}

function createMinutes(ssId, data) {
  try {
    var sh = SpreadsheetApp.openById(ssId).getSheetByName('議事録');
    if (!sh) return { success:false, message:'議事録シートがありません。' };
    var lr = sh.getLastRow();
    var nextNo = 1;
    if (lr >= 2) { var nos = sh.getRange(2,1,lr-1,1).getValues(); nos.forEach(function(r){var n=Number(r[0]);if(n>=nextNo)nextNo=n+1;}); }
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
    var newRow = lr + 1;
    var row = [nextNo, data.date||'', data.startTime||'', data.endTime||'', data.place||'',
      data.attendeesOther||'', data.attendeesUs||'', data.materials||'', data.content||'',
      data.creator||'', data.responsible||'', now, ''];
    sh.getRange(newRow, 1, 1, 13).setValues([row]);
    return { success:true, no:nextNo, sheetRow:newRow, message:'議事録No.'+nextNo+'を作成しました。' };
  } catch(e) { return { success:false, message:e.message }; }
}

function updateMinutes(ssId, mrow, data) {
  try {
    var sh = SpreadsheetApp.openById(ssId).getSheetByName('議事録');
    var rowNum = Number(mrow);
    if (rowNum < 2 || rowNum > sh.getLastRow()) return { success:false, message:'指定行が範囲外です。(行:'+mrow+')' };
    sh.getRange(rowNum, 2, 1, 8).setValues([[data.date||'', data.startTime||'', data.endTime||'', data.place||'', data.attendeesOther||'', data.attendeesUs||'', data.materials||'', data.content||'']]);
    if (data.creator) sh.getRange(rowNum, 10).setValue(data.creator);
    if (data.responsible) sh.getRange(rowNum, 11).setValue(data.responsible);
    return { success:true, message:'更新しました。' };
  } catch(e) { return { success:false, message:e.message }; }
}

// ===== 承認依頼シート管理（トークン方式）=====
// 進捗管理表内に「承認依頼」シートを自動作成・取得
function getApprovalSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID);
  var sh = ss.getSheetByName('承認依頼');
  if (!sh) {
    sh = ss.insertSheet('承認依頼');
    sh.getRange(1,1,1,12).setValues([['token','ssId','mrow','dept','row','code','kenmei','responsible','creator','minutesNo','requestDate','status']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// トークンから承認依頼レコードを検索
function findApprovalRecord_(token) {
  var sh = getApprovalSheet_();
  var lr = sh.getLastRow();
  if (lr < 2) return null;
  var data = sh.getRange(2, 1, lr - 1, 12).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(token)) {
      return {
        sheetRow: i + 2,
        token: data[i][0], ssId: data[i][1], mrow: Number(data[i][2]),
        dept: data[i][3], row: Number(data[i][4]),
        code: data[i][5], kenmei: data[i][6],
        responsible: data[i][7], creator: data[i][8],
        minutesNo: sv(data[i][9]), requestDate: sv(data[i][10]),
        status: data[i][11] || 'pending'
      };
    }
  }
  return null;
}

// ===== 承認画面：接続テスト（デプロイ確認用）=====
function testApprovalConnection(token) {
  return {
    ok: true,
    version: 'v3.4',
    receivedToken: String(token || ''),
    tokenLen: String(token || '').length,
    timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss')
  };
}

// ===== 承認画面用：トークンからデータ一括取得（段階別診断付き）=====
function getApprovalData(token) {
  var log = [];
  try {
    log.push('STEP0: token受信=[' + token + '] type=' + typeof token + ' len=' + String(token).length);

    if (!token) return { error: 'トークンが空です。', debug: log };

    // STEP1: 承認依頼シート取得
    log.push('STEP1: 承認依頼シート取得開始');
    var ash = getApprovalSheet_();
    var lr = ash.getLastRow();
    log.push('STEP1: OK。承認依頼シート行数=' + lr);

    if (lr < 2) return { error: '承認依頼シートにデータがありません(行数=' + lr + ')。先に承認依頼メールを送信してください。', debug: log };

    // STEP2: トークン検索
    log.push('STEP2: トークン検索開始。データ範囲=2行目～' + lr + '行目');
    var data = ash.getRange(2, 1, lr - 1, 12).getValues();
    var rec = null;
    for (var i = 0; i < data.length; i++) {
      var cellToken = String(data[i][0]);
      if (i < 3) log.push('STEP2: 行' + (i+2) + ' token=[' + cellToken.substring(0,8) + '...]');
      if (cellToken === String(token)) {
        rec = {
          sheetRow: i + 2,
          token: data[i][0], ssId: String(data[i][1]), mrow: Number(data[i][2]),
          dept: String(data[i][3]), row: Number(data[i][4]),
          code: String(data[i][5]), kenmei: String(data[i][6]),
          responsible: String(data[i][7]), creator: String(data[i][8]),
          minutesNo: sv(data[i][9]), requestDate: sv(data[i][10]),
          status: String(data[i][11]) || 'pending'
        };
        log.push('STEP2: 一致！行' + (i+2));
        break;
      }
    }
    if (!rec) return { error: '承認依頼が見つかりません。(検索token=' + String(token).substring(0,12) + '..., 全' + data.length + '件検索済)', debug: log };

    // STEP3: 件名SS接続
    log.push('STEP3: 件名SS接続開始。ssId=[' + rec.ssId + ']');
    var ss = SpreadsheetApp.openById(rec.ssId);
    log.push('STEP3: OK。SS名=' + ss.getName());

    // STEP4: 議事録シート取得
    log.push('STEP4: 議事録シート取得');
    var sh = ss.getSheetByName('議事録');
    if (!sh) return { error: '「議事録」シートがありません。(SS=' + ss.getName() + ')', debug: log };
    var shLr = sh.getLastRow();
    log.push('STEP4: OK。議事録シート行数=' + shLr);

    // STEP5: mrow行読み込み
    log.push('STEP5: mrow=' + rec.mrow + ' 読み込み開始');
    if (rec.mrow < 2 || rec.mrow > shLr) {
      return { error: '議事録行' + rec.mrow + 'が範囲外(最終行=' + shLr + ')', debug: log };
    }
    var r = sh.getRange(rec.mrow, 1, 1, 13).getValues()[0];
    log.push('STEP5: OK。No=' + r[0] + ', 日付=' + r[1] + ', 作成者=' + r[9]);

    // STEP6: 返却データ構築（全値をsv()で文字列化。Date型がgoogle.script.run経由で返るとnullになるため）
    var result = {
      code: sv(rec.code), kenmei: sv(rec.kenmei),
      no: sv(r[0]), date: sv(r[1]), startTime: ft(r[2]), endTime: ft(r[3]),
      place: sv(r[4]), attendeesOther: sv(r[5]), attendeesUs: sv(r[6]),
      materials: sv(r[7]), content: sv(r[8]), creator: sv(r[9]),
      responsible: sv(r[10]), createDate: sv(r[11]), approvalDate: sv(r[12]),
      status: sv(rec.status), debug: log
    };
    log.push('STEP6: 完了。返却OK');
    return result;
  } catch (e) {
    log.push('CATCH: ' + e.message + ' stack=' + e.stack);
    return { error: 'エラー: ' + e.message, debug: log };
  }
}

// GASエディタから直接実行して承認依頼シートの中身を確認するテスト関数
function testApprovalSheet() {
  var sh = getApprovalSheet_();
  var lr = sh.getLastRow();
  Logger.log('承認依頼シート行数: ' + lr);
  if (lr >= 2) {
    var data = sh.getRange(2, 1, Math.min(lr - 1, 5), 12).getValues();
    data.forEach(function(r, i) {
      Logger.log('行' + (i+2) + ': token=' + String(r[0]).substring(0,8) + '... ssId=' + r[1] + ' mrow=' + r[2] + ' dept=' + r[3] + ' status=' + r[11]);
    });
  }
}

// ===== 承認ワークフロー =====
function getStaffEmail(name) {
  try {
    var rows = SpreadsheetApp.openById(CONFIG.PHONEBOOK_SS_ID).getSheets()[0].getDataRange().getValues().slice(1);
    for (var i = 0; i < rows.length; i++) { if (String(rows[i][5]).trim() === name) return String(rows[i][11]).trim(); } // L列=送信用
    return '';
  } catch(e) { return ''; }
}

function sendApprovalEmail(ssId, mrow, dept, rowIndex, responsible) {
  try {
    var email = getStaffEmail(responsible);
    if (!email) return { success:false, message:responsible + 'のメールアドレスが見つかりません。' };

    // 件名情報を進捗管理表から取得
    var psh = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID).getSheetByName(dept);
    var code = String(psh.getRange(rowIndex, 4).getValue() || '');
    var kenmei = String(psh.getRange(rowIndex, 5).getValue() || '');

    // 議事録シートからNo取得
    var msh = SpreadsheetApp.openById(ssId).getSheetByName('議事録');
    var mrowNum = Number(mrow);
    var minutesNo = msh.getRange(mrowNum, 1).getValue();

    var creatorEmail = getCurrentUserEmail();
    var creatorName = getNameByEmail(creatorEmail);

    // トークン生成＆承認依頼シートに記録
    var token = Utilities.getUuid();
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
    var ash = getApprovalSheet_();
    ash.appendRow([token, ssId, mrow, dept, rowIndex, code, kenmei, responsible, creatorName, minutesNo, now, 'pending']);

    // URL（英数字トークンのみ）
    var url = getWebAppUrl() + '?mode=approve&token=' + token + '&mrow=' + mrow;

    var plainBody = '実施責任者 ' + responsible + ' 様\n\n'
      + '議事録の承認申請がありました。\n\n'
      + '件名コード: ' + code + '\n件名: ' + kenmei + '\n議事録No: ' + minutesNo + '\n申請者: ' + creatorName + '\n\n'
      + '承認内容確認: ' + url + '\n\n社会基盤ユニットQMS管理システム';
    var htmlBody = '<div style="font-family:sans-serif;max-width:600px;margin:0 auto;color:#2c3e50;">'
      + '<div style="background:linear-gradient(135deg,#0e3a56,#1a5276);color:#fff;padding:20px 24px;border-radius:8px 8px 0 0;">'
      + '<h2 style="margin:0;font-size:18px;">社会基盤ユニットQMS管理システム</h2>'
      + '<p style="margin:4px 0 0;font-size:14px;opacity:.8;">議事録 承認依頼</p></div>'
      + '<div style="background:#fff;border:1px solid #dce1e6;border-top:none;padding:24px;border-radius:0 0 8px 8px;">'
      + '<p style="font-size:16px;margin:0 0 16px;">実施責任者 <strong>' + responsible + '</strong> 様</p>'
      + '<p style="font-size:16px;margin:0 0 20px;">議事録の承認申請がありました。内容をご確認のうえ、承認をお願いいたします。</p>'
      + '<table style="width:100%;border-collapse:collapse;margin:0 0 24px;">'
      + '<tr><th style="text-align:left;padding:10px 14px;background:#e8eaed;border:1px solid #bbb;font-size:14px;width:120px;">件名コード</th>'
      + '<td style="padding:10px 14px;border:1px solid #bbb;font-size:14px;">' + code + '</td></tr>'
      + '<tr><th style="text-align:left;padding:10px 14px;background:#e8eaed;border:1px solid #bbb;font-size:14px;">件名</th>'
      + '<td style="padding:10px 14px;border:1px solid #bbb;font-size:14px;">' + kenmei + '</td></tr>'
      + '<tr><th style="text-align:left;padding:10px 14px;background:#e8eaed;border:1px solid #bbb;font-size:14px;">議事録No</th>'
      + '<td style="padding:10px 14px;border:1px solid #bbb;font-size:14px;font-weight:700;color:#1a5276;">' + minutesNo + '</td></tr>'
      + '<tr><th style="text-align:left;padding:10px 14px;background:#e8eaed;border:1px solid #bbb;font-size:14px;">申請者</th>'
      + '<td style="padding:10px 14px;border:1px solid #bbb;font-size:14px;">' + creatorName + '</td></tr>'
      + '</table>'
      + '<div style="text-align:center;margin:24px 0;">'
      + '<a href="' + url + '" style="display:inline-block;background:#27ae60;color:#fff;text-decoration:none;'
      + 'padding:14px 40px;border-radius:8px;font-size:16px;font-weight:700;letter-spacing:1px;">'
      + '申請内容を確認する</a></div>'
      + '<p style="font-size:12px;color:#7f8c8d;margin:16px 0 0;text-align:center;">ボタンが表示されない場合は下記URLをブラウザに貼り付けてください。<br>'
      + '<a href="' + url + '" style="color:#2471a3;word-break:break-all;">' + url + '</a></p>'
      + '</div></div>';
    MailApp.sendEmail({ to:email, subject:'【承認依頼】議事録 No.' + minutesNo + ' - ' + kenmei, body:plainBody, htmlBody:htmlBody });

    // 作成日を記入（議事録シート mrow行の12列目）
    msh.getRange(mrowNum, 12).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd'));
    writeApprovalRecord_(ssId, '議事録No.'+minutesNo, 'gijiroku', '承認依頼送信', creatorName, '', '', code, kenmei, '議事録');
    return { success:true, message:'承認依頼メールを ' + responsible + '(' + email + ') に送信しました。' };
  } catch(e) { return { success:false, message:e.message }; }
}

function approveMinutes(token) {
  try {
    var rec = findApprovalRecord_(token);
    if (!rec) return { success:false, message:'承認依頼が見つかりません。' };

    var sh = SpreadsheetApp.openById(rec.ssId).getSheetByName('議事録');
    if (!sh) return { success:false, message:'議事録シートがありません。' };
    if (rec.mrow < 2 || rec.mrow > sh.getLastRow()) return { success:false, message:'議事録の行が範囲外です。' };
    var rowData = sh.getRange(rec.mrow, 1, 1, 13).getValues()[0];

    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
    sh.getRange(rec.mrow, 13).setValue(now);

    // ステータス更新
    var ash = getApprovalSheet_();
    ash.getRange(rec.sheetRow, 12).setValue('approved');

    // PDF作成
    try { generateMinutesPDF_(rec.ssId, rec.mrow, rowData, now, rec.code, rec.kenmei); } catch(pe) { Logger.log('PDF error: ' + pe); }
    // 進捗管理表更新
    try { updateProgressMinutes(rec.dept, rec.row, rec.ssId); } catch(ue) { Logger.log('Progress update error: ' + ue); }
    // 承認完了通知
    try {
      var creatorName2 = sv(rowData[9]);
      var creatorEmail2 = getEmailByName_(creatorName2);
      if (creatorEmail2) sendApprovalCompleteEmail_(creatorName2, creatorEmail2, rec.code, rec.kenmei, '議事録No.' + sv(rowData[0]), rec.dept, rec.row);
    } catch(ne) { Logger.log('Notify error: ' + ne); }
    writeApprovalRecord_(rec.ssId, '議事録No.'+sv(rowData[0]), 'gijiroku', '承認（完了）', rec.responsible||'', '', '', rec.code, rec.kenmei, '議事録');
    return { success:true, message:'承認しました。承認日: ' + now };
  } catch(e) { return { success:false, message:e.message }; }
}

function rejectMinutes(token, comment) {
  try {
    var rec = findApprovalRecord_(token);
    if (!rec) return { success:false, message:'承認依頼が見つかりません。' };

    var sh = SpreadsheetApp.openById(rec.ssId).getSheetByName('議事録');
    if (!sh) return { success:false, message:'議事録シートがありません。' };
    var rowData = sh.getRange(rec.mrow, 1, 1, 13).getValues()[0];
    var creatorName = rowData[9] || '';
    var minutesNo = rowData[0] || '';

    // ステータス更新
    var ash = getApprovalSheet_();
    ash.getRange(rec.sheetRow, 12).setValue('rejected');

    if (!creatorName) return { success:false, message:'作成者が見つかりません。' };
    var email = getStaffEmail(creatorName);
    if (!email) email = creatorName;
    var body = '議事録No.' + minutesNo + 'が否決されました。\n\n'
      + '【否決コメント】\n' + comment + '\n\n'
      + '内容を修正して再申請してください。\n\n社会基盤ユニットQMS管理システム';
    MailApp.sendEmail({ to:email, subject:'【否決】議事録 No.' + minutesNo + ' - ' + rec.kenmei, body:body });
    writeApprovalRecord_(rec.ssId, '議事録No.'+minutesNo, 'gijiroku', '否決', rec.responsible||'', '', comment||'', rec.code, rec.kenmei, '議事録');
    return { success:true, message:'否決しました。作成者に通知メールを送信しました。' };
  } catch(e) { return { success:false, message:e.message }; }
}

// ===== 議事録 PDF =====
function generateMinutesPDF_(ssId, mrow, rowData, approvalDate, code, kenmei) {
  var no = rowData[0];
  var dt = rowData[1] ? fd(rowData[1]) : '';
  var st = ft(rowData[2]), et = ft(rowData[3]);
  var pl = sv(rowData[4]);
  var ao = sv(rowData[5]).replace(/\n/g, '<br>');
  var au = sv(rowData[6]).replace(/\n/g, '<br>');
  var mt = sv(rowData[7]).replace(/\n/g, '<br>');
  var ct = sv(rowData[8]).replace(/\n/g, '<br>');
  var cr = sv(rowData[9]);
  var re = sv(rowData[10]);
  var cd = rowData[11] ? fd(rowData[11]) : '';
  var ad = approvalDate || '';

  var html = pdfHead_('議事録', 'Minutes of Meeting', code, kenmei);

  html += '<table class="t"><tr><th colspan="4" style="text-align:center;font-size:14pt;padding:10px;">No.' + no + '</th></tr></table>';
  html += '<table class="t">'
    + '<tr><th>年月日</th><td colspan="3">' + dt + '</td></tr>'
    + '<tr><th>開始時間</th><td>' + st + '</td><th>終了時間</th><td>' + et + '</td></tr>'
    + '<tr><th>場所</th><td colspan="3">' + pl + '</td></tr>'
    + '<tr><th>出席者（相手）</th><td colspan="3">' + ao + '</td></tr>'
    + '<tr><th>出席者（当社）</th><td colspan="3">' + au + '</td></tr>'
    + '<tr><th>資料</th><td colspan="3">' + mt + '</td></tr>'
    + '</table>';

  html += '<div class="sec">協議内容・確認内容・指示内容・報告内容・処置内容</div>';
  html += '<div class="content">' + ct + '</div>';

  html += '<table class="t">'
    + '<tr><th>作成者</th><td>' + cr + '</td><th>実施責任者</th><td>' + re + '</td></tr>'
    + '<tr><th>作成日</th><td class="yn">' + cd + '</td><th>承認日</th><td class="yn">' + ad + '</td></tr>'
    + '</table>';

  html += '</body></html>';
  var pdfFile = createPdfFromHtml_(html, '議事録_No' + no, ssId);
  lockPdfFile_(pdfFile);
}

// ===== PDF共通: HTMLヘッダー + CSSスタイル =====
function pdfHead_(title, subtitle, code, kenmei) {
  var e_ = escHtml_;
  return '<html><head><meta charset="utf-8"><style>' + pdfCss_() + '</style></head>'
    + '<body>'
    + '<div class="header-bar"><span class="header-code">' + e_(code||'') + '</span><span class="header-date">出力: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd') + '</span></div>'
    + '<div class="doc-title">' + title.split('').join('　') + '</div>'
    + '<div class="doc-subtitle">Quality Management System - ' + (subtitle||'') + '</div>'
    + '<table class="t"><tr><th>件名コード</th><td>' + e_(code||'') + '</td><th>件名</th><td class="bold">' + e_(kenmei||'') + '</td></tr></table>';
}

function pdfCss_() {
  return '@page{size:A4 portrait;margin:12mm 10mm 12mm 10mm;}'
    + 'body{font-family:"Noto Sans JP","Hiragino Kaku Gothic Pro",sans-serif;color:#1a1a2e;margin:0;padding:0;font-size:9pt;line-height:1.5;}'
    + '.header-bar{display:flex;justify-content:space-between;font-size:8pt;color:#888;margin-bottom:4px;}'
    + '.header-code{font-weight:600;}'
    + '.doc-title{text-align:center;font-size:16pt;color:#0f3460;letter-spacing:6px;font-weight:bold;margin:0 0 3px;}'
    + '.doc-subtitle{text-align:center;font-size:8pt;color:#666;margin-bottom:12px;padding-bottom:6px;border-bottom:2px solid #0f3460;}'
    + '.sec{background:#0f3460;color:#fff;padding:6px 12px;font-size:9pt;font-weight:bold;margin:10px 0 0;}'
    + '.content{border:1px solid #c0cad4;padding:10px;font-size:9pt;line-height:1.7;white-space:pre-wrap;min-height:80px;margin-bottom:10px;}'
    + 'table.t{width:100%;border-collapse:collapse;margin-bottom:8px;}'
    + 'table.t th{background:#e8eef4;padding:5px 8px;border:1px solid #c0cad4;font-weight:bold;font-size:8pt;color:#0f3460;text-align:left;width:110px;}'
    + 'table.t td{padding:5px 8px;border:1px solid #c0cad4;font-size:8.5pt;vertical-align:top;}'
    + '.yn{font-weight:bold;color:#0f3460;}'
    + '.bold{font-weight:bold;}'
    + '.center{text-align:center;}';
}

// ===== オンデマンドPDF生成（画面の内容をそのまま印刷）=====
function generateOnDemandPDF(formHtml, title, ssId) {
  try {
    var css = onDemandPdfCss_();
    var fullHtml = '<html><head><meta charset="utf-8"><style>' + css + '</style></head>'
      + '<body>' + formHtml + '</body></html>';
    
    var pdfBlob = HtmlService.createHtmlOutput(fullHtml)
      .getBlob()
      .getAs('application/pdf')
      .setName((title||'帳票') + '.pdf');
    
    // ssIdがあれば件名フォルダに保存、なければ一時フォルダ
    var folder;
    if (ssId) {
      var folderId = getKenmeiFolderId_(ssId);
      var parentFolder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
      folder = findOrCreateFolder_(parentFolder, 'QMS帳票');
    } else {
      folder = DriveApp.getRootFolder();
    }
    
    var fileName = (title||'帳票') + '.pdf';
    var existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);
    var file = folder.createFile(pdfBlob);
    
    return { success: true, url: file.getUrl(), message: 'PDFを生成しました。' };
  } catch(e) {
    return { success: false, message: 'PDF生成エラー: ' + e.message };
  }
}

function onDemandPdfCss_() {
  return '@page{size:A4 portrait;margin:12mm 10mm 12mm 10mm;}'
    + '*{box-sizing:border-box;}'
    + 'body{font-family:"Noto Sans JP","Hiragino Kaku Gothic Pro",sans-serif;color:#1a1a2e;margin:0;padding:10px 16px;font-size:9pt;line-height:1.5;max-width:170mm;}'
    // ヘッダー
    + '.inline-section-header{background:#0f3460;color:#fff;padding:8px 14px;margin-bottom:4px;}'
    + '.inline-section-header h3{color:#fff;font-size:12pt;margin:0;}'
    + '.inline-section-header .material-icons{color:#fff;vertical-align:middle;margin-right:6px;}'
    + '.inline-section-body{padding:4px 0;}'
    + '.btn-close-section,.btn-save,.btn-primary,.btn-secondary,.btn-send,.btn-reject,.btn-approve,.gf-actions,.minutes-actions,.approve-bar,.no-print,.form-page-topbar{display:none !important;}'
    // セクション
    + '.gf-section{border-bottom:1px solid #ddd;padding-bottom:6px;margin-bottom:6px;page-break-inside:avoid;}'
    + '.gf-section-title{font-size:10pt;color:#0f3460;border-left:3px solid #0f3460;padding-left:6px;margin-bottom:4px;font-weight:bold;}'
    // フォーム要素→印刷表示
    + 'input[type="text"],input[type="date"],input[type="number"],textarea,select,.form-input,.form-textarea,.form-select,.gf-table-input,.rt-select{border:none;background:transparent;font-size:9pt;color:#000;padding:2px 0;font-family:inherit;}'
    + 'textarea{white-space:pre-wrap;resize:none;min-height:20px;}'
    + 'input[type="radio"],input[type="checkbox"]{-webkit-appearance:none;appearance:none;width:11px;height:11px;border:1.5px solid #333;display:inline-block;vertical-align:middle;margin:0 2px;position:relative;background:#fff;}'
    + 'input[type="radio"]{border-radius:50%;}'
    + 'input[type="checkbox"]{border-radius:2px;}'
    + 'input[type="radio"]:checked::after{content:"";position:absolute;top:2px;left:2px;width:5px;height:5px;border-radius:50%;background:#000;}'
    + 'input[type="checkbox"]:checked::after{content:"";position:absolute;top:0;left:2px;width:4px;height:7px;border:solid #000;border-width:0 2px 2px 0;transform:rotate(45deg);}'
    // ラベル
    + '.form-group{margin-bottom:3px;}'
    + '.form-group label{font-size:8pt;color:#555;font-weight:700;display:block;margin-bottom:1px;}'
    + '.form-row{display:flex;gap:8px;}'
    + '.form-group.half,.form-group.third{flex:1;}'
    // テーブル
    + '.gf-table{width:100%;border-collapse:collapse;font-size:8pt;}'
    + '.gf-table th{background:#e8eef4;color:#0f3460;padding:4px 6px;border:1px solid #aaa;font-weight:bold;font-size:7.5pt;}'
    + '.gf-table td{border:1px solid #bbb;padding:3px 5px;}'
    + '.gf-table-input,.rt-select{width:100%;border:none;padding:0;font-size:8pt;background:transparent;}'
    // チェックリスト・検査
    + '.gf-check-item{display:flex;justify-content:space-between;align-items:center;padding:4px 8px;border:1px solid #ddd;margin-bottom:2px;page-break-inside:avoid;}'
    + '.gf-check-label{flex:1;font-size:8.5pt;}'
    + '.contract-yn{flex-shrink:0;margin-left:8px;}'
    + '.contract-yn label{font-size:8pt;margin:0 4px;}'
    + '.inspection-item{border:1px solid #ddd;margin-bottom:2px;page-break-inside:avoid;}'
    + '.inspection-item-top{display:flex;align-items:center;gap:6px;padding:4px 8px;}'
    + '.inspection-target{font-size:7.5pt;flex-shrink:0;}'
    + '.inspection-label{flex:1;font-size:8.5pt;}'
    + '.inspection-plan{padding:2px 8px 4px 30px;font-size:8pt;display:flex;align-items:center;gap:4px;}'
    // 契約フォーム
    + '.contract-doc-title{font-size:14pt;color:#0f3460;text-align:center;padding:8px 0;border-bottom:2px solid #0f3460;margin-bottom:8px;}'
    + '.contract-section{border:1px solid #ccc;margin-bottom:6px;page-break-inside:avoid;}'
    + '.contract-section-title{background:#0f3460;color:#fff;padding:5px 10px;font-size:9pt;font-weight:bold;}'
    + '.contract-section-body{padding:5px 10px;}'
    + '.contract-item{display:flex;align-items:flex-start;gap:6px;padding:3px 0;}'
    + '.contract-item-label{flex:1;font-size:8.5pt;}'
    + '.contract-check-label{display:flex;align-items:center;gap:4px;font-size:8.5pt;padding:2px 0;}'
    // 議事録
    + '.minutes-doc-title{font-size:14pt;color:#0f3460;text-align:center;padding:8px 0;border-bottom:2px solid #0f3460;margin-bottom:8px;}'
    + '.minutes-table{width:100%;border-collapse:collapse;margin-bottom:6px;}'
    + '.minutes-table th{background:#e8eef4;color:#0f3460;padding:5px 8px;border:1px solid #aaa;font-weight:bold;font-size:8pt;text-align:left;width:100px;}'
    + '.minutes-table td{border:1px solid #bbb;padding:5px 8px;font-size:9pt;}'
    + '.minutes-content-section{page-break-inside:avoid;}'
    + '.minutes-content-label{background:#0f3460;color:#fff;padding:5px 10px;font-size:9pt;font-weight:bold;}'
    + '.minutes-content,.mf-editable{border:1px solid #bbb;padding:8px;font-size:9pt;white-space:pre-wrap;min-height:60px;}'
    // リスクテーブル
    + '.risktable-legend{background:#f5f5f5;border:1px solid #ddd;padding:4px 8px;font-size:7.5pt;page-break-inside:avoid;}'
    + '.risktable-legend-row{margin-bottom:2px;}'
    + '.risktable-legend-label{display:inline-block;width:60px;font-weight:bold;}'
    // ラジオグループ
    + '.gf-radio-group{display:flex;gap:12px;font-size:8.5pt;}'
    + '.gf-radio{display:flex;align-items:center;gap:3px;}'
    // その他
    + '.other-row{display:none;}'
    + '.saved-badge,.readonly-badge{display:none;}'
    + '.fe-node-bar{width:5px;}'
    + 'h2,h3{page-break-after:avoid;}'
    + '.gf-section,.contract-section,.inspection-item{page-break-inside:avoid;}';
}

// ===== PDF共通ヘルパー（HtmlService直接PDF変換）=====
function createPdfFromHtml_(html, fileName, ssId) {
  // HtmlService.createHtmlOutput → Blob → PDF（Google Docs経由不要、CSS完全対応）
  var pdfBlob = HtmlService.createHtmlOutput(html)
    .getBlob()
    .getAs('application/pdf')
    .setName(fileName + '.pdf');

  // 件名フォルダ内の「QMS帳票」サブフォルダに保存
  var folderId = getKenmeiFolderId_(ssId);
  var parentFolder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var qmsFolder = findOrCreateFolder_(parentFolder, 'QMS帳票');

  // 同名ファイルがあれば上書き（旧ファイルをゴミ箱へ）
  var existing = qmsFolder.getFilesByName(fileName + '.pdf');
  while (existing.hasNext()) existing.next().setTrashed(true);

  var pdfFile = qmsFolder.createFile(pdfBlob);
  return pdfFile;
}

function lockPdfFile_(pdfFile) {
  try {
    // コンテンツ制限を設定（閲覧のみ、編集不可）
    Drive.Files.update(
      { contentRestrictions: [{ readOnly: true, reason: 'QMS承認済み文書' }] },
      pdfFile.getId()
    );
  } catch(e) {
    // Drive API v3のcontentRestrictionsが使えない場合はフォールバック
    Logger.log('lockPdfFile_ contentRestrictions failed: ' + e.message + '. Trying setShareableByEditors.');
    try {
      pdfFile.setShareableByEditors(false);
    } catch(e2) { Logger.log('lockPdfFile_ fallback failed: ' + e2.message); }
  }
}

// ===== 進捗管理表の議事録列更新 =====
function updateProgressMinutes(dept, rowIndex, ssId) {
  try {
    var minutes = getMinutesList(ssId);
    if (!minutes.length) return;
    // 承認済みの最新を探す
    var latest = null;
    for (var i = minutes.length - 1; i >= 0; i--) {
      if (minutes[i].approvalDate) { latest = minutes[i]; break; }
    }
    if (!latest) latest = minutes[minutes.length - 1];
    var sh = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID).getSheetByName(dept);
    var colNo = CONFIG.COL_DETAIL_START + 22; // AF列: gijiroku_no
    var colDate = CONFIG.COL_DETAIL_START + 23; // AG列: gijiroku_date
    sh.getRange(rowIndex, colNo).setValue(latest.no);
    sh.getRange(rowIndex, colDate).setValue(latest.approvalDate || '');
  } catch(e) { Logger.log('updateProgressMinutes error: ' + e); }
}

// ===== 受注データインポート =====
function importOrderData() {
  try {
    var folder = DriveApp.getFolderById(CONFIG.ORDER_FOLDER_ID);
    var latestFile = null, latestYM = '';
    ['application/vnd.ms-excel','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'].forEach(function(mime){
      var files = folder.getFilesByType(mime);
      while (files.hasNext()) { var f = files.next(), nm = f.getName(), idx = nm.indexOf('_');
        if (idx > 0) { var ym = nm.substring(0,idx); if (/^\d{6}$/.test(ym) && ym > latestYM) { latestYM=ym; latestFile=f; } } }
    });
    if (!latestFile) return { success:false, message:'Excelファイルが見つかりません。' };
    var tempFile = Drive.Files.create({ name:'temp_'+latestFile.getName(), mimeType:'application/vnd.google-apps.spreadsheet' }, latestFile.getBlob(), { supportsAllDrives:true });
    var tempSsId = tempFile.id;
    var tempSheet = SpreadsheetApp.openById(tempSsId).getSheetByName('累計');
    if (!tempSheet) { DriveApp.getFileById(tempSsId).setTrashed(true); return { success:false, message:'「累計」シートが見つかりません。' }; }
    var dataRows = tempSheet.getDataRange().getValues().slice(1);
    var tD = ['耐震技術部','技術開発部','土木設計部','風力技術部','バックエンド技術部','地下開発技術部'];
    var dD = {}; tD.forEach(function(d){ dD[d] = []; });
    var EC = CONFIG.EXCEL_COLS;
    dataRows.forEach(function(row){
      var bu = String(row[EC.BU_NAME]).trim(); if (tD.indexOf(bu) < 0) return;
      var kk = String(row[EC.KYAKUSAKI_KUBUN]).trim(), kn = String(row[EC.KOKYAKU_NAME]).trim();
      dD[bu].push({ buName:bu, code:String(row[EC.KENMEI_CODE]).trim(), kenmeiName:String(row[EC.KEIYAKU_KENMEI]).trim(),
        kikiStart:row[EC.KOKI_START], kikiEnd:row[EC.KOKI_END],
        juchuAmount:Math.round((Number(row[EC.TOUNENDO_JUCHU])||0)/1000),
        contractor:kk==='民間企業'?kn:kk+' '+kn });
    });
    var pss = SpreadsheetApp.openById(CONFIG.PROGRESS_SS_ID); var total = 0;
    tD.forEach(function(dept){
      var sh = pss.getSheetByName(dept); if (!sh) return;
      var recs = dD[dept]; if (!recs.length) return;
      var lr = sh.getLastRow();
      if (lr >= CONFIG.DATA_START_ROW) sh.getRange(CONFIG.DATA_START_ROW, 2, lr-CONFIG.DATA_START_ROW+1, CONFIG.COL_IMPORT_COUNT).clearContent();
      var wd = recs.map(function(rec,idx){
        var sd = rec.kikiStart, ed = rec.kikiEnd;
        if (sd instanceof Date) sd = Utilities.formatDate(sd, Session.getScriptTimeZone(), 'yyyy/MM/dd');
        if (ed instanceof Date) ed = Utilities.formatDate(ed, Session.getScriptTimeZone(), 'yyyy/MM/dd');
        return [idx+1, rec.buName, rec.code, rec.kenmeiName, sd, ed, rec.juchuAmount, rec.contractor];
      });
      if (wd.length) { sh.getRange(CONFIG.DATA_START_ROW, 2, wd.length, CONFIG.COL_IMPORT_COUNT).setValues(wd); total += wd.length; }
      // 件名フォルダ・SS自動作成
      recs.forEach(function(rec, idx){
        var ri = CONFIG.DATA_START_ROW + idx;
        var existing = sh.getRange(ri, CONFIG.COL_SS_ID).getValue();
        if (!existing) ensureProjectSS(dept, ri, rec.code, rec.kenmeiName);
      });
    });
    DriveApp.getFileById(tempSsId).setTrashed(true);
    return { success:true, message:'インポート完了。' + latestFile.getName() + ' / ' + total + '件' };
  } catch(e) { Logger.log(e); return { success:false, message:'エラー: ' + e.message }; }
}

// ===== 電話帳 =====
function getPhoneBookData(dept) {
  try {
    var rows = SpreadsheetApp.openById(CONFIG.PHONEBOOK_SS_ID).getSheets()[0].getDataRange().getValues().slice(1);
    var honbu = getHonbuByDept(dept);
    var groups=[], staff=[], approvers=[];
    rows.forEach(function(r){
      var bu=String(r[2]).trim(), grp=String(r[3]).trim(), pos=String(r[4]).trim(), nm=String(r[5]).trim(), hb=String(r[1]).trim();
      if (bu===dept && grp && groups.indexOf(grp)<0) groups.push(grp);
      if (bu===dept && nm) staff.push({ name:nm, group:grp, position:pos });
      if (bu===dept && (pos==='部長'||/^部長\s*●$/.test(pos)) && !approvers.find(function(a){return a.name===nm;}))
        approvers.push({ name:nm, position:pos, type:'部長' });
      if (hb===honbu && (pos==='本部長'||/^本部長\s*●$/.test(pos)) && !approvers.find(function(a){return a.name===nm;}))
        approvers.push({ name:nm, position:pos, type:'本部長' });
    });
    return { groups:groups, staff:staff, approvers:approvers, honbuName:honbu };
  } catch(e) { return { error:e.message }; }
}

// ===== トリガー =====
function setupTriggers(){
  ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='importOrderData')ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger('importOrderData').timeBased().everyDays(1).atHour(6).create();
}
function removeAllTriggers(){ ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); }); }
function onOpen(){
  SpreadsheetApp.getUi().createMenu('QMS管理')
    .addItem('受注データ取込','importOrderData')
    .addSeparator().addItem('トリガー設定','setupTriggers').addItem('トリガー全削除','removeAllTriggers').addToUi();
}
