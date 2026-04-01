// ============================================================
//  押なつ申請 承認ワークフロー  Code.gs  v5
// ============================================================

const PHONEBOOK_ID    = '1GbWVn7HZ7fPWpiv2GTdpMXksU-SBzCM0UdAc7Bwb02M';
const PHONEBOOK_SHEET = '社会基盤ユニット_メールアドレス一覧';
const TOKEN_SECRET    = 'DOROCIVIL_WORKFLOW_SECRET_2024';
const TEST_ADMIN_EMAIL = 'takusari@tepsco.co.jp'; // テスト環境アクセス権限

// A(1):押なつ番号 B(2):押なつ年月日 C(3):記番号 D(4):件名
// E(5):種類 F(6):あて先 G(7):部数 H(8):社員保管者印
// I(9):所属 J(10):氏名 K(11):GM L(12):部長
// M(13):依頼年月日 N(14):GM承認年月日 O(15):部長承認年月日
// P(16):本部名 Q(17):申請者メール R(18):ステータス

function getDataSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}
function isValidSheet(name) {
  if (!name) return false;
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name) !== null;
}

// ── ルーティング ─────────────────────────────────────────────
function doGet(e) {
  const user = Session.getActiveUser().getEmail();
  if (!user) return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;padding:40px">Googleアカウントでのログインが必要です。</h2>');
  const action = (e && e.parameter && e.parameter.action) || 'top';

  if (action === 'top') {
    const tpl = HtmlService.createTemplateFromFile('Top');
    tpl.userEmail = user;
    tpl.isTestAdmin = (user.toLowerCase() === TEST_ADMIN_EMAIL);
    return tpl.evaluate().setTitle('押なつ申請ワークフロー')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport','width=device-width, initial-scale=1');
  }
  if (action === 'form') {
    const tpl = HtmlService.createTemplateFromFile('Form');
    tpl.userEmail = user;
    return tpl.evaluate().setTitle('押なつ申請フォーム')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport','width=device-width, initial-scale=1');
  }
  if (action === 'test') {
    if (user.toLowerCase() !== TEST_ADMIN_EMAIL)
      return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;padding:40px;color:#c00">テスト環境へのアクセス権限がありません。</h2>');
    const tpl = HtmlService.createTemplateFromFile('FormTest');
    tpl.userEmail = user;
    return tpl.evaluate().setTitle('【テスト】押なつ申請フォーム')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport','width=device-width, initial-scale=1');
  }
  if (action === 'approve') {
    const row=parseInt(e.parameter.row,10), token=e.parameter.token||'',
          role=e.parameter.role||'', honbu=e.parameter.honbu||'';
    if (!isValidSheet(honbu))
      return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;padding:40px;color:#c00">シート名が不正です。</h2>');
    if (token !== generateToken(row,role,honbu))
      return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;padding:40px;color:#c00">このURLは無効または期限切れです。</h2>');
    const sheet=getSheet(honbu), rv=sheet.getRange(row,1,1,18).getValues()[0];
    const status=String(rv[17]);
    if (role==='gm' && status!=='申請中')
      return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;padding:40px">この申請はすでに処理済みです。（'+status+'）</h2>');
    if (role==='bucho' && status!=='GM承認済')
      return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;padding:40px">この申請はすでに処理済みです。（'+status+'）</h2>');
    const tpl=HtmlService.createTemplateFromFile('Approval');
    tpl.row=row; tpl.token=token; tpl.role=role; tpl.honbu=honbu; tpl.userEmail=user;
    return tpl.evaluate().setTitle('承認画面')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport','width=device-width, initial-scale=1');
  }
  if (action === 'admin') {
    if (!checkAdminAccess(user))
      return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;padding:40px;color:#c00">アクセス権限がありません。</h2>');
    const tpl=HtmlService.createTemplateFromFile('Admin');
    tpl.userEmail=user;
    return tpl.evaluate().setTitle('押なつ時承認（管理用）')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport','width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;padding:40px">不明なアクションです。</h2>');
}

// ── ユーザー情報取得 ─────────────────────────────────────────
function getUserInfoByEmail(email) {
  if (!email) email = Session.getActiveUser().getEmail();
  const pb=SpreadsheetApp.openById(PHONEBOOK_ID).getSheetByName(PHONEBOOK_SHEET);
  const data=pb.getDataRange().getValues();
  for (let i=1;i<data.length;i++) {
    const rowI=String(data[i][8]||'').trim();
    if (rowI.toLowerCase()===email.toLowerCase()) {
      const honbu=String(data[i][1]||'').trim(), dept=String(data[i][2]||'').trim();
      return { name:String(data[i][5]||'').trim(), dept:dept, honbu:honbu,
               sheetName:honbu||dept, mail:rowI, position:String(data[i][4]||'').trim() };
    }
  }
  return null;
}

// ── トップ画面: 自分の申請一覧 ────────────────────────────────
function getMyApplications(email) {
  if (!email) email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet(), results = [];
  for (const sn of getDataSheetNames()) {
    const sheet = ss.getSheetByName(sn); if (!sheet) continue;
    const lr = sheet.getLastRow(); if (lr < 2) continue;
    const range = sheet.getRange(2, 1, lr - 1, 18);
    const data = range.getValues();
    const dv = range.getDisplayValues(); // E列・G列の日付誤変換対策
    for (let i = 0; i < data.length; i++) {
      // Q列(17列目)=申請者メール でフィルタ
      const rowEmail = String(data[i][16] || '').trim().toLowerCase();
      if (rowEmail !== email.toLowerCase()) continue;

      // ★ B列(押なつ年月日) または H列(社員保管者印) が入力済み → 表示しない
      const colB = data[i][1]; // B列: 押なつ年月日
      const colH = data[i][7]; // H列: 社員保管者印
      const hasStampDate = (colB instanceof Date) || (String(colB || '').trim() !== '');
      const hasStampName = String(colH || '').trim() !== '';
      if (hasStampDate || hasStampName) continue;

      results.push({
        A: safeStr(data[i][0]), C: safeStr(data[i][2]),
        D: safeStr(data[i][3]), E: dv[i][4],
        F: safeStr(data[i][5]), G: dv[i][6],
        requestDate: fmtCellDate(data[i][12]),
        gmDate: fmtCellDate(data[i][13]), buchoDate: fmtCellDate(data[i][14]),
        status: safeStr(data[i][17]), honbu: sn
      });
    }
  }
  // ★ 押なつ番号（A列）の若い順（昇順）でソート
  // 番号形式: "土木_2025-001" → プレフィックスを除去して年度-連番で比較
  results.sort(function(a, b) {
    // プレフィックス部分を除去して年度-連番部分を取得
    var numA = String(a.A).replace(/^[^0-9]*/, '');
    var numB = String(b.A).replace(/^[^0-9]*/, '');
    var pa = numA.split('-'), pb = numB.split('-');
    var ya = parseInt(pa[0], 10) || 0, yb = parseInt(pb[0], 10) || 0;
    if (ya !== yb) return ya - yb;
    var na = parseInt(pa[1], 10) || 0, nb = parseInt(pb[1], 10) || 0;
    return na - nb;
  });
  return results.slice(0, 50);
}

// ── テスト用: 電話帳全ユーザーリスト取得 ─────────────────────
function getAllPhonebookUsers() {
  const pb=SpreadsheetApp.openById(PHONEBOOK_ID).getSheetByName(PHONEBOOK_SHEET);
  const data=pb.getDataRange().getValues(), res=[];
  for (let i=1;i<data.length;i++) {
    const f=String(data[i][5]||'').trim(), e=String(data[i][4]||'').trim();
    if (!f || !e) continue;
    const honbu=String(data[i][1]||'').trim(), dept=String(data[i][2]||'').trim();
    res.push({ name:f, dept:dept, honbu:honbu, sheetName:honbu||dept,
               mail:String(data[i][8]||'').trim(), position:e });
  }
  return res;
}

// ── テスト用: 全所属リスト取得 ──────────────────────────────
function getDeptListAll() {
  const pb=SpreadsheetApp.openById(PHONEBOOK_ID).getSheetByName(PHONEBOOK_SHEET);
  const data=pb.getDataRange().getValues(), set=new Set();
  for (let i=1;i<data.length;i++) {
    const c=String(data[i][2]||'').trim();
    if (c) set.add(c);
  }
  return Array.from(set).sort();
}

// ── 管理画面アクセス権 ───────────────────────────────────────
function checkAdminAccess(email) {
  if (!email) email=Session.getActiveUser().getEmail();
  const pb=SpreadsheetApp.openById(PHONEBOOK_ID).getSheetByName(PHONEBOOK_SHEET);
  const data=pb.getDataRange().getValues();
  for (let i=1;i<data.length;i++) {
    if (String(data[i][8]||'').trim().toLowerCase()!==email.toLowerCase()) continue;
    const rowC=String(data[i][2]||'').trim(), rowE=String(data[i][4]||'').trim();
    if (rowC==='社会基盤企画総括部' && (rowE==='部長'||/^GM/i.test(rowE)||rowE==='メンバー')) return true;
  }
  return false;
}

// ── 承認画面: データ取得 (google.script.run用) ───────────────
function getApprovalData(row, token, role, honbu) {
  if (!isValidSheet(honbu)) return { success:false, message:'シートが不正です。' };
  if (token !== generateToken(row, role, honbu)) return { success:false, message:'トークンが不正です。' };
  const sheet = getSheet(honbu);
  const rv = sheet.getRange(row, 1, 1, 18).getValues()[0];
  // E列・G列はDisplayValueで取得（Sheetsの日付自動変換対策）
  const dv = sheet.getRange(row, 1, 1, 18).getDisplayValues()[0];
  return { success:true, data:{
    A:safeStr(rv[0]), B:fmtCellDate(rv[1]), C:safeStr(rv[2]), D:safeStr(rv[3]),
    E:dv[4], F:safeStr(rv[5]), G:dv[6], H:safeStr(rv[7]),
    I:safeStr(rv[8]), J:safeStr(rv[9]), K:safeStr(rv[10]), L:safeStr(rv[11]),
    requestDate:fmtCellDate(rv[12]), gmApproveDate:fmtCellDate(rv[13]),
    buchoApproveDate:fmtCellDate(rv[14]), honbu:safeStr(rv[15])||honbu, status:safeStr(rv[17])
  }};
}

// ── 管理画面: B列未入力リストアップ ──────────────────────────
function getPendingStampApplications() {
  if (!checkAdminAccess(Session.getActiveUser().getEmail()))
    return {success:false,message:'アクセス権限がありません。'};
  const ss=SpreadsheetApp.getActiveSpreadsheet(), result={};
  for (const sn of getDataSheetNames()) {
    const sheet=ss.getSheetByName(sn); if(!sheet) continue;
    const lr=sheet.getLastRow(); if(lr<2) continue;
    const range=sheet.getRange(2,1,lr-1,18);
    const data=range.getValues(), dv=range.getDisplayValues(), items=[];
    for (let i=0;i<data.length;i++) {
      const rowB=data[i][1], bStr=String(rowB||'').trim(), st=String(data[i][17]||'').trim();
      if (!bStr && !(rowB instanceof Date) && st==='承認完了') {
        items.push({row:i+2, A:safeStr(data[i][0]), C:safeStr(data[i][2]),
          D:safeStr(data[i][3]), E:dv[i][4], F:safeStr(data[i][5]),
          G:dv[i][6], I:safeStr(data[i][8]), J:safeStr(data[i][9]),
          K:safeStr(data[i][10]), L:safeStr(data[i][11]),
          requestDate:fmtCellDate(data[i][12]), honbu:sn});
      }
    }
    if (items.length>0) result[sn]=items;
  }
  return {success:true,data:result};
}

// ── 管理画面: 保管者印 ──────────────────────────────────────
function stampApplication(honbu, row, stamperName) {
  try {
    if (!checkAdminAccess(Session.getActiveUser().getEmail()))
      return {success:false,message:'アクセス権限がありません。'};
    if (!isValidSheet(honbu)) return {success:false,message:'シートが不正です。'};
    const sheet=getSheet(honbu), now=new Date();
    sheet.getRange(row,8).setValue(stamperName);
    sheet.getRange(row,2).setValue(now);
    sheet.getRange(row,2).setNumberFormat('yyyy/mm/dd HH:mm');
    return {success:true, message:'押なつ番号 '+sheet.getRange(row,1).getValue()+' の保管者印と押なつ年月日を記録しました。'};
  } catch(err) { return {success:false,message:'エラー: '+err.message}; }
}

// ── フォーム送信 ─────────────────────────────────────────────
function submitApplication(formData) {
  try {
    const user=Session.getActiveUser().getEmail();
    const sheetName=formData.sheetName||formData.honbu;
    if (!sheetName||!isValidSheet(sheetName))
      return {success:false,message:'シート「'+sheetName+'」が見つかりません。管理者にお問い合わせください。'};
    const sheet=getSheet(sheetName), row=sheet.getLastRow()+1;
    const autoNo=generateApplicationNumber(sheet,row,sheetName), now=new Date();
    formData.A=autoNo;
    sheet.getRange(row,1,1,12).setValues([[
      formData.A,formData.B,formData.C,formData.D,formData.E,formData.F,
      formData.G,formData.H,formData.I,formData.J,formData.K,formData.L
    ]]);
    // E列(種類)とG列(部数)をテキスト形式に強制（日付自動変換を防止）
    sheet.getRange(row,5).setNumberFormat('@');
    sheet.getRange(row,7).setNumberFormat('@');
    sheet.getRange(row,5).setValue(formData.E);
    sheet.getRange(row,7).setValue(formData.G);
    sheet.getRange(row,13).setValue(now);
    sheet.getRange(row,13).setNumberFormat('yyyy/mm/dd hh:mm');
    sheet.getRange(row,16).setValue(sheetName);
    sheet.getRange(row,17).setValue(user);
    sheet.getRange(row,18).setValue('申請中');
    const gmInfo=findByName(formData.K);
    if (!gmInfo) return {success:false,message:'GM「'+formData.K+'」が電話帳に見つかりません。'};
    const token=generateToken(row,'gm',sheetName);
    const url=getWebAppUrl()+'?action=approve&row='+row+'&token='+encodeURIComponent(token)+'&role=gm&honbu='+encodeURIComponent(sheetName);
    sendApprovalRequestEmail(gmInfo.sendAddress,gmInfo.name,formData,'【押なつ申請ワークフロー】GM承認依頼',url);
    return {success:true,message:'申請番号 '+autoNo+' を受け付けました。GMに承認依頼メールを送信しました。'};
  } catch(err) { return {success:false,message:'エラー: '+err.message}; }
}

// ── テスト環境用フォーム送信 ─────────────────────────────────
// シート名は固定で「テスト」、GM/部長メールはすべて TEST_ADMIN_EMAIL に送る
function submitTestApplication(formData) {
  try {
    const user = Session.getActiveUser().getEmail();
    if (user.toLowerCase() !== TEST_ADMIN_EMAIL)
      return {success:false, message:'テスト環境の送信権限がありません。'};

    const sheetName = 'テスト';
    if (!isValidSheet(sheetName))
      return {success:false, message:'シート「テスト」が見つかりません。スプレッドシートに作成してください。'};

    const sheet = getSheet(sheetName), row = sheet.getLastRow() + 1;
    const autoNo = generateApplicationNumber(sheet, row, sheetName), now = new Date();
    formData.A = autoNo;
    sheet.getRange(row,1,1,12).setValues([[
      formData.A,formData.B,formData.C,formData.D,formData.E,formData.F,
      formData.G,formData.H,formData.I,formData.J,formData.K,formData.L
    ]]);
    // E列(種類)とG列(部数)をテキスト形式に強制（日付自動変換を防止）
    sheet.getRange(row,5).setNumberFormat('@');
    sheet.getRange(row,7).setNumberFormat('@');
    sheet.getRange(row,5).setValue(formData.E);
    sheet.getRange(row,7).setValue(formData.G);
    sheet.getRange(row,13).setValue(now);
    sheet.getRange(row,13).setNumberFormat('yyyy/mm/dd hh:mm');
    sheet.getRange(row,16).setValue(sheetName);
    sheet.getRange(row,17).setValue(user);
    sheet.getRange(row,18).setValue('申請中');

    // GM/部長 両方のメールを TEST_ADMIN_EMAIL に送る
    const token = generateToken(row,'gm',sheetName);
    const url = getWebAppUrl()+'?action=approve&row='+row+'&token='+encodeURIComponent(token)+'&role=gm&honbu='+encodeURIComponent(sheetName);
    sendApprovalRequestEmail(TEST_ADMIN_EMAIL, formData.K+' (テスト)', formData, '【押なつ申請ワークフロー】【テスト】GM承認依頼', url);

    return {success:true, message:'【テスト】申請番号 '+autoNo+' をシート「テスト」に書き込みました。\nGM承認メールを '+TEST_ADMIN_EMAIL+' に送信しました。'};
  } catch(err) { return {success:false, message:'エラー: '+err.message}; }
}

// ── 承認 / 否決 ──────────────────────────────────────────────
function processApproval(row,token,role,honbu,decision,comment) {
  try {
    if (!isValidSheet(honbu)) return {success:false,message:'シートが不正です。'};
    if (token!==generateToken(row,role,honbu)) return {success:false,message:'トークンが不正です。'};
    const sheet=getSheet(honbu), rv=sheet.getRange(row,1,1,18).getValues()[0];
    const fd={A:rv[0],B:rv[1],C:rv[2],D:rv[3],E:rv[4],F:rv[5],G:rv[6],H:rv[7],
              I:rv[8],J:rv[9],K:rv[10],L:rv[11],honbu:honbu};
    const applicantEmail=String(rv[16]), now=new Date(), nowStr=formatDate(now);
    const isTest = (honbu === 'テスト');  // テストシートの場合は全メールをTEST_ADMIN_EMAILに送る
    if (role==='gm') {
      if (decision==='approve') {
        sheet.getRange(row,14).setValue(now); sheet.getRange(row,14).setNumberFormat('yyyy/mm/dd hh:mm');
        sheet.getRange(row,18).setValue('GM承認済');
        const bi=findByName(fd.L);
        if (!bi) return {success:false,message:'部長「'+fd.L+'」が電話帳に見つかりません。'};
        const bt=generateToken(row,'bucho',honbu);
        const bu=getWebAppUrl()+'?action=approve&row='+row+'&token='+encodeURIComponent(bt)+'&role=bucho&honbu='+encodeURIComponent(honbu);
        const buchoAddr = isTest ? TEST_ADMIN_EMAIL : bi.sendAddress;
        const subjectPrefix = isTest ? '【押なつ申請ワークフロー】【テスト】' : '【押なつ申請ワークフロー】';
        sendApprovalRequestEmail(buchoAddr,bi.name,fd,subjectPrefix+'部長承認依頼（GM承認済）',bu);
        return {success:true,message:'GM承認が完了しました。部長に承認依頼メールを送信しました。'};
      } else {
        sheet.getRange(row,18).setValue('GM否決');
        sendRejectionEmail(isTest ? TEST_ADMIN_EMAIL : applicantEmail,fd,'GM',comment);
        return {success:true,message:'否決しました。申請者に通知しました。'};
      }
    }
    if (role==='bucho') {
      if (decision==='approve') {
        sheet.getRange(row,15).setValue(now); sheet.getRange(row,15).setNumberFormat('yyyy/mm/dd hh:mm');
        sheet.getRange(row,18).setValue('承認完了');
        const gd=fmtCellDate(sheet.getRange(row,14).getValue());
        sendCompletionEmail(isTest ? TEST_ADMIN_EMAIL : applicantEmail,fd,gd,nowStr);
        return {success:true,message:'部長承認が完了しました。申請者に完了通知を送信しました。'};
      } else {
        sheet.getRange(row,18).setValue('部長否決');
        sendRejectionEmail(isTest ? TEST_ADMIN_EMAIL : applicantEmail,fd,'部長',comment);
        return {success:true,message:'否決しました。申請者に通知しました。'};
      }
    }
    return {success:false,message:'不明なロールです。'};
  } catch(err) { return {success:false,message:'エラー: '+err.message}; }
}

// ── 電話帳から人員リスト ─────────────────────────────────────
function getPersonList(dept, matchType) {
  const pb=SpreadsheetApp.openById(PHONEBOOK_ID).getSheetByName(PHONEBOOK_SHEET);
  const data=pb.getDataRange().getValues(), res=[];
  for (let i=1;i<data.length;i++) {
    const rowB=String(data[i][1]||'').trim(), rowC=String(data[i][2]||'').trim();
    const rowE=String(data[i][4]||'').trim(), rowF=String(data[i][5]||'').trim();
    if (!rowF||!rowE) continue;
    if ((rowC||rowB)!==dept) continue;
    if (matchType==='gm' && !/^GM(\s|$)/i.test(rowE)) continue;
    if (matchType==='bucho' && !/^部長(\s|$)/.test(rowE)) continue;
    res.push({name:rowF,position:rowE});
  }
  return res;
}

function getWebAppUrl() { return ScriptApp.getService().getUrl(); }

// ══════════════════════════════════════════════════════════════
//  メール送信
// ══════════════════════════════════════════════════════════════
function sendApprovalRequestEmail(toAddress,toName,fd,subject,approvalUrl) {
  const html=`<div style="font-family:'Hiragino Kaku Gothic Pro',Meiryo,sans-serif;max-width:680px;margin:0 auto">
  <div style="background:#1a3557;color:#fff;padding:20px 28px;border-radius:8px 8px 0 0"><h2 style="margin:0;font-size:20px">押なつ申請 承認依頼</h2></div>
  <div style="background:#f7f9fc;padding:24px 28px;border:1px solid #dde3ec;border-top:none">
    <p style="margin-bottom:12px">${escHtml(toName)} 様</p>
    <p style="margin-bottom:20px">以下の押なつ申請について、ご承認をお願いいたします。</p>
    <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;width:100%;font-size:14px;background:#fff">${buildTableRows(fd)}</table>
    <div style="margin-top:28px;text-align:center"><a href="${approvalUrl}" style="display:inline-block;background:#1a3557;color:#fff;padding:14px 40px;text-decoration:none;border-radius:6px;font-size:16px;font-weight:bold">▶ 承認画面を開く</a></div>
    <p style="margin-top:24px;font-size:12px;color:#888">※ このメールはシステムから自動送信されています。</p>
  </div></div>`;
  GmailApp.sendEmail(toAddress,subject,toName+' 様\n\n押なつ申請の承認依頼です。\n承認URL: '+approvalUrl,{htmlBody:html,name:'押なつ申請ワークフロー'});
}

function sendRejectionEmail(toAddress,fd,rejector,comment) {
  const subject='【押なつ申請ワークフロー】【否決】 - '+fd.D;
  const html=`<div style="font-family:'Hiragino Kaku Gothic Pro',Meiryo,sans-serif;max-width:680px;margin:0 auto">
  <div style="background:#8b1a1a;color:#fff;padding:20px 28px;border-radius:8px 8px 0 0"><h2 style="margin:0">押なつ申請 否決通知</h2></div>
  <div style="background:#fdf7f7;padding:24px 28px;border:1px solid #e8d0d0;border-top:none">
    <p>申請件名「<strong>${escHtml(fd.D)}</strong>」について、<strong>${escHtml(rejector)}</strong>により否決されました。</p>
    <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;width:100%;font-size:14px;background:#fff;margin-top:16px">${buildTableRows(fd)}</table>
    <p style="margin-top:12px"><strong>コメント：</strong>${escHtml(comment||'なし')}</p>
    <p style="font-size:12px;color:#888;margin-top:20px">※ このメールはシステムから自動送信されています。</p>
  </div></div>`;
  GmailApp.sendEmail(toAddress,subject,'否決されました。',{htmlBody:html,name:'押なつ申請ワークフロー'});
}

function sendCompletionEmail(toAddress,fd,gmDate,buchoDate) {
  const subject='【押なつ申請ワークフロー】【承認完了】 - '+fd.D;
  const html=`<div style="font-family:'Hiragino Kaku Gothic Pro',Meiryo,sans-serif;max-width:680px;margin:0 auto">
  <div style="background:#1a5745;color:#fff;padding:20px 28px;border-radius:8px 8px 0 0"><h2 style="margin:0">押なつ申請 承認完了</h2></div>
  <div style="background:#f7fdf9;padding:24px 28px;border:1px solid #c8e8d8;border-top:none">
    <p>申請件名「<strong>${escHtml(fd.D)}</strong>」がすべての承認者に承認されました。</p>
    <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;width:100%;font-size:14px;background:#fff;margin-top:16px">
      ${buildTableRows(fd)}
      <tr><th style="background:#e8f4ec;text-align:left;width:140px;padding:8px 12px;font-weight:bold">GM承認日時</th><td style="padding:8px 12px">${escHtml(gmDate)}</td></tr>
      <tr><th style="background:#e8f4ec;text-align:left;width:140px;padding:8px 12px;font-weight:bold">部長承認日時</th><td style="padding:8px 12px">${escHtml(buchoDate)}</td></tr>
    </table>
    <p style="font-size:12px;color:#888;margin-top:20px">※ このメールはシステムから自動送信されています。</p>
  </div></div>`;
  GmailApp.sendEmail(toAddress,subject,'承認されました。',{htmlBody:html,name:'押なつ申請ワークフロー'});
}

// ══════════════════════════════════════════════════════════════
//  リマインド
// ══════════════════════════════════════════════════════════════
function sendReminders() {
  const ss=SpreadsheetApp.getActiveSpreadsheet(), now=new Date(), DAY=86400000;
  for (const sn of getDataSheetNames()) {
    try {
      const sheet=ss.getSheetByName(sn); if(!sheet) continue;
      const lr=sheet.getLastRow(); if(lr<2) continue;
      const range=sheet.getRange(2,1,lr-1,18);
      const data=range.getValues();
      const dv=range.getDisplayValues();
      for (var i=0;i<data.length;i++) {
        if (!data[i] || data[i].length < 18) continue;
        var st=String(data[i][17]||'').trim(), row=i+2;
        if (!st) continue;

        var fd={A:safeStr(data[i][0]),B:safeStr(data[i][1]),C:safeStr(data[i][2]),D:safeStr(data[i][3]),
                E:dv[i][4],F:safeStr(data[i][5]),G:dv[i][6],H:safeStr(data[i][7]),
                I:safeStr(data[i][8]),J:safeStr(data[i][9]),K:safeStr(data[i][10]),L:safeStr(data[i][11]),honbu:sn};
        var title = fd.D || '（件名なし）';
        var tableHtml = buildTableRows(fd);

        if (st==='申請中' && data[i][12] instanceof Date && (now-data[i][12])>=DAY) {
          var gmName=String(data[i][10]||'').trim();
          var gmInfo=findByName(gmName);
          if (gmInfo) {
            var gmToken=generateToken(row,'gm',sn);
            var gmUrl=getWebAppUrl()+'?action=approve&row='+row+'&token='+encodeURIComponent(gmToken)+'&role=gm&honbu='+encodeURIComponent(sn);
            var gmSubject='【押なつ申請ワークフロー】【リマインド】GM承認依頼 - '+title;
            var gmHtml='<div style="font-family:\'Hiragino Kaku Gothic Pro\',Meiryo,sans-serif;max-width:680px;margin:0 auto">'
              +'<div style="background:#c77600;color:#fff;padding:20px 28px;border-radius:8px 8px 0 0"><h2 style="margin:0;font-size:20px">⏰ 承認リマインド</h2></div>'
              +'<div style="background:#fffbf5;padding:24px 28px;border:1px solid #f0dcc0;border-top:none">'
              +'<p style="margin-bottom:12px">'+escHtml(gmInfo.name)+' 様</p>'
              +'<p style="margin-bottom:20px">未承認のまま1日以上経過しております。ご対応をお願いいたします。</p>'
              +'<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;width:100%;font-size:14px;background:#fff">'+tableHtml+'</table>'
              +'<div style="margin-top:28px;text-align:center"><a href="'+gmUrl+'" style="display:inline-block;background:#c77600;color:#fff;padding:14px 40px;text-decoration:none;border-radius:6px;font-size:16px;font-weight:bold">▶ 承認画面を開く</a></div>'
              +'</div></div>';
            GmailApp.sendEmail(gmInfo.sendAddress,gmSubject,'未承認の申請があります。\n'+gmUrl,{htmlBody:gmHtml,name:'押なつ申請ワークフロー'});
          }
        }
        if (st==='GM承認済' && data[i][13] instanceof Date && (now-data[i][13])>=DAY) {
          var buchoName=String(data[i][11]||'').trim();
          var buchoInfo=findByName(buchoName);
          if (buchoInfo) {
            var buchoToken=generateToken(row,'bucho',sn);
            var buchoUrl=getWebAppUrl()+'?action=approve&row='+row+'&token='+encodeURIComponent(buchoToken)+'&role=bucho&honbu='+encodeURIComponent(sn);
            var buchoSubject='【押なつ申請ワークフロー】【リマインド】部長承認依頼 - '+title;
            var buchoHtml='<div style="font-family:\'Hiragino Kaku Gothic Pro\',Meiryo,sans-serif;max-width:680px;margin:0 auto">'
              +'<div style="background:#c77600;color:#fff;padding:20px 28px;border-radius:8px 8px 0 0"><h2 style="margin:0;font-size:20px">⏰ 承認リマインド</h2></div>'
              +'<div style="background:#fffbf5;padding:24px 28px;border:1px solid #f0dcc0;border-top:none">'
              +'<p style="margin-bottom:12px">'+escHtml(buchoInfo.name)+' 様</p>'
              +'<p style="margin-bottom:20px">未承認のまま1日以上経過しております。ご対応をお願いいたします。</p>'
              +'<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;width:100%;font-size:14px;background:#fff">'+tableHtml+'</table>'
              +'<div style="margin-top:28px;text-align:center"><a href="'+buchoUrl+'" style="display:inline-block;background:#c77600;color:#fff;padding:14px 40px;text-decoration:none;border-radius:6px;font-size:16px;font-weight:bold">▶ 承認画面を開く</a></div>'
              +'</div></div>';
            GmailApp.sendEmail(buchoInfo.sendAddress,buchoSubject,'未承認の申請があります。\n'+buchoUrl,{htmlBody:buchoHtml,name:'押なつ申請ワークフロー'});
          }
        }
      }
    } catch(e) {
      Logger.log('sendReminders error on sheet "'+sn+'": '+e.message);
    }
  }
}
function setupReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(t=>{if(t.getHandlerFunction()==='sendReminders')ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger('sendReminders').timeBased().everyDays(1).atHour(9).create();
}

// ══════════════════════════════════════════════════════════════
//  ユーティリティ
// ══════════════════════════════════════════════════════════════
function getSheet(n) { const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); if(!s) throw new Error('シート「'+n+'」が見つかりません。'); return s; }
function formatDate(d) { return Utilities.formatDate(d,'Asia/Tokyo','yyyy/MM/dd HH:mm'); }
function fmtCellDate(v) { if(!v && v!==0) return ''; if(v instanceof Date) return formatDate(v); return String(v); }
function safeStr(v) { if(v===null||v===undefined||v==='') return ''; if(v instanceof Date) return formatDate(v); return String(v); }
function generateToken(row,role,honbu) {
  const d=Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,row+'-'+role+'-'+honbu+'-'+TOKEN_SECRET);
  return d.map(b=>('0'+(b&0xff).toString(16)).slice(-2)).join('');
}
function getHonbuPrefix(sheetName) {
  const map = {
    '土木本部': '土木_',
    'ジオフロント本部': 'ジオ_',
    '社会基盤企画総括部': '企画総括_',
    'テスト': 'テスト_'
  };
  return map[sheetName] || '';
}
function generateApplicationNumber(sheet,newRow,sheetName) {
  const now=new Date(), fy=now.getMonth()>=3?now.getFullYear():now.getFullYear()-1;
  const prefix=getHonbuPrefix(sheetName);
  const base=fy+'-';
  let c=0;
  if(newRow>2){sheet.getRange(2,1,newRow-2,1).getValues().forEach(r=>{if(String(r[0]||'').startsWith(prefix+base))c++;});}
  return prefix+base+String(c+1).padStart(3,'0');
}
function findByName(name) {
  if(!name) return null;
  const d=SpreadsheetApp.openById(PHONEBOOK_ID).getSheetByName(PHONEBOOK_SHEET).getDataRange().getValues();
  for(let i=1;i<d.length;i++){if(String(d[i][5]).trim()===String(name).trim())
    return {name:String(d[i][5]),sendAddress:String(d[i][11]),mail:String(d[i][8])};}
  return null;
}
function buildTableRows(fd) {
  return [['押なつ番号',fd.A],['記番号',fd.C],['件名',fd.D],['種類',fd.E],['部数',fd.G],['あて先',fd.F],
    ['本部',fd.honbu||''],['所属',fd.I],['氏名',fd.J],['GM',fd.K],['部長',fd.L]]
    .map(([l,v])=>`<tr><th style="background:#eef2f8;text-align:left;width:140px;padding:8px 12px;font-weight:bold">${escHtml(l)}</th><td style="padding:8px 12px">${escHtml(safeStr(v))}</td></tr>`).join('');
}
function escHtml(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function include(f) { return HtmlService.createHtmlOutputFromFile(f).getContent(); }
