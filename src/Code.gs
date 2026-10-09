/**
 * 銀革屋 革在庫管理 Apps Script 完成版
 *
 * このファイルを Googleスプレッドシート → 拡張機能 → Apps Script に貼り付けて使います。
 * 非エンジニア運用前提：在庫数は勝手に増やさず、確認済み・反映OKになったものだけ反映します。
 *
 * 重要ルール：
 * - 革の在庫は枚数ではなく ds数で管理
 * - 1マス＝1行表示を前提
 * - PDF画像読み取り結果は候補。必ず確認してから反映
 */


// ============================================================
// Source: ginkawaya_leather_inventory_apps_script_v2_precision.gs
// ============================================================

const SPREADSHEET_ID = '1JKYaPXUfZvDKApvOY3rQ--YmpeuuXlC_zqVt644pVg0';
const GMAIL_LABEL = '革仕入';

/**
 * 銀革屋 革在庫管理 Apps Script v2
 * 目的：
 * - Gmailの「革仕入」ラベル付きメールから、革・資材らしい行だけを入荷予定へ取込
 * - 送料、漉き代、合計、消費税、返信引用、過去メールを取り込みにくくする
 * - 在庫マスターへは勝手に反映しない安全設計
 *
 * 使う関数：
 * 1. testConnection
 * 2. importLeatherEmails
 * 3. reflectConfirmedArrivals
 * 4. createHourlyTrigger
 * 5. deleteImportTriggers
 */

function testConnection() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const label = GmailApp.getUserLabelByName(GMAIL_LABEL);

  if (!ss.getSheetByName('入荷予定')) throw new Error('「入荷予定」シートが見つかりません。');
  if (!ss.getSheetByName('メール取込ログ')) throw new Error('「メール取込ログ」シートが見つかりません。');
  if (!ss.getSheetByName('在庫マスター')) throw new Error('「在庫マスター」シートが見つかりません。');
  if (!ss.getSheetByName('入出庫履歴')) throw new Error('「入出庫履歴」シートが見つかりません。');
  if (!label) throw new Error('Gmailに「' + GMAIL_LABEL + '」ラベルが見つかりません。');

  Logger.log('接続テスト成功');
  Logger.log('スプレッドシート名: ' + ss.getName());
  Logger.log('Gmailラベル名: ' + label.getName());
}

/**
 * Gmailの「革仕入」ラベル付きメールを「入荷予定」へ取り込む
 * v2：返信引用、送料、漉き代、合計行をより強く除外
 */
function importLeatherEmails() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const importSheet = ss.getSheetByName('入荷予定');
  const logSheet = ss.getSheetByName('メール取込ログ');
  const label = GmailApp.getUserLabelByName(GMAIL_LABEL);

  if (!importSheet) throw new Error('「入荷予定」シートが見つかりません。');
  if (!logSheet) throw new Error('「メール取込ログ」シートが見つかりません。');
  if (!label) throw new Error('Gmailに「' + GMAIL_LABEL + '」ラベルが見つかりません。');

  const loggedIds = getLoggedMessageIds_(logSheet);
  const threads = label.getThreads(0, 60);
  let importedCount = 0;
  let skippedCount = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const messageId = message.getId();
      if (loggedIds.has(messageId)) return;

      const from = message.getFrom();
      const subject = message.getSubject();
      const date = message.getDate();
      const daysAgo = (new Date().getTime() - date.getTime()) / (1000 * 60 * 60 * 24);

      // 直近120日だけ対象
      if (daysAgo > 120) return;

      const body = stripQuotedText_(message.getPlainBody());
      const supplier = detectSupplier_(from + '\n' + body);
      const attachmentStatus = detectAttachmentStatus_(message);
      const status = detectStatus_(subject + '\n' + body);
      const parsedItems = parseItems_(body);

      // 請求書だけ・支払い完了だけ・発送連絡だけなどはログだけ残す
      if (parsedItems.length === 0 && shouldLogOnly_(subject, body, attachmentStatus)) {
        logSheet.appendRow([messageId, new Date(), supplier, from, subject, '参照のみ', '品目なし・取込対象外']);
        loggedIds.add(messageId);
        skippedCount++;
        return;
      }

      if (parsedItems.length === 0) {
        importSheet.appendRow([
          status,
          supplier,
          '',
          '',
          '',
          '',
          '',
          '',
          '',
          date,
          '要確認',
          subject,
          messageId,
          attachmentStatus,
          '自動抽出できないため本文確認',
          '未反映',
          '',
          '未分類',
          '本文確認'
        ]);
        importedCount++;
      } else {
        parsedItems.forEach(item => {
          importSheet.appendRow([
            status,
            supplier,
            item.leatherName || '',
            item.color || '',
            item.thickness || '',
            item.spec || '',
            item.quantity || '',
            item.arrivalDate || '',
            item.price || '',
            date,
            '要確認',
            subject,
            messageId,
            attachmentStatus,
            item.memo || '',
            '未反映',
            '',
            item.category || '',
            item.action || ''
          ]);
          importedCount++;
        });
      }

      logSheet.appendRow([
        messageId,
        new Date(),
        supplier,
        from,
        subject,
        '取込済み',
        importedCount + '行追加 / ' + skippedCount + '件スキップ'
      ]);
      loggedIds.add(messageId);
    });
  });

  Logger.log('取込完了: ' + importedCount + '行追加、' + skippedCount + '件スキップ');
}

/**
 * 入荷予定で「反映OK」の行だけ、在庫マスターと入出庫履歴へ反映
 */
function reflectConfirmedArrivals() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const importSheet = ss.getSheetByName('入荷予定');
  const masterSheet = ss.getSheetByName('在庫マスター');
  const historySheet = ss.getSheetByName('入出庫履歴');

  if (!importSheet) throw new Error('「入荷予定」シートが見つかりません。');
  if (!masterSheet) throw new Error('「在庫マスター」シートが見つかりません。');
  if (!historySheet) throw new Error('「入出庫履歴」シートが見つかりません。');

  const lastRow = importSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('入荷予定にデータがありません。');
    return;
  }

  const rows = importSheet.getRange(2, 1, lastRow - 1, 19).getValues();
  const masterLastRow = masterSheet.getLastRow();
  const masterRows = masterLastRow >= 2 ? masterSheet.getRange(2, 1, masterLastRow - 1, 16).getValues() : [];

  const masterMap = new Map();
  masterRows.forEach((row, index) => {
    const key = makeInventoryKey_(row[1], row[2], row[3], row[4], row[5]);
    if (key) masterMap.set(key, index + 2);
  });

  let reflectedCount = 0;

  rows.forEach((row, i) => {
    const sheetRow = i + 2;

    const status = row[0];
    const supplier = row[1];
    const leatherName = row[2];
    const color = row[3];
    const thickness = row[4];
    const spec = row[5];
    const quantityRaw = row[6];
    const amount = row[8];
    const confirm = row[10];
    const subject = row[11];
    const memo = row[14];
    const reflectFlag = row[15];
    const category = row[17];

    if (reflectFlag !== '反映OK') return;
    if (confirm !== '確認済み') return;
    if (!(status === '入荷済み' || status === '発送済み')) return;
    if (category === 'メール確認のみ') return;
    if (!supplier || !leatherName || !quantityRaw) return;

    const quantity = normalizeQuantity_(quantityRaw);
    if (quantity === 0) return;

    const key = makeInventoryKey_(supplier, leatherName, color, thickness, spec);
    const today = new Date();
    const inventoryId = generateInventoryId_(supplier, leatherName, color, thickness, spec);

    if (masterMap.has(key)) {
      const targetRow = masterMap.get(key);
      const currentQty = Number(masterSheet.getRange(targetRow, 7).getValue()) || 0;
      masterSheet.getRange(targetRow, 7).setValue(currentQty + quantity);
      masterSheet.getRange(targetRow, 12).setValue(today);
      if (amount) masterSheet.getRange(targetRow, 14).setValue(amount);
      if (memo) masterSheet.getRange(targetRow, 16).setValue(memo);
    } else {
      masterSheet.appendRow([
        inventoryId,
        supplier,
        leatherName,
        color,
        thickness,
        spec,
        quantity,
        0,
        '',
        1,
        '',
        today,
        '',
        amount || '',
        '',
        memo || ''
      ]);

      const newRow = masterSheet.getLastRow();
      masterSheet.getRange(newRow, 9).setFormula('=G' + newRow + '-H' + newRow);
      masterSheet.getRange(newRow, 11).setFormula('=IF(I' + newRow + '<=0,"在庫ゼロ",IF(I' + newRow + '<=J' + newRow + ',"発注候補","通常"))');
      masterMap.set(key, newRow);
    }

    historySheet.appendRow([
      today,
      '入荷',
      inventoryId,
      supplier,
      leatherName,
      color,
      thickness,
      quantity,
      getUnit_(quantityRaw),
      '',
      amount || '',
      subject || '',
      memo || ''
    ]);

    importSheet.getRange(sheetRow, 16).setValue('反映済み');
    importSheet.getRange(sheetRow, 17).setValue(today);

    reflectedCount++;
  });

  Logger.log('在庫反映完了: ' + reflectedCount + '行を反映しました。');
}

/**
 * 既に入っている入荷予定のうち、明らかに在庫反映しないものを整理する補助関数
 * 安全のため、自動削除はしません。
 * - 分類が「メール確認のみ」の行だけ
 * - 確認列を「確認済み」
 * - 在庫反映を「反映しない」
 * にします。
 */
function markEmailOnlyRowsAsNoReflect() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('入荷予定');
  if (!sheet) throw new Error('「入荷予定」シートが見つかりません。');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
  let count = 0;

  values.forEach((row, index) => {
    const rowNo = index + 2;
    const confirm = row[10];      // K
    const reflect = row[15];      // P
    const category = row[17];     // R

    if (category === 'メール確認のみ' && confirm === '要確認' && reflect !== '反映済み') {
      sheet.getRange(rowNo, 11).setValue('確認済み');
      sheet.getRange(rowNo, 16).setValue('反映しない');
      count++;
    }
  });

  Logger.log('メール確認のみを反映しないへ整理: ' + count + '行');
}

/**
 * メール本文から品目候補を抽出
 */
function parseItems_(body) {
  const items = [];
  const lines = body
    .split('\n')
    .map(line => normalizeLine_(line))
    .filter(Boolean);

  lines.forEach(line => {
    if (isQuotedLine_(line)) return;
    if (isNonItemLine_(line)) return;
    if (!looksLikeItemLine_(line)) return;

    const item = {
      leatherName: detectLeatherName_(line),
      color: detectColor_(line),
      thickness: detectThickness_(line),
      spec: detectSpec_(line),
      quantity: detectQuantity_(line),
      price: detectPrice_(line),
      memo: line
    };

    item.category = detectCategory_(item, line);
    item.action = item.category === 'メール確認のみ' ? '反映しない候補' : '内容確認';

    if (item.category === 'メール確認のみ') return;

    if (item.leatherName || item.color || item.quantity || item.price) {
      items.push(item);
    }
  });

  return items;
}

/**
 * 返信引用をなるべく除外する
 */
function stripQuotedText_(body) {
  const hardStopMarkers = [
    '-----Original Message-----',
    '差出人:',
    'From:',
    '送信済み:',
    'Original Message',
    'Forwarded message'
  ];

  let text = body || '';

  hardStopMarkers.forEach(marker => {
    const index = text.indexOf(marker);
    if (index !== -1) {
      text = text.substring(0, index);
    }
  });

  return text
    .split('\n')
    .filter(line => !isQuotedLine_(line))
    .join('\n');
}

function normalizeLine_(line) {
  return String(line || '')
    .replace(/[　]/g, ' ')
    .replace(/[㎜ｍ]/g, 'mm')
    .replace(/[①②③④⑤⑥⑦⑧⑨⑩]/g, '')
    .replace(/^[・●○■□\-＊*]\s*/, '')
    .trim();
}

function isQuotedLine_(line) {
  return /^\s*(>|＞|》|&gt;)/.test(String(line || ''));
}

function looksLikeItemLine_(line) {
  return /(INGRASSATO|HAAS|Derby|ALRAN|Sully|ヴォーエプソン|シュリー|コードバン|ガルーシャ|カイラギ|ヌメ革|ブライドル|FENICE|Section Lacquer|枚|ds|デシ|ml|漉き|床なし|床あり)/i.test(line);
}

/**
 * 送料、漉き代、合計、税、返金など在庫品目でない行を除外
 */
function isNonItemLine_(line) {
  return /(送料|往復送料|小\s*計|消\s*費\s*税|合\s*計|税込|税別|振込|銀行|名義|返金|差額|請求額|お支払い|支払い完了|手数料|代引|配送|運賃|漉き代|スキ代|加工代|実費|見積合計|総額|お振込み)/.test(line);
}

function shouldLogOnly_(subject, body, attachmentStatus) {
  const text = subject + '\n' + body + '\n' + attachmentStatus;

  if (/請求書の支払いが完了|新規請求書をお届け|お支払いが完了|領収書|請求書/.test(text) && !hasRealItemKeyword_(text)) {
    return true;
  }

  if (/発送しました|発送致します|送り状|追跡番号|お問い合わせ番号/.test(text) && !hasRealItemKeyword_(text)) {
    return true;
  }

  return false;
}

function hasRealItemKeyword_(text) {
  return /(INGRASSATO|HAAS|Derby|ALRAN|Sully|ヴォーエプソン|シュリー|コードバン|ガルーシャ|カイラギ|ヌメ革|ブライドル|FENICE|Section Lacquer)/i.test(text);
}

function detectSupplier_(text) {
  if (/winsleather|ウインズファクトリー/i.test(text)) return 'ウインズファクトリー';
  if (/k2-japan|株式会社ケイツー|ケイツー/i.test(text)) return 'ケイツー';
  if (/kataoka-leather|片岡商店/i.test(text)) return '片岡商店';
  if (/lig\.co\.jp|株式会社LIG|㈱LIG/i.test(text)) return 'LIG';
  return '要確認';
}

function detectStatus_(text) {
  if (/本日発送|発送予定|発送致します|発送しました|送り状|お問い合わせ番号|追跡番号/.test(text)) return '発送済み';
  if (/入荷待ち|入荷予定|欠品|在庫なし/.test(text)) return '入荷待ち';
  if (/ご注文ありがとうございます|注文ありがとうございます|御請求書|ご請求書|請求書/.test(text)) return '注文済み';
  if (/注文お願|注文をお願い|発注/.test(text)) return '注文予定';
  return '要確認';
}

function detectLeatherName_(line) {
  const patterns = [
    'INGRASSATO',
    'HAAS Derby',
    'ALRAN Sully',
    'ヴォーエプソン',
    'シュリー',
    'コードバン',
    'ガルーシャ',
    'カイラギ',
    'ヌメ革',
    'ブライドルレザー',
    'FENICE Section Lacquer'
  ];

  for (const p of patterns) {
    if (line.indexOf(p) !== -1) return p;
  }

  if (/Derby/i.test(line)) return 'HAAS Derby';
  if (/Sully/i.test(line)) return 'ALRAN Sully';
  if (/FENICE|Section Lacquer/i.test(line)) return 'FENICE Section Lacquer';

  return '';
}

function detectColor_(line) {
  const colors = [
    'ロイヤルブルー',
    'Royal Blue',
    'ダークブラウン',
    'ブラック',
    'ブルーアトール',
    'フェアリーブルー',
    'ブルー',
    'Violet',
    'バイオレット',
    'レッド',
    'ルージュ',
    'グリーン',
    'タン',
    'グレージュ',
    'ウィスキー',
    'ナチュラル',
    'ブラウン',
    'チョコ',
    'オレンジ',
    'イエロー',
    'パープル',
    'モスグリーン',
    'ベースコート'
  ];

  for (const c of colors) {
    if (line.indexOf(c) !== -1) return c;
  }

  return '';
}

function detectThickness_(line) {
  // ml容量は厚みに入れず、容量として扱う
  const ml = line.match(/(\d+)\s?ml/i);
  if (ml) return ml[1] + 'ml';

  const splitMatch = line.match(/(\d+(?:\.\d+)?)\s?mm\s?\/\s?(\d+(?:\.\d+)?)\s?mm/i);
  if (splitMatch) return splitMatch[1] + 'mm/' + splitMatch[2] + 'mm';

  const match = line.match(/(\d+(?:\.\d+)?)\s?mm/i);
  return match ? match[1] + 'mm' : '';
}

function detectSpec_(line) {
  const spec = [];

  if (/床なし/.test(line)) spec.push('床なし');
  if (/床あり/.test(line)) spec.push('床あり');
  if (/フル/.test(line)) spec.push('フル');
  if (/半裁/.test(line)) spec.push('半裁');
  if (/2分割漉き/.test(line)) spec.push('2分割漉き');
  else if (/漉き|スキ/.test(line)) spec.push('漉き');
  if (/FENICE|Section Lacquer/.test(line)) spec.push('資材');

  return spec.join(' / ');
}

function detectQuantity_(line) {
  const ds = line.match(/(\d+)\s?(ds|デシ)/i);
  if (ds) return ds[1] + 'ds';

  const sheets = line.match(/(\d+)\s?枚/);
  if (sheets) return Number(sheets[1]);

  const pieces = line.match(/(\d+)\s?(つ|点|本|個)/);
  if (pieces) return Number(pieces[1]);

  if (/１枚/.test(line)) return 1;
  if (/２枚/.test(line)) return 2;
  if (/３枚/.test(line)) return 3;

  return '';
}

function detectPrice_(line) {
  const total = line.match(/＝\s?([0-9,]+)/);
  if (total) return Number(total[1].replace(/,/g, ''));

  const yen = line.match(/[¥￥\\]\s?([0-9,]+)/);
  if (yen) return Number(yen[1].replace(/,/g, ''));

  return '';
}

function detectCategory_(item, line) {
  if (/送料|漉き代|加工代|消費税|合計|小計|返金|支払い/.test(line)) return 'メール確認のみ';
  if (/FENICE|Section Lacquer|資材/.test(line) || item.leatherName === 'FENICE Section Lacquer') return '資材';

  if (
    /INGRASSATO|HAAS|Derby|ALRAN|Sully|ヴォーエプソン|シュリー|コードバン|ガルーシャ|カイラギ|ヌメ革|ブライドル/i.test(line)
  ) {
    return '革';
  }

  return '未分類';
}

function detectAttachmentStatus_(message) {
  const names = message.getAttachments().map(a => a.getName()).join(' / ');

  if (!names) return '';
  if (/請求書/.test(names) && /領収書/.test(names)) return '請求書・領収書あり';
  if (/請求書/.test(names)) return '請求書あり';
  if (/領収書/.test(names)) return '領収書あり';

  return '添付あり: ' + names;
}

function getLoggedMessageIds_(logSheet) {
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return new Set();

  return new Set(
    logSheet
      .getRange(2, 1, lastRow - 1, 1)
      .getValues()
      .flat()
      .filter(Boolean)
  );
}

function normalizeQuantity_(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;

  const text = String(value);
  const ds = text.match(/(\d+)\s?ds/i);
  if (ds) return Number(ds[1]);

  const num = text.match(/\d+/);
  return num ? Number(num[0]) : 0;
}

function getUnit_(value) {
  const text = String(value || '');
  if (/ds/i.test(text)) return 'ds';
  if (/ml/i.test(text)) return 'ml';
  return '枚';
}

function makeInventoryKey_(supplier, leatherName, color, thickness, spec) {
  if (!supplier || !leatherName) return '';
  return [supplier, leatherName, color || '', thickness || '', spec || ''].join('|');
}

function generateInventoryId_(supplier, leatherName, color, thickness, spec) {
  const base = [supplier, leatherName, color, thickness, spec]
    .filter(Boolean)
    .join('-')
    .replace(/\s+/g, '')
    .replace(/[\/\\]/g, '-');

  return base || 'NO-ID';
}

function createHourlyTrigger() {
  deleteImportTriggers();

  ScriptApp
    .newTrigger('importLeatherEmails')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('1時間ごとの自動取込を設定しました。');
}

function deleteImportTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'importLeatherEmails') {
      ScriptApp.deleteTrigger(t);
    }
  });

  Logger.log('自動取込トリガーを削除しました。');
}


// ============================================================
// Source: ginkawaya_leather_inventory_apps_script_v3_safe_automation_addon.gs
// ============================================================

/**
 * 銀革屋 革在庫管理 Apps Script v3
 * 安全寄りの完全自動化 追加コード
 *
 * 既存の v2 コードの下に、このコードを追加してください。
 *
 * 目的：
 * - Gmail取込
 * - メール確認のみの自動整理
 * - 高信頼データの自動確認
 * - ダッシュボード更新用の実行ログ
 * - 1時間ごとの自動実行
 *
 * 重要：
 * - 在庫マスターへの数量加算は、初期設定では自動実行しません。
 * - 完全に自動反映したい場合だけ、AUTO_REFLECT_STOCK を true に変更してください。
 * - ただし、革は実物確認が大切なので、最初は false 推奨です。
 */

const AUTO_REFLECT_STOCK = false; // 最初は false 推奨。true にすると反映OKの行を在庫へ反映します。
const AUTO_CONFIRM_HIGH_CONFIDENCE = true; // 仕入先・品名・数量が揃っている行を確認済みにする。
const AUTO_NO_REFLECT_EMAIL_ONLY = true; // 送料・請求・発送連絡などを反映しないへ整理する。

/**
 * 完全自動化のメイン関数
 * トリガーはこの関数に設定します。
 */
function runLeatherInventoryAutomation() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const log = [];

  log.push('開始: ' + new Date());

  try {
    importLeatherEmails();
    log.push('Gmail取込: 完了');
  } catch (e) {
    log.push('Gmail取込エラー: ' + e.message);
  }

  if (AUTO_NO_REFLECT_EMAIL_ONLY) {
    try {
      markEmailOnlyRowsAsNoReflect();
      log.push('メール確認のみ整理: 完了');
    } catch (e) {
      log.push('メール確認のみ整理エラー: ' + e.message);
    }
  }

  if (AUTO_CONFIRM_HIGH_CONFIDENCE) {
    try {
      const count = autoConfirmHighConfidenceRows_();
      log.push('高信頼行の自動確認: ' + count + '行');
    } catch (e) {
      log.push('高信頼行の自動確認エラー: ' + e.message);
    }
  }

  if (AUTO_REFLECT_STOCK) {
    try {
      autoSetReflectOkForArrivedRows_();
      reflectConfirmedArrivals();
      log.push('在庫自動反映: 完了');
    } catch (e) {
      log.push('在庫自動反映エラー: ' + e.message);
    }
  } else {
    log.push('在庫自動反映: OFF');
  }

  try {
    writeAutomationLog_(ss, log);
  } catch (e) {
    Logger.log('自動化ログ書き込みエラー: ' + e.message);
  }

  Logger.log(log.join('\n'));
}

/**
 * 高信頼データを自動で確認済みにする
 *
 * 条件：
 * - 確認が「要確認」
 * - 分類が「革」または「資材」
 * - 仕入先が空ではない、かつ「要確認」ではない
 * - 品名がある
 * - 数量がある
 *
 * 在庫反映はしません。
 */
function autoConfirmHighConfidenceRows_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('入荷予定');
  if (!sheet) throw new Error('「入荷予定」シートが見つかりません。');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
  let count = 0;

  values.forEach((row, index) => {
    const rowNo = index + 2;

    const status = row[0];     // A 状態
    const supplier = row[1];   // B 仕入先
    const name = row[2];       // C 革名・品名
    const quantity = row[6];   // G 数量
    const confirm = row[10];   // K 確認
    const reflect = row[15];   // P 在庫反映
    const category = row[17];  // R 分類

    const isHighConfidence =
      confirm === '要確認' &&
      (category === '革' || category === '資材') &&
      supplier &&
      supplier !== '要確認' &&
      name &&
      quantity &&
      reflect !== '反映済み';

    if (isHighConfidence) {
      sheet.getRange(rowNo, 11).setValue('確認済み'); // K
      sheet.getRange(rowNo, 19).setValue('入荷待ち確認'); // S
      count++;
    }
  });

  return count;
}

/**
 * 入荷済み・発送済みで、確認済みの行だけ「反映OK」にする
 *
 * 注意：
 * AUTO_REFLECT_STOCK = true の時だけ使います。
 * 最初は手動確認を推奨します。
 */
function autoSetReflectOkForArrivedRows_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('入荷予定');
  if (!sheet) throw new Error('「入荷予定」シートが見つかりません。');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
  let count = 0;

  values.forEach((row, index) => {
    const rowNo = index + 2;

    const status = row[0];     // A
    const confirm = row[10];   // K
    const reflect = row[15];   // P
    const category = row[17];  // R

    const canReflect =
      (status === '入荷済み' || status === '発送済み') &&
      confirm === '確認済み' &&
      reflect === '未反映' &&
      category !== 'メール確認のみ';

    if (canReflect) {
      sheet.getRange(rowNo, 16).setValue('反映OK'); // P
      count++;
    }
  });

  return count;
}

/**
 * 自動化ログ用シートを作って記録
 */
function writeAutomationLog_(ss, log) {
  let sheet = ss.getSheetByName('自動化ログ');

  if (!sheet) {
    sheet = ss.insertSheet('自動化ログ');
    sheet.appendRow(['実行日時', '内容']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 2).setBackground('#29323c').setFontColor('#ffffff').setFontWeight('bold');
  }

  sheet.appendRow([new Date(), log.join('\n')]);

  // 直近200件だけ残す
  const lastRow = sheet.getLastRow();
  if (lastRow > 201) {
    sheet.deleteRows(2, lastRow - 201);
  }

  sheet.autoResizeColumns(1, 2);
}

/**
 * 1時間ごとの完全自動化トリガーを作成
 */
function createFullAutomationTrigger() {
  deleteFullAutomationTriggers();

  ScriptApp
    .newTrigger('runLeatherInventoryAutomation')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('1時間ごとの安全自動化トリガーを設定しました。');
}

/**
 * 完全自動化トリガーを削除
 */
function deleteFullAutomationTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runLeatherInventoryAutomation') {
      ScriptApp.deleteTrigger(t);
    }
  });

  Logger.log('安全自動化トリガーを削除しました。');
}

/**
 * 現在の自動化設定を確認
 */
function checkAutomationMode() {
  Logger.log('AUTO_REFLECT_STOCK: ' + AUTO_REFLECT_STOCK);
  Logger.log('AUTO_CONFIRM_HIGH_CONFIDENCE: ' + AUTO_CONFIRM_HIGH_CONFIDENCE);
  Logger.log('AUTO_NO_REFLECT_EMAIL_ONLY: ' + AUTO_NO_REFLECT_EMAIL_ONLY);

  if (AUTO_REFLECT_STOCK) {
    Logger.log('注意: 在庫自動反映がONです。実物確認前に在庫へ入る可能性があります。');
  } else {
    Logger.log('安全設定: 在庫自動反映はOFFです。');
  }
}


// ============================================================
// Source: ginkawaya_monthly_inventory_pdf_mail_addon.gs
// ============================================================

/**
 * 棚卸PDFメール送信 追加コード
 * 
 * 今のApps Scriptの一番下に、このコードを追加してください。
 *
 * できること：
 * - 棚卸シートをPDFにする
 * - ginkawaya@gmail.com へメールで送る
 * - 毎月1日 朝9時ごろに自動送信する
 * - 送信結果を「棚卸メールログ」に記録する
 */

const INVENTORY_PDF_EMAIL_TO = 'ginkawaya@gmail.com';
const INVENTORY_PDF_SHEET_NAME = '棚卸';
const INVENTORY_PDF_SETTING_SHEET_NAME = '棚卸メール設定';
const INVENTORY_PDF_LOG_SHEET_NAME = '棚卸メールログ';

/**
 * 今すぐテスト送信する
 * まず最初にこれを1回実行してください。
 */
function sendMonthlyInventoryPdfTest() {
  sendMonthlyInventoryPdf_('テスト送信');
}

/**
 * 毎月送信用
 * 自動予約で動く本番用です。
 */
function sendMonthlyInventoryPdf() {
  sendMonthlyInventoryPdf_('月次自動送信');
}

/**
 * 棚卸シートをPDFにしてメール送信
 */
function sendMonthlyInventoryPdf_(mode) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(INVENTORY_PDF_SHEET_NAME);
  const logSheet = getOrCreateInventoryPdfLogSheet_(ss);

  if (!sheet) {
    throw new Error('「棚卸」シートが見つかりません。');
  }

  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth(); // 前月用。1月なら0
  const reportDate = new Date(year, month - 1, 1);
  const reportYear = reportDate.getFullYear();
  const reportMonth = reportDate.getMonth() + 1;

  const subject = '銀革屋 棚卸月次レポート ' + reportYear + '年' + reportMonth + '月分';
  const fileName = '銀革屋_棚卸月次レポート_' + reportYear + '年' + reportMonth + '月分.pdf';

  const pdfBlob = exportSheetToPdf_(ss, sheet, fileName);

  const body =
    '銀革屋 棚卸月次レポートを送付します。\n\n' +
    '対象月：' + reportYear + '年' + reportMonth + '月分\n' +
    '送信種別：' + mode + '\n\n' +
    '添付PDFをご確認ください。\n\n' +
    '※このメールはスプレッドシートの自動送信です。';

  GmailApp.sendEmail(INVENTORY_PDF_EMAIL_TO, subject, body, {
    attachments: [pdfBlob],
    name: '銀革屋 在庫管理'
  });

  logSheet.appendRow([
    new Date(),
    INVENTORY_PDF_EMAIL_TO,
    subject,
    '送信済み',
    mode
  ]);

  updateInventoryPdfSettingStatus_(ss, '送信済み: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));

  Logger.log('棚卸PDFを送信しました: ' + subject);
}

/**
 * 棚卸シートだけをPDF化する
 */
function exportSheetToPdf_(ss, sheet, fileName) {
  const spreadsheetId = ss.getId();
  const sheetId = sheet.getSheetId();

  const url =
    'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?' +
    'format=pdf' +
    '&gid=' + sheetId +
    '&size=A4' +
    '&portrait=false' +
    '&fitw=true' +
    '&sheetnames=false' +
    '&printtitle=false' +
    '&pagenumbers=true' +
    '&gridlines=true' +
    '&fzr=true' +
    '&top_margin=0.30' +
    '&bottom_margin=0.30' +
    '&left_margin=0.30' +
    '&right_margin=0.30';

  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + token
    },
    muteHttpExceptions: true
  });

  const blob = response.getBlob().setName(fileName);
  return blob;
}

/**
 * 毎月1日 朝9時ごろに自動送信する設定
 */
function createMonthlyInventoryPdfTrigger() {
  deleteMonthlyInventoryPdfTrigger();

  ScriptApp
    .newTrigger('sendMonthlyInventoryPdf')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  updateInventoryPdfSettingStatus_(ss, '毎月1日 朝9時ごろに自動送信');

  Logger.log('棚卸PDFの毎月自動送信を設定しました。');
}

/**
 * 棚卸PDFの毎月自動送信を止める
 */
function deleteMonthlyInventoryPdfTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendMonthlyInventoryPdf') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  updateInventoryPdfSettingStatus_(ss, '自動送信停止中');

  Logger.log('棚卸PDFの毎月自動送信を停止しました。');
}

/**
 * 送信ログシートを取得。なければ作成。
 */
function getOrCreateInventoryPdfLogSheet_(ss) {
  let sheet = ss.getSheetByName(INVENTORY_PDF_LOG_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(INVENTORY_PDF_LOG_SHEET_NAME);
    sheet.appendRow(['送信日時', '送信先', '件名', '結果', 'メモ']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 5).setBackground('#29323c').setFontColor('#ffffff').setFontWeight('bold');
  }

  return sheet;
}

/**
 * 棚卸メール設定シートの状態を更新
 */
function updateInventoryPdfSettingStatus_(ss, statusText) {
  const sheet = ss.getSheetByName(INVENTORY_PDF_SETTING_SHEET_NAME);
  if (!sheet) return;

  // B9 が「送信状態」
  sheet.getRange('B9').setValue(statusText);
}


// ============================================================
// Source: ginkawaya_wins_pdf_ocr_addon.gs
// ============================================================

/**
 * ウインズファクトリー PDF画像読取 追加コード
 *
 * 今のApps Scriptの一番下に追加してください。
 *
 * 目的：
 * - Gmailの「革仕入」ラベル付きメールからウインズファクトリーのPDFを探す
 * - Googleドライブの文字読み取り機能を使って、画像PDFを文字に変換する
 * - 読み取れた「INGRASSATO 〜 ds × 単価 ＝ 金額」の候補を「PDF読取候補」に入れる
 *
 * 注意：
 * - 画像読み取りは100%正確ではありません。
 * - 候補は必ず確認してください。
 * - 在庫マスターへは自動反映しません。
 *
 * 事前にやること：
 * Apps Script左側の「サービス」から Drive API を追加してください。
 * 表示名は「Drive」のままでOKです。
 */

const PDF_OCR_CANDIDATE_SHEET = 'PDF読取候補';
const PDF_OCR_LOG_SHEET = 'PDF読取ログ';
const PDF_OCR_TARGET_LABEL = '革仕入';
const PDF_OCR_TARGET_SUPPLIER_WORD = 'winsleather';
const PDF_OCR_TEMP_FOLDER_NAME = '銀革屋_PDF読取一時フォルダ';

/**
 * まずこれを1回実行
 * ウインズファクトリーのPDFを読んで候補を作ります。
 */
function importWinsFactoryPdfCandidates() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const candidateSheet = getOrCreatePdfOcrCandidateSheet_(ss);
  const logSheet = getOrCreatePdfOcrLogSheet_(ss);

  const label = GmailApp.getUserLabelByName(PDF_OCR_TARGET_LABEL);
  if (!label) throw new Error('Gmailに「' + PDF_OCR_TARGET_LABEL + '」ラベルが見つかりません。');

  const existingKeys = getExistingPdfCandidateKeys_(candidateSheet);
  const folder = getOrCreatePdfOcrTempFolder_();

  const threads = label.getThreads(0, 80);
  let totalCandidates = 0;
  let targetMailCount = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const from = message.getFrom() || '';
      const subject = message.getSubject() || '';
      const messageId = message.getId();
      const body = message.getPlainBody() || '';

      if ((from + subject + body).toLowerCase().indexOf(PDF_OCR_TARGET_SUPPLIER_WORD) === -1 &&
          (from + subject + body).indexOf('ウインズファクトリー') === -1) {
        return;
      }

      const attachments = message.getAttachments();
      const pdfs = attachments.filter(a => {
        const name = a.getName() || '';
        return a.getContentType() === 'application/pdf' || /\.pdf$/i.test(name);
      });

      if (pdfs.length === 0) return;

      targetMailCount++;

      pdfs.forEach(pdf => {
        try {
          const ocrText = ocrPdfByDrive_(pdf, folder);
          const candidates = parseLeatherCandidatesFromOcrText_(ocrText);

          candidates.forEach(item => {
            const key = [
              messageId,
              item.name,
              item.color,
              item.thickness,
              item.ds,
              item.price,
              item.amount
            ].join('|');

            if (existingKeys.has(key)) return;

            candidateSheet.appendRow([
              new Date(),
              'ウインズファクトリー',
              item.name || '',
              item.color || '',
              item.thickness || '',
              item.ds || '',
              item.price || '',
              item.amount || '',
              '未確認',
              subject,
              messageId,
              item.sourceText || ''
            ]);

            existingKeys.add(key);
            totalCandidates++;
          });

          logSheet.appendRow([
            new Date(),
            subject,
            '読取完了',
            candidates.length,
            '',
            pdf.getName()
          ]);
        } catch (e) {
          logSheet.appendRow([
            new Date(),
            subject,
            '読取エラー',
            0,
            e.message,
            pdf.getName()
          ]);
        }
      });
    });
  });

  logSheet.appendRow([
    new Date(),
    '全体',
    '完了',
    totalCandidates,
    '',
    '対象メール数: ' + targetMailCount
  ]);

  Logger.log('PDF読取候補の作成完了: ' + totalCandidates + '件');
}

/**
 * PDFをGoogleドライブに一時保存して、Googleドキュメント変換で文字読み取りする
 *
 * この関数は「Drive API」の追加が必要です。
 */
function ocrPdfByDrive_(pdfBlob, folder) {
  if (typeof Drive === 'undefined') {
    throw new Error('Drive APIが有効になっていません。Apps Script左側の「サービス」→「＋」→「Drive API」を追加してから、もう一度実行してください。');
  }

  const tempPdfFile = folder.createFile(pdfBlob);
  let docFileId = '';

  try {
    const resource = {
      title: 'OCR_' + pdfBlob.getName(),
      mimeType: MimeType.GOOGLE_DOCS,
      parents: [{ id: folder.getId() }]
    };

    const converted = Drive.Files.copy(resource, tempPdfFile.getId(), {
      ocr: true,
      ocrLanguage: 'ja'
    });

    docFileId = converted.id;
    const doc = DocumentApp.openById(docFileId);
    const text = doc.getBody().getText();

    return text || '';
  } finally {
    try {
      tempPdfFile.setTrashed(true);
    } catch (e) {}

    if (docFileId) {
      try {
        DriveApp.getFileById(docFileId).setTrashed(true);
      } catch (e) {}
    }
  }
}

/**
 * OCR文字から候補を抽出
 */
function parseLeatherCandidatesFromOcrText_(text) {
  const results = [];
  const clean = String(text || '')
    .replace(/[　]/g, ' ')
    .replace(/,/g, '')
    .replace(/㎜/g, 'mm')
    .replace(/ｍｍ/g, 'mm');

  const lines = clean.split('\n').map(x => x.trim()).filter(Boolean);

  lines.forEach(line => {
    if (!/INGRASSATO|HAAS|Derby|ALRAN|Sully|ヴォーエプソン|コードバン|ガルーシャ|カイラギ|ヌメ革|ブライドル/i.test(line)) return;
    if (/送料|漉き代|加工代|消費税|小計|合計|振込|手数料/.test(line)) return;

    const dsMatch = line.match(/(\d{1,4})\s*ds/i);
    const priceMatch = line.match(/[x×]\s*@?\s*(\d{2,6})/i);
    const amountMatch = line.match(/[=＝]\s*([0-9]{3,8})/);

    let name = '';
    if (/INGRASSATO/i.test(line)) name = 'INGRASSATO';
    else if (/HAAS|Derby/i.test(line)) name = 'HAAS Derby';
    else if (/ALRAN|Sully/i.test(line)) name = 'ALRAN Sully';
    else if (/ヴォーエプソン/i.test(line)) name = 'ヴォーエプソン';
    else if (/コードバン/i.test(line)) name = 'コードバン';
    else if (/ガルーシャ/i.test(line)) name = 'ガルーシャ';
    else if (/カイラギ/i.test(line)) name = 'カイラギ';
    else if (/ヌメ革/i.test(line)) name = 'ヌメ革';
    else if (/ブライドル/i.test(line)) name = 'ブライドルレザー';

    const color = detectPdfOcrColor_(line);
    const thickness = detectPdfOcrThickness_(line);

    if (name && (dsMatch || amountMatch || priceMatch)) {
      results.push({
        name: name,
        color: color,
        thickness: thickness,
        ds: dsMatch ? Number(dsMatch[1]) : '',
        price: priceMatch ? Number(priceMatch[1]) : '',
        amount: amountMatch ? Number(amountMatch[1]) : '',
        sourceText: line
      });
    }
  });

  return results;
}

function detectPdfOcrColor_(line) {
  const colors = [
    'ダークブラウン',
    'ブラック',
    'ブルー',
    'ネイビーブルー',
    'ロイヤルブルー',
    'フェアリーブルー',
    'ルージュ',
    'Violet',
    'バイオレット',
    'ライム',
    'グリーン',
    'レッド',
    'ブラウン',
    'チョコ',
    'ナチュラル'
  ];

  for (const color of colors) {
    if (line.indexOf(color) !== -1) return color;
  }

  return '';
}

function detectPdfOcrThickness_(line) {
  const match = line.match(/(\d+(?:\.\d+)?)\s*mm/i);
  return match ? match[1] + 'mm' : '';
}

/**
 * PDF候補を入荷予定へコピーする
 *
 * 条件：
 * - PDF読取候補の確認列が「確認済み」
 *
 * 在庫マスターには直接入れません。
 */
function copyConfirmedPdfCandidatesToArrivalPlan() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const candidateSheet = ss.getSheetByName(PDF_OCR_CANDIDATE_SHEET);
  const arrivalSheet = ss.getSheetByName('入荷予定');

  if (!candidateSheet) throw new Error('「PDF読取候補」シートが見つかりません。');
  if (!arrivalSheet) throw new Error('「入荷予定」シートが見つかりません。');

  const lastRow = candidateSheet.getLastRow();
  if (lastRow < 5) {
    Logger.log('PDF読取候補がありません。');
    return;
  }

  const rows = candidateSheet.getRange(5, 1, lastRow - 4, 12).getValues();
  let copied = 0;

  rows.forEach(row => {
    const readDate = row[0];
    const supplier = row[1];
    const name = row[2];
    const color = row[3];
    const thickness = row[4];
    const ds = row[5];
    const price = row[6];
    const amount = row[7];
    const confirm = row[8];
    const subject = row[9];
    const messageId = row[10];
    const sourceText = row[11];

    if (confirm !== '確認済み') return;
    if (!name || !ds) return;

    arrivalSheet.appendRow([
      '入荷済み',
      supplier || 'ウインズファクトリー',
      name,
      color,
      thickness,
      '',
      ds + 'ds',
      '',
      amount || '',
      readDate || new Date(),
      '確認済み',
      subject || '',
      messageId || '',
      'PDF読取',
      sourceText || '',
      '未反映',
      '',
      '革',
      'PDF確認済み'
    ]);

    copied++;
  });

  Logger.log('入荷予定へコピー: ' + copied + '件');
}

function getOrCreatePdfOcrCandidateSheet_(ss) {
  let sheet = ss.getSheetByName(PDF_OCR_CANDIDATE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PDF_OCR_CANDIDATE_SHEET);
    sheet.appendRow(['読取日', '仕入先', '商品名', '色', '厚み', 'ds数', '単価', '金額', '確認', 'メール件名', 'メールID', '元の読取文']);
  }
  return sheet;
}

function getOrCreatePdfOcrLogSheet_(ss) {
  let sheet = ss.getSheetByName(PDF_OCR_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PDF_OCR_LOG_SHEET);
    sheet.appendRow(['実行日時', '対象メール', '結果', '読取件数', 'エラー', 'メモ']);
  }
  return sheet;
}

function getExistingPdfCandidateKeys_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return new Set();

  const rows = sheet.getRange(5, 1, lastRow - 4, 12).getValues();
  const keys = new Set();

  rows.forEach(row => {
    const key = [
      row[10],
      row[2],
      row[3],
      row[4],
      row[5],
      row[6],
      row[7]
    ].join('|');

    keys.add(key);
  });

  return keys;
}

function getOrCreatePdfOcrTempFolder_() {
  const folders = DriveApp.getFoldersByName(PDF_OCR_TEMP_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PDF_OCR_TEMP_FOLDER_NAME);
}
