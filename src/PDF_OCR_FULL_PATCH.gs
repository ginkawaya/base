/**
 * 銀革屋 ウインズファクトリーPDF読取 本体追加パッチ
 *
 * 今のエラー：
 * ReferenceError: importWinsFactoryPdfCandidates is not defined
 *
 * 原因：
 * メニューは出ているが、PDF読取本体の関数がApps Script内に入っていないため。
 *
 * 使い方：
 * Apps Script の一番下に、このファイルの中身をそのまま貼り付けます。
 * 保存後、Drive APIを追加してから「銀革屋 在庫管理 → PDF読取候補を作成」を押してください。
 */

function importWinsFactoryPdfCandidates() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const candidateSheet = getOrCreatePdfOcrCandidateSheet_(ss);
  const logSheet = getOrCreatePdfOcrLogSheet_(ss);

  const labelName = '革仕入';
  const label = GmailApp.getUserLabelByName(labelName);
  if (!label) throw new Error('Gmailに「' + labelName + '」ラベルが見つかりません。');

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
      const searchText = (from + '\n' + subject + '\n' + body).toLowerCase();

      if (searchText.indexOf('winsleather') === -1 && searchText.indexOf('ウインズファクトリー') === -1) {
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

function parseLeatherCandidatesFromOcrText_(text) {
  const results = [];
  const seen = new Set();

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
      const key = [
        name,
        color,
        thickness,
        dsMatch ? dsMatch[1] : '',
        priceMatch ? priceMatch[1] : '',
        amountMatch ? amountMatch[1] : ''
      ].join('|');

      if (seen.has(key)) return;
      seen.add(key);

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

function getOrCreatePdfOcrCandidateSheet_(ss) {
  let sheet = ss.getSheetByName('PDF読取候補');
  if (!sheet) {
    sheet = ss.insertSheet('PDF読取候補');
    sheet.appendRow(['読取日', '仕入先', '商品名', '色', '厚み', 'ds数', '単価', '金額', '確認', 'メール件名', 'メールID', '元の読取文']);
  }
  return sheet;
}

function getOrCreatePdfOcrLogSheet_(ss) {
  let sheet = ss.getSheetByName('PDF読取ログ');
  if (!sheet) {
    sheet = ss.insertSheet('PDF読取ログ');
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
  const folderName = '銀革屋_PDF読取一時フォルダ';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}
