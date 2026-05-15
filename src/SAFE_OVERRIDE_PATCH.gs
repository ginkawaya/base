/**
 * 銀革屋 在庫管理 安全上書きパッチ
 *
 * 使い方：
 * Apps Script の一番下に、このファイルの中身をそのまま貼り付けます。
 * 既存の同名関数を安全版で上書きするためのパッチです。
 *
 * 目的：
 * - スプレッドシート上に「銀革屋 在庫管理」メニューを出す
 * - ds数確認、枚数表記、メール確認のみを在庫反映しない
 * - PDF読取候補を入荷予定へコピーしても、勝手に反映OKにしない
 * - Drive API 未設定時の案内を分かりやすくする
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('銀革屋 在庫管理')
    .addItem('メール取込・整理を実行', 'runLeatherInventoryAutomation')
    .addItem('反映OKを在庫へ反映', 'reflectConfirmedArrivalsSafe')
    .addItem('PDF読取候補を作成', 'importWinsFactoryPdfCandidatesSafe')
    .addItem('棚卸PDFをテスト送信', 'sendMonthlyInventoryPdfTest')
    .addItem('自動化設定を確認', 'checkAutomationMode')
    .addToUi();
}

function reflectConfirmedArrivalsSafe() {
  reflectConfirmedArrivals();
}

function importWinsFactoryPdfCandidatesSafe() {
  importWinsFactoryPdfCandidates();
}

function isSafeToReflectRow_(category, quantityRaw, supplier, leatherName) {
  const qText = String(quantityRaw || '').trim();
  if (!supplier || !leatherName) return false;
  if (!qText) return false;
  if (qText === 'ds数確認') return false;
  if (category === 'メール確認のみ') return false;
  if (category === '革' && /枚/.test(qText)) return false;
  return true;
}

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
  let skippedCount = 0;

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

    if (!isSafeToReflectRow_(category, quantityRaw, supplier, leatherName)) {
      importSheet.getRange(sheetRow, 19).setValue('反映停止：ds数・単位・分類を確認');
      skippedCount++;
      return;
    }

    const quantity = normalizeQuantity_(quantityRaw);
    if (quantity === 0) {
      importSheet.getRange(sheetRow, 19).setValue('反映停止：数量0');
      skippedCount++;
      return;
    }

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
    importSheet.getRange(sheetRow, 19).setValue('処理済み');

    reflectedCount++;
  });

  Logger.log('在庫反映完了: ' + reflectedCount + '行を反映。安全停止: ' + skippedCount + '行。');
}

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

  Logger.log('入荷予定へコピー: ' + copied + '件。P列は未反映のままです。');
}
