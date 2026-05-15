/**
 * 銀革屋 PDF読取 1件だけテスト用パッチ
 *
 * 目的：
 * - 全PDFを一気に読まず、最新のウインズファクトリーPDF 1件だけ読む
 * - ぐるぐる回って止まる問題を避ける
 *
 * 使い方：
 * Apps Script の一番下に貼り付けて保存。
 * 関数選択で importWinsFactoryPdfCandidatesOneTest を選んで実行。
 */

function importWinsFactoryPdfCandidatesOneTest() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const candidateSheet = getOrCreatePdfOcrCandidateSheet_(ss);
  const logSheet = getOrCreatePdfOcrLogSheet_(ss);

  const label = GmailApp.getUserLabelByName('革仕入');
  if (!label) throw new Error('Gmailに「革仕入」ラベルが見つかりません。');

  const folder = getOrCreatePdfOcrTempFolder_();
  const threads = label.getThreads(0, 30);

  for (const thread of threads) {
    const messages = thread.getMessages().reverse();

    for (const message of messages) {
      const from = message.getFrom() || '';
      const subject = message.getSubject() || '';
      const body = message.getPlainBody() || '';
      const messageId = message.getId();
      const searchText = (from + '\n' + subject + '\n' + body).toLowerCase();

      if (searchText.indexOf('winsleather') === -1 && searchText.indexOf('ウインズファクトリー') === -1) {
        continue;
      }

      const pdfs = message.getAttachments().filter(a => {
        const name = a.getName() || '';
        const isPdf = a.getContentType() === 'application/pdf' || /\.pdf$/i.test(name);
        const isInvoice = /請求|見積|御見積|ご請求|御請求/i.test(name);
        const isReceipt = /領収/i.test(name);
        return isPdf && isInvoice && !isReceipt;
      });

      if (pdfs.length === 0) continue;

      const pdf = pdfs[0];

      try {
        const ocrText = ocrPdfByDrive_(pdf, folder);
        const candidates = parseLeatherCandidatesFromOcrText_(ocrText);

        candidates.forEach(item => {
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
        });

        logSheet.appendRow([
          new Date(),
          subject,
          '1件テスト完了',
          candidates.length,
          '',
          pdf.getName()
        ]);

        Logger.log('1件テスト完了: ' + pdf.getName() + ' / 候補 ' + candidates.length + '件');
        return;
      } catch (e) {
        logSheet.appendRow([
          new Date(),
          subject,
          '1件テストエラー',
          0,
          e.message,
          pdf.getName()
        ]);
        throw e;
      }
    }
  }

  logSheet.appendRow([
    new Date(),
    '全体',
    '対象PDFなし',
    0,
    '',
    '最新30スレッド内にウインズファクトリーの請求・見積PDFが見つかりませんでした'
  ]);

  Logger.log('対象PDFなし');
}
