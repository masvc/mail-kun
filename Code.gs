/**
 * スプレッドシートを開いた時に自動実行される関数
 * カスタムメニューを作成します
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // カスタムメニューを作成
  ui.createMenu("📧 Gmail下書き作成")
    .addItem("🚀 下書きを一括作成", "createMultipleDrafts")
    .addItem("🧪 テスト実行（最初の3件）", "testCreateDrafts")
    .addSeparator() // 区切り線
    .addItem("📊 データ確認", "debugSpreadsheetData")
    .addItem("📝 サンプルデータ追加", "addSampleData")
    .addToUi();

  console.log("カスタムメニューを追加しました");
}

function getSpreadsheetData() {
  // 現在のスプレッドシートを取得
  const sheet = SpreadsheetApp.getActiveSheet();

  // データ範囲を取得（1行目はヘッダーなので2行目から）
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 3); // A2からC列の最終行まで
  const values = dataRange.getValues();

  // データをログに出力（確認用）
  console.log(values);
  return values;
}

function createSingleDraft() {
  const recipient = "test@example.com";
  const subject = "テスト件名";
  const body = "テスト本文です。";

  // 下書きを作成
  GmailApp.createDraft(recipient, subject, body);
  console.log("下書きを作成しました");
}

function createMultipleDrafts() {
  try {
    // スプレッドシートのデータを取得
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert(
        "データがありません。先にデータを入力してください。"
      );
      return;
    }

    const dataRange = sheet.getRange(2, 1, lastRow - 1, 3);
    const values = dataRange.getValues();

    let successCount = 0;
    let errorCount = 0;

    // 各行をループ処理
    values.forEach((row, index) => {
      try {
        const [companyName, name, email] = row;

        // 空の行をスキップ
        if (!email || email.trim() === "") {
          console.log(`行 ${index + 2}: メールアドレスが空のためスキップ`);
          return;
        }

        // メールテンプレートを作成
        const subject = createSubject(companyName, name);
        const body = createBody(companyName, name);

        // 下書きを作成
        GmailApp.createDraft(email, subject, body);

        successCount++;
        console.log(`行 ${index + 2}: ${name}さん宛の下書きを作成しました`);

        // APIレート制限対策で少し待機
        Utilities.sleep(100);
      } catch (error) {
        errorCount++;
        console.error(
          `行 ${index + 2}: エラーが発生しました - ${error.message}`
        );
      }
    });

    // 結果をユーザーに表示
    SpreadsheetApp.getUi().alert(
      `処理完了\n成功: ${successCount}件\nエラー: ${errorCount}件`
    );
    console.log(`処理完了: 成功 ${successCount}件, エラー ${errorCount}件`);
  } catch (error) {
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${error.message}`);
    console.error(`メイン処理でエラーが発生しました: ${error.message}`);
  }
}

// ===== 以下、追加が必要な関数 =====

/**
 * メール件名を作成する関数
 * @param {string} companyName - 会社名
 * @param {string} name - 担当者名
 * @returns {string} メール件名
 */
function createSubject(companyName, name) {
  return `【${companyName}】${name}様へのご提案`;
}

/**
 * メール本文を作成する関数
 * @param {string} companyName - 会社名
 * @param {string} name - 担当者名
 * @returns {string} メール本文
 */
function createBody(companyName, name) {
  return `${name}様

いつもお世話になっております。
${companyName}の皆様には、日頃より格別のご愛顧を賜り、誠にありがとうございます。

この度は、弊社サービスについてご提案させていただきたく、ご連絡いたしました。

詳細については、改めてお時間をいただければと思います。
ご都合の良い日時をお聞かせください。

何かご不明な点がございましたら、お気軽にお声がけください。

今後ともよろしくお願いいたします。

---
[あなたの名前]
[会社名]
[連絡先]`;
}

/**
 * テスト実行用関数（最初の3件のみ処理）
 */
function testCreateDrafts() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert("データがありません。");
      return;
    }

    // 最初の3行のみ取得
    const testRows = Math.min(3, lastRow - 1);
    const dataRange = sheet.getRange(2, 1, testRows, 3);
    const values = dataRange.getValues();

    let successCount = 0;

    values.forEach((row, index) => {
      try {
        const [companyName, name, email] = row;

        if (!email || email.trim() === "") {
          console.log(`行 ${index + 2}: メールアドレスが空のためスキップ`);
          return;
        }

        const subject = `[テスト] ${createSubject(companyName, name)}`;
        const body = createBody(companyName, name);

        GmailApp.createDraft(email, subject, body);
        successCount++;

        Utilities.sleep(100);
      } catch (error) {
        console.error(`行 ${index + 2}: エラー - ${error.message}`);
      }
    });

    SpreadsheetApp.getUi().alert(
      `テスト完了: ${successCount}件の下書きを作成しました`
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(`テスト実行エラー: ${error.message}`);
  }
}

/**
 * スプレッドシートのデータを確認する関数
 */
function debugSpreadsheetData() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert("データがありません。");
      return;
    }

    const dataRange = sheet.getRange(2, 1, lastRow - 1, 3);
    const values = dataRange.getValues();

    let message = "現在のデータ:\n\n";

    values.forEach((row, index) => {
      const [companyName, name, email] = row;
      message += `行${index + 2}: ${companyName || "(空)"} | ${
        name || "(空)"
      } | ${email || "(空)"}\n`;
    });

    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    SpreadsheetApp.getUi().alert(`データ確認エラー: ${error.message}`);
  }
}

/**
 * サンプルデータを追加する関数
 */
function addSampleData() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();

    // ヘッダーを設定
    sheet
      .getRange(1, 1, 1, 3)
      .setValues([["会社名", "担当者名", "メールアドレス"]]);

    // サンプルデータを追加
    const sampleData = [
      ["株式会社サンプル", "田中太郎", "tanaka@sample.com"],
      ["テスト商事", "佐藤花子", "sato@test.co.jp"],
      ["例示株式会社", "山田次郎", "yamada@example.jp"],
    ];

    sheet.getRange(2, 1, sampleData.length, 3).setValues(sampleData);

    SpreadsheetApp.getUi().alert("サンプルデータを追加しました！");
  } catch (error) {
    SpreadsheetApp.getUi().alert(`サンプルデータ追加エラー: ${error.message}`);
  }
}
