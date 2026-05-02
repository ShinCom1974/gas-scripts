// ============================================================
// email_classifier.gs
// Gmail の「要処理」ラベル付き未読メールを Gemini API で分類・要約し、
// スプレッドシートに記録して Slack に通知する
// ============================================================

/** スクリプトプロパティのキー名 */
const PROP_GEMINI_API_KEY    = "GEMINI_API_KEY";
const PROP_SLACK_WEBHOOK_URL = "SLACK_WEBHOOK_URL";

/** Gmail ラベル名 */
const LABEL_PENDING = "要処理";
const LABEL_DONE    = "処理済み";

/** シート名 */
const SHEET_MAIL_LOG  = "メールログ";
const SHEET_ERROR_LOG = "エラーログ";

// Gemini モデル名（2025年6月時点の最新 Lite モデル）
const GEMINI_MODEL = "gemini-2.5-flash-lite";

// ------------------------------------------------------------
// メイン処理
// ------------------------------------------------------------

/**
 * 「要処理」ラベルの未読メールを取得し、分類・記録・通知・ラベル更新を行う
 */
function classifyEmails() {
  const props = PropertiesService.getScriptProperties();
  const geminiApiKey    = props.getProperty(PROP_GEMINI_API_KEY);
  const slackWebhookUrl = props.getProperty(PROP_SLACK_WEBHOOK_URL);

  if (!geminiApiKey || !slackWebhookUrl) {
    logError("初期化エラー", "GEMINI_API_KEY または SLACK_WEBHOOK_URL がスクリプトプロパティに設定されていません");
    return;
  }

  // 「処理済み」ラベルは存在しなければ自動作成する
  const pendingLabel = GmailApp.getUserLabelByName(LABEL_PENDING);
  if (!pendingLabel) {
    logError("ラベルエラー", `「${LABEL_PENDING}」ラベルが Gmail に見つかりません`);
    return;
  }
  const doneLabel = GmailApp.getUserLabelByName(LABEL_DONE) || GmailApp.createLabel(LABEL_DONE);

  // 「要処理」かつ未読のスレッドを最大 20 件取得
  const threads = GmailApp.search(`label:${LABEL_PENDING} is:unread`, 0, 20);

  threads.forEach(thread => {
    // スレッド内の未読メッセージだけを処理対象にする
    const unreadMessages = thread.getMessages().filter(m => m.isUnread());

    unreadMessages.forEach(message => {
      try {
        const { classification, summary } = callGeminiApi(geminiApiKey, message);
        appendMailLog(message, classification, summary);
        notifySlack(slackWebhookUrl, message, classification, summary);
        message.markRead();
      } catch (e) {
        logError(message.getSubject() || "（件名なし）", e.message);
      }
    });

    // スレッド全体のラベルを「要処理」→「処理済み」に付け替える
    thread.addLabel(doneLabel);
    thread.removeLabel(pendingLabel);
  });
}

// ------------------------------------------------------------
// Gemini API 呼び出し
// ------------------------------------------------------------

/**
 * メールの件名・本文を Gemini API に送り、分類と要約を取得する
 * @returns {{ classification: string, summary: string }}
 */
function callGeminiApi(apiKey, message) {
  const subject = message.getSubject() || "";
  // 本文が長い場合は 3000 文字で切り詰める（API コスト抑制のため）
  const body = (message.getPlainBody() || "").slice(0, 3000);

  const endpoint =
    `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;

  const prompt = `以下のメールを分析し、指定した JSON 形式のみで回答してください。

件名：${subject}
本文：
${body}

【分類ルール】
- クレーム：苦情・不満・返金要求など
- 質問：問い合わせ・確認依頼など
- 注文：商品・サービスの注文・申し込みなど
- その他：上記に該当しないもの

【出力形式】JSON のみ返答し、他のテキストは一切含めないこと。
{
  "classification": "クレーム or 質問 or 注文 or その他",
  "summary": "50文字以内の要約"
}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: "application/json",
    },
  };

  const response = UrlFetchApp.fetch(endpoint, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(
      `Gemini API エラー (HTTP ${response.getResponseCode()}): ${response.getContentText()}`
    );
  }

  const responseJson = JSON.parse(response.getContentText());
  const text = responseJson?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error("Gemini API から有効なレスポンスが得られませんでした");

  const parsed = JSON.parse(text);
  if (!parsed.classification || !parsed.summary) {
    throw new Error(`Gemini API のレスポンス形式が不正です: ${text}`);
  }

  return { classification: parsed.classification, summary: parsed.summary };
}

// ------------------------------------------------------------
// スプレッドシート記録
// ------------------------------------------------------------

/** 「メールログ」シートにメール情報と分類結果を追記する */
function appendMailLog(message, classification, summary) {
  const sheet = getOrCreateSheet(SHEET_MAIL_LOG, ["受信日時", "送信者", "件名", "分類", "要約"]);
  sheet.appendRow([
    message.getDate(),
    message.getFrom(),
    message.getSubject() || "（件名なし）",
    classification,
    summary,
  ]);
}

/** 「エラーログ」シートにエラー情報を追記する */
function logError(context, errorMessage) {
  try {
    const sheet = getOrCreateSheet(SHEET_ERROR_LOG, ["発生日時", "処理対象", "エラー内容"]);
    sheet.appendRow([new Date(), context, errorMessage]);
  } catch (e) {
    // エラーログ自体の書き込みに失敗した場合は Logger のみに出力する
    Logger.log(`エラーログ書き込み失敗: ${e.message} / 元エラー: ${errorMessage}`);
  }
}

/**
 * 指定名のシートを返す。存在しない場合は作成してヘッダー行を設定する
 */
function getOrCreateSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

// ------------------------------------------------------------
// Slack 通知
// ------------------------------------------------------------

/** Slack Incoming Webhook でメールの分類結果を担当者に通知する */
function notifySlack(webhookUrl, message, classification, summary) {
  const subject = message.getSubject() || "（件名なし）";
  const from    = message.getFrom();
  const date    = Utilities.formatDate(
    message.getDate(),
    Session.getScriptTimeZone(),
    "yyyy/MM/dd HH:mm"
  );

  const text = [
    "*新着メール通知*",
    `*件名：* ${subject}`,
    `*分類：* ${classification}`,
    `*要約：* ${summary}`,
    `*送信者：* ${from}`,
    `*受信日時：* ${date}`,
  ].join("\n");

  const response = UrlFetchApp.fetch(webhookUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ text }),
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(
      `Slack 通知エラー (HTTP ${response.getResponseCode()}): ${response.getContentText()}`
    );
  }
}

// ------------------------------------------------------------
// 5 分おきのトリガー設定
// ------------------------------------------------------------

/**
 * 5 分おきに classifyEmails を実行するトリガーを登録する
 *
 * 使い方：GAS エディタでこの関数を一度だけ手動実行すると、
 * 以降は 5 分ごとに自動でメール分類処理が走るようになる。
 * 二重登録を防ぐため、既存の同名トリガーは事前に削除する。
 */
function setupIntervalTrigger() {
  const TARGET_FUNCTION = "classifyEmails";

  // 既存の同名トリガーを削除して二重登録を防ぐ
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === TARGET_FUNCTION)
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 5 分おきに実行されるトリガーを登録
  ScriptApp.newTrigger(TARGET_FUNCTION)
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log(`トリガーを登録しました：5 分おきに ${TARGET_FUNCTION} を実行します`);
}
