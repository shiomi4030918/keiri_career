/**
 * 経理キャリア LP ── 応募データ受信用 Google Apps Script
 *
 * 役割：
 *   1) LPフォームから送信されたデータを Google スプレッドシートに1行ずつ保存
 *   2) 通知メールを shiomiaki0918403@gmail.com に送信
 *   3) 応募者本人へお礼（自動返信）メールを送信
 *
 * ▼ セットアップ手順
 *   1. 保存用のスプレッドシートを1つ用意し、そのURLの
 *      https://docs.google.com/spreadsheets/d/【ここがID】/edit
 *      の【ここがID】部分を下の SPREADSHEET_ID に貼り付ける。
 *      （空のままにすると、初回実行時に新しいシートを自動作成し、IDをログに出します）
 *   2. 「デプロイ > 新しいデプロイ > 種類：ウェブアプリ」
 *        - 次のユーザーとして実行：自分
 *        - アクセスできるユーザー：全員
 *      でデプロイし、発行された「ウェブアプリのURL」をコピー。
 *   3. index.html の GAS_URL にそのURLを貼り付ける。
 */

// ===== 設定 =====
var SPREADSHEET_ID = "";                       // 保存先スプレッドシートのID（空なら自動作成）
var SHEET_GID      = 0;                         // 保存先タブのgid（URLの #gid=◯◯ の数字）。0なら下のSHEET_NAMEを使用
var SHEET_NAME     = "求職者リスト（経理キャリア）"; // タブ名
var NOTIFY_TO      = "shiomiaki0918403@gmail.com";
var REPLY_TO       = "info@hitotech.jp";        // 応募者へのお礼メールの差出名・返信先（問い合わせ窓口）
var SITE_NAME      = "経理キャリア";
// ================

// 列の定義（順序＝スプレッドシートのヘッダー順）
var FIELDS = [
  ["submittedAt",     "送信日時"],
  ["feeling",         "きっかけ"],
  ["employment",      "就業状況"],
  ["experienceYears", "経理経験年数"],
  ["qualification",   "保有資格"],
  ["workstyle",       "希望勤務形態"],
  ["listedExp",       "上場企業での経理経験"],
  ["region",          "地方"],
  ["pref",            "都道府県"],
  ["city",            "市区町村"],
  ["name",            "お名前"],
  ["birth",           "生まれ年"],
  ["tel",             "電話番号"],
  ["email",           "メールアドレス"]
];

/**
 * ▼ 動作テスト用（エディタから手動実行できます）
 *   GASエディタ上部の関数選択で「testRun」を選び ▶実行 を押すと、
 *   サンプルデータでスプレッドシート保存とメール送信を実際に試します。
 */
function testRun() {
  var sample = {
    feeling:"近いうちに転職したい", employment:"いい転職先があれば辞めたい",
    experienceYears:"2〜3年未満", qualification:"簿記2級", workstyle:"正社員",
    listedExp:"経験なし", region:"関西", pref:"大阪府", city:"大阪市北区",
    name:"テスト 太郎", birth:"1995", tel:"09012345678", email:"sample@example.com",
    submittedAt:new Date().toISOString()
  };
  doPost({ postData: { contents: JSON.stringify(sample) } });
  Logger.log("testRun 完了。スプレッドシートとメールを確認してください。");
}

function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    var data = JSON.parse(raw);

    var sheet = getSheet_();
    var now = new Date();

    // 受信日時（送信側の値が無ければサーバー時刻）
    var jstTime = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
    data.submittedAt = data.submittedAt ? formatIso_(data.submittedAt) : jstTime;

    // 行データを組み立て
    var row = FIELDS.map(function (f) { return data[f[0]] != null ? String(data[f[0]]) : ""; });

    // 書き込み先の行
    var r = sheet.getLastRow() + 1;
    // 電話番号・生まれ年は「テキスト形式」にしてから書き込む（先頭の0が消えないように）
    var birthCol = colIndex_("birth");
    var telCol   = colIndex_("tel");
    sheet.getRange(r, birthCol).setNumberFormat("@");
    sheet.getRange(r, telCol).setNumberFormat("@");
    sheet.getRange(r, 1, 1, row.length).setValues([row]);

    sendMail_(data, jstTime);
    try { sendThanks_(data, jstTime); } catch (e2) { Logger.log("お礼メール送信エラー: " + e2); }

    return jsonOut_({ result: "success" });
  } catch (err) {
    // エラーも通知して原因を追えるようにする
    try {
      MailApp.sendEmail(NOTIFY_TO, "【" + SITE_NAME + "】受信エラー", String(err) + "\n\n" + (e && e.postData ? e.postData.contents : ""));
    } catch (ignore) {}
    return jsonOut_({ result: "error", message: String(err) });
  }
}

// ブラウザで開いた時の動作確認用
function doGet() {
  return jsonOut_({ result: "ok", message: SITE_NAME + " endpoint is running." });
}

function getSheet_() {
  var ss;
  if (SPREADSHEET_ID) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } else {
    ss = SpreadsheetApp.create(SITE_NAME + " 応募データ");
    Logger.log("新しいスプレッドシートを作成しました。ID: " + ss.getId());
    Logger.log("URL: " + ss.getUrl());
  }
  // 保存先タブを決定：SHEET_GID が指定されていれば gid で特定、無ければ名前で取得
  var sheet = null;
  if (SHEET_GID) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === SHEET_GID) { sheet = sheets[i]; break; }
    }
  }
  if (!sheet) {
    sheet = ss.getSheetByName(SHEET_NAME);
  }
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  // タブ名を SHEET_NAME にそろえる
  if (sheet.getName() !== SHEET_NAME) {
    sheet.setName(SHEET_NAME);
  }
  // ヘッダー行を常に最新の項目に合わせる
  sheet.getRange(1, 1, 1, FIELDS.length)
       .setValues([FIELDS.map(function (f) { return f[1]; })])
       .setFontWeight("bold").setBackground("#f5821f").setFontColor("#ffffff");
  sheet.setFrozenRows(1);
  // 電話番号・生まれ年の列はテキスト形式に固定
  sheet.getRange(1, colIndex_("birth"), sheet.getMaxRows(), 1).setNumberFormat("@");
  sheet.getRange(1, colIndex_("tel"),   sheet.getMaxRows(), 1).setNumberFormat("@");
  return sheet;
}

// フィールドのキーから列番号（1始まり）を返す
function colIndex_(key) {
  for (var i = 0; i < FIELDS.length; i++) {
    if (FIELDS[i][0] === key) return i + 1;
  }
  return -1;
}

function sendMail_(data, jstTime) {
  var lines = FIELDS.map(function (f) {
    var v = data[f[0]] != null ? data[f[0]] : "（未入力）";
    if (f[0] === "submittedAt") v = jstTime;
    if (f[0] === "birth" && data.birth) v = data.birth + "年";
    return "■ " + f[1] + "：" + v;
  });

  // 概要（氏名・エリア・経理経験年数・保有資格）。件名と本文先頭サマリに使う
  var area = [data.pref, data.city].filter(function (x) { return x; }).join(" ");
  var parts = [];
  parts.push(data.name ? data.name + "様" : "お名前未入力");
  if (area) parts.push(area);
  if (data.experienceYears) parts.push("経験" + data.experienceYears);
  if (data.qualification) parts.push(data.qualification);
  var summary = parts.join("／");
  var subject = "【" + SITE_NAME + "】新規応募：" + summary;

  // 本文：スマホの通知は本文の改行を詰めて1行表示するため、冒頭の1行で完結させる
  var headline = "新規応募｜" + summary + "｜TEL " + (data.tel || "未入力") + "｜" + (data.email || "未入力");
  var body =
    headline + "\n\n" +
    "──────────────────\n" +
    "【詳細】\n" +
    lines.join("\n") +
    "\n──────────────────\n" +
    SITE_NAME + "（自動送信）";

  // テキスト＋HTML併用（プレーンテキストだと長文が約26文字で強制改行されるため）
  MailApp.sendEmail(NOTIFY_TO, subject, body, { htmlBody: toHtmlBody_(body) });
}

// 応募者本人へのお礼（自動返信）メール
function sendThanks_(data, jstTime) {
  var to = (data.email || "").trim();
  // メールアドレスが未入力／不正な形式のときは送信しない
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) return;

  var lines = FIELDS.map(function (f) {
    if (f[0] === "email") return null; // 控えには本人のメール行は載せない
    var v = (data[f[0]] != null && data[f[0]] !== "") ? data[f[0]] : "（未入力）";
    if (f[0] === "submittedAt") v = jstTime;
    if (f[0] === "birth" && data.birth) v = data.birth + "年";
    return "■ " + f[1] + "：" + v;
  }).filter(function (x) { return x; });

  var greeting = (data.name ? data.name + " 様" : "ご応募者 様");
  var body =
    greeting + "\n\n" +
    "この度は「" + SITE_NAME + "」へご応募いただき、誠にありがとうございます。\n" +
    "以下の内容で承りました。担当アドバイザーより、ご希望に沿った経理・財務の求人を順次ご連絡いたします。\n\n" +
    "──────────────────\n" +
    "【ご応募内容】\n" +
    lines.join("\n") + "\n" +
    "──────────────────\n\n" +
    "ご応募はすべて担当アドバイザーが順次拝見しております。内容を確認のうえ、担当者より順番にご連絡・対応させていただきますので、いましばらくお待ちくださいませ。\n" +
    "※ご応募状況によっては、ご連絡まで数日お時間をいただく場合がございます。あらかじめご了承ください。\n\n" +
    "本メールにお心当たりがない場合は、お手数ですが本メールへご返信ください。\n\n" +
    SITE_NAME + "（運営：Hitotech株式会社）\n" +
    "お問い合わせ：" + REPLY_TO + "\n\n" +
    "※本メールは自動送信です。";

  var subject = "【" + SITE_NAME + "】ご応募ありがとうございます";

  // テキスト＋HTML併用（プレーンテキストだと長文が約26文字で強制改行されるため）
  MailApp.sendEmail(to, subject, body, { name: SITE_NAME, replyTo: REPLY_TO, htmlBody: toHtmlBody_(body) });
}

// プレーンテキスト本文を、自然に折り返されるHTML本文へ変換する
function toHtmlBody_(text) {
  var esc = String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  return '<div style="font-family:\'Hiragino Kaku Gothic ProN\',Meiryo,sans-serif;'
       + 'font-size:14px;line-height:1.8;color:#222222;">'
       + esc.replace(/\n/g, "<br>")
       + '</div>';
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatIso_(iso) {
  try {
    return Utilities.formatDate(new Date(iso), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  } catch (e) {
    return iso;
  }
}
