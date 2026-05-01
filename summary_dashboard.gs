// ============================================================
// summary_dashboard.gs
// 売上データを月ごとに集計し、サマリーシートとグラフを更新する
// ============================================================

/** メイン処理：集計 → サマリー書き込み → グラフ更新 */
function updateSummaryDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const salesData = loadSalesData(ss);
  const monthlySummary = aggregateByMonth(salesData);
  writeSummarySheet(ss, monthlySummary);
  updateBarChart(ss);
}

// ------------------------------------------------------------
// データ読み込み
// ------------------------------------------------------------

/**
 * 「売上データ」シートから全行を読み込む
 * @returns {{date: Date, person: string, product: string, amount: number}[]}
 */
function loadSalesData(ss) {
  const sheet = ss.getSheetByName("売上データ");
  if (!sheet) throw new Error("「売上データ」シートが見つかりません");

  const lastRow = sheet.getLastRow();
  // 1行目はヘッダー行のためスキップ。データが1行もない場合は空を返す
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

  return values
    .filter(row => row[0] && row[3] !== "")  // 日付・金額が空の行をスキップ
    .map(row => ({
      date:    new Date(row[0]),
      person:  row[1],
      product: row[2],
      amount:  Number(row[3]),
    }));
}

// ------------------------------------------------------------
// 月次集計
// ------------------------------------------------------------

/**
 * 売上データを「YYYY年M月」キーで集計する
 * @returns {{month: string, total: number, count: number}[]} 月昇順でソート済み
 */
function aggregateByMonth(salesData) {
  const map = {};

  salesData.forEach(({ date, amount }) => {
    const key = `${date.getFullYear()}年${date.getMonth() + 1}月`;

    // 月ごとに初期化
    if (!map[key]) {
      map[key] = {
        year:  date.getFullYear(),
        month: date.getMonth() + 1,  // 後でソートに使う
        label: key,
        total: 0,
        count: 0,
      };
    }

    map[key].total += amount;
    map[key].count += 1;
  });

  // 年月の昇順にソートして返す
  return Object.values(map).sort((a, b) => {
    if (a.year !== b.year) return a.year - b.year;
    return a.month - b.month;
  });
}

// ------------------------------------------------------------
// サマリーシート書き込み
// ------------------------------------------------------------

/** 「月次サマリー」シートをクリアしてから集計結果を書き込む */
function writeSummarySheet(ss, monthlySummary) {
  let sheet = ss.getSheetByName("月次サマリー");

  // シートが存在しない場合は作成する
  if (!sheet) {
    sheet = ss.insertSheet("月次サマリー");
  }

  // 既存データを全消去
  sheet.clearContents();

  // ヘッダー行
  sheet.getRange(1, 1, 1, 3).setValues([["月", "合計売上", "件数"]]);

  if (monthlySummary.length === 0) return;

  // 集計データを一括書き込み
  const rows = monthlySummary.map(({ label, total, count }) => [label, total, count]);
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  // 合計売上列に通貨書式を適用
  sheet.getRange(2, 2, rows.length, 1).setNumberFormat("¥#,##0");
}

// ------------------------------------------------------------
// 棒グラフ更新
// ------------------------------------------------------------

/**
 * 「月次サマリー」シートのデータを元に棒グラフを作成・更新する
 * 既存のグラフは削除してから新規作成する
 */
function updateBarChart(ss) {
  const summarySheet = ss.getSheetByName("月次サマリー");
  if (!summarySheet) return;

  const lastRow = summarySheet.getLastRow();
  // ヘッダーのみ or データなしの場合はグラフを作らない
  if (lastRow < 2) return;

  // 既存グラフをすべて削除
  summarySheet.getCharts().forEach(chart => summarySheet.removeChart(chart));

  // グラフのデータ範囲：月ラベル（A列）と合計売上（B列）
  const dataRange = summarySheet.getRange(1, 1, lastRow, 2);

  const chart = summarySheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)  // 縦棒グラフ
    .addRange(dataRange)
    .setOption("title", "月次売上推移")
    .setOption("hAxis.title", "月")
    .setOption("vAxis.title", "売上金額（円）")
    .setOption("legend.position", "none")
    .setOption("width", 700)
    .setOption("height", 400)
    .setPosition(2, 5, 0, 0)  // サマリーシートの E2 付近に配置
    .build();

  summarySheet.insertChart(chart);
}

// ------------------------------------------------------------
// 毎朝9時の自動実行トリガー設定
// ------------------------------------------------------------

/**
 * 毎朝9時に updateSummaryDashboard を実行するトリガーを登録する
 *
 * 使い方：GASエディタでこの関数を一度だけ手動実行すると、
 * 以降は毎朝9時に自動で集計処理が走るようになる。
 * 二重登録を防ぐため、既存の同名トリガーは事前に削除する。
 */
function setupDailyTrigger() {
  const TARGET_FUNCTION = "updateSummaryDashboard";

  // 既存の同名トリガーを削除して二重登録を防ぐ
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === TARGET_FUNCTION)
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 毎日9:00〜10:00の間に実行されるトリガーを登録
  ScriptApp.newTrigger(TARGET_FUNCTION)
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log(`トリガーを登録しました：毎朝9時に ${TARGET_FUNCTION} を実行します`);
}
