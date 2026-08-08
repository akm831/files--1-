/**
 * WebアプリのURLを取得します。
 * @return {string} WebアプリのURL。
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * HTMLテンプレート内から共通パーツ（Style.html, Header.htmlなど）を
 * 読み込むためのヘルパー関数。
 * テンプレート変数（url, activePage, pageTitleなど）は明示的に
 * オブジェクトとして渡す。
 * 使い方: <?!= include('Header', {url: url, activePage: activePage, pageTitle: pageTitle}); ?>
 * @param {string} filename 読み込むHTMLファイル名（拡張子なし）。
 * @param {object} vars テンプレートに渡す変数のオブジェクト（省略可）。
 * @return {string} 評価済みのHTML文字列。
 */
function include(filename, vars) {
  var tmpl = HtmlService.createTemplateFromFile(filename);
  vars = vars || {};
  for (var key in vars) {
    if (Object.prototype.hasOwnProperty.call(vars, key)) {
      tmpl[key] = vars[key];
    }
  }
  return tmpl.evaluate().getContent();
}

/**
 * ページごとの表示タイトルを返します。
 * @param {string} page ページ識別子。
 * @return {string} 画面ヘッダーに表示するタイトル。
 */
function getPageTitle(page) {
  var titles = {
    index: "昇任試験問題集",
    search: "問題を検索する",
    stats: "学習統計",
    bookmarks: "ブックマーク一覧",
  };
  return titles[page] || "昇任試験問題集";
}

/**
 * WebアプリのGETリクエストを処理し、ページを振り分けます。
 * @param {object} e イベントオブジェクト。
 * @return {HtmlOutput} HTMLオブジェクト。
 */
function doGet(e) {
  var page = "index"; // デフォルトページ

  if (e && e.parameter && e.parameter.page) {
    var pageParam = e.parameter.page;
    if (pageParam === "stats") {
      page = "stats";
    } else if (pageParam === "bookmarks") {
      page = "bookmarks";
    } else if (pageParam === "search") {
      // 検索ページへのルーティングを追加
      page = "search";
    }
  }

  var appUrl = ScriptApp.getService().getUrl();
  // 末尾に "/" が付いている場合、"?page=xxx" と連結すると
  // "/?" のような不正なURLになりリンク切れの原因になるため正規化する
  if (appUrl && appUrl.charAt(appUrl.length - 1) === "/") {
    appUrl = appUrl.slice(0, -1);
  }

  var html = HtmlService.createTemplateFromFile(page);
  html.url = appUrl;
  html.activePage = page;
  html.pageTitle = getPageTitle(page);

  // HTMLを生成し、レスポンシブ対応のためのviewportメタタグを追加して返す
  return html
    .evaluate()
    .addMetaTag("viewport", "content=width=device-width, initial-scale=1")
    .setTitle(getPageTitle(page));
}

/**
 * 【新規】キーワードで問題を検索する関数。
 * @param {string} keyword 検索キーワード。
 * @return {Array<object>} 検索結果の問題データ配列。
 */
function searchQuestions(keyword) {
  if (!keyword || keyword.trim() === "") {
    return [];
  }
  const lowerCaseKeyword = keyword.toLowerCase();
  const allSheetNames = getSheetNames();
  let results = [];

  allSheetNames.forEach((sheetName) => {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const question = row[2] || "";
      const commentary = row[4] || "";

      if (
        question.toLowerCase().includes(lowerCaseKeyword) ||
        commentary.toLowerCase().includes(lowerCaseKeyword)
      ) {
        results.push({
          sheetName: sheetName,
          genre: row[1] || "",
          question: question,
          answer: row[3] || "",
          commentary: commentary,
          article: row[5] || "",
        });
      }
    }
  });
  return results;
}

/**
 * 指定した名前のシートが存在することを確認し、なければ作成します。
 * @param {string} sheetName 確認または作成するシート名。
 * @param {Array<string>} headers シートが新規作成される場合に設定するヘッダー行。
 * @return {Sheet} 存在する、または作成されたシートオブジェクト。
 */
function ensureSheet(sheetName, headers) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
    }
  }
  return sheet;
}

/**
 * ブックマークを追加します。
 * @param {string} sheetName 対象のシート名。
 * @param {number} rowNum 対象の行番号。
 * @return {string} 処理結果のメッセージ。
 */
function addBookmark(sheetName, rowNum) {
  var sheet = ensureSheet("ブックマーク", ["ユーザー", "シート", "行"]);
  var user = Session.getActiveUser().getEmail();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (
      data[i][0] === user &&
      data[i][1] === sheetName &&
      data[i][2] == rowNum
    ) {
      return "Already bookmarked";
    }
  }
  sheet.appendRow([user, sheetName, rowNum]);
  return "Bookmarked";
}

/**
 * 【一覧ページ用】ユーザーの全てのブックマークを取得します。
 * @return {Array<object>} ブックマークされた全問題のデータ配列。
 */
function getAllBookmarkedQuestionsForUser() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("ブックマーク");
  if (!sheet) return [];
  var user = Session.getActiveUser().getEmail();
  var bookmarks = sheet
    .getDataRange()
    .getValues()
    .filter(function (row, index) {
      return index > 0 && row[0] === user;
    });
  if (bookmarks.length === 0) return [];

  var questionsBySheet = {};
  bookmarks.forEach(function (bm) {
    var sheetName = bm[1];
    var rowNum = bm[2];
    if (!questionsBySheet[sheetName]) {
      questionsBySheet[sheetName] = [];
    }
    questionsBySheet[sheetName].push(rowNum);
  });

  var res = [];
  for (var sheetName in questionsBySheet) {
    var questionSheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!questionSheet) continue;
    var data = questionSheet.getDataRange().getValues();
    var rowNums = questionsBySheet[sheetName];
    rowNums.forEach(function (rowNum) {
      if (rowNum > 0 && rowNum <= data.length) {
        var rowData = data[rowNum - 1];
        res.push({
          sheetName: sheetName,
          genre: rowData[1],
          question: rowData[2],
          answer: rowData[3],
          commentary: rowData[4],
          article: rowData[5],
          rowNum: rowNum,
        });
      }
    });
  }
  return res;
}

/**
 * 【一覧ページ用】ブックマークを削除します。
 * @param {string} sheetName 対象のシート名。
 * @param {number} rowNum 対象の行番号。
 * @return {string} 処理結果のメッセージ。
 */
function removeBookmark(sheetName, rowNum) {
  var sheet = SpreadsheetApp.getActive().getSheetByName("ブックマーク");
  if (!sheet) return "Sheet not found";
  var user = Session.getActiveUser().getEmail();
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (
      data[i][0] === user &&
      data[i][1] === sheetName &&
      data[i][2] == rowNum
    ) {
      sheet.deleteRow(i + 1);
      return "Removed";
    }
  }
  return "Not found";
}

/**
 * 【問題出題用】指定シートのブックマークされた問題を取得します。
 * @param {string} sheetName 対象のシート名。
 * @return {Array<object>} ブックマークされた問題のデータ配列。
 */
function getBookmarkedQuestions(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName("ブックマーク");
  if (!sheet) return [];
  var user = Session.getActiveUser().getEmail();
  var bookmarks = sheet
    .getDataRange()
    .getValues()
    .filter(function (row, index) {
      return index > 0 && row[0] === user && row[1] === sheetName;
    })
    .map(function (row) {
      return row[2];
    });
  if (bookmarks.length === 0) return [];

  var questionSheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var data = questionSheet.getDataRange().getValues();
  var res = [];
  bookmarks.forEach(function (rowNum) {
    if (rowNum > 0 && rowNum <= data.length) {
      var rowData = data[rowNum - 1];
      res.push({
        genre: rowData[1],
        question: rowData[2],
        answer: rowData[3],
        commentary: rowData[4],
        article: rowData[5],
        rowNum: rowNum,
      });
    }
  });
  return res;
}

/**
 * シート名の一覧を取得します。
 * @return {Array<string>} 履歴とブックマークシートを除いたシート名の配列。
 */
function getSheetNames() {
  var sheets = SpreadsheetApp.getActive().getSheets();
  var names = [];
  var excludedSheets = ["履歴", "ブックマーク"];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (excludedSheets.indexOf(name) === -1) {
      names.push(name);
    }
  }
  return names;
}

/**
 * ジャンルの一覧を取得します。
 * @param {string} sheetName 対象のシート名。
 * @return {Array<string>} ジャンルの配列。
 */
function getGenres(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var genres = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) genres[data[i][1]] = true;
  }
  var arr = Object.keys(genres);
  arr.unshift("すべてのジャンル");
  return arr;
}

/**
 * 指定されたシートから全ての問題を取得します。
 * @param {string} sheetName 対象のシート名。
 * @param {string} genreFilter ジャンルのフィルタ。
 * @return {Array<object>} 問題データの配列。
 */
function getAllQuestions(sheetName, genreFilter) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var res = [];
  for (var i = 1; i < data.length; i++) {
    if (
      !genreFilter ||
      genreFilter === "すべてのジャンル" ||
      data[i][1] === genreFilter
    ) {
      res.push({
        genre: data[i][1],
        question: data[i][2],
        answer: data[i][3],
        commentary: data[i][4],
        article: data[i][5],
        rowNum: i + 1,
      });
    }
  }
  return res;
}

/**
 * 解答履歴を記録します。
 * @param {string} sheetName 対象のシート名。
 * @param {number} rowNum 対象の行番号。
 * @param {string} genre ジャンル。
 * @param {string} question 問題文。
 * @param {string} userAns ユーザーの回答。
 * @param {boolean} correct 正誤。
 * @param {string} correctAns 正解。
 * @param {string} article 根拠条文。
 */
function recordHistory(
  sheetName,
  rowNum,
  genre,
  question,
  userAns,
  correct,
  correctAns,
  article
) {
  var sheet = ensureSheet("履歴", [
    "日時",
    "シート",
    "行",
    "ジャンル",
    "問題文",
    "自分の回答",
    "正否",
    "正解",
    "根拠条文",
  ]);
  var now = new Date();
  sheet.appendRow([
    Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss"),
    sheetName,
    rowNum,
    genre,
    question,
    userAns,
    correct ? "○" : "×",
    correctAns,
    article,
  ]);
}

/**
 * シート全体の統計情報を取得します。
 * @return {object} シートごとの統計情報。
 */
function getStats() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("履歴");
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var stats = {};
  for (var i = 1; i < data.length; i++) {
    var s = data[i][1];
    var correct = data[i][6];
    if (!stats[s]) stats[s] = { correct: 0, total: 0 };
    stats[s].total++;
    if (correct === "○") stats[s].correct++;
  }
  return stats;
}

/**
 * ジャンルごとの統計情報を取得します。
 * @param {string} sheetName 対象のシート名。
 * @return {object} ジャンルごとの統計情報。
 */
function getGenreStats(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName("履歴");
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var stats = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] !== sheetName) continue;
    var genre = data[i][3];
    var correct = data[i][6];
    if (!stats[genre]) stats[genre] = { correct: 0, total: 0 };
    stats[genre].total++;
    if (correct === "○") stats[genre].correct++;
  }
  return stats;
}

/**
 * 問題ごとの統計情報を取得します。
 * @param {string} sheetName 対象のシート名。
 * @param {string} genreFilter ジャンルのフィルタ。
 * @return {object} 問題ごとの統計情報。
 */
function getQuestionStats(sheetName, genreFilter) {
  var sheet = SpreadsheetApp.getActive().getSheetByName("履歴");
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] !== sheetName) continue;
    var genre = data[i][3];
    if (
      genreFilter &&
      genreFilter !== "すべてのジャンル" &&
      genre !== genreFilter
    )
      continue;
    var question = data[i][4];
    var correct = data[i][6];
    if (!map[question]) map[question] = { total: 0, wrong: 0, correct: 0 };
    map[question].total++;
    if (correct === "○") map[question].correct++;
    if (correct === "×") map[question].wrong++;
  }
  return map;
}

/**
 * 誤答率の高い問題をランダムに取得します。
 * @param {string} sheetName 対象のシート名。
 * @param {string} genreFilter ジャンルのフィルタ。
 * @param {number} n 取得する問題数。
 * @return {Array<object>} 誤答率の高い問題のデータ配列。
 */
function getWorstQuestions(sheetName, genreFilter, n) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var qmap = {};
  for (var i = 1; i < data.length; i++) {
    var genre = data[i][1],
      question = data[i][2],
      answer = data[i][3],
      commentary = data[i][4],
      article = data[i][5];
    if (
      genreFilter &&
      genreFilter !== "すべてのジャンル" &&
      genre !== genreFilter
    )
      continue;
    qmap[question] = {
      question,
      answer,
      commentary,
      article,
      total: 0,
      wrong: 0,
    };
  }
  var hist = ss.getSheetByName("履歴");
  if (!hist) return [];
  var hdata = hist.getDataRange().getValues();
  for (var i = 1; i < hdata.length; i++) {
    if (hdata[i][1] !== sheetName) continue;
    var question = hdata[i][4];
    if (!qmap[question]) continue;
    qmap[question].total++;
    if (hdata[i][6] === "×") qmap[question].wrong++;
  }
  var arr = Object.values(qmap).filter(
    (x) => x.total > 0 && x.wrong / x.total > 0.4
  );
  if (arr.length <= n) return arr;
  arr.sort((a, b) => b.wrong / b.total - a.wrong / a.total);
  let topArr = arr.slice(0, Math.max(5, n * 2));
  let picked = [];
  while (picked.length < n && topArr.length > 0) {
    let idx = Math.floor(Math.random() * topArr.length);
    picked.push(topArr[idx]);
    topArr.splice(idx, 1);
  }
  return picked;
}
