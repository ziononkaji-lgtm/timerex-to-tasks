// ============================================================
// TimeREX → Google Tasks 自動化スクリプト（無料版）
// Claude APIなし・正規表現のみで解析
// ============================================================

// ============================================================
// ★ 設定（ここだけ変更OK）
// ============================================================
const CONFIG = {
  TASK_LIST_NAME: "TimeREX予定",        // Google Tasksのリスト名
  PROCESSED_LABEL: "TimeREX処理済",     // 処理済みGmailラベル名
  TIMEREX_SENDER: "notifications@timerex.net",
};

// ============================================================
// メイン処理（トリガーで自動実行）
// ============================================================
function runTimerexToTasks() {
  const threads = searchUnprocessedEmails();

  if (threads.length === 0) {
    Logger.log("新しいTimeREXメールはありません");
    return;
  }

  Logger.log(threads.length + "件のメールを処理します");

  const taskListId = getOrCreateTaskList(CONFIG.TASK_LIST_NAME);
  const processedLabel = getOrCreateLabel(CONFIG.PROCESSED_LABEL);

  threads.forEach(function(thread) {
    try {
      const message = thread.getMessages()[0];
      const body = message.getPlainBody();
      const subject = message.getSubject();

      Logger.log("処理中: " + subject);

      // 正規表現で解析
      const parsed = parseEmail(body, subject);

      if (!parsed) {
        Logger.log("解析失敗（スキップ）: " + subject);
        return;
      }

      // 定型文を生成
      const draft = buildDraft(parsed);

      // Google Tasksに追加
      const taskTitle = "【TimeREX】" + parsed.person + " " + parsed.dateFormatted;
      addTask(taskListId, taskTitle, draft, parsed.dueDate);

      // 処理済みラベルを付ける
      thread.addLabel(processedLabel);

      Logger.log("完了: " + taskTitle);

    } catch(e) {
      Logger.log("エラー: " + e.message);
    }
  });
}

// ============================================================
// メール解析（正規表現）
// ============================================================
function parseEmail(body, subject) {
  // 1. 誰からか（件名 or 本文から）
  // 例: "TimeREXから田中太郎さんが予定を追加しました"
  var personMatch = subject.match(/TimeREXから(.+?)さんが/) ||
                    body.match(/TimeREXから(.+?)さんが/);
  var person = personMatch ? personMatch[1] + "さん" : "（名前不明）";

  // 2. 日時
  // 例: "2026年6月29日 (月) 22:30 - 23:30"
  var dateMatch = body.match(/(\d{4})年(\d{1,2})月(\d{1,2})日\s*[（(]([月火水木金土日])[)）]\s*(\d{1,2}:\d{2})\s*[-－]\s*(\d{1,2}:\d{2})/);
  if (!dateMatch) {
    Logger.log("日時が見つかりません");
    return null;
  }

  var year    = dateMatch[1];
  var month   = dateMatch[2];
  var day     = dateMatch[3];
  var weekday = dateMatch[4];
  var startTime = dateMatch[5];

  // 表示用: "29日(月)22:30～"
  var dateFormatted = day + "日(" + weekday + ")" + startTime + "～";

  // Tasks期限用: "2026-06-29"
  var dueDate = year + "-" + month.padStart(2, "0") + "-" + day.padStart(2, "0");

  // 3. Web会議URL
  // Google Meet / Zoom / Teams など
  var urlMatch = body.match(/(https?:\/\/(?:meet\.google\.com|zoom\.us|teams\.microsoft\.com|us\d+web\.zoom\.us)[^\s\n]+)/);
  var url = urlMatch ? urlMatch[1].trim() : null;

  return {
    person: person,
    dateFormatted: dateFormatted,
    dueDate: dueDate,
    url: url
  };
}

// ============================================================
// 定型文を生成
// ============================================================
function buildDraft(parsed) {
  var lines = [
    "お世話になっております。",
    parsed.dateFormatted + "よろしくお願いいたします。",
  ];

  if (parsed.url) {
    lines.push("お時間になりましたらこちらからご入室ください。");
    lines.push(parsed.url);
  }

  return lines.join("\n");
}

// ============================================================
// Gmail: 未処理メールを検索
// ============================================================
function searchUnprocessedEmails() {
  var query = "from:" + CONFIG.TIMEREX_SENDER +
              " -label:" + CONFIG.PROCESSED_LABEL;
  return GmailApp.search(query, 0, 20);
}

// ============================================================
// Google Tasks: リストを取得 or 作成
// ============================================================
function getOrCreateTaskList(name) {
  var lists = Tasks.Tasklists.list().getItems() || [];
  for (var i = 0; i < lists.length; i++) {
    if (lists[i].getTitle() === name) return lists[i].getId();
  }
  var newList = Tasks.Tasklists.insert({ title: name });
  Logger.log("タスクリスト作成: " + name);
  return newList.getId();
}

// ============================================================
// Google Tasks: タスクを追加
// ============================================================
function addTask(taskListId, title, notes, dueDateStr) {
  var task = { title: title, notes: notes };
  if (dueDateStr) {
    task.due = new Date(dueDateStr).toISOString();
  }
  Tasks.Tasks.insert(task, taskListId);
  Logger.log("タスク追加: " + title);
}

// ============================================================
// Gmail: ラベルを取得 or 作成
// ============================================================
function getOrCreateLabel(name) {
  var labels = GmailApp.getUserLabels();
  for (var i = 0; i < labels.length; i++) {
    if (labels[i].getName() === name) return labels[i];
  }
  return GmailApp.createLabel(name);
}

// ============================================================
// ★ セットアップ確認用（最初に一度だけ手動で実行）
// ============================================================
function setup() {
  Logger.log("=== セットアップ確認 ===");

  // Tasks API確認
  try {
    Tasks.Tasklists.list();
    Logger.log("✅ Tasks API: OK");
  } catch(e) {
    Logger.log("❌ Tasks API: サービスを追加してください（手順③参照）");
  }

  // Gmail確認
  try {
    var results = GmailApp.search("from:" + CONFIG.TIMEREX_SENDER, 0, 5);
    Logger.log("✅ Gmail: OK（TimeREXメール " + results.length + "件検出）");
  } catch(e) {
    Logger.log("❌ Gmail: " + e.message);
  }

  Logger.log("=== 確認完了 ===");
}

// ============================================================
// ★ テスト用（サンプルデータで動作確認）
// ============================================================
function testParse() {
  var sampleBody = [
    "TimeREXから田中太郎さんが予定を追加しました。",
    "",
    "予定名: 初回ご相談",
    "日時: 2026年6月29日 (月) 22:30 - 23:30",
    "Web会議URL: https://meet.google.com/abc-defg-hij",
    "ミーティングID: 123-456-789"
  ].join("\n");

  var sampleSubject = "TimeREXから田中太郎さんが予定を追加しました";

  var parsed = parseEmail(sampleBody, sampleSubject);

  if (parsed) {
    Logger.log("✅ 解析成功");
    Logger.log("相手: " + parsed.person);
    Logger.log("日時: " + parsed.dateFormatted);
    Logger.log("期限: " + parsed.dueDate);
    Logger.log("URL: " + parsed.url);
    Logger.log("---定型文---");
    Logger.log(buildDraft(parsed));
  } else {
    Logger.log("❌ 解析失敗");
  }
}
