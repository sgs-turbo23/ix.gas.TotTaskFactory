const SHEET_NAME = "main";
let slack;

function main() {
  slack = new slackNotifier(secret.getSlackId, "Task Factory");
  try {
    // スケジュール用スプシを読み込み
    const book = SpreadsheetApp.openById(secret.getBookId());
    const sheet = book.getSheetByName(SHEET_NAME);

    // 本日
    const today = new Date();
    const notifies = sheet.getDataRange().getValues();
    const columns = notifies[0];
    for (var i = 1; i < notifies.length; i++) {
      const row = makeDictionary(columns, notifies[i]);

      if (today.getDate() === row['NotificationDate']) {
        addTask(
          row['Name'],
          row['Note'],
          new Date(`${today.getFullYear()}/${today.getMonth() + 1}/${row['DueDate']}`)
          );
        slack.postToSlack(`<@ma.iw>定期的なタスクを保存しました\n${row['Name']}`);
      }
    }
  } catch (error) {
    slack.postToSlack(`エラーが発生しました。\n${error}`);
  }
}

function makeDictionary(columns, row) {
  var ret = {};
  for (var i = 0; i < columns.length; i++) {
    ret[columns[i]] = row[i];
  }
  return ret;
}

function addTask(title, notes, due) {
 const dueStr = Utilities.formatDate(due, "Asia/Tokyo", "yyyy-MM-dd");
 const task = {
   title: title,
   notes: notes,
   due: dueStr + "T00:00:00.000Z"
 };
 
 // タスク追加
 Tasks.Tasks.insert(task, secret.getTasklistId());
 Utilities.sleep(500);
}

