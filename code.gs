
/**
 * ページを開いた時に最初に呼ばれるルートメソッド
 */
function doGet(e) {
  // ログアウトのURLパラメータをチェック
  if (e.parameter.action == 'logout') {
    PropertiesService.getUserProperties().deleteAllProperties();
    return HtmlService.createHtmlOutputFromFile('view_login').setTitle("ログイン");
  }

  // ログイン情報を取得
  var properties = PropertiesService.getUserProperties();
  var userId = properties.getProperty('userId');
  
   if (userId) {
    // ログイン済みの場合、役割情報をHTMLに渡してメイン画面を表示
    var template = HtmlService.createTemplateFromFile('view_main');
    template.role = properties.getProperty('role'); // HTMLへ役割を渡す
    template.userId = userId; // HTMLへIDを渡す
    template.userName = properties.getProperty('userName'); //HTMLへ名前を渡す
    return template.evaluate().setTitle("メイン画面");

  }else {
    // 未ログインの場合、ログイン画面を表示
    return HtmlService.createHtmlOutputFromFile('view_login').setTitle("ログイン");
  }
}

/**
 * このアプリのURLを返す
 */
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

//ログイン処理
function loginUser(userId) {
  var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("従業員名簿（スプシから編集可）");
  var lastRow = empSheet.getLastRow();
  // 従業員名簿の2行目から最後までをチェック
  var empRange = empSheet.getRange(2, 1, lastRow - 1, 3); // A列からC列(role)まで取得
  var employees = empRange.getValues();

  for (var i = 0; i < employees.length; i++) {
    // 1列目(従業員番号)が一致するかチェック
    if (employees[i][0] == userId) {
      var role = employees[i][2]; // 3列目(role)を取得
      var name = employees[i][1]; // 名前を取得するコード
      
      // ログイン情報を保存
      var properties = PropertiesService.getUserProperties();
      properties.setProperty('userId', userId);
      properties.setProperty('role', role);
      properties.setProperty('userName', name); //
      
      // ログイン成功。メイン画面のURLを返す
      return getAppUrl();
    }
  }

  // ユーザーが見つからなかった
  Logger.log("→ 全データを確認しましたが、一致しませんでした。");
  return null;
}


/**
 * 従業員一覧　2025-08-05-13:07 利用予定なし
 */
function getEmployees() {
  var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1]// 「従業員名簿」のシート
  var last_row = empSheet.getLastRow()
  var empRange = empSheet.getRange(2, 1, last_row, 2);// シートの中のヘッダーを除く範囲を取得
  var employees = [];
  var i = 1;
  while (true) {
    var empId =empRange.getCell(i, 1).getValue();
    var empName =empRange.getCell(i, 2).getValue();
    if (empId === ""){ //　値を取得できなくなったら終了
      break;
    }
    employees.push({
      'id': empId,
      'name': empName
    })
    i++
  }
  return employees
}

/**
 * 従業員情報の取得　2025-08-05-13:07 利用予定なし
 * ※ デバッグするときにはselectedEmpIdを存在するIDで書き換えてください
 */
function getEmployeeName() {
  var selectedEmpId =PropertiesService.getUserProperties().getProperty('selectedEmpId') // ※デバッグするにはこの変数を直接書き換える必要があります
  var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1]// 「従業員名簿」のシート
  var last_row = empSheet.getLastRow()
  var empRange = empSheet.getRange(2, 1, last_row, 2);// シートの中のヘッダーを除く範囲を取得
  var i = 1;
  var empName = ""
  while (true) {
    var id =empRange.getCell(i, 1).getValue();
    var name =empRange.getCell(i, 2).getValue();
    if (id === ""){ 
      break;
    }
    if(id == selectedEmpId){
      empName = name
    }
    i++
  }

  return empName
}

/**
 * 勤怠情報の取得
 * 今月における今日までの勤怠情報が取得される
 */
function getTimeClocks() {
  var userId =PropertiesService.getUserProperties().getProperty('userId') // ※デバッグするにはこの変数を直接書き換える必要があります
  if(!userId){
    // ログインしていないときは空の配列を返す
    return [];
  }
  var timeClocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3]// 「打刻履歴」のシート
  var last_row = timeClocksSheet.getLastRow()
  var timeClocksRange = timeClocksSheet.getRange(2, 1, last_row, 3);// シートの中のヘッダーを除く範囲を取得
  var empTimeClocks = [];
  var i = 1;
  while (true) {
    var empId =timeClocksRange.getCell(i, 1).getValue();
    var type =timeClocksRange.getCell(i, 2).getValue();
    var datetime =timeClocksRange.getCell(i, 3).getValue();
    if (empId === ""){ //　値を取得できなくなったら終了
      break;
    }
    if (empId == userId){
      empTimeClocks.push({
        'date': Utilities.formatDate(datetime, "Asia/Tokyo", "yyyy-MM-dd HH:mm"),
        'type': type
    })
    }
    i++
  }
  return empTimeClocks
}

/**
 * 勤怠情報登録　　
 */
function saveWorkRecord(form) {
  var userId = PropertiesService.getUserProperties().getProperty('userId') // ※デバッグするにはこの変数を直接書き換える必要があります
  // inputタグのnameで取得
  var targetDate = form.target_date
  var targetTime = form.target_time
  var targetType = ''
  switch (form.target_type) {
    case 'clock_in':
      targetType = '出勤'
      break
    case 'break_begin':
      targetType = '休憩開始'
      break
    case 'break_end':
      targetType = '休憩終了'
      break
    case 'clock_out':
      targetType = '退勤'
      break;
  }
  var timeClocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3]// 「打刻履歴」のシート
  var targetRow = timeClocksSheet.getLastRow() + 1
   // ▼▼ 退勤ガードを追加：同日・退勤時刻以前の活動記録が必須 ▼▼
  if (targetType === '退勤') {
    var activitySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("活動履歴");
    var hasActivity = false;

    if (activitySheet && activitySheet.getLastRow() >= 2) {
      var actValues = activitySheet.getRange(2, 1, activitySheet.getLastRow() - 1, 4).getValues(); // [userId, datetime, title, content]
      var cutoff = new Date(targetDate + ' ' + targetTime); // 退勤の打刻時刻

      for (var i = 0; i < actValues.length; i++) {
        if (actValues[i][0] == userId) {
          var ts = new Date(actValues[i][1]);
          var sameDay = Utilities.formatDate(ts, "Asia/Tokyo", "yyyy-MM-dd") === targetDate;
          if (sameDay && ts <= cutoff) {
            hasActivity = true;
            break;
          }
        }
      }
    }

    if (!hasActivity) {
      throw new Error('ACTIVITY_REQUIRED'); // クライアントで拾って誘導
    }
  }
  timeClocksSheet.getRange(targetRow, 1).setValue(userId)
  timeClocksSheet.getRange(targetRow, 2).setValue(targetType)
  timeClocksSheet.getRange(targetRow, 3).setValue(targetDate + ' ' + targetTime)
  return '登録しました'
}



/**
 * spreadSheetに保存されている指定のemployee_idの行番号を返す
 */
function getTargetEmpRowNumber(empId) {
  // 開いているシートを取得
  var sheet = SpreadsheetApp.getActiveSheet()
  // 最終行取得
  var last_row = sheet.getLastRow()
  // 2行目から最終行までの1列目(emp_id)の範囲を取得
  var data_range = sheet.getRange(1, 1, last_row, 1);
  // 該当範囲のデータを取得
  var sheetRows = data_range.getValues();
  // ループ内で検索
  for (var i = 0; i <= sheetRows.length - 1; i++) {
    var row = sheetRows[i]
    if (row[0] == empId) {
      // spread sheetの行番号は1から始まるが配列のindexは0から始まるため + 1して行番号を返す
      return i + 1;
    }
  }
  // 見つからない場合にはnullを返す
  return null
}

/**
 * 【追加】活動内容をスプレッドシートに保存する
 */
function saveActivity(form) {
  var userId = PropertiesService.getUserProperties().getProperty('userId');
  var activitySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("活動履歴");
  
  // スプレッドシートの最終行に追記
  activitySheet.appendRow([
    userId,
    new Date(),
    form.activity_title,
    form.activity_content
  ]);
  
  return "活動を記録しました。";
}

/**
 * 【★デバッグ用】タイムライン用に、"全員"の活動履歴をページ指定で取得する
 * @param {number} page - 取得したいページ番号 (1から始まる)
 */
function getAllActivityLogs(page) {
  try {
    Logger.log("1. getAllActivityLogs関数が開始されました。リクエストされたページ: " + page); // ★追加

    const POSTS_PER_PAGE = 5;

    const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("従業員名簿（スプシから編集可）");
    if (!empSheet) {
      throw new Error("シート「従業員名簿（スプシから編集可）」が見つかりません。");
    }
    Logger.log("2. 従業員名簿シートの取得に成功しました。"); // ★追加

    const staffValues = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 3).getValues(); 
    const nameMap = {}, roleMap = {};
    
    for (let i = 0; i < staffValues.length; i++) {
      const id = staffValues[i][0];
      if(id) {
        nameMap[id] = staffValues[i][1];
        roleMap[id] = staffValues[i][2];
      }
    }
    Logger.log("3. 従業員名簿の読み込みが完了しました。登録者数: " + Object.keys(nameMap).length + "名"); // ★追加

    const activitySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("活動履歴");
    if (!activitySheet || activitySheet.getLastRow() < 1) {
      Logger.log("4. 「活動履歴」シートが存在しないか、空です。処理を終了します。"); // ★追加
      return { posts: [], hasMore: false };
    }
    Logger.log("4. 活動履歴シートの取得に成功しました。"); // ★追加

    const logs = activitySheet.getDataRange().getValues();
    Logger.log("5. 活動履歴シートの全データ（ヘッダー含む）を " + logs.length + "行取得しました。"); // ★追加
    
    const allPosts = [];

    for (let i = 1; i < logs.length; i++) {
      const userId = logs[i][0];
      if (userId) {
        allPosts.push({
          userId:   userId,
          userName: nameMap[userId] || '不明なユーザー',
          userRole: roleMap[userId] || '',
          datetime: Utilities.formatDate(new Date(logs[i][1]), "Asia/Tokyo", "yyyy-MM-dd HH:mm"),
          title:    logs[i][2],
          content:  logs[i][3]
        });
      }
    }
    Logger.log("6. 処理対象の投稿を " + allPosts.length + "件見つけました。"); // ★追加
    
    allPosts.sort((a, b) => new Date(b.datetime) - new Date(a.datetime));

    const pageNum = parseInt(page, 10) || 1;
    const startIndex = (pageNum - 1) * POSTS_PER_PAGE;
    const endIndex = startIndex + POSTS_PER_PAGE;
    
    const postsForPage = allPosts.slice(startIndex, endIndex);
    
    const result = { // ★一度変数に入れる
      posts: postsForPage,
      hasMore: allPosts.length > endIndex
    };
    
    Logger.log("7. 最終的な返り値: " + JSON.stringify(result)); // ★追加

    return result;

  } catch (e) {
    Logger.log("【致命的なエラー】処理中にクラッシュしました: " + e.message); // ★修正
    return null;
  }
}


// ===================================================================
//【ここから追加】フィードバック機能関連
// ===================================================================

/**
 * フィードバック送信先の「無給スタッフ」一覧を取得する
 * @returns {Array<Object>} 無給スタッフのIDと名前の配列 [{id: 'user002', name: '鈴木 花子'}, ...]
 */
function getStaffListForFeedback() {
  try {
    const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("従業員名簿（スプシから編集可）");
    const values = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 3).getValues(); // A列(ID), B列(名前), C列(権限) を取得

    const unpaidStaff = values.filter(function(row) {
      // 権限が「無給」のスタッフのみを抽出
      return row[2] === '無給';
    }).map(function(row) {
      // フロントエンドで使いやすいようにオブジェクト形式に変換
      return { id: row[0], name: row[1] };
    });

    return unpaidStaff;

  } catch (e) {
    Logger.log('無給スタッフ一覧の取得に失敗しました: ' + e.message);
    return []; // エラーが発生した場合は空の配列を返す
  }
}


/**
 * 受け取ったフィードバックをスプレッドシートに保存する
 * @param {Object} data - フロントエンドから渡されるデータ {recipientId: '...', content: '...'}
 * @returns {string} 処理結果のメッセージ
 */
function sendFeedback(data) {
  try {
    const senderId = PropertiesService.getUserProperties().getProperty('userId');
    if (!senderId) {
      throw new Error("ログインしていません。");
    }

    const feedbackSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フィードバック");
    if (!feedbackSheet) {
      // もし「フィードバック」シートがなければ作成する
      SpreadsheetApp.getActiveSpreadsheet().insertSheet("フィードバック").getRange("A1:E1").setValues([["フィードバックID", "送信者ID", "受信者ID", "フィードバック内容", "送信日時"]]);
    }
    
    // シートを再取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フィードバック");
    
    // スプレッドシートに新しい行として追記
    sheet.appendRow([
      Utilities.getUuid(), // 一意のフィードバックIDを自動生成
      senderId,
      data.recipientId,
      data.content,
      new Date() // 現在日時
    ]);

    return "フィードバックを送信しました。";

  } catch (e) {
    Logger.log('フィードバックの保存に失敗しました: ' + e.message);
    return "エラーが発生しました。フィードバックを送信できませんでした。";
  }
}


/**
 * ログイン中の無給スタッフ宛のフィードバックをすべて取得する
 * @returns {Array<Object>} 自分宛のフィードバックの配列
 */
function getMyFeedbacks() {
  try {
    const recipientId = PropertiesService.getUserProperties().getProperty('userId');
    if (!recipientId) {
      return []; // ログインしていなければ空の配列を返す
    }

    // --- 送信者名を取得するための準備 ---
    const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("従業員名簿（スプシから編集可）");
    const nameValues = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 2).getValues();
    const nameMap = {};
    for (var i = 0; i < nameValues.length; i++) {
      nameMap[nameValues[i][0]] = nameValues[i][1]; // { 'user001': '田中 太郎', ... } のような対応表を作成
    }

    // --- 自分宛のフィードバックを取得・整形 ---
    const feedbackSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フィードバック");
    if (!feedbackSheet) {
      return []; // シートがなければ空の配列を返す
    }
    const values = feedbackSheet.getDataRange().getValues();
    const myFeedbacks = [];
    
    // ヘッダー行を除いてループ (i=1から)
    for (var i = 1; i < values.length; i++) {
      const row = values[i];
      // 受信者IDが自分と一致するかチェック (C列、インデックス2)
      if (row[2] == recipientId) {
        const senderId = row[1];
        myFeedbacks.push({
          senderName: nameMap[senderId] || '不明なユーザー', // 送信者IDを名前に変換
          content: row[3], // 内容
          datetime: Utilities.formatDate(new Date(row[4]), "Asia/Tokyo", "yyyy-MM-dd HH:mm") // 日付をフォーマット
        });
      }
    }
    
    // 新しいフィードバックが上にくるように並び替え
    myFeedbacks.sort(function(a, b) {
      return new Date(b.datetime) - new Date(a.datetime);
    });

    return myFeedbacks;

  } catch (e) {
    Logger.log('フィードバックの取得に失敗しました: ' + e.message);
    return [];
  }
}
