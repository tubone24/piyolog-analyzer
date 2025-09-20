// ==========================================
// ぴよログ育児支援GAS - セキュア版
// ==========================================

// ==========================================
// 環境変数管理クラス
// ==========================================
function EnvironmentConfig() {
  this.scriptProperties = PropertiesService.getScriptProperties();
  this.userProperties = PropertiesService.getUserProperties();
}

EnvironmentConfig.getInstance = function() {
  if (!EnvironmentConfig._instance) {
    EnvironmentConfig._instance = new EnvironmentConfig();
  }
  return EnvironmentConfig._instance;
};

EnvironmentConfig.prototype.get = function(key) {
  return this.scriptProperties.getProperty(key) || 
         this.userProperties.getProperty(key);
};

EnvironmentConfig.prototype.set = function(key, value, isUserSpecific) {
  if (isUserSpecific) {
    this.userProperties.setProperty(key, value);
  } else {
    this.scriptProperties.setProperty(key, value);
  }
};

EnvironmentConfig.prototype.setAll = function(properties, isUserSpecific) {
  if (isUserSpecific) {
    this.userProperties.setProperties(properties);
  } else {
    this.scriptProperties.setProperties(properties);
  }
};

EnvironmentConfig.prototype.has = function(key) {
  return this.get(key) !== null;
};

EnvironmentConfig.prototype.validateRequired = function() {
  var required = [
    'SLACK_WEBHOOK_URL',
    'ANTHROPIC_API_KEY', 
    'SPREADSHEET_ID'
  ];
  
  var missing = [];
  for (var i = 0; i < required.length; i++) {
    if (!this.has(required[i])) {
      missing.push(required[i]);
    }
  }
  
  return {
    valid: missing.length === 0,
    missing: missing
  };
};

EnvironmentConfig.prototype.getConfig = function() {
  return {
    SLACK_WEBHOOK_URL: this.get('SLACK_WEBHOOK_URL') || '',
    ANTHROPIC_API_KEY: this.get('ANTHROPIC_API_KEY') || '',
    ANTHROPIC_MODEL: this.get('ANTHROPIC_MODEL') || 'claude-sonnet-4-20250514',
    GMAIL_SEARCH_QUERY: this.get('GMAIL_SEARCH_QUERY') || 'subject:"【ぴよログ】" is:unread',
    SPREADSHEET_ID: this.get('SPREADSHEET_ID') || '',
    ALERT_THRESHOLDS: JSON.parse(this.get('ALERT_THRESHOLDS') || '{}'),
    TIMEZONE: this.get('TIMEZONE') || 'Asia/Tokyo',
    EXECUTION_HOURS: JSON.parse(this.get('EXECUTION_HOURS') || '[7, 19]'),
    DEBUG_MODE: this.get('DEBUG_MODE') === 'true'
  };
};

// ==========================================
// 初期設定関数
// ==========================================
function setupEnvironment() {
  try {
    // スプレッドシートのUIが利用可能かチェック
    var ui;
    try {
      ui = SpreadsheetApp.getUi();
    } catch (uiError) {
      // スプレッドシートのUIが利用できない場合はブラウザで開く
      console.log('スプレッドシートのUIが利用できません。ブラウザで設定画面を開きます。');
      var htmlOutput = HtmlService.createHtmlOutput(getSetupHtml())
        .setWidth(700)
        .setHeight(600)
        .setTitle('ぴよログGAS環境設定');
      
      // ブラウザで開く
      var htmlUrl = htmlOutput.getUrl();
      console.log('設定画面URL: ', htmlUrl);
      throw new Error('GASエディタではなく、Googleスプレッドシートから実行してください。または新しいスプレッドシートを作成して、その中でApps Scriptを開いて実行してください。');
    }
    
    var env = EnvironmentConfig.getInstance();
    var validation = env.validateRequired();
    
    if (validation.valid) {
      var result = ui.alert(
        '設定確認',
        'すでに設定が完了しています。再設定しますか？',
        ui.ButtonSet.YES_NO
      );
      
      if (result !== ui.Button.YES) {
        return;
      }
    }
    
    var html = HtmlService.createHtmlOutput(getSetupHtml())
      .setWidth(700)
      .setHeight(600)
      .setTitle('ぴよログGAS環境設定');
    
    ui.showModalDialog(html, '環境変数設定');
  } catch (error) {
    console.error('setupEnvironment error: ' + error);
    
    // UIエラーの場合は詳細な案内を表示
    if (error.message && error.message.indexOf('getUi') !== -1) {
      console.log('=== 実行方法案内 ===');
      console.log('1. Googleスプレッドシートを新規作成してください');
      console.log('2. 作成したスプレッドシート内で「拡張機能」→「Apps Script」を選択');
      console.log('3. このコードをコピー&ペーストしてください');
      console.log('4. スプレッドシート内からsetupEnvironment()を実行してください');
      console.log('==================');
      return; // エラーをthrowせずにreturn
    }
    
    console.log('エラー: スプレッドシートを開いてからこの関数を実行してください。');
    return; // エラーをthrowせずにreturn
  }
}

// ==========================================
// スプレッドシート用のUI関数
// ==========================================
function createExecutionButton() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    
    // A1セルにボタン用の図形を作成する説明を追加
    sheet.getRange('A1').setValue('ぴよログ分析システム');
    sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
    
    sheet.getRange('A3').setValue('実行方法:');
    sheet.getRange('A3').setFontWeight('bold');
    
    sheet.getRange('A4').setValue('1. 「挿入」→「図形描画」でボタンを作成');
    sheet.getRange('A5').setValue('2. 図形を選択→「⋮」→「スクリプトを割り当て」');
    sheet.getRange('A6').setValue('3. 関数名に「main」と入力してOK');
    
    sheet.getRange('A8').setValue('または以下のメニューから実行:');
    sheet.getRange('A8').setFontWeight('bold');
    
    sheet.getRange('A9').setValue('• 手動実行: runPiyologAnalysis()');
    sheet.getRange('A10').setValue('• 設定変更: setupEnvironment()');
    sheet.getRange('A11').setValue('• テスト実行: testSystemWithSampleData()');
    
    // カスタムメニューを追加
    createCustomMenu();
    
    console.log('ボタン作成用の説明を追加しました。');
    console.log('カスタムメニューも追加されました。');
    
  } catch (error) {
    console.error('ボタン作成エラー:', error);
  }
}

function createCustomMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ぴよログ分析')
    .addItem('分析実行', 'runPiyologAnalysis')
    .addSeparator()
    .addItem('環境設定', 'setupEnvironment')
    .addItem('サンプルテスト', 'testSystemWithSampleData')
    .addSeparator()
    .addItem('実行ログ確認', 'showExecutionStatus')
    .addToUi();
}

function runPiyologAnalysis() {
  try {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
      'ぴよログ分析実行',
      '分析を開始しますか？（最新のメールから分析を行います）',
      ui.ButtonSet.YES_NO
    );
    
    if (result === ui.Button.YES) {
      ui.alert('分析開始', '分析を開始しました。完了まで少々お待ちください。', ui.ButtonSet.OK);
      main();
      ui.alert('完了', '分析が完了しました！Slackとスプレッドシートを確認してください。', ui.ButtonSet.OK);
    }
  } catch (error) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('エラー', 'エラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
  }
}

function showExecutionStatus() {
  var env = EnvironmentConfig.getInstance();
  var config = env.getConfig();
  
  var ui = SpreadsheetApp.getUi();
  var message = '=== 設定状況 ===\n';
  message += 'Slack Webhook: ' + (config.SLACK_WEBHOOK_URL ? '設定済み' : '未設定') + '\n';
  message += 'Claude API: ' + (config.ANTHROPIC_API_KEY ? '設定済み' : '未設定') + '\n';
  message += 'スプレッドシート: ' + (config.SPREADSHEET_ID ? '設定済み' : '未設定') + '\n';
  message += 'モデル: ' + config.ANTHROPIC_MODEL + '\n';
  message += '実行時刻: ' + JSON.parse(config.EXECUTION_HOURS || '[7,19]').join('時, ') + '時\n';
  message += 'デバッグモード: ' + (config.DEBUG_MODE ? 'ON' : 'OFF');
  
  ui.alert('システム状況', message, ui.ButtonSet.OK);
}

// スプレッドシートを開いた時に自動でメニューを作成
function onOpen() {
  createCustomMenu();
}

// ==========================================
// GASエディタ用の代替設定関数
// ==========================================
function setupEnvironmentInEditor() {
  console.log('=== ぴよログGAS環境設定（エディタ版） ===');
  console.log('');
  console.log('この関数はGASエディタから実行できますが、手動設定が必要です。');
  console.log('');
  console.log('以下の環境変数を手動で設定してください：');
  console.log('PropertiesService.getScriptProperties().setProperties({');
  console.log('  "SLACK_WEBHOOK_URL": "your_slack_webhook_url",');
  console.log('  "ANTHROPIC_API_KEY": "your_anthropic_api_key",');
  console.log('  "SPREADSHEET_ID": "your_spreadsheet_id",');
  console.log('  "GMAIL_SEARCH_QUERY": "subject:\\"piyolog\\" is:unread",');
  console.log('  "ANTHROPIC_MODEL": "claude-sonnet-4-20250514",');
  console.log('  "EXECUTION_HOURS": "[7, 19]",');
  console.log('  "ALERT_THRESHOLDS": "{\\"minMilkPerDay\\": 500, \\"maxMilkPerDay\\": 1200, \\"minSleepHours\\": 10, \\"maxTemperature\\": 37.5}",');
  console.log('  "DEBUG_MODE": "false"');
  console.log('});');
  console.log('');
  console.log('または、より簡単な方法：');
  console.log('1. 新しいGoogleスプレッドシートを作成');
  console.log('2. スプレッドシート内で「拡張機能」→「Apps Script」');
  console.log('3. このコードをコピー&ペースト');
  console.log('4. setupEnvironment()を実行');
  console.log('=======================================');
}

// ==========================================
// 設定画面HTML
// ==========================================
function getSetupHtml() {
  return '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<base target="_top">' +
    '<style>' +
    'body { font-family: Roboto, sans-serif; padding: 20px; background: #f5f5f5; margin: 0; }' +
    '.container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }' +
    'h2 { color: #1a73e8; margin-bottom: 20px; text-align: center; }' +
    '.form-group { margin-bottom: 20px; }' +
    'label { display: block; margin-bottom: 8px; font-weight: 500; color: #5f6368; }' +
    'input[type="text"], input[type="password"], select, textarea { width: 100%; padding: 12px 16px; border: 1px solid #dadce0; border-radius: 8px; font-size: 14px; box-sizing: border-box; }' +
    'input:focus, select:focus, textarea:focus { outline: none; border-color: #1a73e8; box-shadow: 0 0 0 2px rgba(26,115,232,0.1); }' +
    '.help-text { font-size: 12px; color: #5f6368; margin-top: 4px; line-height: 1.4; }' +
    '.required { color: #d93025; }' +
    'button { background-color: #1a73e8; color: white; border: none; padding: 12px 24px; border-radius: 8px; font-size: 14px; font-weight: 500; cursor: pointer; margin-right: 8px; transition: background-color 0.2s; }' +
    'button:hover { background-color: #1765cc; }' +
    'button:disabled { background-color: #dadce0; cursor: not-allowed; }' +
    '.secondary-btn { background-color: #ffffff; color: #1a73e8; border: 1px solid #dadce0; }' +
    '.secondary-btn:hover { background-color: #f8f9fa; }' +
    '.success { color: #188038; padding: 12px; background: #e6f4ea; border-radius: 8px; margin-bottom: 20px; }' +
    '.error { color: #d93025; padding: 12px; background: #fce8e6; border-radius: 8px; margin-bottom: 20px; }' +
    '.tabs { display: flex; border-bottom: 1px solid #dadce0; margin-bottom: 24px; }' +
    '.tab { padding: 12px 24px; cursor: pointer; border-bottom: 2px solid transparent; color: #5f6368; transition: all 0.2s; }' +
    '.tab.active { color: #1a73e8; border-bottom-color: #1a73e8; }' +
    '.tab:hover { background-color: #f8f9fa; }' +
    '.tab-content { display: none; }' +
    '.tab-content.active { display: block; }' +
    '.advanced-settings { background: #f8f9fa; padding: 20px; border-radius: 8px; margin-top: 20px; }' +
    '.grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }' +
    '.button-group { margin-top: 30px; padding-top: 20px; border-top: 1px solid #dadce0; display: flex; justify-content: space-between; align-items: center; }' +
    '.test-section { background: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 20px; }' +
    '</style>' +
    '</head>' +
    '<body>' +
    '<div class="container">' +
    '<h2>ぴよログGAS環境設定</h2>' +
    '<div id="message"></div>' +
    '<div class="tabs">' +
    '<div class="tab active" onclick="switchTab(\'basic\')">基本設定</div>' +
    '<div class="tab" onclick="switchTab(\'advanced\')">詳細設定</div>' +
    '<div class="tab" onclick="switchTab(\'test\')">テスト</div>' +
    '</div>' +
    '<div id="basic" class="tab-content active">' +
    '<div class="form-group">' +
    '<label>Slack Webhook URL <span class="required">*</span></label>' +
    '<input type="password" id="slackWebhookUrl" placeholder="https://hooks.slack.com/services/...">' +
    '<div class="help-text"><a href="https://api.slack.com/messaging/webhooks" target="_blank">取得方法はこちら</a></div>' +
    '</div>' +
    '<div class="form-group">' +
    '<label>Anthropic API Key <span class="required">*</span></label>' +
    '<input type="password" id="anthropicApiKey" placeholder="sk-ant-api03-...">' +
    '<div class="help-text"><a href="https://console.anthropic.com/settings/keys" target="_blank">取得方法はこちら</a> - 月額$1-3程度の利用料金がかかります</div>' +
    '</div>' +
    '<div class="form-group">' +
    '<label>スプレッドシートID <span class="required">*</span></label>' +
    '<input type="text" id="spreadsheetId" placeholder="1a2b3c4d5e6f...">' +
    '<div class="help-text">URLの/d/と/editの間の文字列をコピー</div>' +
    '</div>' +
    '<div class="form-group">' +
    '<label>Gmail検索クエリ</label>' +
    '<input type="text" id="gmailQuery" value="subject:\\"【ぴよログ】\\" is:unread">' +
    '<div class="help-text">ぴよログメールを検索するための条件</div>' +
    '</div>' +
    '</div>' +
    '<div id="advanced" class="tab-content">' +
    '<div class="form-group">' +
    '<label>Claudeモデル</label>' +
    '<select id="claudeModel">' +
    '<option value="claude-sonnet-4-20250514">Claude Sonnet 4 (推奨・高精度)</option>' +
    '<option value="claude-opus-4-20250514">Claude Opus 4 (最高性能)</option>' +
    '<option value="claude-3-5-sonnet-20241022">Claude 3.5 Sonnet (コスト重視)</option>' +
    '</select>' +
    '</div>' +
    '<div class="form-group">' +
    '<label>実行時刻 (カンマ区切り)</label>' +
    '<input type="text" id="executionHours" value="7,19" placeholder="7,12,19">' +
    '<div class="help-text">24時間形式で指定 (例: 7,19 = 朝7時と夜7時)</div>' +
    '</div>' +
    '<div class="advanced-settings">' +
    '<h4>アラート闾値設定</h4>' +
    '<div class="grid">' +
    '<div class="form-group">' +
    '<label>最少ミルク量/日 (ml)</label>' +
    '<input type="number" id="minMilk" value="500" placeholder="500">' +
    '</div>' +
    '<div class="form-group">' +
    '<label>最大ミルク量/日 (ml)</label>' +
    '<input type="number" id="maxMilk" value="1200" placeholder="1200">' +
    '</div>' +
    '<div class="form-group">' +
    '<label>最少睡眠時間/日 (時間)</label>' +
    '<input type="number" id="minSleep" value="10" placeholder="10">' +
    '</div>' +
    '<div class="form-group">' +
    '<label>体温警告値 (度)</label>' +
    '<input type="number" id="maxTemp" value="37.5" step="0.1" placeholder="37.5">' +
    '</div>' +
    '</div>' +
    '</div>' +
    '<div class="form-group">' +
    '<label><input type="checkbox" id="debugMode"> デバッグモード有効</label>' +
    '<div class="help-text">詳細なログを出力します</div>' +
    '</div>' +
    '</div>' +
    '<div id="test" class="tab-content">' +
    '<div class="test-section">' +
    '<h4>接続テスト</h4>' +
    '<p>設定した情報でサービスへの接続を確認します。</p>' +
    '<button onclick="testConnection(\'slack\')">Slackテスト</button>' +
    '<button onclick="testConnection(\'claude\')">Claude APIテスト</button>' +
    '<button onclick="testConnection(\'gmail\')">Gmailテスト</button>' +
    '</div>' +
    '<div class="test-section">' +
    '<h4>サンプルデータテスト</h4>' +
    '<p>サンプルデータを使用して完全な処理フローをテストします。</p>' +
    '<button onclick="testSampleData()">サンプルデータで実行</button>' +
    '</div>' +
    '<div id="testResults"></div>' +
    '</div>' +
    '<div class="button-group">' +
    '<div>' +
    '<button onclick="saveSettings()">保存</button>' +
    '<button class="secondary-btn" onclick="google.script.host.close()">キャンセル</button>' +
    '</div>' +
    '<button class="secondary-btn" onclick="clearSettings()">設定をクリア</button>' +
    '</div>' +
    '</div>' +
    '<script>' +
    'function switchTab(tabName) {' +
    'var tabs = document.querySelectorAll(".tab");' +
    'for (var i = 0; i < tabs.length; i++) {' +
    'tabs[i].classList.remove("active");' +
    '}' +
    'var contents = document.querySelectorAll(".tab-content");' +
    'for (var i = 0; i < contents.length; i++) {' +
    'contents[i].classList.remove("active");' +
    '}' +
    'event.target.classList.add("active");' +
    'document.getElementById(tabName).classList.add("active");' +
    '}' +
    'function saveSettings() {' +
    'var executionHours = document.getElementById("executionHours").value.split(",");' +
    'var hoursArray = [];' +
    'for (var i = 0; i < executionHours.length; i++) {' +
    'hoursArray.push(parseInt(executionHours[i].trim()));' +
    '}' +
    'var settings = {' +
    'SLACK_WEBHOOK_URL: document.getElementById("slackWebhookUrl").value,' +
    'ANTHROPIC_API_KEY: document.getElementById("anthropicApiKey").value,' +
    'SPREADSHEET_ID: document.getElementById("spreadsheetId").value,' +
    'GMAIL_SEARCH_QUERY: document.getElementById("gmailQuery").value,' +
    'ANTHROPIC_MODEL: document.getElementById("claudeModel").value,' +
    'EXECUTION_HOURS: JSON.stringify(hoursArray),' +
    'ALERT_THRESHOLDS: JSON.stringify({' +
    'minMilkPerDay: parseInt(document.getElementById("minMilk").value) || 500,' +
    'maxMilkPerDay: parseInt(document.getElementById("maxMilk").value) || 1200,' +
    'minSleepHours: parseInt(document.getElementById("minSleep").value) || 10,' +
    'maxTemperature: parseFloat(document.getElementById("maxTemp").value) || 37.5' +
    '}),' +
    'DEBUG_MODE: document.getElementById("debugMode").checked.toString()' +
    '};' +
    'if (!settings.SLACK_WEBHOOK_URL || !settings.ANTHROPIC_API_KEY || !settings.SPREADSHEET_ID) {' +
    'showMessage("必須項目を入力してください", "error");' +
    'return;' +
    '}' +
    'google.script.run' +
    '.withSuccessHandler(function(result) {' +
    'if (result.success) {' +
    'showMessage("設定を保存しました！自動実行も設定されました。", "success");' +
    'setTimeout(function() { google.script.host.close(); }, 3000);' +
    '} else {' +
    'showMessage("保存に失敗しました: " + result.error, "error");' +
    '}' +
    '})' +
    '.withFailureHandler(function(error) {' +
    'showMessage("エラー: " + error, "error");' +
    '})' +
    '.saveEnvironmentSettings(settings);' +
    '}' +
    'function testConnection(service) {' +
    'document.getElementById("testResults").innerHTML = "<p>テスト中...</p>";' +
    'google.script.run' +
    '.withSuccessHandler(function(result) {' +
    'var color = result.success ? "success" : "error";' +
    'var icon = result.success ? "OK" : "NG";' +
    'document.getElementById("testResults").innerHTML = ' +
    '"<div class=\\"" + color + "\\">" + icon + " " + result.message + "</div>";' +
    '})' +
    '.withFailureHandler(function(error) {' +
    'document.getElementById("testResults").innerHTML = ' +
    '"<div class=\\"error\\">NG テスト失敗: " + error + "</div>";' +
    '})' +
    '.testService(service);' +
    '}' +
    'function testSampleData() {' +
    'document.getElementById("testResults").innerHTML = "<p>サンプルデータテスト実行中...</p>";' +
    'google.script.run' +
    '.withSuccessHandler(function(result) {' +
    'if (result.success) {' +
    'document.getElementById("testResults").innerHTML = ' +
    '"<div class=\\"success\\">OK " + result.message + "</div>";' +
    '} else {' +
    'document.getElementById("testResults").innerHTML = ' +
    '"<div class=\\"error\\">NG " + result.message + "</div>";' +
    '}' +
    '})' +
    '.withFailureHandler(function(error) {' +
    'document.getElementById("testResults").innerHTML = ' +
    '"<div class=\\"error\\">NG エラー: " + error + "</div>";' +
    '})' +
    '.testSystemWithSampleData();' +
    '}' +
    'function clearSettings() {' +
    'if (confirm("すべての設定をクリアしますか？この作業は元に戻せません。")) {' +
    'google.script.run' +
    '.withSuccessHandler(function() {' +
    'showMessage("設定をクリアしました", "success");' +
    'setTimeout(function() { location.reload(); }, 1000);' +
    '})' +
    '.clearEnvironmentSettings();' +
    '}' +
    '}' +
    'function showMessage(text, type) {' +
    'document.getElementById("message").innerHTML = ' +
    '"<div class=\\"" + type + "\\">" + text + "</div>";' +
    'setTimeout(function() {' +
    'document.getElementById("message").innerHTML = "";' +
    '}, 5000);' +
    '}' +
    'google.script.run' +
    '.withSuccessHandler(function(config) {' +
    'if (config.SLACK_WEBHOOK_URL) ' +
    'document.getElementById("slackWebhookUrl").value = config.SLACK_WEBHOOK_URL;' +
    'if (config.ANTHROPIC_API_KEY) ' +
    'document.getElementById("anthropicApiKey").value = config.ANTHROPIC_API_KEY;' +
    'if (config.SPREADSHEET_ID) ' +
    'document.getElementById("spreadsheetId").value = config.SPREADSHEET_ID;' +
    'if (config.GMAIL_SEARCH_QUERY) ' +
    'document.getElementById("gmailQuery").value = config.GMAIL_SEARCH_QUERY;' +
    'if (config.ANTHROPIC_MODEL) ' +
    'document.getElementById("claudeModel").value = config.ANTHROPIC_MODEL;' +
    'if (config.EXECUTION_HOURS) {' +
    'var hours = JSON.parse(config.EXECUTION_HOURS);' +
    'document.getElementById("executionHours").value = hours.join(",");' +
    '}' +
    'if (config.ALERT_THRESHOLDS) {' +
    'var thresholds = JSON.parse(config.ALERT_THRESHOLDS);' +
    'if (thresholds.minMilkPerDay) ' +
    'document.getElementById("minMilk").value = thresholds.minMilkPerDay;' +
    'if (thresholds.maxMilkPerDay) ' +
    'document.getElementById("maxMilk").value = thresholds.maxMilkPerDay;' +
    'if (thresholds.minSleepHours) ' +
    'document.getElementById("minSleep").value = thresholds.minSleepHours;' +
    'if (thresholds.maxTemperature) ' +
    'document.getElementById("maxTemp").value = thresholds.maxTemperature;' +
    '}' +
    'if (config.DEBUG_MODE === "true") ' +
    'document.getElementById("debugMode").checked = true;' +
    '})' +
    '.loadCurrentSettings();' +
    '</script>' +
    '</body>' +
    '</html>';
}

// ==========================================
// 設定管理用の公開関数
// ==========================================
function saveEnvironmentSettings(settings) {
  try {
    var env = EnvironmentConfig.getInstance();
    env.setAll(settings);
    
    if (settings.EXECUTION_HOURS) {
      var hours = JSON.parse(settings.EXECUTION_HOURS);
      updateTriggers(hours);
    }
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function loadCurrentSettings() {
  var env = EnvironmentConfig.getInstance();
  var config = env.getConfig();
  
  if (config.SLACK_WEBHOOK_URL) {
    config.SLACK_WEBHOOK_URL = maskSecret(config.SLACK_WEBHOOK_URL);
  }
  if (config.ANTHROPIC_API_KEY) {
    config.ANTHROPIC_API_KEY = maskSecret(config.ANTHROPIC_API_KEY);
  }
  
  return config;
}

function clearEnvironmentSettings() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteAllProperties();
  
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties();
  
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function maskSecret(secret) {
  if (secret.length <= 8) return '****';
  return secret.substring(0, 4) + '****' + secret.substring(secret.length - 4);
}

// ==========================================
// トリガー管理
// ==========================================
function updateTriggers(hours) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  for (var i = 0; i < hours.length; i++) {
    ScriptApp.newTrigger('main')
      .timeBased()
      .atHour(hours[i])
      .everyDays(1)
      .create();
  }
  
  console.log('トリガーを設定しました: ' + hours.join(', ') + '時');
}

// ==========================================
// メイン処理
// ==========================================
function main() {
  var env = EnvironmentConfig.getInstance();
  
  var validation = env.validateRequired();
  if (!validation.valid) {
    console.error('必須の環境変数が設定されていません:', validation.missing);
    console.log('setupEnvironment() を実行して設定してください');
    return;
  }
  
  var config = env.getConfig();
  
  if (config.DEBUG_MODE) {
    console.log('デバッグモード: 処理開始', {
      timestamp: new Date().toISOString(),
      model: config.ANTHROPIC_MODEL,
      executionHours: config.EXECUTION_HOURS
    });
  }
  
  try {
    executeMainProcess(config);
  } catch (error) {
    console.error('メイン処理エラー:', error);
    sendErrorNotification(config, error);
  }
}

function sendErrorNotification(config, error) {
  try {
    if (config.SLACK_WEBHOOK_URL) {
      var payload = {
        text: '⚠️ ぴよログGASエラー',
        attachments: [{
          color: 'danger',
          fields: [{
            title: 'エラー詳細',
            value: error.toString(),
            short: false
          }, {
            title: '発生時刻',
            value: new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' }),
            short: true
          }]
        }]
      };
      
      UrlFetchApp.fetch(config.SLACK_WEBHOOK_URL, {
        method: 'post',
        headers: { 'Content-Type': 'application/json' },
        payload: JSON.stringify(payload)
      });
    }
  } catch (notificationError) {
    console.error('エラー通知送信失敗:', notificationError);
  }
}

function executeMainProcess(config) {
  var startTime = new Date();
  
  var piyologData = getPiyologDataFromGmail(config.GMAIL_SEARCH_QUERY);
  
  if (!piyologData || piyologData.length === 0) {
    if (config.DEBUG_MODE) {
      console.log('新しいぴよログデータがありません');
    }
    return;
  }
  
  var analyzedData = analyzePiyologData(piyologData, config.ALERT_THRESHOLDS);
  
  saveToSpreadsheet(config.SPREADSHEET_ID, analyzedData);
  
  var charts = generateCharts(config.SPREADSHEET_ID, analyzedData);
  
  var predictions = getPredictionsFromClaude(config, analyzedData);
  
  sendToSlack(config, analyzedData, charts, predictions);
  
  logExecution(config.SPREADSHEET_ID, {
    timestamp: startTime,
    dataCount: piyologData.length,
    alerts: analyzedData.alerts.length,
    duration: new Date().getTime() - startTime.getTime(),
    success: true
  });
  
  markEmailsAsRead(config.GMAIL_SEARCH_QUERY);
  
  if (config.DEBUG_MODE) {
    console.log('処理完了', {
      duration: new Date().getTime() - startTime.getTime(),
      dataCount: piyologData.length,
      alertCount: analyzedData.alerts.length
    });
  }
}

// ==========================================
// データ取得・解析関数
// ==========================================
function getPiyologDataFromGmail(query) {
  var threads = GmailApp.search(query, 0, 10);
  var allData = [];
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var body = messages[j].getPlainBody();
      var date = messages[j].getDate();
      
      var parsedData = parsePiyologText(body, date);
      if (parsedData) {
        allData.push(parsedData);
      }
    }
  }
  
  return allData;
}

function parsePiyologText(text, emailDate) {
  var lines = text.split('\n');
  var data = {
    date: null,
    babyName: '',
    age: '',
    events: [],
    summary: {
      milk: { count: 0, total: 0, max: 0 },
      breastMilk: { left: 0, right: 0 },
      sleep: { total: 0, sessions: [] },
      diaper: { pee: 0, poop: 0 }
    }
  };
  
  var currentDate = null;
  
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    
    if (line.match(/^\d{4}\/\d{1,2}\/\d{1,2}/)) {
      var dateMatch = line.match(/^(\d{4}\/\d{1,2}\/\d{1,2})/);
      if (dateMatch) {
        currentDate = dateMatch[1];
        data.date = currentDate;
      }
    }
    
    if (line.indexOf('歳') !== -1 && line.indexOf('か月') !== -1) {
      var nameMatch = line.match(/^(.+?)\s+\(/);
      if (nameMatch) {
        data.babyName = nameMatch[1];
      }
      var ageMatch = line.match(/\((.+?)\)/);
      if (ageMatch) {
        data.age = ageMatch[1];
      }
    }
    
    var timeEventMatch = line.match(/^(\d{2}:\d{2})\s+(.+)/);
    if (timeEventMatch) {
      var time = timeEventMatch[1];
      var event = timeEventMatch[2];
      
      data.events.push({
        time: time,
        event: event,
        type: categorizeEvent(event)
      });
      
      if (event.indexOf('ミルク') !== -1) {
        var mlMatch = event.match(/(\d+)ml/);
        if (mlMatch) {
          var amount = parseInt(mlMatch[1]);
          data.summary.milk.count++;
          data.summary.milk.total += amount;
          data.summary.milk.max = Math.max(data.summary.milk.max, amount);
        }
      }
      
      if (event.indexOf('寝る') !== -1) {
        data.summary.sleep.sessions.push({ start: time, end: null });
      }
      if (event.indexOf('起きる') !== -1) {
        var durationMatch = event.match(/\((.+?)\)/);
        if (durationMatch && data.summary.sleep.sessions.length > 0) {
          var lastSession = data.summary.sleep.sessions[data.summary.sleep.sessions.length - 1];
          if (lastSession) {
            lastSession.end = time;
            lastSession.duration = durationMatch[1];
          }
        }
      }
    }
    
    if (line.indexOf('母乳合計') !== -1) {
      var leftMatch = line.match(/左\s*(\d+)分/);
      var rightMatch = line.match(/右\s*(\d+)分/);
      if (leftMatch) data.summary.breastMilk.left = parseInt(leftMatch[1]);
      if (rightMatch) data.summary.breastMilk.right = parseInt(rightMatch[1]);
    }
    
    if (line.indexOf('睡眠合計') !== -1) {
      var sleepMatch = line.match(/(\d+)時間(\d+)分/);
      if (sleepMatch) {
        data.summary.sleep.total = parseInt(sleepMatch[1]) * 60 + parseInt(sleepMatch[2]);
      }
    }
    
    if (line.indexOf('おしっこ合計') !== -1) {
      var peeMatch = line.match(/(\d+)回/);
      if (peeMatch) data.summary.diaper.pee = parseInt(peeMatch[1]);
    }
    
    if (line.indexOf('うんち合計') !== -1) {
      var poopMatch = line.match(/(\d+)回/);
      if (poopMatch) data.summary.diaper.poop = parseInt(poopMatch[1]);
    }
  }
  
  return data.date ? data : null;
}

function categorizeEvent(event) {
  if (event.indexOf('ミルク') !== -1 || event.indexOf('母乳') !== -1) return 'feeding';
  if (event.indexOf('寝る') !== -1 || event.indexOf('起きる') !== -1) return 'sleep';
  if (event.indexOf('おしっこ') !== -1 || event.indexOf('うんち') !== -1) return 'diaper';
  if (event.indexOf('お風呂') !== -1 || event.indexOf('沐浴') !== -1) return 'bath';
  if (event.indexOf('体温') !== -1) return 'temperature';
  return 'other';
}

function analyzePiyologData(dataArray, alertThresholds) {
  var analysis = {
    period: {
      start: null,
      end: null,
      days: 0
    },
    averages: {
      milk: { perDay: 0, perFeeding: 0, maxPerDay: 0 },
      sleep: { hoursPerDay: 0, sessionsPerDay: 0 },
      diaper: { peePerDay: 0, poopPerDay: 0 }
    },
    trends: {
      milkVolume: [],
      sleepDuration: [],
      feedingIntervals: []
    },
    alerts: []
  };
  
  if (dataArray.length === 0) return analysis;
  
  var dates = [];
  for (var i = 0; i < dataArray.length; i++) {
    if (dataArray[i].date) {
      dates.push(dataArray[i].date);
    }
  }
  dates.sort();
  
  analysis.period.start = dates[0];
  analysis.period.end = dates[dates.length - 1];
  analysis.period.days = dataArray.length;
  
  var totalMilk = 0;
  var totalFeedings = 0;
  var totalSleep = 0;
  var totalPee = 0;
  var totalPoop = 0;
  var maxMilkPerDay = 0;
  
  for (var i = 0; i < dataArray.length; i++) {
    var data = dataArray[i];
    totalMilk += data.summary.milk.total;
    totalFeedings += data.summary.milk.count;
    totalSleep += data.summary.sleep.total;
    totalPee += data.summary.diaper.pee;
    totalPoop += data.summary.diaper.poop;
    maxMilkPerDay = Math.max(maxMilkPerDay, data.summary.milk.total);
    
    analysis.trends.milkVolume.push({
      date: data.date || '',
      total: data.summary.milk.total,
      max: data.summary.milk.max
    });
    
    analysis.trends.sleepDuration.push({
      date: data.date || '',
      total: data.summary.sleep.total,
      sessions: data.summary.sleep.sessions.length
    });
    
    checkAlerts(data, alertThresholds, analysis.alerts);
  }
  
  var days = analysis.period.days || 1;
  analysis.averages.milk.perDay = Math.round(totalMilk / days);
  analysis.averages.milk.perFeeding = totalFeedings > 0 ? Math.round(totalMilk / totalFeedings) : 0;
  analysis.averages.milk.maxPerDay = maxMilkPerDay;
  analysis.averages.sleep.hoursPerDay = Math.round((totalSleep / days) / 60 * 10) / 10;
  analysis.averages.diaper.peePerDay = Math.round(totalPee / days * 10) / 10;
  analysis.averages.diaper.poopPerDay = Math.round(totalPoop / days * 10) / 10;
  
  return analysis;
}

function checkAlerts(data, thresholds, alerts) {
  if (thresholds.minMilkPerDay && data.summary.milk.total < thresholds.minMilkPerDay) {
    alerts.push({
      type: 'milk_low',
      message: 'ミルク摂取量が少なめです (' + data.summary.milk.total + 'ml < ' + thresholds.minMilkPerDay + 'ml)',
      severity: 'warning'
    });
  }
  
  if (thresholds.maxMilkPerDay && data.summary.milk.total > thresholds.maxMilkPerDay) {
    alerts.push({
      type: 'milk_high',
      message: 'ミルク摂取量が多めです (' + data.summary.milk.total + 'ml > ' + thresholds.maxMilkPerDay + 'ml)',
      severity: 'warning'
    });
  }
  
  var sleepHours = data.summary.sleep.total / 60;
  if (thresholds.minSleepHours && sleepHours < thresholds.minSleepHours) {
    alerts.push({
      type: 'sleep_low',
      message: '睡眠時間が少なめです (' + sleepHours.toFixed(1) + '時間 < ' + thresholds.minSleepHours + '時間)',
      severity: 'warning'
    });
  }
  
  for (var i = 0; i < data.events.length; i++) {
    var event = data.events[i];
    if (event.type === 'temperature') {
      var tempMatch = event.event.match(/(\d+\.?\d*)度/);
      if (tempMatch) {
        var temp = parseFloat(tempMatch[1]);
        if (thresholds.maxTemperature && temp > thresholds.maxTemperature) {
          alerts.push({
            type: 'temperature_high',
            message: '体温が高めです (' + temp + '度 > ' + thresholds.maxTemperature + '度) [' + event.time + ']',
            severity: 'error'
          });
        }
      }
    }
  }
}

// ==========================================
// Claude API連携
// ==========================================
function getPredictionsFromClaude(config, data) {
  var predictions = {
    nextFeeding: '',
    milkAmount: '',
    sleepTime: '',
    insights: [],
    recommendations: [],
    confidence: 0
  };
  
  try {
    var prompt = createPrompt(data);
    var response = callClaudeAPI(config, prompt);
    
    if (response) {
      return parseClaudeResponse(response, data);
    }
    
  } catch (error) {
    console.error('Claude APIエラー:', error);
  }
  
  return generateFallbackPredictions(data);
}

function createPrompt(data) {
  var alertsText = '';
  if (data.alerts.length > 0) {
    alertsText = '\n注意点:\n';
    for (var i = 0; i < data.alerts.length; i++) {
      alertsText += '- ' + data.alerts[i].message + '\n';
    }
  }
  
  var recentMilk = '';
  if (data.trends.milkVolume.length > 0) {
    var recent = data.trends.milkVolume.slice(-3);
    var milkTexts = [];
    for (var i = 0; i < recent.length; i++) {
      milkTexts.push(recent[i].date + ': ' + recent[i].total + 'ml');
    }
    recentMilk = milkTexts.join(', ');
  }
    
  return 'あなたは育児支援AIアシスタントです。以下のぴよログデータを分析して科学的根拠に基づく予測とアドバイスを提供してください。\n\n' +
    '分析期間:\n' + data.period.start + ' から ' + data.period.end + ' (' + data.period.days + ' 日間)\n\n' +
    '現在の平均値:\n' +
    '- ミルク: 1日あたり ' + data.averages.milk.perDay + 'ml, 1回あたり ' + data.averages.milk.perFeeding + 'ml\n' +
    '- 睡眠: 1日あたり ' + data.averages.sleep.hoursPerDay + ' 時間\n' +
    '- おしっこ: 1日あたり ' + data.averages.diaper.peePerDay + ' 回\n' +
    '- うんち: 1日あたり ' + data.averages.diaper.poopPerDay + ' 回\n\n' +
    'トレンド:\n' +
    '最近3日間のミルク摂取量: ' + recentMilk + '\n' +
    alertsText + '\n\n' +
    '以下の形式で回答してください:\n\n' +
    '次回授乳予測\n[時間の予測とその根拠]\n\n' +
    '推奨ミルク量\n[量の予測とその根拠]\n\n' +
    '睡眠パターン分析\n[次の睡眠時間の予測]\n\n' +
    '成長に関する洞察\n1. [観察されたパターン]\n2. [成長の兆候]\n3. [注意すべき点]\n\n' +
    '今日のお世話のポイント\n1. [具体的なアドバイス]\n2. [タイミングの提案]\n3. [観察すべき点]\n\n' +
    '回答は具体的で実用的な内容にしてください。';
}

function callClaudeAPI(config, prompt) {
  var apiUrl = 'https://api.anthropic.com/v1/messages';
  var headers = {
    'Content-Type': 'application/json',
    'x-api-key': config.ANTHROPIC_API_KEY,
    'anthropic-version': '2023-06-01'
  };
  
  var payload = {
    model: config.ANTHROPIC_MODEL,
    max_tokens: 2000,
    temperature: 0.3,
    messages: [{
      role: 'user',
      content: prompt
    }]
  };
  
  var options = {
    method: 'post',
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(apiUrl, options);
  
  if (response.getResponseCode() === 200) {
    var responseData = JSON.parse(response.getContentText());
    return responseData.content[0].text;
  } else {
    console.error('Claude API error:', response.getResponseCode(), response.getContentText());
    return null;
  }
}

function parseClaudeResponse(response, data) {
  var predictions = {
    nextFeeding: '',
    milkAmount: '',
    sleepTime: '',
    insights: [],
    recommendations: [],
    confidence: 85
  };
  
  var sections = response.split('\n\n');
  
  for (var i = 0; i < sections.length; i++) {
    var section = sections[i];
    var lines = section.trim().split('\n');
    var title = lines[0];
    var content = lines.slice(1).join('\n').trim();
    
    if (title.indexOf('次回授乳') !== -1) {
      predictions.nextFeeding = content.substring(0, 100);
    } else if (title.indexOf('ミルク量') !== -1) {
      predictions.milkAmount = content.substring(0, 100);
    } else if (title.indexOf('睡眠') !== -1) {
      predictions.sleepTime = content.substring(0, 100);
    } else if (title.indexOf('洞察') !== -1) {
      var insights = content.split('\n');
      var validInsights = [];
      for (var j = 0; j < insights.length && validInsights.length < 3; j++) {
        if (insights[j].indexOf('.') !== -1) {
          validInsights.push(insights[j]);
        }
      }
      predictions.insights = validInsights;
    } else if (title.indexOf('ポイント') !== -1) {
      var recommendations = content.split('\n');
      var validRecommendations = [];
      for (var j = 0; j < recommendations.length && validRecommendations.length < 3; j++) {
        if (recommendations[j].indexOf('.') !== -1) {
          validRecommendations.push(recommendations[j]);
        }
      }
      predictions.recommendations = validRecommendations;
    }
  }
  
  return predictions;
}

function generateFallbackPredictions(data) {
  return {
    nextFeeding: '約3時間後',
    milkAmount: data.averages.milk.perFeeding + 'ml前後',
    sleepTime: '1-2時間後',
    insights: ['順調に成長しています', 'パターンが安定してきています'],
    recommendations: ['規則正しいリズムを心がけましょう', '体調の変化に注意してください'],
    confidence: 60
  };
}