// ==========================================
// ぴよログGAS - ユーティリティ関数群
// ==========================================

// ==========================================
// スプレッドシート関連
// ==========================================
function saveToSpreadsheet(spreadsheetId, data) {
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    
    var dataSheet = ss.getSheetByName('育児データ') || ss.insertSheet('育児データ');
    initializeDataSheet(dataSheet);
    saveDataToSheet(dataSheet, data);
    
    var logSheet = ss.getSheetByName('実行ログ') || ss.insertSheet('実行ログ');
    initializeLogSheet(logSheet);
    
  } catch (error) {
    console.error('スプレッドシート保存エラー:', error);
    throw error;
  }
}

function initializeDataSheet(sheet) {
  if (sheet.getLastRow() === 0) {
    var headers = [
      '記録日時', '分析期間開始', '分析期間終了', '分析日数',
      'ミルク合計(ml)', 'ミルク回数', '1回平均(ml)', '最大ミルク量(ml)',
      '睡眠時間(時間)', '睡眠セッション数', 'おしっこ回数', 'うんち回数',
      'アラート数', 'アラート詳細'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
  }
}

function saveDataToSheet(sheet, data) {
  var now = new Date();
  var alertsText = '';
  for (var i = 0; i < data.alerts.length; i++) {
    if (i > 0) alertsText += '; ';
    alertsText += '[' + data.alerts[i].severity.toUpperCase() + '] ' + data.alerts[i].message;
  }
  
  var row = [
    now,
    data.period.start,
    data.period.end,
    data.period.days,
    data.averages.milk.perDay,
    Math.round(data.averages.milk.perDay / (data.averages.milk.perFeeding || 1)),
    data.averages.milk.perFeeding,
    data.averages.milk.maxPerDay,
    data.averages.sleep.hoursPerDay,
    data.averages.sleep.sessionsPerDay,
    data.averages.diaper.peePerDay,
    data.averages.diaper.poopPerDay,
    data.alerts.length,
    alertsText
  ];
  
  sheet.appendRow(row);
  
  var lastRow = sheet.getLastRow();
  if (data.alerts.length > 0) {
    sheet.getRange(lastRow, 1, 1, row.length).setBackground('#fff3cd');
  }
}

function initializeLogSheet(sheet) {
  if (sheet.getLastRow() === 0) {
    var headers = [
      '実行日時', 'データ件数', 'アラート数', '実行時間(ms)', 'ステータス', 'エラー詳細'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#34a853');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
  }
}

function logExecution(spreadsheetId, logData) {
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var logSheet = ss.getSheetByName('実行ログ') || ss.insertSheet('実行ログ');
    
    initializeLogSheet(logSheet);
    
    var row = [
      logData.timestamp,
      logData.dataCount,
      logData.alerts,
      logData.duration,
      logData.success ? '成功' : 'エラー',
      logData.error || ''
    ];
    
    logSheet.appendRow(row);
    
    if (!logData.success) {
      var lastRow = logSheet.getLastRow();
      logSheet.getRange(lastRow, 1, 1, row.length).setBackground('#f8d7da');
    }
    
  } catch (error) {
    console.error('実行ログ記録エラー:', error);
  }
}

// ==========================================
// グラフ生成
// ==========================================
function generateCharts(spreadsheetId, data) {
  var charts = [];
  
  try {
    // スプレッドシートIDの検証
    if (!spreadsheetId || spreadsheetId.trim() === '') {
      console.warn('スプレッドシートIDが設定されていません。グラフ生成をスキップします。');
      return charts;
    }
    
    console.log('スプレッドシートにアクセス中... ID: ' + spreadsheetId);
    var ss = SpreadsheetApp.openById(spreadsheetId);
    
    // アクセス権限の確認
    try {
      var name = ss.getName();
      console.log('スプレッドシート名: ' + name);
    } catch (accessError) {
      console.error('スプレッドシートへのアクセス権限がありません:', accessError);
      throw new Error('スプレッドシートID: ' + spreadsheetId + ' にアクセスできません。IDが正しいか、共有設定を確認してください。');
    }
    
    var chartSheet = ss.getSheetByName('グラフ') || ss.insertSheet('グラフ');
    
    // 既存のチャートを削除
    var existingCharts = chartSheet.getCharts();
    for (var i = 0; i < existingCharts.length; i++) {
      chartSheet.removeChart(existingCharts[i]);
    }
    
    // データの存在確認
    if (!data || !data.trends) {
      console.warn('グラフ用データが不足しています。');
      return charts;
    }
    
    // グラフ生成
    if (data.trends.milkVolume && data.trends.milkVolume.length > 0) {
      createMilkVolumeChart(chartSheet, data.trends.milkVolume, charts);
    }
    
    if (data.trends.sleepDuration && data.trends.sleepDuration.length > 0) {
      createSleepDurationChart(chartSheet, data.trends.sleepDuration, charts);
    }
    
    createSummaryChart(chartSheet, data, charts);
    
    console.log('グラフ生成完了: ' + charts.length + '個');
    
  } catch (error) {
    console.error('グラフ生成エラー:', error);
    console.error('エラー詳細:', error.message);
    console.log('グラフ生成をスキップして処理を続行します。');
  }
  
  return charts;
}

function createMilkVolumeChart(sheet, milkData, charts) {
  var trendData = [['日付', 'ミルク合計(ml)', '最大量(ml)']];
  for (var i = 0; i < milkData.length; i++) {
    trendData.push([milkData[i].date, milkData[i].total, milkData[i].max]);
  }
  
  var range = sheet.getRange(1, 1, trendData.length, 3);
  range.setValues(trendData);
  
  var milkChart = sheet.newChart()
    .addRange(range)
    .setChartType(Charts.ChartType.LINE)
    .setPosition(2, 5, 0, 0)
    .setOption('title', 'ミルク摂取量の推移')
    .setOption('width', 600)
    .setOption('height', 400)
    .setOption('curveType', 'function')
    .setOption('legend.position', 'bottom')
    .setOption('hAxis.title', '日付')
    .setOption('vAxis.title', 'ミルク量 (ml)')
    .build();
  
  sheet.insertChart(milkChart);
  charts.push(milkChart);
}

function createSleepDurationChart(sheet, sleepData, charts) {
  var trendData = [['日付', '睡眠時間(時間)', 'セッション数']];
  for (var i = 0; i < sleepData.length; i++) {
    trendData.push([sleepData[i].date, sleepData[i].total / 60, sleepData[i].sessions]);
  }
  
  var range = sheet.getRange(1, 7, trendData.length, 3);
  range.setValues(trendData);
  
  var sleepChart = sheet.newChart()
    .addRange(range)
    .setChartType(Charts.ChartType.COMBO)
    .setPosition(2, 11, 0, 0)
    .setOption('title', '睡眠パターンの推移')
    .setOption('width', 600)
    .setOption('height', 400)
    .setOption('series.0.type', 'columns')
    .setOption('series.1.type', 'line')
    .setOption('series.1.targetAxisIndex', 1)
    .setOption('legend.position', 'bottom')
    .setOption('hAxis.title', '日付')
    .setOption('vAxes.0.title', '睡眠時間 (時間)')
    .setOption('vAxes.1.title', 'セッション数')
    .build();
  
  sheet.insertChart(sleepChart);
  charts.push(sleepChart);
}

function createSummaryChart(sheet, data, charts) {
  var summaryData = [
    ['指標', '値', '単位'],
    ['平均ミルク量/日', data.averages.milk.perDay, 'ml'],
    ['平均睡眠時間/日', data.averages.sleep.hoursPerDay, '時間'],
    ['平均おしっこ回数/日', data.averages.diaper.peePerDay, '回'],
    ['平均うんち回数/日', data.averages.diaper.poopPerDay, '回']
  ];
  
  var range = sheet.getRange(15, 1, summaryData.length, 3);
  range.setValues(summaryData);
  
  var summaryChart = sheet.newChart()
    .addRange(sheet.getRange(16, 1, summaryData.length - 1, 2))
    .setChartType(Charts.ChartType.BAR)
    .setPosition(20, 1, 0, 0)
    .setOption('title', '現在の平均値サマリー')
    .setOption('width', 500)
    .setOption('height', 300)
    .setOption('legend.position', 'none')
    .setOption('hAxis.title', '値')
    .build();
  
  sheet.insertChart(summaryChart);
  charts.push(summaryChart);
}

// ==========================================
// Slack通知機能
// ==========================================
function sendToSlack(config, data, charts, predictions) {
  var now = new Date();
  var formattedDate = Utilities.formatDate(now, config.TIMEZONE, 'yyyy/MM/dd HH:mm');
  
  var message = createSlackMessage(data, predictions, formattedDate);
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(message),
    muteHttpExceptions: true
  };
  
  try {
    var response = UrlFetchApp.fetch(config.SLACK_WEBHOOK_URL, options);
    
    if (response.getResponseCode() === 200) {
      if (config.DEBUG_MODE) {
        console.log('Slack送信成功');
      }
    } else {
      console.error('Slack送信エラー:', response.getResponseCode(), response.getContentText());
    }
    
  } catch (error) {
    console.error('Slack送信エラー:', error);
    throw error;
  }
}

function createSlackMessage(data, predictions, formattedDate) {
  var alertsAttachment = data.alerts.length > 0 ? createAlertsAttachment(data.alerts) : null;
  var insightsAttachment = createInsightsAttachment(predictions);
  
  var attachments = [
    {
      color: getMainColor(data.alerts),
      title: '🍼 育児データ分析結果 (' + formattedDate + ')',
      fields: [
        {
          title: '📊 分析期間',
          value: (data.period.start || '不明') + ' ~ ' + (data.period.end || '不明') + ' (' + data.period.days + '日間)',
          short: false
        },
        {
          title: '🍼 ミルク摂取',
          value: '平均: ' + data.averages.milk.perDay + 'ml/日\n1回平均: ' + data.averages.milk.perFeeding + 'ml',
          short: true
        },
        {
          title: '😴 睡眠',
          value: '平均: ' + data.averages.sleep.hoursPerDay + '時間/日',
          short: true
        },
        {
          title: '💧 おしっこ',
          value: '平均: ' + data.averages.diaper.peePerDay + '回/日',
          short: true
        },
        {
          title: '💩 うんち',
          value: '平均: ' + data.averages.diaper.poopPerDay + '回/日',
          short: true
        }
      ],
      footer: 'ぴよログGAS',
      ts: Math.floor(Date.now() / 1000)
    },
    insightsAttachment
  ];
  
  if (alertsAttachment) {
    attachments.splice(1, 0, alertsAttachment);
  }
  
  var filteredAttachments = [];
  for (var i = 0; i < attachments.length; i++) {
    if (attachments[i] !== null) {
      filteredAttachments.push(attachments[i]);
    }
  }
  
  return {
    text: getMainEmoji(data.alerts) + ' 育児データレポート',
    attachments: filteredAttachments
  };
}

function getMainColor(alerts) {
  for (var i = 0; i < alerts.length; i++) {
    if (alerts[i].severity === 'error') return 'danger';
  }
  for (var i = 0; i < alerts.length; i++) {
    if (alerts[i].severity === 'warning') return 'warning';
  }
  return 'good';
}

function getMainEmoji(alerts) {
  for (var i = 0; i < alerts.length; i++) {
    if (alerts[i].severity === 'error') return '🚨';
  }
  for (var i = 0; i < alerts.length; i++) {
    if (alerts[i].severity === 'warning') return '⚠️';
  }
  return '✅';
}

function createAlertsAttachment(alerts) {
  var alertTexts = [];
  for (var i = 0; i < alerts.length; i++) {
    var icon = alerts[i].severity === 'error' ? '🔴' : '🟡';
    alertTexts.push(icon + ' ' + alerts[i].message);
  }
  
  return {
    color: 'warning',
    title: '⚠️ 注意事項',
    text: alertTexts.join('\n'),
    footer: '健康状態に不安がある場合は医師にご相談ください'
  };
}

function createInsightsAttachment(predictions) {
  return {
    color: '#4a90e2',
    title: '🔮 AI予測・アドバイス',
    fields: [
      {
        title: '⏰ 次回授乳予測',
        value: predictions.nextFeeding || '約3時間後',
        short: true
      },
      {
        title: '🍼 推奨量',
        value: predictions.milkAmount || '通常量',
        short: true
      },
      {
        title: '😴 睡眠予測',
        value: predictions.sleepTime || '1-2時間後',
        short: true
      },
      {
        title: '📈 信頼度',
        value: (predictions.confidence || 60) + '%',
        short: true
      }
    ],
    text: createInsightsText(predictions)
  };
}

function createInsightsText(predictions) {
  var text = '';
  
  if (predictions.insights && predictions.insights.length > 0) {
    text += '*📝 今日の観察ポイント:*\n';
    for (var i = 0; i < predictions.insights.length; i++) {
      text += (i + 1) + '. ' + predictions.insights[i] + '\n';
    }
    text += '\n';
  }
  
  if (predictions.recommendations && predictions.recommendations.length > 0) {
    text += '*💡 おすすめアクション:*\n';
    for (var i = 0; i < predictions.recommendations.length; i++) {
      text += (i + 1) + '. ' + predictions.recommendations[i] + '\n';
    }
  }
  
  return text.trim();
}

function sendErrorNotification(config, error) {
  var message = {
    text: '🚨 ぴよログ処理エラー',
    attachments: [{
      color: 'danger',
      title: 'システムエラーが発生しました',
      fields: [
        {
          title: 'エラー内容',
          value: error.toString(),
          short: false
        },
        {
          title: '発生時刻',
          value: new Date().toLocaleString('ja-JP', { timeZone: config.TIMEZONE }),
          short: true
        }
      ],
      footer: '設定や接続を確認してください'
    }]
  };
  
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(message),
    muteHttpExceptions: true
  };
  
  try {
    UrlFetchApp.fetch(config.SLACK_WEBHOOK_URL, options);
  } catch (e) {
    console.error('エラー通知の送信に失敗:', e);
  }
}

// ==========================================
// メール処理
// ==========================================
function markEmailsAsRead(query) {
  try {
    var threads = GmailApp.search(query, 0, 10);
    for (var i = 0; i < threads.length; i++) {
      threads[i].markRead();
    }
    
    if (threads.length > 0) {
      console.log(threads.length + '件のメールを既読にしました');
    }
    
  } catch (error) {
    console.error('メール既読処理エラー:', error);
  }
}

// ==========================================
// テスト関数群
// ==========================================
function testService(service) {
  var env = EnvironmentConfig.getInstance();
  var config = env.getConfig();
  
  try {
    switch(service) {
      case 'slack':
        return testSlackConnection(config.SLACK_WEBHOOK_URL);
      case 'claude':
        return testClaudeConnection(config.ANTHROPIC_API_KEY, config.ANTHROPIC_MODEL);
      case 'gmail':
        return testGmailConnection(config.GMAIL_SEARCH_QUERY);
      default:
        return { success: false, message: '不明なサービスです' };
    }
  } catch (error) {
    return { success: false, message: 'テストエラー: ' + error.toString() };
  }
}

function testSlackConnection(webhookUrl) {
  if (!webhookUrl) {
    return { success: false, message: 'Webhook URLが設定されていません' };
  }
  
  var testMessage = {
    text: '🔧 テストメッセージ',
    attachments: [{
      color: 'good',
      title: 'ぴよログGAS接続テスト',
      fields: [
        {
          title: '状態',
          value: '正常に接続できました',
          short: true
        },
        {
          title: 'テスト時刻',
          value: new Date().toLocaleString('ja-JP'),
          short: true
        }
      ],
      footer: 'このメッセージはテスト用です'
    }]
  };
  
  try {
    var response = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(testMessage),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      return { success: true, message: 'Slack接続成功！チャンネルにテストメッセージが送信されました。' };
    } else {
      return { success: false, message: '接続エラー (' + response.getResponseCode() + '): ' + response.getContentText() };
    }
  } catch (error) {
    return { success: false, message: '接続エラー: ' + error };
  }
}

function testClaudeConnection(apiKey, model) {
  if (!apiKey) {
    return { success: false, message: 'APIキーが設定されていません' };
  }
  
  var testPayload = {
    model: model,
    max_tokens: 50,
    messages: [{
      role: 'user',
      content: 'テスト接続です。「接続成功」と日本語で返答してください。'
    }]
  };
  
  try {
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(testPayload),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      var reply = data.content && data.content[0] ? data.content[0].text : 'レスポンス取得';
      return { 
        success: true, 
        message: 'Claude API接続成功！\nモデル: ' + model + '\nレスポンス: ' + reply.substring(0, 50) + '...' 
      };
    } else {
      var errorData = JSON.parse(response.getContentText());
      return { 
        success: false, 
        message: 'APIエラー (' + response.getResponseCode() + '): ' + (errorData.error && errorData.error.message ? errorData.error.message : 'Unknown error') 
      };
    }
  } catch (error) {
    return { success: false, message: '接続エラー: ' + error };
  }
}

function testGmailConnection(query) {
  if (!query) {
    return { success: false, message: '検索クエリが設定されていません' };
  }
  
  try {
    var threads = GmailApp.search(query, 0, 5);
    var messageCount = threads.reduce(function(count, thread) {
      return count + thread.getMessageCount();
    }, 0);
    
    return { 
      success: true, 
      message: 'Gmail接続成功！\n検索クエリ: ' + query + '\n見つかったスレッド: ' + threads.length + '件\nメッセージ総数: ' + messageCount + '件' 
    };
  } catch (error) {
    return { success: false, message: 'Gmailエラー: ' + error };
  }
}

function testSystemWithSampleData() {
  var env = EnvironmentConfig.getInstance();
  var validation = env.validateRequired();
  
  if (!validation.valid) {
    return {
      success: false,
      message: '設定が完了していません。基本設定タブで必須項目を入力してください。'
    };
  }
  
  try {
    var config = env.getConfig();
    
    var sampleData = {
      period: { 
        start: '2025/01/15', 
        end: '2025/01/20', 
        days: 5 
      },
      averages: {
        milk: { perDay: 750, perFeeding: 125, maxPerDay: 180 },
        sleep: { hoursPerDay: 14.5, sessionsPerDay: 6 },
        diaper: { peePerDay: 8.2, poopPerDay: 2.4 }
      },
      trends: {
        milkVolume: [
          { date: '2025/01/18', total: 720, max: 150 },
          { date: '2025/01/19', total: 780, max: 160 },
          { date: '2025/01/20', total: 750, max: 140 }
        ],
        sleepDuration: [
          { date: '2025/01/18', total: 840, sessions: 6 },
          { date: '2025/01/19', total: 900, sessions: 5 },
          { date: '2025/01/20', total: 870, sessions: 6 }
        ],
        feedingIntervals: []
      },
      alerts: [{
        type: 'test_alert',
        message: 'これはサンプルデータによるテストです',
        severity: 'warning'
      }]
    };
    
    var samplePredictions = {
      nextFeeding: '約3時間後（サンプル予測）',
      milkAmount: '120-140ml（サンプル推奨量）',
      sleepTime: '14:00-16:00頃（サンプル睡眠予測）',
      insights: [
        'サンプルデータによる分析です',
        '実際のデータでより精度の高い予測が可能です',
        'システムは正常に動作しています'
      ],
      recommendations: [
        'このテストが成功すれば設定完了です',
        '実際のぴよログデータで運用開始できます',
        '定期実行が自動で開始されます'
      ],
      confidence: 95
    };
    
    sendToSlack(config, sampleData, [], samplePredictions);
    
    if (config.SPREADSHEET_ID) {
      try {
        saveToSpreadsheet(config.SPREADSHEET_ID, sampleData);
      } catch (error) {
        console.warn('スプレッドシート保存に失敗しましたが、テストは継続します:', error);
      }
    }
    
    return {
      success: true,
      message: 'サンプルデータテスト完了！Slackチャンネルを確認してください。問題なければ実運用開始可能です。'
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'テスト実行エラー: ' + error.toString()
    };
  }
}