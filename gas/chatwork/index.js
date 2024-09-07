"use strict";

// Chatwork APIの設定
var CHATWORK_API_TOKEN = PropertiesService.getScriptProperties().getProperty("CHATWORK_API_TOKEN");
var CHATWORK_ROOM_ID = PropertiesService.getScriptProperties().getProperty("CHATWORK_ROOM_ID");
var CHATWORK_API_ENDPOINT = PropertiesService.getScriptProperties().getProperty("CHATWORK_API_ENDPOINT");

// OpenAI APIの設定
var OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
var OPENAI_API_ENDPOINT = PropertiesService.getScriptProperties().getProperty("OPENAI_API_ENDPOINT");

// スプレッドシートの設定
var SPREADSHEET_URL = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL");
var SHEET_NAME = PropertiesService.getScriptProperties().getProperty("SHEET_NAME");

// メイン処理
// Chatworkのメッセージを取得して、問い合わせメッセージ文字列が含まれている場合、OpenAIで回答文を生成して応答する
function main() {
  var url = CHATWORK_API_ENDPOINT + '/rooms/' + CHATWORK_ROOM_ID + '/messages';
  var options = {
    method: 'get',
    headers: {
      'X-ChatWorkToken': CHATWORK_API_TOKEN
    }
  };

  var response = UrlFetchApp.fetch(url, options);
  var responseText = response.getContentText();

  Logger.log(response);
  try {
    var messages = JSON.parse(responseText);
    
    if (!messages || messages.length === 0) {
      Logger.log('No new messages found.');
      return;
    }

    messages.forEach(function(message) {
      Logger.log(message.body);
      if (message.body.includes('問い合わせです')) {
        logMessageToSheet(message);
        var question = message.body.replace('ChatGPT', '').trim();
        var answer = getChatGPTResponse(question);
        sendMessage('【AI回答】' + '\n' + '問い合わせありがとうございます。OpenAIが自動回答します。' + '\n' + answer);
      }
    });

  } catch (e) {
    Logger.log('Failed to parse JSON response: ' + e.message);
    Logger.log('Response text: ' + responseText);
  }
}

// OpenAIに質問を送信し、回答を取得する関数
function getChatGPTResponse(question) {
//var openAiModel = 'gpt-3.5-turbo';  
  var openAiModel = 'gpt-4o';
  var openAiContext = '貴方はヘルプデスクの専門家です。質問について、10行未満で、日本語で簡潔に回答してください.';

  var data = {
    model: openAiModel,
    messages: [
      { role: 'system', content: openAiContext },
      { role: 'user', content: question }
    ],
    max_tokens: 100
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + OPENAI_API_KEY
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: true // エラーメッセージの詳細を確認
  };

  var response = UrlFetchApp.fetch(OPENAI_API_ENDPOINT, options);
  var responseData = JSON.parse(response.getContentText());
  return responseData.choices[0].message.content.trim();
}

// スプレッドシートにメッセージを記録する関数
function logMessageToSheet(message) {
  var spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  var lastRow = sheet.getLastRow();
  sheet.appendRow([new Date(), message.account.name, message.body]);
}

// Chatworkにメッセージを送信する関数
function sendMessage(message) {
  if (!message || message.trim() === '') {
    Logger.log('メッセージが空です。');
    return;
  }

  var url = CHATWORK_API_ENDPOINT + '/rooms/' + CHATWORK_ROOM_ID + '/messages';
  var payload = {
    body: message
  };

  var options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {
      'X-ChatWorkToken': CHATWORK_API_TOKEN
    },
    payload: payload
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}

// 定期的に実行するためのトリガー設定
function createTimeDrivenTriggers() {
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyMinutes(5)  // 5分ごとに実行
    .create();
}

// 初期設定を行う関数
function setup() {
  createTimeDrivenTriggers();
}
