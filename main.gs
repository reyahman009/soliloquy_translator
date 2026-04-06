/**
 * 語学学習サポートAI (スプレッドシート連携)
 * バックエンド処理スクリプト (main.gs)
 */

// GASの「スクリプトプロパティ」から環境変数（APIキーなど）を読み込むためのオブジェクト
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

// --- カスタムメニュー ---

/**
 * スプレッドシートを開いた時に自動でカスタムメニューを追加するトリガー関数
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('語学学習AI')
    .addItem('初期設定（APIキーの登録）', 'showApiKeyDialog')
    .addItem('アプリを公開（デプロイ）', 'showDeployGuide')
    .addToUi();
}

/**
 * APIキー入力ダイアログを表示する
 */
function showApiKeyDialog() {
  const ui = SpreadsheetApp.getUi();
  const current = SCRIPT_PROPERTIES.getProperty('GEMINI_API_KEY');
  const status = current ? '（登録済み）' : '（未登録）';

  const result = ui.prompt(
    'Gemini APIキーの登録 ' + status,
    'Google AI Studio（https://aistudio.google.com/）で取得したAPIキーを貼り付けてください。',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const key = result.getResponseText().trim();
    if (key) {
      SCRIPT_PROPERTIES.setProperty('GEMINI_API_KEY', key);
      ui.alert('APIキーを登録しました。\n\n次に「語学学習AI」メニュー →「アプリを公開（デプロイ）」に進んでください。');
    } else {
      ui.alert('APIキーが空です。もう一度お試しください。');
    }
  }
}

/**
 * デプロイ手順をガイドするダイアログを表示する
 */
function showDeployGuide() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: sans-serif; font-size: 14px; line-height: 1.7; color: #333; padding: 8px; }
      h3 { margin-top: 0; color: #1a73e8; }
      ol { padding-left: 20px; }
      li { margin-bottom: 8px; }
      .highlight { background: #FEF3C7; padding: 2px 6px; border-radius: 4px; font-weight: bold; }
      .note { background: #f0f4ff; padding: 12px; border-radius: 8px; margin-top: 12px; font-size: 13px; }
    </style>
    <h3>アプリの公開（デプロイ）手順</h3>
    <ol>
      <li>このダイアログを閉じて、下の<span class="highlight">「エディタを開く」</span>ボタンをクリック</li>
      <li>Apps Scriptエディタが開いたら、右上の<span class="highlight">「デプロイ」</span>ボタンをクリック</li>
      <li><span class="highlight">「新しいデプロイ」</span>をクリック</li>
      <li>左上の<span class="highlight">歯車アイコン</span>をクリック →<span class="highlight">「ウェブアプリ」</span>を選ぶ</li>
      <li>設定を確認:
        <ul>
          <li>実行するユーザー → <span class="highlight">「自分」</span></li>
          <li>アクセスできるユーザー → <span class="highlight">「自分のみ」</span></li>
        </ul>
      </li>
      <li><span class="highlight">「デプロイ」</span>ボタンを押す</li>
      <li>「アクセスを承認」→ アカウントを選択 → 「詳細」→ 「〇〇（安全ではないページ）に移動」→「許可」</li>
      <li>表示されたURLがアプリのアドレスです。ブックマークしておきましょう！</li>
    </ol>
    <div style="text-align:center; margin-top: 16px;">
      <a href="EDITOR_URL_PLACEHOLDER" target="_blank"
         style="display:inline-block; background:#1a73e8; color:#fff; padding:10px 24px; border-radius:8px; text-decoration:none; font-weight:bold; font-size:14px;">
        エディタを開く
      </a>
    </div>
  `)
    .setWidth(520)
    .setHeight(440);

  // スクリプトIDからエディタURLを生成し、ダイアログ内のリンクに埋め込む
  const scriptId = ScriptApp.getScriptId();
  const editorUrl = 'https://script.google.com/d/' + scriptId + '/edit';
  const finalHtml = HtmlService.createHtmlOutput(
    html.getContent().replace('EDITOR_URL_PLACEHOLDER', editorUrl)
  )
    .setWidth(520)
    .setHeight(440);

  SpreadsheetApp.getUi().showModalDialog(finalHtml, 'アプリの公開手順');
}

/**
 * WebアプリのURLにアクセスがあった際（GETリクエスト時）に最初に呼ばれる関数
 * フロントエンドのUIとなる index.html を生成してブラウザに返します
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  // スプレッドシートのURLを動的に取得してHTMLテンプレートに渡す
  template.sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  return template.evaluate()
    .setTitle('語学学習サポートAI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * フロントエンドから非同期通信(google.script.run)経由で呼び出されるメイン処理関数
 * @param {string} message ユーザーが画面（テキストボックス）から入力したチャットテキスト
 * @param {string} targetLanguage ユーザーが選択した学習対象言語
 * @returns {Object} 画面に表示するAIのメッセージを含んだオブジェクト
 */
function processChat(message, targetLanguage = "自動判定", modelName = "gemini-2.5-flash-lite") {
  // 1. バリデーション: メッセージが空でないか確認
  if (!message || message.trim() === '') {
    throw new Error("メッセージが空です。");
  }

  try {
    // 2. Gemini APIの呼び出し：文章の生成と「構造化データ」の抽出を行う
    const aiResponse = callGeminiApi(message, targetLanguage, modelName);

    // AIからのレスポンス形式が正しいか（設定したJSONスキーマ通りか）確認
    if (!aiResponse || !aiResponse.reply_message || !aiResponse.extracted_data) {
      throw new Error("APIから正しい形式のレスポンスが返されませんでした。");
    }

    // 3. スプレッドシートへの保存：抽出された構造化データのみをシートの行に追記保存する
    saveToSheet(aiResponse.extracted_data);

    // 4. フロントエンドへの返却処理：UI表示用に抽出データも一緒に返す
    return {
      reply_message: aiResponse.reply_message,
      phrase: aiResponse.extracted_data.phrase || "",
      examples: aiResponse.extracted_data.examples || [],
      language: aiResponse.extracted_data.language || targetLanguage
    };
  } catch (error) {
    console.error(error);
    // エラー時はフロントエンド側（withFailureHandler）に向けてエラーメッセージを送る
    throw new Error(error.message || "エラーが発生しました。");
  }
}

/**
 * Gemini APIと実際に通信を行い、生成結果（レスポンス）を取得する関数
 * @param {string} userMessage ユーザーからの入力テキスト
 * @param {string} targetLanguage 対象とする言語
 * @returns {Object} APIからパースされたJSON形式のレスポンス（回答と抽出データ）
 */
function callGeminiApi(userMessage, targetLanguage, modelName = "gemini-2.5-flash-lite") {
  // スクリプトプロパティからGemini APIキーを取得
  const apiKey = SCRIPT_PROPERTIES.getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error("Gemini APIキーが設定されていません。スクリプトプロパティを確認してください。");
  }

  // 利用するAPIのエンドポイント（フロントから選択されたモデルを使用）
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

  // AIに与える指示（役割設定、出力フォーマットの制限など）
  const prompt = `
あなたは${targetLanguage === "自動判定" ? "語学" : targetLanguage}学習をサポートする優秀なAIアシスタントです。
ユーザーから${targetLanguage === "自動判定" ? "対象となる言語" : targetLanguage}の表現や意味についての質問が届きます。
ユーザーへの親切で自然な日本語の回答（reply_message）と、その表現から抽出した構造化データ（extracted_data）を生成してください。
返信は必ずJSONのみを出力してください。マークダウンの\`\`\`jsonなどは含めず、純粋なJSONオブジェクトを返してください。

ユーザーの入力:
「${userMessage}」
  `;

  // APIに送信するパラメータ群
  const payload = {
    "contents": [{
      "parts": [{"text": prompt}]
    }],
    "generationConfig": {
      // 構造化出力（JSON Schema）機能を利用し、必ず指定したキーを持ったJSONで返すようGeminiに強制する
      "response_mime_type": "application/json",
      "responseSchema": {
        "type": "OBJECT",
        "properties": {
          "reply_message": {
            "type": "STRING",
            "description": "ユーザーへの自然な日本語での回答テキスト（マークダウン形式推奨）"
          },
          "extracted_data": {
            "type": "OBJECT",
            "description": "スプレッドシートに保存するための構造化データ",
            "properties": {
              "phrase": { "type": "STRING", "description": "対象の言語表現" },
              "language": { "type": "STRING", "description": "対象の言語名（例: 英語、中国語、アラビア語など）" },
              "translation": { "type": "STRING", "description": "日本語訳" },
              "nuance_context": { "type": "STRING", "description": "ニュアンスや使用される文脈" },
              "examples": {
                "type": "ARRAY",
                "items": { "type": "STRING", "description": "対象言語の例文と日本語訳（例: 'Bonjour. - こんにちは。')" }
              },
              "tags": {
                "type": "ARRAY",
                "items": { "type": "STRING", "description": "カテゴリタグ" }
              }
            },
            "required": ["phrase", "language", "translation", "nuance_context", "examples", "tags"]
          }
        },
        "required": ["reply_message", "extracted_data"]
      }
    }
  };

  // HTTPリクエストのオプション設定
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true // エラー発生時もスクリプトを強制終了させず、自前でステータスコードを確認するための設定
  };

  // 実際にAPIへリクエストを送信し、レスポンスを受け取る
  const response = UrlFetchApp.fetch(endpoint, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  // 成功(HTTPステータス 200)以外の場合は直ちにエラーを投げる
  if (responseCode !== 200) {
    throw new Error(`Gemini API Error: [${responseCode}] ${responseBody}`);
  }

  // 正常な場合、まずAPIのラッパー全体のJSONをパースする
  const jsonResponse = JSON.parse(responseBody);
  // Geminiが生成したコンテンツ（我々が要求したJSON文字列）を抽出する
  const textContent = jsonResponse.candidates[0].content.parts[0].text;
  
  // 生成された文字列を再度オブジェクトに変換して return する
  return JSON.parse(textContent);
}

/**
 * 抽出された構造化データをGoogleスプレッドシートに追記（蓄積）保存する処理
 * @param {Object} extractedData APIから返されたシート保存用のデータオブジェクト
 */
function saveToSheet(extractedData) {
  // このGASプロジェクトに直接紐づいているスプレッドシートを開き、左から1つ目のシートを取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheets()[0];

  // シートが完全に空（1行目すらない）場合は、初回データ保存前に自動でヘッダー行を作成
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', 'Timestamp', 'Language', 'Phrase', 'Translation', 'Nuance/Context', 'Examples', 'Tags']);
  }

  // 個別のID(UUID)と、保存日時のためのタイムスタンプ(ISO文字列)を生成
  const id = generateUuid();
  const timestamp = new Date().toISOString();
  
  // 配列型のデータ（例文やタグ）は1つのセルに収めるためにJSON文字列に変換
  const examplesStr = JSON.stringify(extractedData.examples || []);
  const tagsStr = JSON.stringify(extractedData.tags || []);

  // シートの各列（A列〜H列）の並びに合わせた1行分のデータ配列を作成
  const rowData = [
    id,
    timestamp,
    extractedData.language || "",
    extractedData.phrase || "",
    extractedData.translation || "",
    extractedData.nuance_context || "",
    examplesStr,
    tagsStr
  ];

  // スプレッドシートの最終データ行の次の行に、配列のデータを一発で追記保存する
  sheet.appendRow(rowData);
}

/**
 * ランダムで一意なUUID（ユニークな文字と数字の羅列）を生成するユーティリティ関数
 * @returns {string} UUIDの文字列（例: "123e4567-e89b-12d3..."）
 */
function generateUuid() {
  return Utilities.getUuid();
}
