/**
 * @fileoverview Google Drive上の会議メモからGemini APIを利用して議事録を自動生成するスクリプト
 * クラスベース設計で責任分離とエラーハンドリングを統一
 */

//================================================================
// 設定とエラーハンドリング
//================================================================

/**
 * スクリプトの設定を管理するオブジェクト
 * @const
 */
const CONFIG = {
  // ★★★ 1. ここに取得したGemini APIキーを貼り付け ★★★
  API_KEY: 'ここに取得したGemini APIキーを貼り付け',

  // ★★★ 2. Meetのメモが保存されるフォルダのIDを貼り付け ★★★
  SOURCE_FOLDER_ID: 'Meetのメモが保存されるフォルダのIDを貼り付け',

  // ★★★ 3. 議事録の保存先フォルダのIDを貼り付け ★★★
  DESTINATION_FOLDER_ID: '議事録の保存先フォルダのIDを貼り付け',

  // ★★★ 4. 使用するGeminiモデルを選択 ★★★
  // 利用可能モデル（記事投稿時点）: 
  // - 'gemini-2.5-flash': 高速、低コスト（推奨）
  // - 'gemini-1.5-pro': 高性能、やや高コスト
  // - 'gemini-1.5-flash': バランス型
  GEMINI_MODEL: 'gemini-2.5-flash',

  // 処理済みファイルリストを保存するプロパティサービスのキー
  PROCESSED_FILES_KEY: 'processedFiles',

  // APIリトライ設定
  RETRY_COUNT: 3,
  RETRY_DELAY: 2000,

  // AIへの命令文（プロンプト）
  PROMPT: `あなたは議事録作成の専門家です。会議の文字起こしから、実用的で読みやすい議事録を作成してください。

# 議事録_YYYY-MM-DD_会議名

## 会議概要
- **会議名**: [会議のタイトル]
- **開催日時**: [日付と時間]
- **参加者**: [参加者リスト]
- **目的**: [会議の目的]

## 決定事項
- [決定された内容を箇条書きで記載]
- [実施時期や担当者も含める]

## 議論・検討事項
- [話し合われた内容]
- [検討が必要な項目]
- [保留事項があれば記載]

## 課題・問題点
- [課題内容と対応方針]
- [期限や担当者を明記]

## アクションアイテム
| 担当者 | 期限 | タスク内容 |
|--------|------|------------|
| [名前] | [日付] | [具体的なタスク] |

## 次回会議
- **日時**: [次回予定]
- **議題**: [予定議題]

---

以下の文字起こしデータから上記形式の議事録を作成してください：
<<transcript>>

【注意事項】
- 不明確な内容は「要確認」と記載
- 日付・数値・固有名詞は正確に抽出
- 決定事項と検討事項を明確に分離`
};

/**
 * 統一エラーハンドリング用のカスタムエラークラス
 */
class ProcessingError extends Error {
  constructor(message, type = 'UNKNOWN', originalError = null) {
    super(message);
    this.name = 'ProcessingError';
    this.type = type;
    this.originalError = originalError;
    this.timestamp = new Date().toISOString();
  }
}

/**
 * ログ管理クラス
 */
class Logger {
  static info(message, data = null) {
    const logMessage = data ? `${message} - データ: ${JSON.stringify(data)}` : message;
    console.log(`[INFO] ${new Date().toISOString()} - ${logMessage}`);
  }

  static warn(message, data = null) {
    const logMessage = data ? `${message} - データ: ${JSON.stringify(data)}` : message;
    console.warn(`[WARN] ${new Date().toISOString()} - ${logMessage}`);
  }

  static error(message, error = null) {
    const errorInfo = error ? ` - エラー詳細: ${error.message || error}` : '';
    console.error(`[ERROR] ${new Date().toISOString()} - ${message}${errorInfo}`);
  }
}

//================================================================
// 処理済みファイル管理クラス
//================================================================

/**
 * 処理済みファイルの状態管理を行うクラス
 */
class ProcessedFilesManager {
  constructor() {
    this.scriptProperties = PropertiesService.getScriptProperties();
  }

  /**
   * 処理済みファイルリストを取得
   * @return {Object} 処理済みファイルの情報
   */
  getProcessedFiles() {
    try {
      const jsonString = this.scriptProperties.getProperty(CONFIG.PROCESSED_FILES_KEY);
      return jsonString ? JSON.parse(jsonString) : {};
    } catch (error) {
      Logger.error('処理済みファイルリストの取得に失敗', error);
      return {};
    }
  }

  /**
   * ファイルが処理済みかチェック
   * @param {string} fileId ファイルID
   * @return {boolean} 処理済みの場合true
   */
  isProcessed(fileId) {
    const processedFiles = this.getProcessedFiles();
    return !!processedFiles[fileId];
  }

  /**
   * ファイルの処理状況を更新
   * @param {string} fileId ファイルID
   * @param {string} status 処理状況
   */
  updateStatus(fileId, status) {
    try {
      const processedFiles = this.getProcessedFiles();
      processedFiles[fileId] = `${status} - ${new Date().toISOString()}`;
      this.scriptProperties.setProperty(CONFIG.PROCESSED_FILES_KEY, JSON.stringify(processedFiles));
      Logger.info(`ファイル処理状況を更新`, { fileId, status });
    } catch (error) {
      Logger.error('処理状況の更新に失敗', error);
      throw new ProcessingError('処理状況の更新に失敗しました', 'STORAGE_ERROR', error);
    }
  }
}

//================================================================
// Google Docs テキスト抽出クラス
//================================================================

/**
 * Google Docsからテキストを抽出するクラス
 */
class GoogleDocsExtractor {
  /**
   * ドキュメント要素からテキストを再帰的に抽出
   * @param {Array<Object>} elements ドキュメント要素の配列
   * @return {string} 抽出されたテキスト
   */
  readTextFromElements(elements) {
    if (!elements || !Array.isArray(elements)) {
      return '';
    }

    let text = '';
    elements.forEach(structuralElement => {
      if (structuralElement.paragraph) {
        structuralElement.paragraph.elements.forEach(element => {
          if (element.textRun && element.textRun.content) {
            text += element.textRun.content;
          }
        });
      } else if (structuralElement.table) {
        structuralElement.table.tableRows.forEach(row => {
          row.tableCells.forEach(cell => {
            text += this.readTextFromElements(cell.content);
          });
        });
      }
    });
    return text;
  }

  /**
   * Google Docs APIを使用して「文字起こし」タブからテキストを抽出
   * @param {string} documentId ドキュメントID
   * @return {string|null} 抽出されたテキスト、見つからない場合はnull
   */
  extractTranscriptionText(documentId) {
    if (!documentId) {
      throw new ProcessingError('ドキュメントIDが指定されていません', 'INVALID_INPUT');
    }

    try {
      Logger.info('ドキュメントから文字起こしテキストを抽出開始', { documentId });
      
      const doc = Docs.Documents.get(documentId, { includeTabsContent: true });
      const transcriptionTabTitle = '文字起こし';

      if (!doc.tabs || !Array.isArray(doc.tabs)) {
        Logger.warn('ドキュメントにタブが見つかりません');
        return null;
      }

      const transcriptionTab = doc.tabs.find(tab => 
        tab.tabProperties && tab.tabProperties.title === transcriptionTabTitle
      );

      if (!transcriptionTab) {
        Logger.warn(`「${transcriptionTabTitle}」タブが見つかりません`);
        return null;
      }

      Logger.info(`「${transcriptionTabTitle}」タブを検出、内容を抽出中`);
      
      if (transcriptionTab.documentTab && 
          transcriptionTab.documentTab.body && 
          transcriptionTab.documentTab.body.content) {
        const extractedText = this.readTextFromElements(transcriptionTab.documentTab.body.content);
        Logger.info('テキスト抽出完了', { textLength: extractedText.length });
        return extractedText;
      } else {
        Logger.warn(`「${transcriptionTabTitle}」タブに内容が見つかりません`);
        return null;
      }

    } catch (error) {
      Logger.error('Google Docs APIからのテキスト抽出に失敗', error);
      throw new ProcessingError('テキスト抽出に失敗しました', 'DOCS_API_ERROR', error);
    }
  }
}

//================================================================
// Gemini API クライアントクラス
//================================================================

/**
 * Gemini APIとの通信を管理するクラス
 */
class GeminiAPIClient {
  constructor(apiKey, modelName) {
    if (!apiKey || apiKey === 'ここに取得したGemini APIキーを貼り付け') {
      throw new ProcessingError('Gemini APIキーが設定されていません', 'CONFIG_ERROR');
    }
    if (!modelName) {
      throw new ProcessingError('Geminiモデル名が設定されていません', 'CONFIG_ERROR');
    }
    this.apiKey = apiKey;
    this.modelName = modelName;
    this.baseUrl = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent`;
  }

  /**
   * プロンプトを準備する
   * @param {string} transcriptionText 文字起こしテキスト
   * @return {string} 完成したプロンプト
   */
  preparePrompt(transcriptionText) {
    const today = new Date();
    const todayString = `${today.getFullYear()}年${today.getMonth() + 1}月${today.getDate()}日`;
    
    const promptWithMemo = CONFIG.PROMPT.replace('<<transcript>>', transcriptionText);
    return `本日は${todayString}です。\n\n${promptWithMemo}`;
  }

  /**
   * APIリクエストのペイロードを作成
   * @param {string} prompt プロンプト
   * @return {Object} APIリクエストペイロード
   */
  createRequestPayload(prompt) {
    return {
      "contents": [{
        "parts": [{
          "text": prompt
        }]
      }],
      "generationConfig": {
        "temperature": 0.7,
        "topP": 0.8,
        "topK": 40,
        "maxOutputTokens": 8192
      }
    };
  }

  /**
   * APIレスポンスを処理する
   * @param {number} responseCode HTTPステータスコード
   * @param {string} responseBody レスポンスボディ
   * @return {string} 生成されたテキスト
   */
  processApiResponse(responseCode, responseBody) {
    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      const generatedText = jsonResponse?.candidates?.[0]?.content?.parts?.[0]?.text;
      
      if (!generatedText) {
        throw new ProcessingError('APIレスポンスの形式が不正です', 'API_RESPONSE_ERROR');
      }
      
      return generatedText;
    } else if (responseCode === 429) {
      throw new ProcessingError('レート制限に達しました', 'RATE_LIMIT_ERROR');
    } else if (responseCode === 403) {
      throw new ProcessingError('API認証エラー: APIキーを確認してください', 'AUTH_ERROR');
    } else if (responseCode >= 500) {
      throw new ProcessingError(`Geminiサーバーエラー: ${responseCode}`, 'SERVER_ERROR');
    } else {
      throw new ProcessingError(`Gemini APIエラー: ${responseCode}`, 'API_ERROR');
    }
  }

  /**
   * ファイル名を抽出する
   * @param {string} generatedText 生成されたテキスト
   * @param {string} originalFileName 元のファイル名
   * @return {string} 議事録ファイル名
   */
  extractFileName(generatedText, originalFileName) {
    try {
      // 入力値の検証
      if (!generatedText || typeof generatedText !== 'string') {
        Logger.warn('生成テキストが無効です', { generatedText: generatedText, type: typeof generatedText });
        return this.createFallbackFileName(originalFileName, 'invalid_text');
      }

      // 複数のパターンでファイル名を抽出を試行
      const patterns = [
        /^#\s+(議事録_\d{4}-\d{2}-\d{2}_.+?)$/m,  // 元のパターン
        /^#\s+(.+議事録.+)$/m,                     // より柔軟なパターン
        /^#\s+(.+)$/m                              // 最初の見出し
      ];

      for (const pattern of patterns) {
        const match = generatedText.match(pattern);
        if (match && match[1] && match[1].trim()) {
          const extractedName = match[1].trim();
          Logger.info('ファイル名を抽出', { extractedName, pattern: pattern.source });
          const fileName = extractedName.endsWith('.md') ? extractedName : `${extractedName}.md`;
          
          // 最終的な安全性チェック
          if (fileName && typeof fileName === 'string' && fileName.length > 0) {
            return fileName;
          }
        }
      }

      // すべてのパターンマッチングに失敗した場合
      Logger.warn('すべてのパターンマッチングに失敗', { 
        originalFileName, 
        textPreview: generatedText.substring(0, 200) 
      });
      
      return this.createFallbackFileName(originalFileName, 'pattern_failed');
      
    } catch (error) {
      Logger.error('ファイル名抽出処理でエラー', error);
      return this.createFallbackFileName(originalFileName, 'error');
    }
  }

  /**
   * フォールバックファイル名を作成
   * @param {string} originalFileName 元のファイル名
   * @param {string} reason 理由
   * @return {string} フォールバックファイル名
   */
  createFallbackFileName(originalFileName, reason) {
    const today = new Date().toISOString().split('T')[0];
    const safeName = originalFileName && typeof originalFileName === 'string' 
      ? originalFileName.replace(/\.gdoc$/, '').replace(/[^\w\-_]/g, '_').substring(0, 30)
      : 'meeting';
    
    const fallbackName = `議事録_${today}_${safeName}_${reason}.md`;
    Logger.warn('フォールバックファイル名を生成', { fallbackName, originalFileName, reason });
    
    return fallbackName;
  }

  /**
   * リトライ付きでAPIを呼び出す
   * @param {string} transcriptionText 文字起こしテキスト
   * @param {string} originalFileName 元のファイル名
   * @param {number} retryCount リトライ回数
   * @return {Object} {fileName, content}
   */
  generateMinutesWithRetry(transcriptionText, originalFileName, retryCount = 0) {
    try {
      Logger.info('Gemini API呼び出し開始', { retryCount, originalFileName });
      
      const prompt = this.preparePrompt(transcriptionText);
      const payload = this.createRequestPayload(prompt);
      
      const options = {
        'method': 'POST',
        'contentType': 'application/json',
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
      };

      const response = UrlFetchApp.fetch(`${this.baseUrl}?key=${this.apiKey}`, options);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();

      const generatedText = this.processApiResponse(responseCode, responseBody);
      
      // デバッグログ追加
      Logger.info('生成されたテキストの確認', { 
        textLength: generatedText ? generatedText.length : 0,
        textPreview: generatedText ? generatedText.substring(0, 100) : 'null',
        originalFileName 
      });
      
      const fileName = this.extractFileName(generatedText, originalFileName);
      
      // ファイル名の最終確認
      if (!fileName) {
        Logger.error('extractFileNameがnullを返しました', { generatedText: generatedText ? generatedText.substring(0, 200) : 'null', originalFileName });
        throw new ProcessingError('ファイル名の生成に失敗しました', 'FILENAME_GENERATION_ERROR');
      }

      Logger.info('Gemini API呼び出し成功', { fileName, contentLength: generatedText ? generatedText.length : 0 });
      return { fileName, content: generatedText };

    } catch (error) {
      if (retryCount < CONFIG.RETRY_COUNT && 
          (error.type === 'RATE_LIMIT_ERROR' || 
           error.message.includes('timeout') || 
           error.message.includes('network'))) {
        
        const delay = CONFIG.RETRY_DELAY * Math.pow(2, retryCount); // 指数バックオフ
        Logger.warn(`API呼び出し失敗、${delay}ms後にリトライ`, { retryCount, error: error.message });
        
        Utilities.sleep(delay);
        return this.generateMinutesWithRetry(transcriptionText, originalFileName, retryCount + 1);
      }

      Logger.error('Gemini API呼び出しが最終的に失敗', error);
      throw error;
    }
  }

  /**
   * 議事録を生成する（メインメソッド）
   * @param {string} transcriptionText 文字起こしテキスト
   * @param {string} originalFileName 元のファイル名
   * @return {Object|null} {fileName, content} または null
   */
  generateMinutes(transcriptionText, originalFileName) {
    try {
      return this.generateMinutesWithRetry(transcriptionText, originalFileName);
    } catch (error) {
      Logger.error('議事録生成に失敗', error);
      return null;
    }
  }
}

//================================================================
// ファイル管理クラス
//================================================================

/**
 * Google Driveのファイル操作を管理するクラス
 */
class FileManager {
  constructor(sourceFolderId, destinationFolderId) {
    if (!sourceFolderId || sourceFolderId === 'Meetのメモが保存されるフォルダのIDを貼り付け') {
      throw new ProcessingError('ソースフォルダIDが設定されていません', 'CONFIG_ERROR');
    }
    if (!destinationFolderId || destinationFolderId === '議事録の保存先フォルダのIDを貼り付け') {
      throw new ProcessingError('保存先フォルダIDが設定されていません', 'CONFIG_ERROR');
    }

    this.sourceFolderId = sourceFolderId;
    this.destinationFolderId = destinationFolderId;
  }

  /**
   * ソースフォルダからGoogle Docsファイルを取得
   * @return {GoogleAppsScript.Drive.FileIterator} ファイルイテレータ
   */
  getSourceFiles() {
    try {
      const sourceFolder = DriveApp.getFolderById(this.sourceFolderId);
      return sourceFolder.getFilesByType(MimeType.GOOGLE_DOCS);
    } catch (error) {
      Logger.error('ソースフォルダへのアクセスに失敗', error);
      throw new ProcessingError('ソースフォルダにアクセスできません', 'FOLDER_ACCESS_ERROR', error);
    }
  }

  /**
   * 議事録ファイルを作成
   * @param {string} fileName ファイル名
   * @param {string} content ファイル内容
   */
  createMinutesFile(fileName, content) {
    try {
      // 入力値の検証
      if (!fileName || typeof fileName !== 'string') {
        throw new ProcessingError(`無効なファイル名: ${fileName}`, 'INVALID_FILENAME');
      }
      if (!content || typeof content !== 'string') {
        throw new ProcessingError('ファイル内容が空です', 'EMPTY_CONTENT');
      }

      const destinationFolder = DriveApp.getFolderById(this.destinationFolderId);
      destinationFolder.createFile(fileName, content, MimeType.PLAIN_TEXT);
      Logger.info('議事録ファイルを作成', { fileName, contentLength: content.length });
    } catch (error) {
      Logger.error('議事録ファイルの作成に失敗', error);
      throw new ProcessingError('ファイル作成に失敗しました', 'FILE_CREATE_ERROR', error);
    }
  }
}

//================================================================
// メイン処理クラス
//================================================================

/**
 * 会議議事録処理のメインクラス
 */
class MeetingMinutesProcessor {
  constructor() {
    this.processedFilesManager = new ProcessedFilesManager();
    this.docsExtractor = new GoogleDocsExtractor();
    this.geminiClient = new GeminiAPIClient(CONFIG.API_KEY, CONFIG.GEMINI_MODEL);
    this.fileManager = new FileManager(CONFIG.SOURCE_FOLDER_ID, CONFIG.DESTINATION_FOLDER_ID);
  }

  /**
   * 単一ファイルを処理
   * @param {GoogleAppsScript.Drive.File} file 処理対象ファイル
   */
  processSingleFile(file) {
    const fileId = file.getId();
    const fileName = file.getName();
    
    Logger.info('新規ファイル処理開始', { fileName, fileId });

    try {
      // 文字起こしテキストを抽出
      const transcriptionText = this.docsExtractor.extractTranscriptionText(fileId);
      
      if (!transcriptionText) {
        Logger.info('文字起こしタブが見つからないためスキップ', { fileName });
        this.processedFilesManager.updateStatus(fileId, 'Skipped (No transcription tab)');
        return;
      }

      // Gemini APIで議事録生成
      const result = this.geminiClient.generateMinutes(transcriptionText, fileName);
      
      if (!result) {
        this.processedFilesManager.updateStatus(fileId, 'Failed (API Error)');
        return;
      }

      // 議事録ファイル作成
      this.fileManager.createMinutesFile(result.fileName, result.content);
      
      Logger.info('議事録作成完了', { originalFile: fileName, minutesFile: result.fileName });
      this.processedFilesManager.updateStatus(fileId, 'Success');

    } catch (error) {
      Logger.error('ファイル処理中にエラーが発生', error);
      this.processedFilesManager.updateStatus(fileId, `Error: ${error.message}`);
      
      // 重要なエラーの場合は再スロー
      if (error.type === 'CONFIG_ERROR') {
        throw error;
      }
    }
  }

  /**
   * 新しい会議メモを処理するメインメソッド
   */
  processNewMeetingMemos() {
    try {
      Logger.info('会議メモ処理開始');
      
      const files = this.fileManager.getSourceFiles();
      let processedCount = 0;
      let skippedCount = 0;

      while (files.hasNext()) {
        const file = files.next();
        const fileId = file.getId();
        
        if (this.processedFilesManager.isProcessed(fileId)) {
          skippedCount++;
          continue;
        }
        
        this.processSingleFile(file);
        processedCount++;
      }

      Logger.info('会議メモ処理完了', { processedCount, skippedCount });
      
    } catch (error) {
      Logger.error('会議メモ処理中に致命的なエラーが発生', error);
      throw error;
    }
  }
}

//================================================================
// エントリーポイント
//================================================================

/**
 * メイン実行関数（トリガーで呼び出される）
 */
function processNewMeetingMemos() {
  try {
    const processor = new MeetingMinutesProcessor();
    processor.processNewMeetingMemos();
  } catch (error) {
    Logger.error('処理が中断されました', error);
    
    // 設定エラーの場合はユーザーに分かりやすいメッセージを表示
    if (error.type === 'CONFIG_ERROR') {
      console.error('==== 設定エラー ====');
      console.error('CONFIG オブジェクトの設定を確認してください:');
      console.error('- API_KEY: Gemini APIキー');
      console.error('- SOURCE_FOLDER_ID: Meetメモフォルダのドライブ ID');
      console.error('- DESTINATION_FOLDER_ID: 議事録保存先フォルダのドライブ ID');
      console.error('- GEMINI_MODEL: 使用するGeminiモデル名');
    }
    
    throw error; // GASランタイムにエラーを報告
  }
}

/**
 * 手動テスト用関数
 */
function testSingleFile() {
  const processor = new MeetingMinutesProcessor();
  const files = processor.fileManager.getSourceFiles();
  
  if (files.hasNext()) {
    const file = files.next();
    Logger.info('テスト実行: 単一ファイル処理', { fileName: file.getName() });
    processor.processSingleFile(file);
  } else {
    Logger.warn('テスト対象ファイルが見つかりません');
  }
}
  