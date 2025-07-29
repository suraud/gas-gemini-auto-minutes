# Gas-Gemini-Auto-Minutes

Google Meet の会議メモから Gemini AI を利用して議事録を自動生成する Google Apps Script システムです。

## 概要

このシステムは、Google Meet で自動生成された文字起こしテキストから、構造化された読みやすい議事録を自動作成します。Gemini AI の自然言語処理能力を活用し、会議の内容を整理・要約して実用的な議事録を生成します。

## 主な機能

- 📝 **自動議事録生成**: Google Docs の「文字起こし」タブから議事録を自動作成
- 🔄 **バッチ処理**: 複数のファイルを一括処理
- 📋 **構造化出力**: 決定事項、アクションアイテム、課題などを整理
- 🚫 **重複処理防止**: 処理済みファイルを自動管理
- 🔧 **エラーハンドリング**: 包括的なエラー処理とリトライ機能
- 📊 **ログ機能**: 詳細な処理ログとデバッグ情報

## システム構成

### クラス構造

- **`MeetingMinutesProcessor`**: メイン処理クラス
- **`GoogleDocsExtractor`**: Google Docs からテキスト抽出
- **`GeminiAPIClient`**: Gemini API との通信管理
- **`FileManager`**: Google Drive ファイル操作
- **`ProcessedFilesManager`**: 処理済みファイル状態管理
- **`Logger`**: ログ出力管理

### 生成される議事録の形式

```markdown
# 議事録_YYYY-MM-DD_会議名

## 会議概要
- **会議名**: [会議のタイトル]
- **開催日時**: [日付と時間]
- **参加者**: [参加者リスト]
- **目的**: [会議の目的]

## 決定事項
- [決定された内容を箇条書きで記載]

## 議論・検討事項
- [話し合われた内容]

## 課題・問題点
- [課題内容と対応方針]

## アクションアイテム
| 担当者 | 期限 | タスク内容 |
|--------|------|------------|
| [名前] | [日付] | [具体的なタスク] |

## 次回会議
- **日時**: [次回予定]
- **議題**: [予定議題]
```

## セットアップ手順

### 1. 前提条件

- Google アカウント
- Google Apps Script へのアクセス権限
- Gemini API キー

### 2. Gemini API キーの取得

1. [Google AI Studio](https://aistudio.google.com/) にアクセス
2. API キーを作成
3. キーをコピーして保存

### 3. Google Drive フォルダの準備

1. **ソースフォルダ**: Google Meet のメモが保存されるフォルダ
2. **保存先フォルダ**: 生成された議事録を保存するフォルダ

各フォルダの ID を取得してください（URL の最後の部分）。

### 4. Google Apps Script プロジェクトの設定

1. [Google Apps Script](https://script.google.com/) で新しいプロジェクトを作成
2. `Code.gs` と `appsscript.json` をプロジェクトにコピー
3. `Code.gs` の `CONFIG` オブジェクトを編集：

```javascript
const CONFIG = {
  // ★★★ 1. 取得した Gemini API キーを設定 ★★★
  API_KEY: 'your-gemini-api-key-here',

  // ★★★ 2. Meet メモフォルダの ID を設定 ★★★
  SOURCE_FOLDER_ID: 'your-source-folder-id',

  // ★★★ 3. 議事録保存先フォルダの ID を設定 ★★★
  DESTINATION_FOLDER_ID: 'your-destination-folder-id',

  // ★★★ 4. 使用する Gemini モデルを選択 ★★★
  GEMINI_MODEL: 'gemini-2.5-flash', // 推奨: 高速・低コスト
};
```

### 5. Google Docs API の有効化

1. Apps Script エディターで「サービス」→「Google Docs API」を追加
2. バージョン v1 を選択

### 6. 権限の承認

初回実行時に以下の権限を承認してください：
- Google Drive へのアクセス
- Google Documents へのアクセス
- 外部 API（Gemini）への接続

## 使用方法

### 手動実行

```javascript
// 全ての新規ファイルを処理
processNewMeetingMemos();

// 単一ファイルのテスト処理
testSingleFile();
```

### 自動実行（トリガー設定）

1. Apps Script エディターで「トリガー」を選択
2. 新しいトリガーを作成：
   - 実行する関数: `processNewMeetingMemos`
   - イベントのソース: 時間主導型
   - 時間ベースのトリガーのタイプ: 定期的（例：1時間ごと）

## 利用可能な Gemini モデル

| モデル名 | 特徴 | 推奨用途 |
|----------|------|----------|
| `gemini-2.5-flash` | 高速・低コスト | 一般的な議事録作成（推奨） |
| `gemini-1.5-pro` | 高性能・やや高コスト | 複雑な会議内容 |
| `gemini-1.5-flash` | バランス型 | 中程度の複雑さ |

## エラー対応

### よくあるエラーと対処法

1. **設定エラー**: `CONFIG` オブジェクトの値を確認
2. **API 認証エラー**: Gemini API キーを確認
3. **フォルダアクセスエラー**: フォルダ ID と権限を確認
4. **レート制限エラー**: 自動リトライされますが、API 使用量を確認

### ログの確認

Apps Script エディターの「実行」→「実行トランスクリプト」でログを確認できます。

## 注意事項

- **API 使用量**: Gemini API の使用量に注意してください
- **プライバシー**: 機密情報を含む会議メモの取り扱いに注意
- **文字起こし品質**: Google Meet の文字起こし精度に依存します
- **処理時間**: 大きなファイルは処理に時間がかかる場合があります

## システム要件

- Google Apps Script 環境
- Google Docs API v1
- Gemini API アクセス権限
- Google Drive の適切なアクセス権限

## トラブルシューティング

### 議事録が生成されない場合

1. ソースファイルに「文字起こし」タブが存在するか確認
2. Gemini API キーが正しく設定されているか確認
3. フォルダ ID が正しく設定されているか確認
4. 実行ログでエラーメッセージを確認

### API エラーが発生する場合

1. API キーの有効性を確認
2. API 使用量制限を確認
3. ネットワーク接続を確認

## ライセンス

このプロジェクトは個人・商用利用可能です。Gemini API の利用規約に従ってご使用ください。

## 貢献

バグレポートや機能追加の提案は Issue または Pull Request でお知らせください。

---

**開発者向け情報**: このシステムはクラスベース設計を採用し、各機能が独立してテスト・拡張可能な構造になっています。 