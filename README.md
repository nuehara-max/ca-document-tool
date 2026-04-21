# CA書類自動生成ツール

人材紹介会社のCA（キャリアアドバイザー）が、求職者情報から
**社内共有サマリー／職務経歴書本文／クライアント提案メール**
を Anthropic Claude で一括生成し、Google Drive に Google ドキュメントとして保存するためのWebアプリです。

## 機能

1. **入力フォーム**（氏名・年齢・学歴・職歴・スキル・転職理由・希望条件・面談メモ・提案先業界チップ）
2. **ファイルアップロード**（PDF / .docx をサーバー側で解析し、不足情報の補完に利用）
3. **AI書類生成**（Claude `claude-sonnet-4-20250514` で3種類の書類をJSONで一括生成）
4. **Google Drive連携**（OAuth 2.0／Google Docs として保存・共有リンクを表示）

## セットアップ

### 1. 依存インストール
```bash
cd "CA書類自動生成ツール"
npm install
```

### 2. 環境変数の設定
`.env.example` をコピーして `.env` を作成し、以下を記入します。

```bash
cp .env.example .env
```

| 変数 | 取得方法 |
| --- | --- |
| `ANTHROPIC_API_KEY` | https://console.anthropic.com/ で発行 |
| `GOOGLE_CLIENT_ID` / `GOOGLE_CLIENT_SECRET` | https://console.cloud.google.com/apis/credentials で OAuth 2.0 クライアント（種類: **ウェブアプリケーション**）を作成 |
| `GOOGLE_REDIRECT_URI` | `http://localhost:3000/api/auth/google/callback`（上記クライアントの「承認済みのリダイレクトURI」に同じ値を登録） |
| `GOOGLE_DRIVE_FOLDER_ID` | （任意）保存先フォルダのURL末尾のID。未指定ならマイドライブ直下 |
| `SESSION_SECRET` | 任意のランダム文字列 |

Google Cloud Console 側で **Google Drive API** を有効化しておく必要があります。

### 3. 起動
```bash
npm start
```

ブラウザで http://localhost:3000 にアクセス。
社内の複数スタッフが同じPCのブラウザから利用するか、同一LANなら
`http://<起動PCのIP>:3000` でも接続できます。

## 使い方

1. （任意）履歴書PDF / Word をドロップ → 自動でテキスト抽出
2. フォームに追加情報を入力し、提案先業界チップを選択
3. 「✨ AIで3書類をまとめて生成」をクリック
4. タブで3つの書類を確認・編集・コピー
5. ヘッダー右上の「Googleに接続」で OAuth 認証後、「💾 Google Driveに3件まとめて保存」

## ディレクトリ構成
```
CA書類自動生成ツール/
├── server.js          # Express バックエンド
├── public/index.html  # フロントエンド（単一HTML）
├── package.json
├── .env.example
└── README.md
```

## 注意事項
- 個人情報を扱うため、本番利用時は HTTPS 化・アクセス制限を必ず行ってください
- アップロードしたファイルはメモリ上で処理し、ディスクには保存しません
- Claude への入力には本文テキストが含まれるため、社内のデータ取扱ルールに従ってください
