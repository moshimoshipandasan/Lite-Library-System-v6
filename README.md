# Lite Library System v6

軽量な図書館管理システム（Google Apps Script版）

## 概要

このプロジェクトは、Google Apps Script（GAS）を使用した図書館管理システムです。書籍の貸出・返却・登録などの基本的な図書館業務をWebアプリケーションとして提供します。

## 機能

- 📚 **図書貸出システム** - バーコードスキャンまたはID入力による書籍の貸出
- 📖 **図書返却システム** - 貸出中の書籍の返却処理
- 🔍 **貸出書籍検索システム** - 貸出中の書籍の検索
- 👤 **利用者別返却システム** - 利用者ごとの返却処理
- ➕ **書籍登録システム** - 新規書籍の登録

## 技術スタック

- **バックエンド**: Google Apps Script (V8ランタイム)
- **フロントエンド**: HTML5, CSS3, JavaScript
- **データベース**: Google Sheets
- **バーコード読み取り**: QuaggaJS
- **デプロイツール**: CLASP (Command Line Apps Script Projects)

## セットアップ

### 前提条件

- Node.js がインストールされていること
- Googleアカウントを持っていること
- Google Apps Scriptへのアクセス権限があること

### インストール手順

1. リポジトリをクローン
```bash
git clone [repository-url]
cd Lite-Library-System-v6
```

2. 依存関係をインストール
```bash
npm install
```

3. CLASPでログイン
```bash
npx clasp login
```

4. Google Apps Scriptにプッシュ
```bash
npx clasp push
```

5. Webアプリとしてデプロイ
```bash
npx clasp open
```
Apps Script エディタで「デプロイ」→「新しいデプロイ」を選択

## データ構造

### Google Sheetsの構成

#### 書籍DB
| 列 | フィールド名 | 説明 |
|---|------------|-----|
| A | 書籍ID | 書籍の一意識別子 |
| B | 書籍名 | 書籍のタイトル |

#### 利用者DB
| 列 | フィールド名 | 説明 |
|---|------------|-----|
| A | 利用者ID | 利用者の一意識別子 |
| B | 利用者名 | 利用者の氏名 |
| C | メールアドレス | 連絡先メールアドレス |

#### 貸出記録
| 列 | フィールド名 | 説明 |
|---|------------|-----|
| A | 書籍ID | 貸出書籍のID |
| B | 書籍名 | 貸出書籍のタイトル |
| C | 利用者ID | 借りた利用者のID |
| D | 利用者名 | 借りた利用者の氏名 |
| E | 貸出日時 | 貸出した日時 |
| F | 返却予定日 | 返却予定日 |
| G | 返却状況 | "未返却" または "返却済み" |
| H | 返却日時 | 実際の返却日時 |

## 使用方法

### アクセス方法

デプロイ後、発行されたWebアプリケーションURLにアクセスします。

### ページ遷移

- デフォルト: メニュー画面（各機能への入口）
- `?page=checkout`: 図書貸出システム
- `?page=return`: 図書返却システム
- `?page=finder`: 貸出書籍検索システム
- `?page=user_returns`: 利用者別返却システム
- `?page=register`: 書籍登録システム

## 開発

### コマンド

```bash
# コードをGoogle Apps Scriptにアップロード
npx clasp push

# コードをローカルにダウンロード
npx clasp pull

# Apps Scriptエディタを開く
npx clasp open

# ログを確認
npx clasp logs
```

### ファイル構成

```
├── コード.js           # バックエンドロジック
├── menu.html          # メニューページ
├── lending.html       # 貸出ページ
├── returning.html     # 返却ページ
├── rental_books_finder.html  # 検索ページ
├── user_returns.html  # 利用者別返却ページ
├── book_register.html # 書籍登録ページ
├── appsscript.json   # GAS設定ファイル
├── .clasp.json       # CLASPプロジェクト設定
└── package.json      # Node.js設定
```

## 注意事項

- このシステムは誰でもアクセス可能な設定（`ANYONE_ANONYMOUS`）になっています
- 本番環境で使用する場合は、適切なアクセス制限を設定してください
- バーコードスキャン機能を使用する場合は、HTTPSでホストされている必要があります

## ライセンス

[ライセンス情報を追加してください]

## 作者

[作者情報を追加してください]