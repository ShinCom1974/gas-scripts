# gas-scripts

Google Apps Script (GAS) のスクリプト集。Google スプレッドシート・ドキュメント・フォームなどの自動化スクリプトを管理するプロジェクト。

## プロジェクト構成

```
gas-scripts/
├── CLAUDE.md
├── src/           # 各スクリプトのソースコード
└── README.md      # (必要に応じて作成)
```

## 開発環境

- **言語**: JavaScript (GAS 環境)
- **ランタイム**: Google Apps Script (V8)
- **デプロイ**: Google スプレッドシート / Google ドライブ上のスクリプトエディタ
- **CLIツール**: [clasp](https://github.com/google/clasp) (任意)

## コーディング規約

- 変数・関数名は camelCase を使用する
- GAS 固有のグローバルオブジェクト (`SpreadsheetApp`, `DriveApp` など) は直接使用する
- スクリプトは機能単位でファイルを分割する

## Git 運用ルール

**コードを変更するたびに、必ず GitHub にプッシュすること。**

### 手順

1. **変更をステージングしてコミット**
   ```bash
   git add <変更ファイル>
   git commit -m "変更内容の要約"
   ```

2. **GitHub にプッシュ**
   ```bash
   git push origin main
   ```

### コミットメッセージの形式

```
<種別>: <要約>

# 種別の例:
# feat: 新機能追加
# fix: バグ修正
# refactor: リファクタリング
# docs: ドキュメント更新
# chore: その他の変更
```

### ルール

- 1 つの変更ごとに 1 コミットを作成する（まとめすぎない）
- コミット後は必ず `git push` を実行する。省略しない
- `main` ブランチへの直接プッシュを基本とする（小規模プロジェクトのため）
- force push は禁止。履歴は必ず保持する

## 初回セットアップ（未実施の場合）

```bash
git init
git remote add origin https://github.com/ShinCom1974/gas-scripts.git
git add .
git commit -m "chore: initial commit"
git push -u origin main
```
