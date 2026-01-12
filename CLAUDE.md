# SHUSEI-AUTOMATION Project Rules

> **Claude Code向け設定ファイル**
> コード生成・修正前に必ずこのファイルと `docs/` を参照すること。

---

## 📌 プロジェクト概要

守成クラブ福岡飯塚の月例会運営ダッシュボードシステム。

**ミッション**: 手作業ゼロ、属人化ゼロ → 誰でも1時間で引き継げる状態

**設計思想**: 判断は人、処理はシステム

### 機能一覧

| 機能 | 概要 | LIFF |
|------|------|------|
| 出欠連絡 | 月例会の出欠を登録 | ✅ |
| 売上報告 | メンバー間売上を報告 | ✅ |
| ゲスト申請 | ゲスト参加申請 | ✅ |
| 自動配席 | 座席表を自動生成 | - |
| ダッシュボード | 統合管理画面 | - |
| スケジュール | 会場・日程管理 | - |

---

## 🤖 Claude Code 指示

### 必読ドキュメント（コード書く前に確認）

```
docs/requirements.md  # 要件定義
docs/tasks.md         # タスク管理（何をやるか）
docs/errors.md        # 過去のエラーと解決策
PROGRESS.md           # 全体進捗
```

### 作業ルール

1. **タスク着手前**: `docs/tasks.md` で対象タスクを確認
2. **タスク完了時**: `docs/tasks.md` のチェックボックスを `[x]` に更新
3. **機能完了時**: `PROGRESS.md` の進捗率を更新
4. **エラー解決時**: `docs/errors.md` に追記
5. **大きな変更時**: 該当するドキュメントも更新

### コンテキスト管理

- 長い作業の前に `/compact` でコンテキスト整理を検討
- 複雑なタスクはサブエージェント活用を検討
- 不明点は推測せず確認を求める

### 禁止事項

- ドキュメント未確認でのコード生成
- 既存コードの意図を確認せずに大規模リファクタ
- マジックナンバーの直書き
- `docs/errors.md` に載ってるエラーパターンの再発

---

## 🏗 ディレクトリ構造

```
SHUSEI-AUTOMATION/
├── gas/                 # Google Apps Script（バックエンド）
│   ├── attendance/      # 出欠管理
│   ├── dashboard/       # ダッシュボード本体
│   ├── guest-form/      # ゲスト申請フォーム
│   ├── sales-report/    # 売上報告
│   ├── schedule/        # スケジュール管理
│   └── seat-maker/      # 自動座席配置
│
├── web/                 # フロントエンド（LIFF / Web）
│   ├── attendance/      # 出欠LIFF
│   ├── dashboard/       # ダッシュボード画面
│   ├── guest-form/      # ゲスト申請LIFF
│   ├── sales-report/    # 売上報告LIFF
│   └── seat-maker/      # 座席配置画面
│
├── docs/                # ドキュメント
│   ├── requirements.md  # 要件定義
│   ├── tasks.md         # タスク管理
│   └── errors.md        # エラー解決履歴
│
├── CLAUDE.md            # このファイル
├── PROGRESS.md          # 進捗サマリー
└── README.md
```

---

## 🔧 技術スタック

| レイヤー | 技術 |
|----------|------|
| フロント | LIFF SDK + HTML + Vanilla JS + CSS |
| バックエンド | Google Apps Script |
| データ | Google Sheets |
| デプロイ（GAS） | clasp |
| デプロイ（Web） | GitHub Pages / Vercel 等 |
| 出力 | PDF生成（式次第など） |

---

## 📏 コーディングルール

### GAS（.gs）

```javascript
// ✅ シート名・設定は冒頭で定数化
const SHEET_NAME = {
  ATTENDANCE: '出欠管理',
  MEMBERS: 'メンバー一覧',
  SALES: '売上報告'
};

// ✅ 関数名: camelCase
function getAttendanceData() { ... }

// ✅ Spreadsheet は引数で渡す or 関数冒頭で1回取得
function processData(sheet) { ... }
```

- マジックナンバー禁止（列番号も定数化）
- 日付は JST 明示: `Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd')`
- `try-catch` でエラーハンドリング必須（API系）

### フロントエンド（web/）

```javascript
// ✅ LIFF初期化パターン
liff.init({ liffId: 'xxxx-xxxx' })
  .then(() => { ... })
  .catch((err) => { console.error(err); });
```

- 外部ライブラリ最小限
- ES6+ OK（モダンブラウザ前提）
- CSS は シンプルに（BEM不要）

### 共通

- コメントは「なぜ」を書く（「何」はコードで分かる）
- 1関数1責務
- 命名で意図を伝える

---

## 🔄 開発フロー（Claude Code用）

### 新機能追加

```
1. docs/requirements.md で仕様確認
2. docs/tasks.md にタスク追加（なければ）
3. 実装
4. docs/tasks.md のチェックボックス更新
5. PROGRESS.md の進捗率更新
```

### バグ修正

```
1. docs/errors.md で類似エラー検索
2. 修正
3. docs/errors.md に解決方法を追記
4. docs/tasks.md にバグ修正を記録
```

### ファイル編集時の注意

```bash
# GAS編集時
gas/[機能名]/*.gs  # バックエンド
↓ 連携
web/[機能名]/*     # フロントエンド（API呼び出し側）

# 両方の整合性を確認すること
```

### デプロイ前チェック

```
1. 該当シートでテスト実行
2. LIFFの場合はLINEアプリ内で動作確認
3. エラーがあれば docs/errors.md に記録
```

---

## ⚠️ 注意事項

### Google Sheets
- シート名変更は影響範囲大 → 必ず全コード検索
- 列追加・削除時は定数ファイルも更新

### LIFF
- `liff.getProfile()` は init 完了後に呼ぶ
- LINE内ブラウザと外部ブラウザで挙動が違う場合あり

### PDF生成
- レイアウト崩れやすい → 既存フォーマットに厳密に合わせる
- フォント・余白の微調整は手動確認必須

### 日付処理
- JST前提で統一
- 月跨ぎ・年跨ぎのエッジケース注意

---

## 🎯 優先順位

1. **動くこと** （MVP最優先）
2. **引き継ぎやすさ** （1時間で理解できる）
3. **パフォーマンス・コード美** （後から磨く）

---

## 📎 関連リンク

- Google Sheets: （URLを記載）
- LIFF管理: https://developers.line.biz/console/
- clasp: https://github.com/google/clasp
