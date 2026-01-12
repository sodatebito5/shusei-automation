# SYSTEM_CONFIG - システム構成管理

> **最終更新**: 2026-01-12
> **現在のバージョン**: v3.0（配席API統合済）

---

## 1. プロジェクト概要

**守成クラブ福岡飯塚 例会運営自動化システム**

- **ミッション**: 手作業ゼロ、属人化ゼロ → 誰でも1時間で引き継げる状態
- **設計思想**: 判断は人、処理はシステム

---

## 2. GASプロジェクト一覧

| プロジェクト名 | scriptId | デプロイID | 用途 | 状態 |
|---------------|----------|------------|------|------|
| 守成ダッシュボード_v3 | `1shW-t-_5b9j22faoQB7qCnqZ6w6t_WrFT-hZPULT_xVCPnM7VrgnHStG` | `AKfycbzgVIzNjfW_UHihZS5bwrWM7xix0U4dnodZSlq7nPC8eGXGu_Fj6haCzivxiARVDPGL` | 本番ダッシュボード・配席API | ✅ 使用中 |
| ~~管理ダッシュボード v2~~ | `1nQ2lLHhR7HuDD6ygiua_kmE_GF_1zYWoPiy4r2XWjVuGXM1XVtkNAxLC` | `AKfycbwBT...` | 旧バージョン | 🗑️ 削除済 |

### GASエディタリンク

- **本番**: https://script.google.com/d/1shW-t-_5b9j22faoQB7qCnqZ6w6t_WrFT-hZPULT_xVCPnM7VrgnHStG/edit

---

## 3. デプロイ先一覧

| アプリ | URL | ホスティング | 用途 |
|--------|-----|--------------|------|
| 配席アプリ | https://seat-maker2.pages.dev | Cloudflare Pages | 座席自動配置 |
| ダッシュボード | https://script.google.com/macros/s/AKfycbzgVIzNjfW_UHihZS5bwrWM7xix0U4dnodZSlq7nPC8eGXGu_Fj6haCzivxiARVDPGL/exec | GAS WebApp | 管理画面 |

---

## 4. スプレッドシート一覧

| シート名 | ID | 主な用途 |
|----------|-----|----------|
| 出欠確認シート | `1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic` | 出欠状況、ゲスト出欠、会員名簿、配席アーカイブ |
| 設定シート | `1R4GR1GZg6mJP9zPX5MTE0IsYEAdVLNIM314o7vBqrg8` | ダッシュボード設定、イベント設定 |
| 売上報告シート | `1QYLHr7wMj0jQW5ApgWf-l1Bk9_9M5EV54FDHQp2Rkic` | 売上報告データ |

### シート構成詳細

**出欠確認シート** (`1IPyjDi3...`)
- 出欠状況（自動）
- ゲスト出欠状況（自動）
- 会員名簿マスター
- 受付名簿
- 配席アーカイブ
- 紹介記録

**設定シート** (`1R4GR1GZ...`)
- 設定
- ダッシュボード集計

---

## 5. ローカルフォルダ構成

```
C:\Users\Owner\
├── shusei-automation/          # 配席アプリ、LIFFアプリ、ドキュメント
│   ├── web/
│   │   ├── seat-maker/         # 配席アプリ（Cloudflare Pages）
│   │   ├── attendance/         # 出欠LIFF
│   │   ├── guest-form/         # ゲスト申請LIFF
│   │   └── dashboard/          # ダッシュボード画面（未使用）
│   ├── gas/                    # 旧GASスクリプト（deprecated）
│   │   ├── seat-maker/         # 配席API（ダッシュボードGASに統合済）
│   │   ├── attendance/
│   │   ├── sales-report/
│   │   └── schedule/
│   └── docs/                   # ドキュメント
│
└── shusei-dashboard-v3/        # ダッシュボードGAS（clasp管理）
    ├── .clasp.json             # clasp設定（scriptId指定）
    └── src/
        ├── dashboard.js        # メインロジック
        ├── MemberMonthly.js    # 月次処理
        ├── index.html          # 画面
        ├── style.html          # スタイル
        └── appsscript.json     # GAS設定
```

---

## 6. 重要なコマンド

### ダッシュボードGAS（shusei-dashboard-v3）

```bash
# ディレクトリ移動
cd C:/Users/Owner/shusei-dashboard-v3

# コードをGASにプッシュ
clasp push

# 強制プッシュ（確認なし）
clasp push -f

# 既存デプロイを更新（URLを変えずにバージョン更新）
clasp deploy -i AKfycbzgVIzNjfW_UHihZS5bwrWM7xix0U4dnodZSlq7nPC8eGXGu_Fj6haCzivxiARVDPGL -d "説明文"

# デプロイ一覧を確認
clasp deployments

# GASエディタを開く
clasp open

# ログ確認
clasp logs
```

### 配席アプリ（Cloudflare Pages）

```bash
# ディレクトリ: shusei-automation

# 本番デプロイ
wrangler pages deploy web/seat-maker --project-name=seat-maker2

# コミット済みのみでデプロイ（警告なし）
wrangler pages deploy web/seat-maker --project-name=seat-maker2 --commit-dirty=true
```

### Git

```bash
# 変更確認
git status

# コミット
git add -A && git commit -m "メッセージ"

# プッシュ
git push
```

---

## 7. API エンドポイント

**ベースURL**: `https://script.google.com/macros/s/AKfycbzgVIzNjfW_UHihZS5bwrWM7xix0U4dnodZSlq7nPC8eGXGu_Fj6haCzivxiARVDPGL/exec`

| action | メソッド | 用途 |
|--------|----------|------|
| `getSeatingParticipants` | GET | 配席対象者を取得 |
| `syncSeats` | POST | 座席確定・受付名簿反映・アーカイブ保存 |
| `listSeatingArchives` | GET | 過去の配席一覧を取得 |
| `getSeatingArchive` | GET | 特定イベントの配席を取得 |

---

## 8. 変更履歴

| 日付 | バージョン | 変更内容 |
|------|------------|----------|
| 2026-01-12 | v3.0 | 配席API統合、ゲスト表示修正、ID重複バグ修正、GASプロジェクト統合 |
| 2026-01-12 | - | SYSTEM_CONFIG.md 作成 |
| 2026-01-12 | - | 管理ダッシュボード v2 削除 |

---

## 9. 注意事項

### デプロイ時

- **GASデプロイ**: `clasp deploy -i <デプロイID>` を使うこと（URLが変わらない）
- **新規デプロイ禁止**: `clasp deploy`（ID指定なし）は新しいURLが発行されるので使わない
- **アクセス権限**: Webアプリは「全員」アクセスに設定すること（CORS対策）

### スプレッドシート

- シート名を変更する場合は全コード検索が必要（定数で管理されている）
- 列の追加・削除時は `dashboard.js` の定数も更新

### タイムゾーン

- すべて `Asia/Tokyo` で統一
- `appsscript.json` の `timeZone` を確認

---

## 10. 関連ドキュメント

| ファイル | 内容 |
|----------|------|
| `CLAUDE.md` | Claude Code向け作業ルール |
| `PROGRESS.md` | 開発進捗サマリー |
| `docs/tasks.md` | タスク管理 |
| `docs/requirements.md` | 要件定義 |
| `docs/errors.md` | エラー解決履歴 |
| `docs/seat-maker-redesign.md` | 配席アプリ再設計書 |
