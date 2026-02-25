# 守成クラブ福岡飯塚 システム全体概要

> **最終更新**: 2026-02-01
> **管理者向け**: システム全体を把握するためのドキュメント

---

## 1. システム概要

**守成クラブ福岡飯塚 例会運営自動化システム**

| 項目 | 内容 |
|------|------|
| ミッション | 手作業ゼロ、属人化ゼロ → 誰でも1時間で引き継げる状態 |
| 設計思想 | 判断は人、処理はシステム |
| 技術スタック | LIFF + GAS + Google Sheets + Cloudflare Pages |

---

## 2. アプリ一覧・デプロイ情報

### Webアプリ（Cloudflare Pages）

| アプリ名 | 本番URL | Cloudflare Project | 用途 |
|----------|---------|-------------------|------|
| 出欠確認LIFF | https://shukketsu1.pages.dev | `shukketsu1` | 月例会の出欠登録 |
| ゲスト申請LIFF | https://guest-form1.pages.dev | `guest-form1` | ゲスト参加申請 |
| 売上報告LIFF | https://uriage5.pages.dev | `uriage5` | メンバー間売上報告 |
| 自動配席アプリ | https://shusei-seat-manager.pages.dev | `shusei-seat-manager` | 座席表自動生成 |

### GAS WebApp

| アプリ名 | 本番URL | 用途 |
|----------|---------|------|
| ダッシュボード | https://script.google.com/macros/s/AKfycbzgVIzNjfW_UHihZS5bwrWM7xix0U4dnodZSlq7nPC8eGXGu_Fj6haCzivxiARVDPGL/exec | 統合管理画面 |
| ダッシュボード（テスト） | https://script.google.com/macros/s/AKfycbz5j-qZV2RW5nU2PoYEQYUQKooGSWtboHMPOgjSIQI/dev | テスト環境 |

---

## 3. GASプロジェクト一覧

| プロジェクト名 | scriptId | 用途 | ローカルパス |
|---------------|----------|------|-------------|
| 守成ダッシュボード_v3 | `1shW-t-_5b9j22faoQB7qCnqZ6w6t_WrFT-hZPULT_xVCPnM7VrgnHStG` | ダッシュボード本体・配席API | `gas/dashboard/` |
| 出欠管理GAS | `1iIXIXSjIeQmboirKGj2GsIY0vfpGoYxYpFTmJsXavlYN-q_fOqfWArm_` | 出欠LIFF用バックエンド | `gas/attendance/` |
| ゲスト申請GAS | `1Vmoiqom-KxSpxA29_E4PmV0ay-pcy4yYj_eAIHSGaxMcw7LgxJjVHWQQ` | ゲスト申請LIFF用バックエンド | `gas/guest-form/` |
| 売上報告GAS | `1sHyTgDX-cFBoWzOC4UP6AdhJrb2SoUkQCheWeEhIzr6lw7orOvTA7vuj` | 売上報告LIFF用バックエンド | `gas/sales-report/` |

### GASエディタ直リンク

| プロジェクト | URL |
|-------------|-----|
| ダッシュボード | https://script.google.com/d/1shW-t-_5b9j22faoQB7qCnqZ6w6t_WrFT-hZPULT_xVCPnM7VrgnHStG/edit |
| 出欠管理 | https://script.google.com/d/1iIXIXSjIeQmboirKGj2GsIY0vfpGoYxYpFTmJsXavlYN-q_fOqfWArm_/edit |
| ゲスト申請 | https://script.google.com/d/1Vmoiqom-KxSpxA29_E4PmV0ay-pcy4yYj_eAIHSGaxMcw7LgxJjVHWQQ/edit |
| 売上報告 | https://script.google.com/d/1sHyTgDX-cFBoWzOC4UP6AdhJrb2SoUkQCheWeEhIzr6lw7orOvTA7vuj/edit |

---

## 4. スプレッドシート一覧

| シート名 | ID | 主な用途 |
|----------|-----|----------|
| 出欠確認シート | `1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic` | 出欠状況、ゲスト出欠、会員名簿、配席アーカイブ |
| 設定シート | `1R4GR1GZg6mJP9zPX5MTE0IsYEAdVLNIM314o7vBqrg8` | ダッシュボード設定、イベント設定 |
| 売上報告シート | `1QYLHr7wMj0jQW5ApgWf-l1Bk9_9M5EV54FDHQp2Rkic` | 売上報告データ |

### スプレッドシート直リンク

| シート | URL |
|--------|-----|
| 出欠確認シート | https://docs.google.com/spreadsheets/d/1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic/edit |
| 設定シート | https://docs.google.com/spreadsheets/d/1R4GR1GZg6mJP9zPX5MTE0IsYEAdVLNIM314o7vBqrg8/edit |
| 売上報告シート | https://docs.google.com/spreadsheets/d/1QYLHr7wMj0jQW5ApgWf-l1Bk9_9M5EV54FDHQp2Rkic/edit |

### シート構成詳細

**出欠確認シート**
- 出欠状況（自動）
- ゲスト出欠状況（自動）
- 会員名簿マスター
- 受付名簿
- 配席アーカイブ
- 紹介記録

**設定シート**
- 設定
- ダッシュボード集計

**売上報告シート**
- sales（売上報告データ）
  - A: timestamp（送信日時）
  - B: line_user_id（LINEユーザーID）
  - C: display_name（LINE表示名）
  - D: report_date（報告日）
  - E: deals_count（成約件数）
  - F: sales_amount（売上金額）
  - G: join_next_seat（同席希望）
  - H: desired_industry（入会希望業種）
  - I: next_guest（次回ゲスト紹介）
  - N: 氏名（数式で自動取得）
  - O: event_key（例: `2026年2月_売上報告`）
  - P: meetings_count（商談件数）

---

## 5. LINE LIFF アプリ情報

| アプリ名 | LIFF ID | バックエンドGAS URL |
|----------|---------|---------------------|
| 出欠確認 | `2008446831-256wkWVY` | `AKfycby-rXLMy1lKIhyl60ZaS_tCob20cAuCgrkzf9B0Ky738DyvTWNBh_7AuDb1iNrEhf7ElA` |
| ゲスト申請 | `2008498539-Ak3Vxy1X` | `AKfycbxvtSiPkQSl1Gf3Xc8bQsUeBnrw6NT8QYQZbMYBISENkTB9ahk16GPcevEW_bNVpEYq` |
| 売上報告 | `2008299790-J3WnAP3j` | `AKfycbw_K0031SKFJbEeaG0OUYqW4grvWNnQ5briAdAvDlKpvrcB5Uek5h-SxtxxC9qdqDPC2w` |

**LINE Developers Console**: https://developers.line.biz/console/

---

## 6. GitHubリポジトリ連携

| リポジトリ | URL | 用途 | Cloudflare連携 |
|-----------|-----|------|----------------|
| shusei-automation | https://github.com/sodatebito5/shusei-automation | メインリポジトリ（LIFF、ドキュメント） | - |
| shusei-dashboard-v3 | https://github.com/sodatebito5/shusei-dashboard-v3 | ダッシュボードGAS | - |
| shusei-seat-manager | https://github.com/sodatebito5/shusei-seat-manager | 自動配席アプリ | `shusei-seat-manager` に自動デプロイ |

### ローカルフォルダとリポジトリの対応

| ローカルパス | リポジトリ |
|-------------|-----------|
| `C:\Users\Owner\shusei-automation` | sodatebito5/shusei-automation |
| `C:\Users\Owner\shusei-dashboard-v3` | sodatebito5/shusei-dashboard-v3 |

---

## 7. デプロイ方法まとめ

### ダッシュボード（GAS）

```bash
cd gas/dashboard
npm run push      # テスト環境に反映
npm run deploy    # 本番デプロイ
```

> **注意**: デプロイ後、GASエディタで「アクセスできるユーザー」が「全員」になっているか確認

### 出欠確認LIFF（Cloudflare Pages）

```bash
cd web/attendance
npx wrangler pages deploy . --project-name=shukketsu1
```

### ゲスト申請LIFF（Cloudflare Pages）

```bash
cd web/guest-form
npx wrangler pages deploy . --project-name=guest-form1
```

### 売上報告LIFF（Cloudflare Pages）

```bash
cd web/sales-report/liff
npx wrangler pages deploy . --project-name=uriage5
```

### 自動配席アプリ（Cloudflare Pages）

```bash
cd web/seat-maker
npx wrangler pages deploy . --project-name=shusei-seat-manager
```

または GitHub push で自動デプロイ（`sodatebito5/shusei-seat-manager`）

---

## 8. API エンドポイント

### ダッシュボードAPI

**ベースURL**: `https://script.google.com/macros/s/AKfycbzgVIzNjfW_UHihZS5bwrWM7xix0U4dnodZSlq7nPC8eGXGu_Fj6haCzivxiARVDPGL/exec`

| action | メソッド | 用途 |
|--------|----------|------|
| `getSeatingParticipants` | GET | 配席対象者を取得 |
| `syncSeats` | POST | 座席確定・受付名簿反映・アーカイブ保存 |
| `listSeatingArchives` | GET | 過去の配席一覧を取得 |
| `getSeatingArchive` | GET | 特定イベントの配席を取得 |

---

## 9. 機能一覧

| 機能 | 概要 | フロント | バックエンド |
|------|------|----------|-------------|
| 出欠連絡 | 月例会の出欠を登録 | LIFF（Cloudflare） | GAS |
| 売上報告 | メンバー間売上を報告 | LIFF（Cloudflare） | GAS |
| ゲスト申請 | ゲスト参加申請 | LIFF（Cloudflare） | GAS |
| 自動配席 | 座席表を自動生成 | Web（Cloudflare） | ダッシュボードGAS |
| ダッシュボード | 統合管理画面 | GAS WebApp | GAS |
| 配席表表示 | 配席をモーダルで閲覧 | ダッシュボード内 | ダッシュボードGAS |

---

## 10. 売上報告システム詳細

### データフロー

```
[LIFF] → POST → [GAS] → [スプレッドシート sales シート]
```

### eventKey形式

`{西暦}年{月}月_売上報告`

例: `2026年2月_売上報告`

※ 同一ユーザー＋同一eventKeyの組み合わせで既存行を上書き（UPSERT）

### UPSERT検索ロジック

1. B列（line_user_id）とO列（event_key）の両方が一致する行を検索
2. 一致すれば上書き（mode: updated）
3. なければ新規追加（mode: inserted）

### GAS デプロイ情報

- **GASプロジェクト**: スプレッドシートにバインド
- **スプレッドシートID**: `1QYLHr7wMj0jQW5ApgWf-l1Bk9_9M5EV54FDHQp2Rkic`
- **シート名**: `sales`（小文字）
- **データ開始行**: 7行目

### GAS WebApp URL

`https://script.google.com/macros/s/AKfycbw_K0031SKFJbEeaG0OUYqW4grvWNnQ5briAdAvDlKpvrcB5Uek5h-SxtxxC9qdqDPC2w/exec`

---

## 11. 関連ドキュメント

| ファイル | 内容 |
|----------|------|
| `CLAUDE.md` | Claude Code向け作業ルール |
| `PROGRESS.md` | 開発進捗サマリー |
| `docs/SYSTEM_CONFIG.md` | 技術構成詳細 |
| `docs/tasks.md` | タスク管理 |
| `docs/requirements.md` | 要件定義 |
| `docs/errors.md` | エラー解決履歴 |
| `docs/dashboard-spec.md` | ダッシュボード詳細仕様 |
