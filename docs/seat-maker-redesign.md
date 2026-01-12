# 自動配席アプリ 再設計 詳細設計書

> 作成日: 2025-01-12
> ステータス: 設計完了、実装待ち

## 概要

自動配席アプリのアーキテクチャを刷新し、ダッシュボードGASに統合する。

### 目標
- 「自動配席用」シートを廃止
- ダッシュボードGASにAPI統合
- 例会アーカイブ機能を追加

---

## 1. 新しいシート構成

### スプレッドシート（ID: 1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic）

| シート名 | 用途 | 変更 |
|----------|------|------|
| 出欠状況（自動） | LIFFからの出欠回答 | そのまま |
| ゲスト出欠状況（自動） | ゲスト申請データ | そのまま |
| 会員名簿マスター | 会員情報 | そのまま |
| 他会場名簿マスター | 他会場参加者 | そのまま |
| 受付名簿（自動） | 会員の受付名簿（座席反映先） | そのまま |
| 受付名簿（他会場・ゲスト） | 他会場/ゲストの受付名簿 | そのまま |
| **配席アーカイブ** | 過去の配席履歴 | **新規作成** |
| ~~自動配席用~~ | ~~中間データ~~ | **廃止** |

### 配席アーカイブ シート構造

| 列 | 項目 | 型 | 説明 |
|----|------|-----|------|
| A | 例会キー | string | "202501_01" 形式 |
| B | 例会日 | date | 2025-01-15 |
| C | 参加者ID | string | userId または生成ID |
| D | 氏名 | string | |
| E | 区分 | string | 会員/ゲスト/他会場 |
| F | 所属 | string | 福岡飯塚/他会場名 |
| G | 役割 | string | 代表世話人/世話人等 |
| H | チーム | string | A/B/C等 |
| I | テーブル | string | A/B/C.../PA/MC |
| J | 席番 | number | 0=マスター, 1-7=一般 |
| K | 確定日時 | datetime | 保存した日時 |
| L | 確定者 | string | userId（操作者） |

---

## 2. API設計

### 2.1 getSeatingParticipants(eventKey?)

**用途**: 配席アプリが参加者一覧を取得

```javascript
/**
 * 配席用参加者データを取得
 * @param {string} eventKey - 例会キー（省略時は現在の例会）
 * @returns {Object} 参加者データ
 */
function getSeatingParticipants(eventKey) {
  // eventKeyがなければ設定シートから現在の例会キーを取得

  return {
    success: true,
    eventKey: "202501_01",
    eventDate: "2025-01-15",
    eventTitle: "第90回例会",
    participants: [
      {
        id: "U1234567890abcdef",  // userIdまたは生成ID
        name: "山田 太郎",
        category: "会員",         // 会員/ゲスト/他会場
        affiliation: "福岡飯塚",  // 所属会場
        role: "代表世話人",       // 役割
        team: "A",                // チーム
        business: "建設業",       // 営業内容
        lastTable: "A",           // 前回のテーブル（参考用）
        lastSeat: 0
      }
    ],
    summary: {
      total: 45,
      members: 30,
      guests: 10,
      otherVenue: 5
    }
  };
}
```

### 2.2 getSeatingArchive(eventKey)

**用途**: 過去の配席を取得

```javascript
/**
 * 過去の配席データを取得
 * @param {string} eventKey - 例会キー
 * @returns {Object} 配席アーカイブ
 */
function getSeatingArchive(eventKey) {
  return {
    success: true,
    eventKey: "202501_01",
    eventDate: "2025-01-15",
    confirmedAt: "2025-01-14T10:30:00",
    confirmedBy: "U1234567890abcdef",
    assignments: [
      {
        id: "U1234567890abcdef",
        name: "山田 太郎",
        category: "会員",
        table: "A",
        seat: 0  // 0=マスター
      }
    ],
    tables: {
      "A": { master: "山田 太郎", members: ["佐藤 花子", "..."] },
      "B": { master: "鈴木 一郎", members: ["..."] }
    }
  };
}
```

### 2.3 syncSeats(payload)

**用途**: 座席を確定して保存

```javascript
/**
 * 座席情報を保存（受付名簿反映 + アーカイブ保存）
 * @param {Object} payload - 座席データ
 * @returns {Object} 結果
 */
function syncSeats(payload) {
  // payload構造
  const input = {
    eventKey: "202501_01",
    userId: "U1234567890abcdef",  // 操作者
    assignments: [
      { id: "xxx", name: "山田 太郎", category: "会員", table: "A", seat: 0 },
      { id: "yyy", name: "佐藤 花子", category: "会員", table: "A", seat: 1 }
    ]
  };

  // 1. 受付名簿に反映（既存ロジック）
  // 2. 配席アーカイブに保存（新規）

  return {
    success: true,
    eventKey: "202501_01",
    updatedMain: 25,      // 会員受付名簿の更新数
    updatedOther: 15,     // 他会場/ゲスト受付名簿の更新数
    archivedCount: 40,    // アーカイブ保存数
    unmatched: []         // マッチしなかった人
  };
}
```

### 2.4 listSeatingArchives()

**用途**: アーカイブ一覧を取得

```javascript
/**
 * 配席アーカイブの一覧を取得
 * @returns {Object} アーカイブ一覧
 */
function listSeatingArchives() {
  return {
    success: true,
    archives: [
      { eventKey: "202501_01", eventDate: "2025-01-15", confirmedAt: "...", count: 45 },
      { eventKey: "202412_01", eventDate: "2024-12-18", confirmedAt: "...", count: 42 }
    ]
  };
}
```

---

## 3. データフロー図

```
┌─────────────────────────────────────────────────────────────────────┐
│                    スプレッドシート                                  │
├─────────────────────────────────────────────────────────────────────┤
│                                                                      │
│  ┌──────────────────┐  ┌──────────────────┐  ┌──────────────────┐  │
│  │ 出欠状況（自動）  │  │ゲスト出欠状況    │  │他会場名簿マスター │  │
│  │ ← LIFF出欠       │  │（自動）← LIFF    │  │                  │  │
│  └────────┬─────────┘  └────────┬─────────┘  └────────┬─────────┘  │
│           │                     │                     │            │
│           └─────────────────────┼─────────────────────┘            │
│                                 ▼                                   │
│                    ┌────────────────────────┐                      │
│                    │ getSeatingParticipants │                      │
│                    │ （GAS内で統合処理）     │                      │
│                    └────────────┬───────────┘                      │
│                                 │                                   │
│  ┌──────────────────┐          │          ┌──────────────────┐    │
│  │ 受付名簿（自動）  │◀─────────┼──────────│ 配席アーカイブ    │    │
│  │ ← syncSeats()    │          │          │ ← syncSeats()    │    │
│  └──────────────────┘          │          └──────────────────┘    │
│  ┌──────────────────┐          │                                   │
│  │受付名簿（他会場） │◀─────────┘                                   │
│  │ ← syncSeats()    │                                              │
│  └──────────────────┘                                              │
│                                                                      │
└─────────────────────────────────────────────────────────────────────┘
                                 │
                                 │ 統合API
                                 ▼
                    ┌────────────────────────┐
                    │   ダッシュボードGAS     │
                    │   (SHUSEI-DASHBOARD-V3) │
                    ├────────────────────────┤
                    │ 既存:                   │
                    │  - getDashboardDataFast │
                    │  - getReceptionRoster   │
                    │                         │
                    │ 新規:                   │
                    │  - getSeatingParticipants│
                    │  - getSeatingArchive    │
                    │  - syncSeats            │
                    │  - listSeatingArchives  │
                    └────────────┬────────────┘
                                 │
              ┌──────────────────┴──────────────────┐
              ▼                                      ▼
    ┌──────────────────┐                  ┌──────────────────┐
    │   ダッシュボード   │                  │   配席アプリ      │
    │   (例会事前準備)  │                  │  (Cloudflare)    │
    │                   │                  │                  │
    │ - リンクボタン     │ ────────────────▶│ - データ読込     │
    │ - 配席履歴表示    │                  │ - ドラッグ配席   │
    │   （将来）        │                  │ - 座席確定       │
    └──────────────────┘                  └──────────────────┘
```

---

## 4. 移行手順

### Phase 1: 準備（Day 1）

| Step | 作業 |
|------|------|
| 1-1 | 「配席アーカイブ」シート作成 |
| 1-2 | 現在の「自動配席用」シートの内容をバックアップ |
| 1-3 | 「自動配席用」シートの更新元を確認・停止 |

### Phase 2: API実装（Day 2-3）

| Step | 作業 | ファイル |
|------|------|----------|
| 2-1 | `getSeatingParticipants()` 実装 | dashboard.js |
| 2-2 | `getSeatingArchive()` 実装 | dashboard.js |
| 2-3 | `syncSeats()` 移植・拡張 | dashboard.js |
| 2-4 | `listSeatingArchives()` 実装 | dashboard.js |
| 2-5 | doGet/doPost にルーティング追加 | dashboard.js |
| 2-6 | clasp push & デプロイ | - |

### Phase 3: 配席アプリ改修（Day 4）

| Step | 作業 | ファイル |
|------|------|----------|
| 3-1 | API_URL を ダッシュボードGAS に変更 | main.js |
| 3-2 | データ形式の調整 | main.js |
| 3-3 | 座席確定時に syncSeats() 呼び出し追加 | main.js |
| 3-4 | ステージングデプロイ・テスト | - |

### Phase 4: 検証（Day 5）

| Step | 作業 |
|------|------|
| 4-1 | 参加者データ取得テスト |
| 4-2 | 座席確定→アーカイブ保存テスト |
| 4-3 | 受付名簿への反映テスト |
| 4-4 | 過去配席の取得テスト |

### Phase 5: 本番移行（Day 6）

| Step | 作業 |
|------|------|
| 5-1 | 配席アプリ本番デプロイ |
| 5-2 | 「自動配席用」シートを非表示 |
| 5-3 | gas/seat-maker/ を deprecated マーク |
| 5-4 | ドキュメント更新 |

---

## 5. 実装優先度

| 優先度 | 機能 | 理由 |
|--------|------|------|
| **P0** | getSeatingParticipants | 配席アプリの動作に必須 |
| **P0** | syncSeats | 座席確定に必須 |
| **P1** | 配席アーカイブ保存 | 履歴管理 |
| **P2** | getSeatingArchive | 過去配席の参照 |
| **P2** | listSeatingArchives | 履歴一覧表示 |

---

## 6. 関連ファイル

| ファイル | 用途 | 変更 |
|----------|------|------|
| shusei-dashboard-v3/src/dashboard.js | ダッシュボードGAS | API追加 |
| shusei-automation/web/seat-maker/main.js | 配席アプリ | API_URL変更 |
| shusei-automation/gas/seat-maker/ | 旧API | 廃止予定 |
