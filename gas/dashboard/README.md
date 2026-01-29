# 守成ダッシュボード GAS

管理ダッシュボードのGoogle Apps Scriptコード。

## デプロイ

```bash
cd gas/dashboard
npm run push      # テスト環境に反映
npm run deploy    # 本番デプロイ
```

## ファイル構成

| ファイル | 説明 |
|---------|------|
| dashboard.js | メインAPI（出欠、配席、受付名簿等） |
| MemberMonthly.js | 会員月次処理 |
| index.html | ダッシュボードUI |
| style.html | CSS |

## 関連ドキュメント

- 詳細仕様: `docs/dashboard-spec.md`
- システム全体: `docs/SYSTEM_CONFIG.md`

## GAS情報

- scriptId: `1shW-t-_5b9j22faoQB7qCnqZ6w6t_WrFT-hZPULT_xVCPnM7VrgnHStG`
- テストURL: https://script.google.com/macros/s/AKfycbz5j-qZV2RW5nU2PoYEQYUQKooGSWtboHMPOgjSIQI/dev
