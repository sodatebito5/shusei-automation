# Claude Code 実践ナレッジ

> ShinCode氏のYouTube動画4本から抽出した実践的知見

---

## 1. コンテキストエンジニアリング：4つの原則

LangChainが提唱するコンテキスト管理の基本原則。Claude Codeの潜在能力を10倍〜20倍に引き上げる。

### 1.1 Write（記述）- 外部メモを残す

**目的**: LLMの記憶限界（20万トークン）を補う

```
プロジェクト/
├── docs/
│   ├── requirements.md    # 要件定義
│   ├── architecture.md    # アーキテクチャ設計
│   ├── data-structure.md  # データ構造
│   ├── implementation-tasks.md  # タスク・進捗管理 ★重要
│   └── errors.md          # エラー解決履歴
├── CLAUDE.md              # コーディングルール
└── src/
```

**ポイント**:
- 要件定義・仕様書は最初にがっちり固める（スペック駆動開発）
- タスク管理ファイルでチェックボックス管理 → AIの進捗が一目瞭然
- エラー解決履歴を残す → 同じミスを繰り返さない

### 1.2 Select（選択）- 適切なタイミングで参照

**CLAUDE.mdに記載すべきルール例**:
```markdown
# Project Rules

## ドキュメント参照
- コードを書く前にdocs/を確認すること
- タスク完了時はimplementation-tasks.mdにチェックを入れること

## コーディングルール
- Server Componentsを積極的に使用
- useEffectは極力避ける
- エラー解決時はdocs/errors.mdに追記
```

### 1.3 Compress（圧縮）- コンテキストの自動圧縮

- Claude Codeは95%到達で**オートコンパクト**が自動発動
- 手動: `/compact` コマンド
- 会話履歴をクリアしつつ要約は保持

### 1.4 Isolate（分離）- サブエージェントでコンテキスト分離

**独立したコンテキストで動作** → メインコンテキストを圧迫しない

```
メインコンテキスト
    │
    ├── サブエージェント（パフォーマンス改善）
    │       └── MCP呼び出し可能
    │       └── 完了後に要約を返す
    │
    └── サブエージェント（セキュリティチェック）
```

**作成方法**:
```bash
# エージェント一覧確認
/agents

# 新規作成
/agents → Create new agent
```

**配置場所**: `.claude/agents/エージェント名.md`

---

## 2. おすすめMCPサーバー（優先度順）

### 2.1 Supabase MCP ★最重要

**できること**:
- テーブル構造・スキーマ確認
- 実データの検索・表示
- RLS設定の確認・作成
- Database Function / Trigger作成

**メリット**: 存在しないスキーマへのアクセスエラーを防ぐ

```json
{
  "mcpServers": {
    "supabase": {
      "command": "npx",
      "args": ["-y", "@supabase/mcp-server"],
      "env": {
        "SUPABASE_ACCESS_TOKEN": "your-token"
      }
    }
  }
}
```

### 2.2 Stripe MCP

**できること**:
- 商品の検索・作成
- 取引履歴確認
- ドキュメント検索（最新API情報）

### 2.3 Serena MCP - コードベース探索

**用途**: 巨大なコードベースから正確・高速に探索
- 指示した箇所を的確に特定して修正
- カスタムスラッシュコマンドと組み合わせて使用

### 2.4 Context7 MCP - ドキュメント検索

**用途**: 最新の公式ドキュメントを参照
- Next.js, Supabase, shadcn/ui, Tailwind CSS v4, Better Auth など
- 古い情報での実装を防ぐ

### 2.5 Playwright MCP - ブラウザ自動操作

**用途**:
- 動作テスト（AIが実際にブラウザを操作）
- スクリーンショット取得 → UIブラッシュアップ
- リリース前の最終確認

### 2.6 Notion MCP

**用途**:
- ブログ記事の自動執筆
- 学習メモの自動生成
- ナレッジベース構築

### 2.7 Figma MCP

**用途**: デザインの忠実な再現
- FigmaデザインをそのままHTML/CSS/Reactに変換
- **注意**: Claude 4.1 Opus推奨（Sonnetだと精度低下）

### 2.8 GitHub CLI (gh)

**用途**:
- リポジトリ作成
- ブランチ操作
- GitHub Pagesデプロイ

---

## 3. マルチエージェント組織構築

### 3.1 組織構造

```
プレジデント（社長）= 自分が指示
    │
    └── ボス1（プロジェクトマネージャー）
            │
            ├── ワーカー1（フロントエンド）
            ├── ワーカー2（バックエンド）
            └── ワーカー3（テスト/デプロイ）
```

### 3.2 セットアップ手順

```bash
# 1. リポジトリクローン
git clone https://github.com/xxx/claude-code-communication.git

# 2. セットアップ実行
sh setup.sh

# 3. マルチエージェントセッション起動（TMUX）
tmux attach -t multi-agent

# 4. プレジデントセッション起動
tmux attach -t president

# 5. 全Claude Code起動（自動承認モード）
claude --dangerously-skip-permissions
```

### 3.3 実践的な注意点

| 問題 | 対策 |
|------|------|
| 自律対話がうまくいかない | まず自分で1つずつ指示出し → 慣れたら組織化 |
| Next.js新規作成でTailwind当たらない | 手動で`npx create-next-app`してから任せる |
| レートリミット | $100プランだとすぐ到達。$200プラン推奨 |
| 指示書が弱い | CLAUDE.mdとinstructions/の内容を充実させる |

### 3.4 推奨フロー

```
1. 要件定義書作成
    ↓
2. 並行開発しやすいチケット分割を依頼
    ↓
3. 各チケットを1つずつ自分で指示出し
    ↓
4. うまくいったら組織化・自動化
```

---

## 4. 実践Tips

### 4.1 コンテキスト状態確認

```bash
/context  # シンキングモード(Tab)で実行
```

**確認項目**:
- Messages: 会話履歴の占有率
- System Prompt/Tools: デフォルトで約7%
- Auto-compact buffer: ここに達すると自動圧縮

### 4.2 便利なスラッシュコマンド

| コマンド | 用途 |
|----------|------|
| `/init` | CLAUDE.md初期化 |
| `/compact` | 手動でコンテキスト圧縮 |
| `/context` | コンテキスト状態確認 |
| `/mcp` | MCP接続状態確認 |
| `/agents` | サブエージェント管理 |
| `/memory` | メモリ編集 |

### 4.3 シンキングモード

**Tab**キーで切り替え → 複雑なタスク・設計・要件定義で使用

### 4.4 カスタムスラッシュコマンド

**配置場所**: `.claude/commands/コマンド名.md`

**例**: Serena探索用
```markdown
---
name: serena
description: Serena MCPでコードベース探索
---

Serena MCPを使用してコードベースを探索してください。
サブエージェントも起動してください。

探索対象: $ARGUMENTS
```

### 4.5 開発効率の目安

| フェーズ | Claude Code貢献度 |
|----------|-------------------|
| 要件定義〜MVP | 60-70% |
| 細かい調整・文言修正 | 20-30%（手動） |
| 課金実装（Stripe） | MCP活用で大幅効率化 |

---

## 5. マインドセット

### 5.1 バイブコーディングの本質

- **技術力の差はほぼなくなった** → 差別化は「やり切る力」
- セキュリティは必須で抑える。それ以外（パフォーマンス、コード綺麗さ）は後回しでOK
- ユーザーにとって大事なのは「価値」であって技術スタックではない

### 5.2 継続のコツ

```
小さな成功体験 → ドーパミン放出 → 継続したくなる
```

- 1人でもユーザー獲得 → 成功体験
- リリースできた → 成功体験
- いいね・コメントもらえた → 成功体験

### 5.3 学習性無力感の突破

「どうせできない」を突破するには**結果を出すしかない**
→ 小さくてもいいから成功体験を積む

---

## 6. 推奨ワークフロー

```
【開発開始】
1. docs/ディレクトリ作成
2. /init でCLAUDE.md初期化
3. 要件定義をAIと壁打ち → docs/requirements.md
4. タスク分割 → docs/implementation-tasks.md

【開発中】
5. フェーズごとに指示出し
6. 定期的に /context で状態確認
7. 必要に応じて /compact
8. MCP活用（Supabase, Stripe等）

【高度な活用】
9. サブエージェント作成（リサーチ、セキュリティ等）
10. カスタムスラッシュコマンド作成
11. 組織構築（慣れてから）

【リリース】
12. Playwright MCPで動作テスト
13. GitHub CLI / Vercel MCPでデプロイ
```

---

## 参考リンク

- [LangChain Context Engineering Blog](https://blog.langchain.dev/context-engineering/)
- [Claude Code 公式ドキュメント](https://docs.anthropic.com/claude-code)
- [Claude Code マルチエージェント構築](https://github.com/xxx/claude-code-communication)

---

*最終更新: 2026-01-12*
*出典: ShinCode YouTube チャンネル*
