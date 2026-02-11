# SlideKit

[Claude Code](https://docs.anthropic.com/en/docs/claude-code) 用のプレゼンテーションスライド制作ツールキットです。

HTMLスライドの新規作成から、既存PDFのHTML化、さらにPowerPoint（PPTX）への変換まで、スライド制作の一連のワークフローをカバーします。

## Claude Code スキルとは

Claude Code のスキルは、Claude に特定の専門知識やワークフローを教えるための仕組みです。スキルをインストールすると、Claude Code のセッション内でスラッシュコマンドとして呼び出せるようになります。

スキルの実体は `SKILL.md` というMarkdownファイルで、Claude が従うべき手順・制約・デザインルールなどが記述されています。

## 収録スキル一覧

| コマンド | 概要 |
|---------|------|
| [`/slidekit-create`](#slidekitcreate) | HTMLスライドを新規作成 |
| [`/slidekit-templ`](#slidekittempl) | PDFからHTMLスライドテンプレートを作成 |

2つのスキルは連携して使えます:

```
slidekit-templ でPDFからテンプレートを作成
        ↓
  templates/ に配置（サブディレクトリで複数管理も可能）
        ↓
slidekit-create でテンプレートを選択 → テンプレートモードでスライドを新規作成
        ↓
  /pptx で PowerPoint に変換（任意）
```

## インストール

全スキルを一括でインストール:

```bash
claude install-skill https://github.com/nogataka/claude-code-skills
```

個別にインストールする場合:

```bash
# slidekit-create のみ
claude install-skill https://github.com/nogataka/claude-code-skills/tree/main/skills/slidekit-create

# slidekit-templ のみ
claude install-skill https://github.com/nogataka/claude-code-skills/tree/main/skills/slidekit-templ
```

---

## slidekit-create

HTMLスライドプレゼンテーションを新規作成するスキルです。

### 特徴

- **1スライド = 1 HTMLファイル**（1280 x 720px）
- **Tailwind CSS** + **Font Awesome** + **Google Fonts** をCDN経由で使用
- 基本は純粋な HTML + CSS。データ可視化が必要な場合のみ Chart.js を自動で使用
- PPTX変換を考慮したDOM構造（`/pptx` スキルと連携可能）
- `print.html` で全スライドの一覧表示・印刷に対応
- **一問一答形式** — 1メッセージにつき1質問で、回答を待ってから次へ進むシンプルな対話フロー
- **番号選択方式** — 選択肢がある質問はすべて番号入力で回答可能（テキスト入力が必要な項目のみ自由入力）

### 対応スタイル（5種類）

| スタイル | 特徴 |
|---------|------|
| **Creative** | 大胆な配色、装飾要素、グラデーション、遊び心のあるレイアウト |
| **Elegant** | 落ち着いたパレット（ゴールド系）、セリフ寄りのタイポグラフィ、広めの余白 |
| **Modern** | フラットデザイン、鮮やかなアクセントカラー、シャープなエッジ、テック志向 |
| **Professional** | ネイビー/グレー系、構造的なレイアウト、情報密度高め |
| **Minimalist** | 少ない色数、極端な余白、タイポグラフィ主導、最小限の装飾 |

### 対応テーマ（5種類）

| テーマ | 用途 |
|-------|------|
| **Marketing** | 製品発表、キャンペーン提案、市場分析 |
| **Portfolio** | ケーススタディ、実績紹介、作品集 |
| **Business** | 事業計画、経営レポート、戦略提案、投資家ピッチ |
| **Technology** | SaaS紹介、技術提案、DX推進、AI/データ分析 |
| **Education** | 研修資料、セミナー、ワークショップ、社内勉強会 |

### スライド枚数

10 / 15 / 20（推奨）/ 25 / 自動 から番号で選択できます。「自動」を選ぶと、内容のボリュームに応じて最適な枚数が自動決定されます。

### 20種類のレイアウトパターン

カバー、セクション区切り、2カラム、3カラム、タイムライン、KPIダッシュボード、ファネル、グリッドテーブル、2x2グリッド、TAM/SAM/SOM、チャプター区切り、コンタクト、VS比較など、20種類のレイアウトパターンを使い分けて多彩なスライドを生成します。

### カスタムテンプレート機能

自作のHTMLスライドや `slidekit-templ` で作成したテンプレートをスタイルの参考資料として登録できます。`/slidekit-create` 実行時にテンプレートが検出されると、使用するかどうかをユーザーに確認します。使用を選択すると、カラーパレット・フォント・ヘッダー/フッター・装飾要素などのデザインが自動的に抽出され、デザイン関連のヒアリング（スタイル選択・テーマ選択・カラー希望・背景画像）がスキップされます。

**使い方:**

```
~/.claude/skills/slidekit-create/references/templates/
```

上記ディレクトリにHTMLファイルを配置します。

**単一テンプレートの場合** — ディレクトリ直下にHTMLファイルを配置:

```
templates/
├── 001.html
├── 002.html
└── README.md
```

**複数テンプレートセットの場合** — サブディレクトリで分類:

```
templates/
├── navy-gold/
│   ├── 001.html
│   └── 002.html
├── modern-tech/
│   ├── 001.html
│   └── 002.html
└── README.md
```

複数のテンプレートセットがある場合、一覧が表示されどれを使用するか選択できます。「使用しない」を選ぶと通常モード（フルヒアリング）で作成されます。

**注意事項:**

- HTMLファイルのみ対応（画像やPDFは無視されます）
- 1テンプレートセットあたり最大5ファイル（超過分はスキップされます）
- テキスト内容はコピーされません（抽出されるのはビジュアルデザインのみ）
- 元テンプレートにJavaScriptや`<table>`レイアウトがあっても、出力はスキルの制約ルールに従います

### ワークフロー

`/slidekit-create` は2つのモードで動作します:

**テンプレートモード** — `references/templates/` にカスタムテンプレートがあり、ユーザーが使用を選択した場合。デザイン関連のヒアリングがスキップされ、コンテンツに関する質問のみが行われます。テンプレートのデザインをそのまま使うか、一部変更するかを確認します。

**通常モード** — テンプレートがない場合、またはユーザーが使用しないと選択した場合。スタイル・テーマ・カラーなどを含むフルヒアリングが行われます。

どちらのモードでも、ヒアリングは**一問一答形式**で進行します。選択肢のある質問は番号で回答し、タイトルや会社名など自由記述が必要な項目のみテキストで入力します。

| Phase | 名前 | 説明 |
|-------|------|------|
| 0 | テンプレート検出・選択 | `templates/` を確認し、テンプレートモード or 通常モードを決定 |
| 1 | ヒアリング | モードに応じた質問でスライドの要件を確認 |
| 2 | デザイン決定 | カラーパレット（3〜4色）、フォントペア（日本語＋欧文）、ブランドアイコンを決定 |
| 3 | スライド構成の設計 | 各スライドの役割・レイアウトパターンを計画 |
| 4 | HTML生成 | 全スライドを `001.html` 〜 `NNN.html` として出力 |
| 5 | print.html生成 | 全スライドをiframeで並べた印刷用ページを出力 |
| 6 | チェックリスト確認 | 制約・品質基準への適合を検証 |
| 7 | PPTX変換（任意） | `/pptx` スキルがあれば PowerPoint に変換（後述） |

### 使い方

Claude Code で以下のように呼び出します:

```
/slidekit-create
```

または自然言語で依頼できます:

```
プレゼン資料を作ってください
```

---

## slidekit-templ

既存のPDFプレゼンテーションからHTMLスライドテンプレートを作成するスキルです。

生成したテンプレートは `slidekit-create` のカスタムテンプレートとして登録でき、既存資料のデザインを踏襲した新しいスライドを作成できます。

### 特徴

- PDFの各ページをスクリーンショット化し、Claudeがそれを見ながらHTMLを再現
- `slidekit-create` のデザインシステムに準拠したHTML出力
- 色・テキスト・レイアウトを元のPDFに忠実に再現
- 変換後のビジュアルQAプロセスを内蔵

### 変換パイプライン

```
PDF → (pdftoppm) → スライド画像（JPEG）
                        ↓
    Claude が各画像を読み取り → HTMLを作成
                        ↓
         001.html, 002.html, ... + print.html
```

### 前提条件

**Poppler**（`pdftoppm` コマンド）が必要です:

```bash
# macOS
brew install poppler

# Ubuntu / Debian
sudo apt-get install poppler-utils

# Windows (chocolatey)
choco install poppler
```

### 使い方

Claude Code で以下のように呼び出します:

```
/slidekit-templ
```

または自然言語で依頼できます:

```
このPDFからテンプレートを作ってください: ./presentation.pdf
```

### ワークフロー

1. **スライド画像の生成** — PDFを `pdftoppm` でJPEG画像に変換
2. **デッキ分析** — 全スライド画像を読み取り、カラーパレット・フォント・ヘッダー/フッターのパターンを特定
3. **デザインシステムの読み込み** — `slidekit-create` のルール・パターンを参照
4. **HTML生成** — 各スライド画像を見ながら、最も近いレイアウトパターンを選択してHTMLを作成
5. **print.html生成** — 全スライドの一覧表示用ページを出力
6. **ビジュアルQA** — 生成したHTMLと元のスクリーンショットを比較し、差異があれば修正
7. **PPTX変換（任意）** — `/pptx` スキルがあれば PowerPoint に変換可能（手順は下記セクションを参照）

### テンプレートとして登録する

`slidekit-templ` で生成したHTMLの中から、デザインの参考にしたいファイルを選んでコピーします:

```bash
# 直接配置する場合
cp output/templ/003.html ~/.claude/skills/slidekit-create/references/templates/

# サブディレクトリで管理する場合（複数テンプレートセット向き）
mkdir -p ~/.claude/skills/slidekit-create/references/templates/my-design
cp output/templ/*.html ~/.claude/skills/slidekit-create/references/templates/my-design/
```

登録後、`/slidekit-create` を実行するとテンプレートが検出され、使用するかどうかを確認されます。

---

## HTML から PowerPoint（PPTX）への変換

`slidekit-create` および `slidekit-templ` が生成するのはHTMLファイルです。最終的にPowerPointファイル（`.pptx`）として使いたい場合は、別途 `/pptx` スキルを使って変換します。

### 事前準備

`/pptx` スキルをインストールしておきます（SlideKitには含まれていません）:

```bash
claude install-skill https://github.com/anthropics/claude-code-agent-skills/tree/main/skills/pptx
```

### 変換手順

**方法A: ワークフロー内で自動変換**

`/slidekit-create` のワークフロー完了後、Claude が「PPTXに変換しますか？」と確認します。「はい」と答えれば `/pptx` スキルが自動で呼び出され、HTMLからPowerPointに変換されます。

**方法B: 手動で変換**

既にHTMLスライドが生成済みの場合は、Claude Code で直接依頼できます:

```
output/slide-page01/ のHTMLスライドをPPTXに変換してください
```

または `/pptx` スキルを直接呼び出します:

```
/pptx
```

### 変換の全体像

```
/slidekit-create                         /pptx
        │                                    │
  テンプレート検出                       HTMLファイルを読み取り
        ↓                                    ↓
  ヒアリング（モードに応じ簡略化）      各スライドのDOM・CSSを解析
        ↓                                    ↓
  HTMLスライド生成                       PowerPoint要素に変換
        ↓                                    ↓
  001.html ~ NNN.html ──────→  presentation.pptx を出力
  + print.html
```

PDF からの場合:

```
PDF ──→ /slidekit-templ ──→ HTML ──→ /pptx ──→ PPTX
```

> **なぜHTMLを経由するのか:**
> PDFから直接PPTXに変換するのではなく、一度HTMLを経由することで、PPTX変換しやすいDOM構造を確保しつつ、HTMLの段階で内容の確認・修正が行えます。

### 注意事項

- HTMLのDOM構造がPPTX変換精度に直結します。SlideKit はPPTX変換を前提としたルール（テキストは `<p>` / `<h*>` を使う、flexテーブルを使う、`::before` / `::after` にテキストを入れない等）を守ってHTMLを生成するため、高い変換精度が得られます
- 複雑なCSSグラデーションや Chart.js のチャート（`<canvas>` 要素）など一部の装飾はスクリーンショットとして埋め込まれる場合があります
- `/pptx` スキルがインストールされていない場合、変換はスキップされます

---

## ディレクトリ構成

```
claude-code-skills/              （リポジトリ名）
├── README.md
├── skills/
│   ├── slidekit-create/         # /slidekit-create
│   │   ├── SKILL.md             # スキル定義（ルール・制約・ワークフロー）
│   │   └── references/
│   │       ├── patterns.md      # 20レイアウトパターンのDOMツリーとコンポーネント集
│   │       └── templates/       # カスタムテンプレート置き場
│   │           ├── README.md    # テンプレートの使い方
│   │           └── {name}/      # サブディレクトリで複数テンプレートセットを管理可能
│   └── slidekit-templ/          # /slidekit-templ
│       ├── SKILL.md             # スキル定義（変換パイプライン・QAプロセス）
│       └── scripts/
│           └── pdf_to_images.py # PDF→JPEG変換スクリプト（pdftoppm利用）
└── .gitignore
```

## 動作環境

- [Claude Code](https://docs.anthropic.com/en/docs/claude-code) がインストール済みであること
- `slidekit-templ` を使う場合は Poppler（`pdftoppm`）が必要
- PPTX変換を行う場合は別途 [`pptx` スキル](https://github.com/anthropics/claude-code-agent-skills/tree/main/skills/pptx) のインストールが必要

## よくある質問

### スキルはどこにインストールされますか？

`~/.claude/skills/` 配下にインストールされます:

```
~/.claude/skills/slidekit-create/SKILL.md
~/.claude/skills/slidekit-templ/SKILL.md
```

### スキルをアンインストールするには？

`~/.claude/skills/` 内の該当ディレクトリを削除してください:

```bash
rm -rf ~/.claude/skills/slidekit-create
rm -rf ~/.claude/skills/slidekit-templ
```

### 生成されたHTMLをPowerPointに変換できますか？

はい。詳しい手順は [HTML から PowerPoint（PPTX）への変換](#html-から-powerpointpptxへの変換) セクションをご覧ください。

### HTMLスライドを印刷・PDF化するには？

生成される `print.html` をブラウザで開き、ブラウザの印刷機能（Ctrl/Cmd + P）から「PDFとして保存」を選択してください。

### slidekit-templ と slidekit-create の違いは？

| | slidekit-templ | slidekit-create |
|---|---|---|
| **入力** | 既存のPDFファイル | ユーザーへのヒアリング |
| **目的** | PDFのデザインをHTMLで再現し、テンプレート化 | 新しいスライドをゼロから作成 |
| **用途** | 既存資料のデザインを流用したいとき | 新しいプレゼン資料を作りたいとき |

2つを組み合わせると、既存PDFのデザインを踏襲しつつ新しい内容のスライドを作成できます。

## ライセンス

MIT
