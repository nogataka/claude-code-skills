# Claude Code Skills

[Claude Code](https://docs.anthropic.com/en/docs/claude-code) 用のカスタムスキル集です。

プレゼンテーション資料の作成・変換に特化したスキルを提供しています。

## Claude Code スキルとは

Claude Code のスキルは、Claude に特定の専門知識やワークフローを教えるための仕組みです。スキルをインストールすると、Claude Code のセッション内でスラッシュコマンド（例: `/create-slide-template`）として呼び出せるようになります。

スキルの実体は `SKILL.md` というMarkdownファイルで、Claude が従うべき手順・制約・デザインルールなどが記述されています。

## 収録スキル一覧

| スキル名 | 概要 |
|---------|------|
| [create-slide-template](#create-slide-template) | HTMLスライドテンプレートを新規作成 |
| [pdf-to-html](#pdf-to-html) | PDF資料をHTMLスライドに変換 |

---

## create-slide-template

HTMLスライドプレゼンテーションのテンプレートを新規作成するスキルです。

### 特徴

- **1スライド = 1 HTMLファイル**（1280 x 720px）
- **Tailwind CSS** + **Font Awesome** + **Google Fonts** をCDN経由で使用
- JavaScript不使用 — 純粋なHTML + CSSのみ
- PPTX変換を考慮したDOM構造（`/pptx` スキルと連携可能）
- `print.html` で全スライドの一覧表示・印刷に対応

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

10 / 15 / 20（推奨）/ 25枚から選択できます。

### 15種類のレイアウトパターン

カバー、セクション区切り、2カラム、3カラム、タイムライン、KPIダッシュボード、ファネル、グリッドテーブル、2x2グリッドなど、15種類のレイアウトパターンを使い分けて多彩なスライドを生成します。

### カスタムテンプレート機能

自作のHTMLスライドをスタイルの参考資料として登録できます。登録すると、新しいデッキ生成時にカラーパレット・フォント・ヘッダー/フッター・装飾要素などのデザインが自動的に抽出・反映されます。

**使い方:**

```
~/.claude/skills/create-slide-template/references/templates/
```

上記ディレクトリに自作HTMLファイルを配置するだけです。`/create-slide-template` 実行時に自動で検出されます。

**注意事項:**

- HTMLファイルのみ対応（画像やPDFは無視されます）
- 最大5ファイルまで（超過分はスキップされます）
- テキスト内容はコピーされません（抽出されるのはビジュアルデザインのみ）
- 元テンプレートにJavaScriptや`<table>`レイアウトがあっても、出力はスキルの制約ルールに従います

### ワークフロー

1. **ヒアリング** — 出力先ディレクトリ、スタイル、テーマ、タイトル、枚数、カラーなどを確認
2. **カスタムテンプレート読み込み** — `references/templates/` にHTMLがあればデザインを抽出
3. **デザイン決定** — カラーパレット（3〜4色）、フォントペア（日本語＋欧文）、ブランドアイコンを決定
4. **HTML生成** — 全スライドを `001.html` 〜 `NNN.html` として出力
5. **print.html生成** — 全スライドをiframeで並べた印刷用ページを出力
6. **PPTX変換（任意）** — `/pptx` スキルがあれば PowerPoint に変換可能

### インストール

```bash
claude install-skill https://github.com/nogataka/claude-code-skills/tree/main/skills/create-slide-template
```

### 使い方

Claude Code で以下のように呼び出します:

```
/create-slide-template
```

または自然言語で依頼できます:

```
プレゼン資料のHTMLテンプレートを作ってください
```

---

## pdf-to-html

既存のPDFプレゼンテーションを高品質なHTMLスライドファイルに変換するスキルです。

### 特徴

- PDFの各ページをスクリーンショット化し、Claudeがそれを見ながらHTMLを再現
- `create-slide-template` のデザインシステムに準拠したHTML出力
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

### インストール

```bash
claude install-skill https://github.com/nogataka/claude-code-skills/tree/main/skills/pdf-to-html
```

### 使い方

Claude Code で以下のように呼び出します:

```
/pdf-to-html
```

または自然言語で依頼できます:

```
このPDFをHTMLスライドに変換してください: ./presentation.pdf
```

### ワークフロー

1. **スライド画像の生成** — PDFを `pdftoppm` でJPEG画像に変換
2. **デッキ分析** — 全スライド画像を読み取り、カラーパレット・フォント・ヘッダー/フッターのパターンを特定
3. **デザインシステムの読み込み** — `create-slide-template` のルール・パターンを参照
4. **HTML生成** — 各スライド画像を見ながら、最も近いレイアウトパターンを選択してHTMLを作成
5. **print.html生成** — 全スライドの一覧表示用ページを出力
6. **ビジュアルQA** — 生成したHTMLと元のスクリーンショットを比較し、差異があれば修正

---

## ディレクトリ構成

```
claude-code-skills/
├── README.md
├── skills/
│   ├── create-slide-template/
│   │   ├── SKILL.md              # スキル定義（ルール・制約・ワークフロー）
│   │   └── references/
│   │       ├── patterns.md       # 15レイアウトパターンのDOMツリーとコンポーネント集
│   │       └── templates/        # カスタムテンプレート置き場（自作HTMLを配置）
│   │           └── README.md     # テンプレートの使い方
│   └── pdf-to-html/
│       ├── SKILL.md              # スキル定義（変換パイプライン・QAプロセス）
│       └── scripts/
│           └── pdf_to_images.py  # PDF→JPEG変換スクリプト（pdftoppm利用）
└── .gitignore
```

## 動作環境

- [Claude Code](https://docs.anthropic.com/en/docs/claude-code) がインストール済みであること
- `pdf-to-html` を使う場合は Poppler（`pdftoppm`）が必要
- PPTX変換を行う場合は別途 [`pptx` スキル](https://github.com/anthropics/claude-code-agent-skills/tree/main/skills/pptx) のインストールが必要

## よくある質問

### スキルはどこにインストールされますか？

`~/.claude/skills/` 配下にインストールされます。例:

```
~/.claude/skills/create-slide-template/SKILL.md
~/.claude/skills/pdf-to-html/SKILL.md
```

### スキルをアンインストールするには？

`~/.claude/skills/` 内の該当ディレクトリを削除してください:

```bash
rm -rf ~/.claude/skills/create-slide-template
rm -rf ~/.claude/skills/pdf-to-html
```

### 生成されたHTMLをPowerPointに変換できますか？

はい。`/pptx` スキルが別途インストールされていれば、`create-slide-template` のワークフロー最後にPPTX変換を提案します。以下でインストールできます:

```bash
claude install-skill https://github.com/anthropics/claude-code-agent-skills/tree/main/skills/pptx
```

### HTMLスライドを印刷・PDF化するには？

生成される `print.html` をブラウザで開き、ブラウザの印刷機能（Ctrl/Cmd + P）から「PDFとして保存」を選択してください。

## ライセンス

MIT
