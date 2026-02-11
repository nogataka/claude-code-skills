# slide-templates

SlideKit 用の HTML スライドテンプレート集です。各テンプレートは 1280×720px のスライド形式で、`/slidekit-create` スキルのカスタムテンプレートとして利用できます。

---

## テンプレート一覧

| ニックネーム | ブランド/テーマ | 概要 | スライド数 |
|-------------|-----------------|------|-----------|
| **abc-navy** | ABC Corp. | ビジネスプレゼンテーション基本テンプレート集（Navy+Gold） | 20枚 |
| **venture-split** | OUTDOOR+ | ベンチャー風・左ダーク右白スプリット（Orange+Green） | 18枚 |
| **biz-plan-blue** | OUTDOOR+ | 事業計画書・青基調フォーマル（BIZ UDGothic, Blue+Amber） | 20枚 |
| **greenfield** | GreenField Inc. | 新規事業提案（Forest Green） | 20枚 |
| **novatech** | NovaTech Solutions | スタートアップ風（Navy+Orange） | 20枚 |
| **skyline** | SkyLine Corp. | 次世代ビジネス戦略（Cyan+Red） | 20枚 |
| **ai-proposal** | NexTech Solutions | AI導入プロジェクト提案書 | 20枚 |
| **customer-experience** | - | 顧客関係の変革（Customer Experience, M+1） | 13枚 |
| **ai-tech** | TechVantage | 現代AI技術ビジネスプレゼンテーション | 11枚 |
| **marketing-research** | - | Modern Corporate Marketing Research | 10枚 |
| **digital-report** | - | Strategic Corporate Digital Business Report（Navy+Gold） | 11枚 |

---

## テンプレート説明

### abc-navy
- **用途**: 一般的なビジネスプレゼンテーション
- **配色**: Navy + Gold
- **雰囲気**: 堅実・信頼感のあるフォーマル

### venture-split
- **用途**: ベンチャー向けピッチ、スタートアップ提案
- **配色**: Orange + Green
- **雰囲気**: 左右スプリットレイアウト、左ダーク・右白の大胆な構成

### biz-plan-blue
- **用途**: 事業計画書、経営計画
- **フォント**: BIZ UDGothic
- **配色**: Blue + Amber
- **雰囲気**: 青基調のフォーマルなビジネス資料

### greenfield
- **用途**: 新規事業提案、グリーンフィールド戦略
- **配色**: Forest Green
- **雰囲気**: 成長・新規性をイメージした緑系

### novatech
- **用途**: スタートアップ紹介、技術提案
- **配色**: Navy + Orange
- **雰囲気**: テック・イノベーション寄り

### skyline
- **用途**: 次世代ビジネス戦略、ロードマップ
- **配色**: Cyan + Red
- **雰囲気**: モダンでダイナミック

### ai-proposal
- **用途**: AI導入プロジェクト提案書
- **ブランド**: NexTech Solutions
- **雰囲気**: AI・DX 向けの提案資料

### customer-experience
- **用途**: 顧客体験の変革、CX 提案
- **フォント**: M+1
- **雰囲気**: 顧客中心設計の説明向け

### ai-tech
- **用途**: AI技術のビジネスプレゼンテーション
- **ブランド**: TechVantage
- **雰囲気**: 現代AI技術の解説・紹介

### marketing-research
- **用途**: 市場調査レポート、マーケティングリサーチ
- **雰囲気**: Modern Corporate スタイル

### digital-report
- **用途**: デジタルビジネス戦略レポート
- **配色**: Navy + Gold
- **雰囲気**: 企業向け戦略報告書

---

## スキルキットクリエイト（slidekit-create）へのテンプレート配置

`/slidekit-create` は `~/.claude/skills/slidekit-create/references/templates/` 配下の HTML をテンプレートとして認識します。`slide-templates` の各テンプレートをここに配置することで、スキル実行時にテンプレートモードで利用できます。

### 配置の仕方

#### 1. 単一テンプレートを配置する場合

使用したいニックネームのテンプレートを `templates/` 直下にコピーします。**1テンプレートセットあたり最大5ファイル**まで読み込まれるため、代表的なスライド（例: 001〜005）を選んで配置するか、サブディレクトリでまとめて配置します。

```bash
# 例: abc-navy を配置（001〜005 のみを使う場合）
mkdir -p ~/.claude/skills/slidekit-create/references/templates/abc-navy
cp slide-templates/abc-navy/001.html ~/.claude/skills/slidekit-create/references/templates/abc-navy/
cp slide-templates/abc-navy/002.html ~/.claude/skills/slidekit-create/references/templates/abc-navy/
cp slide-templates/abc-navy/003.html ~/.claude/skills/slidekit-create/references/templates/abc-navy/
cp slide-templates/abc-navy/004.html ~/.claude/skills/slidekit-create/references/templates/abc-navy/
cp slide-templates/abc-navy/005.html ~/.claude/skills/slidekit-create/references/templates/abc-navy/
```

#### 2. 複数テンプレートをまとめて配置する場合

各ニックネームをサブディレクトリ名として、`templates/` の下に配置します。`/slidekit-create` 実行時に一覧が表示され、使用するテンプレートを選択できます。

```bash
TEMPLATES_DIR=~/.claude/skills/slidekit-create/references/templates
SLIDE_TEMPLATES=/path/to/claude-code-skills/slide-templates

# 全テンプレートを配置（各テンプレートは先頭5ファイルのみが読み込まれる）
for name in abc-navy venture-split biz-plan-blue greenfield novatech skyline \
            ai-proposal customer-experience ai-tech marketing-research digital-report; do
  mkdir -p "$TEMPLATES_DIR/$name"
  cp "$SLIDE_TEMPLATES/$name"/00[1-5].html "$TEMPLATES_DIR/$name/" 2>/dev/null || true
done
```

#### 3. 配置後のディレクトリ構成イメージ

```
~/.claude/skills/slidekit-create/references/templates/
├── README.md
├── abc-navy/
│   ├── 001.html
│   ├── 002.html
│   ├── 003.html
│   ├── 004.html
│   └── 005.html
├── venture-split/
│   ├── 001.html
│   └── ...
├── biz-plan-blue/
│   └── ...
└── ... （他のニックネーム）
```

### 注意事項

- **最大5ファイル**: 1サブディレクトリあたり、アルファベット順で先頭5つの `.html` のみが読み込まれます。それ以上あってもスキップされます
- **HTML のみ**: 画像や PDF は無視されます
- **デザインのみ参照**: テキスト内容はコピーされず、カラーパレット・フォント・ヘッダー/フッター・装飾などのビジュアル要素のみが抽出されます
- **1280×720px**: スライドサイズは固定です

### テンプレートモード時の動作

テンプレートを配置して `/slidekit-create` を実行すると:

1. Phase 0 で `templates/` をスキャンし、サブディレクトリ一覧を表示
2. 使用するテンプレートを番号で選択（または「使用しない」で通常モード）
3. **テンプレートモード**の場合: スタイル・テーマ・カラー等のヒアリングをスキップし、コンテンツに関する質問のみ実施
4. 選んだテンプレートのデザインを踏襲してスライドを生成

---

## ディレクトリ構成（slide-templates）

```
slide-templates/
├── README.md                 # 本ドキュメント
├── abc-navy/
│   ├── 001.html ～ 020.html
│   └── print.html            # 印刷用（全スライド一覧）
├── venture-split/
│   ├── 001.html ～ 018.html
│   └── print.html
├── biz-plan-blue/
│   ├── 001.html ～ 020.html
│   └── print.html
├── greenfield/
│   └── ...
├── novatech/
│   └── ...
├── skyline/
│   └── ...
├── ai-proposal/
│   └── ...
├── customer-experience/
│   └── ...
├── ai-tech/
│   └── ...
├── marketing-research/
│   └── ...
└── digital-report/
    └── ...
```

各テンプレートには `print.html` が含まれ、ブラウザで開くと全スライドを一覧表示できます。印刷（Ctrl/Cmd + P）で PDF 化も可能です。
