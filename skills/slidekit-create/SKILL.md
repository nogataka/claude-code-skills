---
name: slidekit-create
description: Generate HTML slide presentations (1 slide = 1 HTML file, 1280x720px) using Tailwind CSS, Font Awesome, and Google Fonts. Use when the user asks to create a new presentation deck or slide HTML files. Covers design guidelines, 15 layout patterns, component library, and PPTX conversion compatibility rules. Supports style selection (creative, elegant, modern, professional, minimalist) and theme selection (marketing, portfolio, business, technology, education).
---

# SlideKit Create

**All communication with the user MUST be in Japanese.** Questions, confirmations, progress updates, and any other messages — always use Japanese.

Generate HTML files forming a complete presentation deck. Each file is a self-contained HTML document rendered at **1280 x 720 px**. No JavaScript is used — all content is pure HTML + CSS.

For DOM snippets and component patterns, see [references/patterns.md](references/patterns.md).
For user-provided custom templates, see [references/templates/](references/templates/).

---

## Workflow Overview

| Phase | Name | Description |
|-------|------|-------------|
| 1 | ヒアリング | ユーザーに質問してスライドの要件を確認 |
| 2 | カスタムテンプレート読み込み | templates/ にHTMLがあればデザインを抽出 |
| 3 | デザイン決定 | カラーパレット・フォント・アイコンを確定 |
| 4 | スライド構成の設計 | 各スライドの役割・レイアウトパターンを計画 |
| 5 | HTML生成 | 全スライドを 001.html 〜 NNN.html として出力 |
| 6 | print.html 生成 | 全スライド一覧表示用ページを出力 |
| 7 | チェックリスト確認 | 制約・品質基準への適合を検証 |
| 8 | PPTX変換（任意） | /pptx スキルで PowerPoint に変換 |

---

## Phase 1: ヒアリング

Before generating any files, ask the user the following questions to capture intent. **All questions must be asked in Japanese.** Questions can be grouped into 2-3 messages rather than asking one by one.

### 1-1. 出力ディレクトリ

Ask the user for the target directory name.

- Default: `output/slide-page{NN}/` (NN = next sequential number)
- Custom: any path the user specifies

### 1-2. スタイル選択

| Style | Characteristics |
|-------|----------------|
| **Creative** | Bold colors, decorative elements, gradients, playful layouts |
| **Elegant** | Subdued palette (gold tones), serif-leaning type, generous whitespace |
| **Modern** | Flat design, vivid accent colors, sharp edges, tech-oriented |
| **Professional** | Navy/gray palette, structured layouts, higher info density |
| **Minimalist** | Few colors, extreme whitespace, typography-driven, minimal decoration |

### 1-3. テーマ選択

| Theme | Use Case |
|-------|----------|
| **Marketing** | Product launches, campaign proposals, market analysis |
| **Portfolio** | Case studies, work showcases, creative collections |
| **Business** | Business plans, executive reports, strategy proposals, investor pitches |
| **Technology** | SaaS introductions, tech proposals, DX initiatives, AI/data analysis |
| **Education** | Training materials, seminars, workshops, internal study sessions |

### 1-4. スライド内容のソース

Ask the user how they want to provide the content for the slides. This determines what text, data, and structure will appear in the deck.

| Option | Description |
|--------|-------------|
| **Reference file** | User provides a file (Markdown, text, Word, etc.) containing the content. Read the file and use its structure and text as the basis for the slides. |
| **Direct text input** | User types or pastes the content directly in the chat. |
| **Topic only** | User provides just a topic/title and lets Claude generate the content. Clarify the target audience and key messages before generating. |

**When a reference file is provided:**

1. Read the entire file
2. Extract the logical structure (headings, sections, bullet points, data)
3. Map the structure to the slide sequence — each major section becomes a section divider + content slides
4. Preserve key text, numbers, and data points faithfully
5. Adapt the content to fit the slide format (concise bullet points, not full paragraphs)

**When direct text is provided:**

1. Organize the text into a logical presentation flow
2. Ask clarifying questions if the structure is ambiguous

**When only a topic is given:**

1. Ask about the target audience (executives, engineers, clients, etc.)
2. Ask about key messages or points the user wants to convey
3. Generate content based on the answers

### 1-5. プレゼンタイトル

Ask for the presentation title. Skip if already clear from the content source in 1-4.

### 1-6. スライド枚数

Ask the user to choose: **10 / 15 / 20 (recommended) / 25 / Auto**

- **Auto**: Determine the optimal slide count based on the content source. Guidelines:
  - Reference file: count the major sections/headings → each section ≈ 1 divider + 2-3 content slides, plus cover/agenda/summary/closing
  - Direct text: estimate from the volume and structure of the provided text
  - Topic only: default to 15-20 based on topic complexity
- When Auto is selected, inform the user of the determined count before proceeding (e.g., "内容から判断して18枚で作成します。よろしいですか？")

### 1-7. 会社名・ブランド名

Ask for the company or brand name to display in the header/footer.

### 1-8. カラーの希望

Ask if the user has color preferences.

- If yes: use the specified colors as the basis for the palette
- If no: auto-suggest based on the selected style × theme combination

### 1-9. 背景画像の使用

Ask whether to use background images.

- Default: none (CSS gradients only)
- If yes: user must provide or approve specific images

---

## Phase 2: カスタムテンプレート読み込み

Check `references/templates/` for user-provided HTML template files.

1. **List** all `.html` files in `references/templates/`
2. If no files exist, skip to Phase 3
3. If files exist, **read each file** and extract:
   - Color palette (CSS custom properties / Tailwind classes)
   - Font pair (primary JP + accent Latin)
   - Header/footer structure and style
   - Decorative elements and visual motifs
   - Layout patterns used
4. Use the extracted design language as the **primary style reference** for the new deck
5. If the user also selected a style/theme in Phase 1, **blend** the custom template's design with the chosen style/theme — custom template takes priority for colors, fonts, and header/footer

### Cautions

- **Mandatory Constraints still apply.** Even if a custom template uses `<table>`, JavaScript, or non-CDN assets, the generated output must follow all rules in the Mandatory Constraints section below. Extract only the visual design (colors, fonts, spacing, decorative style) — not non-compliant implementation details.
- **Slide size must remain 1280x720px.** Ignore any different dimensions in custom templates.
- **Do not copy text content.** Custom templates are style references only. All text content comes from Phase 1 hearing.
- **Maximum 5 template files.** If more than 5 files exist, read only the first 5 (sorted alphabetically) to limit context usage. Warn the user that remaining files were skipped.
- **Supported format: HTML only.** Ignore non-HTML files (images, PDFs, etc.) in the templates directory.

---

## Phase 3: デザイン決定

Based on Phase 1 hearing results and Phase 2 custom templates (if loaded), determine the following **before generating any HTML**:

1. **Color palette** (3-4 custom colors — prefer custom template palette if available)
2. **Font pair** (1 Japanese + 1 Latin — prefer custom template fonts if available)
3. **Brand icon** (1 Font Awesome icon)

Present the design decisions to the user for confirmation before proceeding.

### Proven Palette Examples

| Template | Style | Primary Dark | Accent | Secondary | Fonts |
|----------|-------|-------------|--------|-----------|-------|
| 01 Navy & Gold | Elegant | `#0F2027` | `#C5A065` | `#2C5364` | Noto Sans JP + Lato |
| 02 Casual Biz | Professional | `#1f2937` | Indigo | `#F97316` | Noto Sans JP |
| 03 Blue & Orange | Professional | `#333333` | `#007BFF` | `#F59E0B` | BIZ UDGothic |
| 04 Green Forest | Modern | `#1B4332` | `#40916C` | `#52B788` | Noto Sans JP + Inter |
| 05 Dark Tech | Creative | `#0F172A` | `#F97316` | `#3B82F6` | Noto Sans JP + Inter |

---

## Phase 4: スライド構成の設計

Plan the full deck structure before writing any HTML. This phase produces a slide map.

### 4-1. Required Slides (All Decks)

| Position | Type | Pattern | Category |
|----------|------|---------|----------|
| First | Cover | Center | `cover` |
| Second | Agenda | HBF | `agenda` |
| Second to last | Summary | HBF | `conclusion` |
| Last | Closing | Full-bleed / Center | `conclusion` |

### 4-2. Section Dividers by Slide Count

| Slides | Section Dividers | Content Slides |
|--------|-----------------|----------------|
| **10** | 2 | 4 |
| **15** | 3 | 8 |
| **20** | 4 | 12 |
| **25** | 5 | 16 |

### 4-3. Build the Slide Map

For each slide, determine:

1. **File number** (`001.html`, `002.html`, ...)
2. **Type** (Cover / Agenda / Section Divider / Content / Summary / Closing)
3. **Layout pattern** (from the 15 patterns — see Phase 5 reference)
4. **Content summary** (what text/data goes on this slide)

Rules:
- Never use the same layout pattern for 3 or more consecutive slides
- Match content from Phase 1-4 (reference file / text / topic) to appropriate slide types
- Use good variety across the 15 layout patterns

### 4-4. Standard Composition for 20 Slides (Reference)

| File | Type | Pattern | Purpose |
|------|------|---------|---------|
| `001.html` | Cover | Center | Title, subtitle, presenter, date |
| `002.html` | Agenda | HBF | Numbered section list |
| `003.html` | Section Divider 1 | Left-Right Split | Section 1 introduction |
| `004.html` | Content | HBF + Top-Bottom Split | Challenges vs. solutions + KPI |
| `005.html` | Content | HBF + 2-Column | Comparison / contrast |
| `006.html` | Section Divider 2 | Left-Right Split | Section 2 introduction |
| `007.html` | Content | HBF + 3-Column | 3-item cards |
| `008.html` | Content | HBF + Grid Table | Competitive comparison table |
| `009.html` | Content | HBF + 2x2 Grid | Risk analysis / SWOT |
| `010.html` | Section Divider 3 | Left-Right Split | Section 3 introduction |
| `011.html` | Content | HBF + N-Column | Process flow |
| `012.html` | Content | HBF + Timeline/Roadmap | Quarterly roadmap |
| `013.html` | Content | HBF + KPI Dashboard | KPI cards + CSS bar chart |
| `014.html` | Section Divider 4 | Left-Right Split | Section 4 introduction |
| `015.html` | Content | HBF + Funnel | Conversion funnel |
| `016.html` | Content | HBF + Vertical Stack | Architecture / org chart |
| `017.html` | Content | HBF + 3-Column | Strategy / policy (3 pillars) |
| `018.html` | Content | HBF + 2-Column | Detailed analysis / data |
| `019.html` | Summary | HBF | Key takeaways + next actions |
| `020.html` | Closing | Full-bleed / Center | Thank-you slide + contact info |

---

## Phase 5: HTML生成

Generate all slide HTML files based on the slide map from Phase 4.

### 5-1. For Each Slide

1. Use the HTML Boilerplate (below)
2. Apply the design decisions from Phase 3 (colors, fonts, icon)
3. Use the layout pattern assigned in Phase 4
4. Fill in the content from Phase 1-4 (reference file / text / generated)
5. Save as `{output_dir}/{NNN}.html` (zero-padded: 001.html, 002.html, ...)

### HTML Boilerplate

```html
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="utf-8" />
    <meta content="width=device-width, initial-scale=1.0" name="viewport" />
    <title>{Slide Title}</title>
    <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet" />
    <link href="https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free@6.4.0/css/all.min.css" rel="stylesheet" />
    <link href="https://fonts.googleapis.com/css2?family={PrimaryFont}:wght@300;400;500;700;900&family={AccentFont}:wght@400;600;700&display=swap" rel="stylesheet" />
    <style>
        body { margin: 0; padding: 0; font-family: '{PrimaryFont}', sans-serif; overflow: hidden; }
        .font-accent { font-family: '{AccentFont}', sans-serif; }
        .slide { width: 1280px; height: 720px; position: relative; overflow: hidden; background: #FFFFFF; }
        /* Custom color classes: .bg-brand-dark, .bg-brand-accent, .bg-brand-warm, etc. */
    </style>
</head>
<body>
    <div class="slide {layout-classes}">
        <!-- Content -->
    </div>
</body>
</html>
```

### 15 Layout Patterns

Use one pattern per slide. For full DOM trees and component snippets, see [references/patterns.md](references/patterns.md).

| # | Pattern | Root classes | When to use |
|---|---------|-------------|-------------|
| 1 | **Center** | `flex flex-col items-center justify-center` | Cover, thank-you slides |
| 2 | **Left-Right Split** | `flex` with `w-1/3` + `w-2/3` | Chapter dividers, concept + detail |
| 3 | **Header-Body-Footer** | `flex flex-col` with header + `flex-1` + footer | Most content slides (default) |
| 4 | **HBF + 2-Column** | Pattern 3 body with two `w-1/2` | Comparison, data + explanation |
| 5 | **HBF + 3-Column** | Pattern 3 body with `grid grid-cols-3` | Card listings, 3-way comparison |
| 6 | **HBF + N-Column** | Pattern 3 body with `grid grid-cols-{N}` | Process flows (max 5 cols) |
| 7 | **Full-bleed** | `relative` with `absolute inset-0` layers | Impact covers (CSS gradient default) |
| 8 | **HBF + Top-Bottom Split** | Pattern 3 body with `flex flex-col` two sections | Content top + KPI/summary bar bottom |
| 9 | **HBF + Timeline/Roadmap** | Pattern 3 body with timeline bar + `grid grid-cols-4` | Quarterly roadmaps, phased plans |
| 10 | **HBF + KPI Dashboard** | Pattern 3 body with KPI `grid` + `flex-1` chart area | KPI cards + chart/progress visualization |
| 11 | **HBF + Grid Table** | Pattern 3 body with flex-based rows (`w-1/N`) | Feature comparison, competitive analysis |
| 12 | **HBF + Funnel** | Pattern 3 body with decreasing-width centered bars | Conversion funnel, sales pipeline |
| 13 | **HBF + Vertical Stack** | Pattern 3 body with stacked full-width cards + separators | Architecture diagrams, layered systems |
| 14 | **HBF + 2x2 Grid** | Pattern 3 body with `grid grid-cols-2` (2 rows) | Risk analysis, SWOT, feature overview |
| 15 | **HBF + Stacked Cards** | Pattern 3 body with vertically stacked full-width cards + numbered badges | FAQ, Q&A, numbered key points, interview summary |

### Heading Convention

Bilingual: small English label above, larger Japanese title below.

```html
<p class="text-xs uppercase tracking-widest text-gray-400 mb-1 font-accent">Market Analysis</p>
<h1 class="text-3xl font-bold text-brand-dark">市場分析</h1>
```

### Number Emphasis Convention

Large digits + small unit span:

```html
<p class="text-4xl font-black font-accent">415<span class="text-sm font-normal ml-1">M</span></p>
```

---

## Phase 6: print.html 生成

After all slide HTML files are generated, create `{output_dir}/print.html` for viewing and printing all slides:

```html
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="utf-8" />
    <title>View for Print</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { background: #FFFFFF; }
        .slide-frame {
            width: 1280px; height: 720px;
            margin: 20px auto;
            box-shadow: 0 2px 10px rgba(0,0,0,0.15);
            border: 1px solid #e2e8f0;
            overflow: hidden;
        }
        .slide-frame iframe { width: 1280px; height: 720px; border: none; }
        @media print {
            body { background: #FFFFFF; }
            .slide-frame {
                page-break-after: always; box-shadow: none; border: none;
                margin: 0 auto;
                transform: scale(0.85); transform-origin: top center;
            }
        }
    </style>
</head>
<body>
    <!-- One iframe per slide -->
    <div class="slide-frame"><iframe src="001.html"></iframe></div>
    <!-- ... repeat for all slides ... -->
</body>
</html>
```

Add one `<div class="slide-frame"><iframe src="{NNN}.html"></iframe></div>` per slide.

---

## Phase 7: チェックリスト確認

After Phase 5 and Phase 6 are complete, verify the following. Fix any issues before delivering to the user.

- [ ] All files use identical CDN links
- [ ] Custom colors defined identically in every `<style>`
- [ ] Root is single `<div>` under `<body>` with `overflow: hidden`
- [ ] Slide size exactly 1280 x 720
- [ ] No external images (unless user approved)
- [ ] No JavaScript whatsoever (no `<script>` tags)
- [ ] All files for the chosen slide count are present
- [ ] Font sizes follow hierarchy
- [ ] Consistent header/footer on content slides
- [ ] Page numbers increment correctly
- [ ] `Confidential` in footer
- [ ] Decorative elements use low z-index and low opacity
- [ ] File naming: zero-padded 3 digits (`001.html`, `002.html`, ...)
- [ ] Text uses `<p>` / `<h*>` (not `<div>`)
- [ ] No visible text in `::before` / `::after`
- [ ] No one-off colors outside palette
- [ ] Content density guidelines followed
- [ ] `print.html` generated with iframes for all slides

---

## Phase 8: PPTX変換（任意）

After all checks pass, ask the user in Japanese:

> 「HTMLスライドの生成が完了しました。PowerPoint（PPTX）に変換しますか？」

If the user declines, the workflow ends here.

### Prerequisite Check

Before invoking the pptx skill, verify it is available by checking the list of available skills in the current session. The pptx skill will appear in the system's skill list if installed.

**If the pptx skill is NOT available**, inform the user and provide installation instructions:

> `/pptx` スキルが必要ですが、現在インストールされていません。
>
> 以下のコマンドでインストールできます:
>
> ```
> claude install-skill https://github.com/anthropics/claude-code-agent-skills/tree/main/skills/pptx
> ```
>
> インストール後、新しいセッションで `/pptx` を実行し、出力ディレクトリ `{output_dir}` を指定してください。

Then end the workflow. Do not attempt conversion without the pptx skill.

### Invocation (pptx skill available)

If the pptx skill is available, invoke `/pptx` using the Skill tool. Pass the following context:

1. **Source directory** — the output path containing all `NNN.html` files
2. **Slide count** — total number of HTML files
3. **Presentation title** — from Phase 1 hearing
4. **Color palette** — the 3-4 brand colors chosen in Phase 3
5. **Font pair** — primary (JP) and accent (Latin) fonts

Example invocation prompt for the Skill tool:

```
Convert the HTML slide deck in {output_dir} to a single PPTX file.
- {N} slides (001.html through {NNN}.html)
- Title: {title}
- Colors: {primary_dark}, {accent}, {secondary}
- Fonts: {primary_font} + {accent_font}
- Output: {output_dir}/presentation.pptx
```

**Important:** Do not attempt HTML-to-PPTX conversion yourself. Always delegate to the `/pptx` skill, which has its own specialized workflow, QA process, and conversion tools.

---

## Reference: Mandatory Constraints

| Rule | Value |
|------|-------|
| Slide size | `width: 1280px; height: 720px` |
| CSS framework | Tailwind CSS 2.2.19 via CDN |
| Icons | Font Awesome 6.4.0 via CDN |
| Fonts | Google Fonts (1 JP primary + 1 Latin accent) |
| Language | `lang="ja"` |
| Root DOM | `<body>` -> single wrapper `<div>` (no siblings) |
| Overflow | `overflow: hidden` on root wrapper |
| External images | None by default. Explicit user approval required |
| JavaScript | **Strictly forbidden.** No Chart.js or any JS library. All data visualization must be CSS-only |
| Custom CSS | Inline `<style>` in `<head>` only; no external CSS files |

### PPTX Conversion Rules (Critical)

These directly affect PPTX conversion accuracy. **Always follow.**

- Prefer `<p>` over `<div>` for text (tree-walkers may miss `<div>` text)
- Never put visible text in `::before` / `::after`
- Separate decorative elements with `-z-10` / `z-0`
- Max DOM nesting: 5-6 levels
- Font Awesome icons in `<i>` tags (converter detects `fa-` on `<i>`)
- Use flex-based tables over `<table>`
- `linear-gradient(...)` supported; complex multi-stop may fall back to screenshot
- `box-shadow`, `border-radius`, `opacity` all extractable

### Anti-Patterns (Avoid)

- Purposeless wrapper `<div>`s (increases nesting)
- One-off colors outside the palette
- `<table>` for layout
- Inline styles that Tailwind can replace
- Text in `::before` / `::after`
- `<div>` for text (use `<p>`, `<h1>`-`<h6>`)

---

## Reference: Design Guidelines

### Color Palette

Define 3-4 custom colors as Tailwind-style utility classes in `<style>`:

| Role | Class Example | Purpose |
|------|---------------|---------|
| Primary Dark | `.bg-brand-dark` | Dark backgrounds, titles |
| Primary Accent | `.bg-brand-accent` | Borders, highlights, icons |
| Warm/Secondary | `.bg-brand-warm` | CTAs, emphasis, badges |
| Body text | Tailwind grays | Body, captions |

Keep palette consistent across all slides. No one-off colors.

### Font Pair

| Role | Examples | Usage |
|------|----------|-------|
| Primary (JP) | Noto Sans JP, BIZ UDGothic | Body, headings |
| Accent (Latin) | Lato, Inter, Roboto | Numbers, English labels, page numbers |

Set primary on `body`; define `.font-accent` for the accent font.

### Font Size Hierarchy

| Purpose | Tailwind class |
|---------|---------------|
| Main title | `text-3xl` - `text-6xl` + `font-bold`/`font-black` |
| Section heading | `text-xl` - `text-2xl` + `font-bold` |
| Card heading | `text-lg` + `font-bold` |
| Body text | `text-sm` - `text-base` |
| Caption/label | `text-xs` |

### Content Density Guidelines

| Element | Recommended Max |
|---------|----------------|
| Bullet points | 5-6 |
| Cards per row | 3-4 |
| Body text lines | 6-8 |
| KPI boxes | 4-6 |
| Process steps | 4-5 |
