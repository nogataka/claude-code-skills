---
name: create-slide-template
description: Generate HTML slide presentation templates (1 slide = 1 HTML file, 1280x720px) using Tailwind CSS, Font Awesome, and Google Fonts. Use when the user asks to create a new slide template, presentation deck, or slide HTML files. Covers design guidelines, 15 layout patterns, component library, and PPTX conversion compatibility rules derived from template01-07 analysis. Supports style selection (creative, elegant, modern, professional, minimalist) and theme selection (marketing, portfolio, business, technology, education).
---

# Create Slide Template

Generate HTML files forming a complete presentation deck. Slide count is selected by the user (10 / 15 / 20 / 25). Each file is a self-contained HTML document rendered at **1280 x 720 px**. No JavaScript is used — all content is pure HTML + CSS.

For DOM snippets and component patterns, see [references/patterns.md](references/patterns.md).

---

## Phase 1: User Hearing

Before generating any files, ask the user the following questions to capture intent.

### 1. Output Directory

Ask the user for the target directory name.

- Default: `output/slide-page{NN}/` (NN = next sequential number)
- Custom: any path the user specifies

### 2. Style Selection

| Style | Characteristics |
|-------|----------------|
| **Creative** | Bold colors, decorative elements, gradients, playful layouts |
| **Elegant** | Subdued palette (gold tones), serif-leaning type, generous whitespace |
| **Modern** | Flat design, vivid accent colors, sharp edges, tech-oriented |
| **Professional** | Navy/gray palette, structured layouts, higher info density |
| **Minimalist** | Few colors, extreme whitespace, typography-driven, minimal decoration |

### 3. Theme Selection

| Theme | Use Case |
|-------|----------|
| **Marketing** | Product launches, campaign proposals, market analysis |
| **Portfolio** | Case studies, work showcases, creative collections |
| **Business** | Business plans, executive reports, strategy proposals, investor pitches |
| **Technology** | SaaS introductions, tech proposals, DX initiatives, AI/data analysis |
| **Education** | Training materials, seminars, workshops, internal study sessions |

### 4. Additional Questions

- Presentation title/topic
- **Slide count** — ask the user to choose: **10 / 15 / 20 (recommended) / 25 slides**
- Company or brand name (for header/footer)
- Color preferences (if any; otherwise auto-suggest from style x theme)
- Background image usage (default: none)

---

## Phase 2: Template Design

Based on hearing results, determine these before generating code:

1. **Color palette** (3-4 custom colors)
2. **Font pair** (1 Japanese + 1 Latin)
3. **Brand icon** (1 Font Awesome icon)

### Proven Palette Examples (template01-05)

| Template | Style | Primary Dark | Accent | Secondary | Fonts |
|----------|-------|-------------|--------|-----------|-------|
| 01 Navy & Gold | Elegant | `#0F2027` | `#C5A065` | `#2C5364` | Noto Sans JP + Lato |
| 02 Casual Biz | Professional | `#1f2937` | Indigo | `#F97316` | Noto Sans JP |
| 03 Blue & Orange | Professional | `#333333` | `#007BFF` | `#F59E0B` | BIZ UDGothic |
| 04 Green Forest | Modern | `#1B4332` | `#40916C` | `#52B788` | Noto Sans JP + Inter |
| 05 Dark Tech | Creative | `#0F172A` | `#F97316` | `#3B82F6` | Noto Sans JP + Inter |

---

## Mandatory Constraints

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

### CDN Links (Unified)

```html
<link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet" />
<link href="https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free@6.4.0/css/all.min.css" rel="stylesheet" />
<link href="https://fonts.googleapis.com/css2?family={PrimaryFont}:wght@300;400;500;700;900&family={AccentFont}:wght@400;600;700&display=swap" rel="stylesheet" />
```

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

## Design Guidelines

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

---

## HTML Boilerplate

```html
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="utf-8" />
    <meta content="width=device-width, initial-scale=1.0" name="viewport" />
    <title>{Slide Title}</title>
    <!-- CDN links here -->
    <style>
        body { margin: 0; padding: 0; font-family: '{PrimaryFont}', sans-serif; overflow: hidden; }
        .font-accent { font-family: '{AccentFont}', sans-serif; }
        .slide { width: 1280px; height: 720px; position: relative; overflow: hidden; background: #FFFFFF; }
        /* Custom color classes */
    </style>
</head>
<body>
    <div class="slide {layout-classes}">
        <!-- Content -->
    </div>
</body>
</html>
```

---

## Slide Sequence (with Category Mapping)

Structure the deck according to the user's chosen slide count. The **required slides** below must be included in every deck. Adjust the number of content slides accordingly.

### Required Slides (All Counts)

| Position | Type | Pattern | Category |
|----------|------|---------|----------|
| First | Cover | Center | `cover` |
| Second | Agenda | HBF | `agenda` |
| Second to last | Summary | HBF | `conclusion` |
| Last | Closing | Full-bleed / Center | `conclusion` |

### Section Dividers by Slide Count

| Slides | Section Dividers | Content Slides |
|--------|-----------------|----------------|
| **10** | 2 | 4 |
| **15** | 3 | 8 |
| **20** | 4 | 12 |
| **25** | 5 | 16 |

### Standard Composition for 20 Slides (Reference)

| File | Type | Pattern | Category | Purpose |
|------|------|---------|----------|---------|
| `001.html` | Cover | Center | `cover` | Title, subtitle, presenter, date |
| `002.html` | Agenda | HBF | `agenda` | Numbered section list |
| `003.html` | Section Divider 1 | Left-Right Split | `content` | Section 1 introduction |
| `004.html` | Content | HBF + Top-Bottom Split | `content` | Challenges vs. solutions + KPI |
| `005.html` | Content | HBF + 2-Column | `content` | Comparison / contrast |
| `006.html` | Section Divider 2 | Left-Right Split | `content` | Section 2 introduction |
| `007.html` | Content | HBF + 3-Column | `content` | 3-item cards |
| `008.html` | Content | HBF + Grid Table | `content` | Competitive comparison table |
| `009.html` | Content | HBF + 2x2 Grid | `content` | Risk analysis / SWOT |
| `010.html` | Section Divider 3 | Left-Right Split | `content` | Section 3 introduction |
| `011.html` | Content | HBF + N-Column | `content` | Process flow |
| `012.html` | Content | HBF + Timeline/Roadmap | `content` | Quarterly roadmap |
| `013.html` | Content | HBF + KPI Dashboard | `content` | KPI cards + CSS bar chart |
| `014.html` | Section Divider 4 | Left-Right Split | `content` | Section 4 introduction |
| `015.html` | Content | HBF + Funnel | `content` | Conversion funnel |
| `016.html` | Content | HBF + Vertical Stack | `content` | Architecture / org chart |
| `017.html` | Content | HBF + 3-Column | `content` | Strategy / policy (3 pillars) |
| `018.html` | Content | HBF + 2-Column | `content` | Detailed analysis / data |
| `019.html` | Summary | HBF | `conclusion` | Key takeaways + next actions |
| `020.html` | Closing | Full-bleed / Center | `conclusion` | Thank-you slide + contact info |

Use the 15 layout patterns with good variety across content slides. Never use the same pattern for 3 or more consecutive slides.

---

## 15 Layout Patterns

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

---

## Heading Convention

Bilingual: small English label above, larger Japanese title below.

```html
<p class="text-xs uppercase tracking-widest text-gray-400 mb-1 font-accent">Market Analysis</p>
<h1 class="text-3xl font-bold text-brand-dark">市場分析</h1>
```

## Number Emphasis Convention

Large digits + small unit span:

```html
<p class="text-4xl font-black font-accent">415<span class="text-sm font-normal ml-1">M</span></p>
```

---

## Pre-Delivery Checklist

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

## Phase 2.5: Generate print.html

After all slides pass the checklist, generate `{output_dir}/print.html` for viewing/printing all slides:

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
            .slide-frame { page-break-after: always; box-shadow: none; border: none; margin: 0 auto; }
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

## Phase 3: PPTX Conversion (Optional)

After all HTML files pass the Pre-Delivery Checklist, ask the user:

> "Would you like to convert these slides to a PPTX file?"

If the user declines, the workflow ends here.

### Prerequisite Check

Before invoking the pptx skill, verify it is available by checking the list of available skills in the current session. The pptx skill will appear in the system's skill list if installed.

**If the pptx skill is NOT available**, inform the user and provide installation instructions:

> The `/pptx` skill is required for PPTX conversion but is not currently installed.
>
> To install it, run the following command in your terminal:
>
> ```
> claude install-skill https://github.com/anthropics/claude-code-agent-skills/tree/main/skills/pptx
> ```
>
> After installation, you can convert the HTML slides by running `/pptx` in a new session and pointing it to the output directory: `{output_dir}`

Then end the workflow. Do not attempt conversion without the pptx skill.

### Invocation (pptx skill available)

If the pptx skill is available, invoke `/pptx` using the Skill tool. Pass the following context:

1. **Source directory** — the output path containing all `NNN.html` files
2. **Slide count** — total number of HTML files
3. **Presentation title** — from Phase 1 hearing
4. **Color palette** — the 3-4 brand colors chosen in Phase 2
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
