# Reference: Layout Patterns & Component Snippets

Copy-paste DOM structures for each layout pattern and reusable component.

---

## Layout Pattern DOM Trees

### Pattern 1: Center (Cover / Thank-you)

```html
<div class="slide flex flex-col items-center justify-center">
    <!-- Decorative background -->
    <div class="absolute top-{n} right-{n} w-{n} h-{n} rounded-full bg-{color} opacity-{low} -z-10"></div>

    <div class="text-center z-10 relative">
        <p class="text-sm uppercase tracking-widest text-brand-accent mb-4">SUBTITLE LABEL</p>
        <h1 class="text-6xl font-black text-gray-900 mb-6">メインタイトル</h1>
        <div class="w-20 h-1 bg-brand-accent mx-auto mb-8"></div>
        <p class="text-xl text-gray-600">サブタイトルまたはキャッチコピー</p>

        <div class="mt-12 flex items-center justify-center gap-8 text-sm text-gray-500">
            <div class="flex items-center">
                <i class="fas fa-calendar-alt mr-2 text-brand-accent"></i>
                <span>2026/XX/XX</span>
            </div>
            <div class="flex items-center">
                <i class="fas fa-user mr-2 text-brand-accent"></i>
                <span>発表者名</span>
            </div>
        </div>
    </div>
</div>
```

### Pattern 2: Left-Right Split

```html
<div class="slide flex">
    <!-- Left Panel (dark) -->
    <div class="w-1/3 bg-brand-dark text-white p-10 flex flex-col justify-between relative overflow-hidden">
        <div class="absolute top-0 right-0 w-40 h-40 bg-brand-accent rounded-bl-full opacity-20"></div>
        <div class="relative z-10">
            <div class="flex items-center space-x-3 mb-6">
                <div class="w-10 h-10 bg-brand-accent rounded-lg flex items-center justify-center">
                    <i class="fas fa-{icon}"></i>
                </div>
                <span class="text-brand-accent font-bold uppercase text-sm tracking-widest">Section Label</span>
            </div>
            <h1 class="text-4xl font-black leading-tight mb-6">セクション<br />タイトル</h1>
            <p class="text-sm text-gray-300 leading-relaxed">説明テキスト</p>
        </div>
        <p class="relative z-10 text-xs text-gray-500">Confidential</p>
    </div>

    <!-- Right Panel (light) -->
    <div class="w-2/3 bg-white p-10 flex flex-col">
        <!-- Content -->
    </div>
</div>
```

### Pattern 3: Header-Body-Footer (Default Content)

```html
<div class="slide flex flex-col">
    <!-- Header -->
    <div class="px-16 pt-10 pb-4 flex justify-between items-end border-b border-gray-200 mx-16">
        <div class="flex items-center space-x-4">
            <div class="w-1.5 h-10 bg-brand-accent"></div>
            <div>
                <p class="text-xs text-gray-400 font-accent tracking-widest uppercase mb-1">English Label</p>
                <h1 class="text-3xl font-bold text-brand-dark tracking-tight">スライドタイトル</h1>
            </div>
        </div>
        <div class="flex items-center space-x-2 text-brand-dark opacity-50">
            <i class="fas fa-{brand-icon} text-lg"></i>
            <p class="text-xs font-bold tracking-widest uppercase font-accent">BRAND</p>
        </div>
    </div>

    <!-- Body -->
    <div class="flex-1 px-16 py-8">
        <!-- Slide content -->
    </div>

    <!-- Footer -->
    <div class="h-12 w-full flex justify-between items-center px-16 bg-white border-t border-gray-100">
        <p class="text-xs text-gray-400 tracking-wider">会社名 - Confidential</p>
        <div class="flex items-center space-x-2">
            <span class="text-xs text-gray-400">Page</span>
            <span class="text-sm font-bold text-brand-accent font-accent">{NN}</span>
        </div>
    </div>
</div>
```

### Pattern 4: HBF + 2-Column Body

Same header/footer as Pattern 3. Body:

```html
<div class="flex-1 px-16 py-8 flex gap-8">
    <div class="w-1/2 flex flex-col gap-4">
        <!-- Left column -->
    </div>
    <div class="w-1/2 flex flex-col gap-4">
        <!-- Right column -->
    </div>
</div>
```

### Pattern 5: HBF + 3-Column Body

```html
<div class="flex-1 px-16 py-8 grid grid-cols-3 gap-6">
    <div class="bg-white rounded-xl p-6 border-t-4 border-{color1} shadow-sm flex flex-col">
        <!-- Card -->
    </div>
    <div class="bg-white rounded-xl p-6 border-t-4 border-{color2} shadow-sm flex flex-col">
        <!-- Card -->
    </div>
    <div class="bg-white rounded-xl p-6 border-t-4 border-{color3} shadow-sm flex flex-col">
        <!-- Card -->
    </div>
</div>
```

### Pattern 6: HBF + N-Column Process Flow

```html
<div class="flex-1 px-16 py-8 grid grid-cols-{N} gap-4">
    <div class="bg-white rounded-lg shadow-sm border-t-4 border-brand-accent p-5 flex flex-col relative">
        <div class="flex justify-between items-start mb-3">
            <div class="w-10 h-10 rounded-full bg-{light} flex items-center justify-center text-brand-accent">
                <i class="fas fa-{icon}"></i>
            </div>
            <span class="text-xs font-bold bg-brand-accent text-white px-2 py-0.5 rounded">Step {N}</span>
        </div>
        <h3 class="text-lg font-bold text-gray-900 mb-2">ステップ名</h3>
        <p class="text-xs text-gray-500 leading-relaxed">説明テキスト</p>
        <!-- Arrow (except last) -->
        <div class="absolute top-1/2 -right-3 transform -translate-y-1/2 z-20 text-gray-300 text-xl">
            <i class="fas fa-chevron-right"></i>
        </div>
    </div>
</div>
```

### Pattern 7: Full-Bleed

**Default: CSS gradient (no external images)**

```html
<div class="slide relative flex flex-col overflow-hidden">
    <!-- CSS gradient background -->
    <div class="absolute inset-0 z-0" style="background: linear-gradient(135deg, {dark1}, {dark2});"></div>
    <!-- Accent -->
    <div class="absolute top-0 left-0 w-3 h-full bg-brand-accent z-20"></div>
    <!-- Content -->
    <div class="relative z-20 w-full h-full flex flex-col justify-center px-24">
        <h1 class="text-7xl font-black text-white leading-none mb-8">タイトル</h1>
        <p class="text-3xl text-gray-300 font-light">サブタイトル</p>
    </div>
</div>
```

**With image (user-approved only):**

```html
<div class="absolute inset-0 z-0">
    <img class="w-full h-full object-cover" src="{url}" alt="" />
</div>
<div class="absolute inset-0 z-10" style="background-color: rgba(0,0,0,0.7);"></div>
```

---

## Component Snippets

### Icon Card

```html
<div class="flex items-start">
    <div class="flex-shrink-0 w-10 h-10 rounded-full bg-{light} flex items-center justify-center text-brand-accent mr-4">
        <i class="fas fa-{icon}"></i>
    </div>
    <div>
        <h3 class="font-bold text-gray-900 text-lg mb-1">見出し</h3>
        <p class="text-sm text-gray-600">説明テキスト</p>
    </div>
</div>
```

### Border-Left Highlight

```html
<div class="border-l-4 border-brand-accent pl-6 py-2">
    <h3 class="text-lg font-bold text-gray-900 mb-1">見出し</h3>
    <p class="text-sm text-gray-600 leading-relaxed">説明テキスト</p>
</div>
```

### KPI Metric Box (Light)

```html
<div class="bg-white border border-gray-200 rounded-xl p-5 shadow-sm flex items-center justify-between relative overflow-hidden">
    <div class="absolute left-0 top-0 bottom-0 w-1 bg-brand-accent"></div>
    <div>
        <p class="text-xs text-gray-500 font-bold uppercase tracking-wider mb-1">Metric Label</p>
        <p class="text-2xl font-bold text-gray-900">415<span class="text-sm font-normal text-gray-500 ml-1">M</span></p>
    </div>
    <div class="w-10 h-10 rounded-full bg-{light} flex items-center justify-center text-brand-accent">
        <i class="fas fa-{icon}"></i>
    </div>
</div>
```

### KPI Metric Box (Dark)

```html
<div class="bg-brand-dark rounded-xl p-5 text-white text-center shadow-md">
    <p class="text-xs opacity-80 mb-1">ラベル</p>
    <p class="text-3xl font-black">数値</p>
    <p class="text-xs mt-1 opacity-70">補足テキスト</p>
</div>
```

### Badge / Tag

```html
<span class="inline-block bg-brand-accent text-white text-xs font-bold px-2 py-0.5 rounded">Label</span>
```

### Decorative Background Circle

```html
<div class="absolute top-{n} right-{n} w-{size} h-{size} rounded-full bg-{color} opacity-{5-30} -z-10"></div>
```

### Flex Table (no `<table>`)

```html
<div class="bg-white rounded-lg overflow-hidden shadow-sm border border-gray-200">
    <!-- Header Row -->
    <div class="flex bg-brand-accent text-white text-xs font-bold">
        <div class="flex-1 px-4 py-2">Column A</div>
        <div class="flex-1 px-4 py-2">Column B</div>
        <div class="flex-1 px-4 py-2">Column C</div>
    </div>
    <!-- Data Rows -->
    <div class="flex text-sm border-b border-gray-100">
        <div class="flex-1 px-4 py-2 font-bold text-gray-800">Value</div>
        <div class="flex-1 px-4 py-2 text-gray-600">Value</div>
        <div class="flex-1 px-4 py-2 text-gray-600">Value</div>
    </div>
</div>
```

### CSS Bar Chart (No JavaScript)

```html
<div class="space-y-3">
    <!-- Bar Item -->
    <div class="flex items-center gap-3">
        <p class="text-xs text-gray-500 w-16 text-right font-accent">FY2024</p>
        <div class="flex-1 bg-gray-100 rounded-full h-6 relative overflow-hidden">
            <div class="bg-brand-accent h-6 rounded-full flex items-center justify-end pr-2" style="width: 65%;">
                <span class="text-xs font-bold text-white font-accent">¥32億</span>
            </div>
        </div>
    </div>
    <div class="flex items-center gap-3">
        <p class="text-xs text-gray-500 w-16 text-right font-accent">FY2025</p>
        <div class="flex-1 bg-gray-100 rounded-full h-6 relative overflow-hidden">
            <div class="bg-brand-accent h-6 rounded-full flex items-center justify-end pr-2" style="width: 80%;">
                <span class="text-xs font-bold text-white font-accent">¥48億</span>
            </div>
        </div>
    </div>
    <div class="flex items-center gap-3">
        <p class="text-xs text-gray-500 w-16 text-right font-accent">FY2026E</p>
        <div class="flex-1 bg-gray-100 rounded-full h-6 relative overflow-hidden">
            <div class="bg-brand-warm h-6 rounded-full flex items-center justify-end pr-2" style="width: 95%;">
                <span class="text-xs font-bold text-white font-accent">¥65億</span>
            </div>
        </div>
    </div>
</div>
```

### CSS Donut / Pie (No JavaScript)

Use conic-gradient for pie/donut charts:

```html
<div class="flex items-center gap-8">
    <!-- Donut -->
    <div class="w-40 h-40 rounded-full relative" style="background: conic-gradient(#00B4D8 0% 55%, #FF6B6B 55% 80%, #94A3B8 80% 100%);">
        <div class="absolute inset-4 bg-white rounded-full flex items-center justify-center">
            <p class="text-lg font-black text-brand-dark font-accent">100%</p>
        </div>
    </div>
    <!-- Legend -->
    <div class="space-y-2">
        <div class="flex items-center gap-2">
            <div class="w-3 h-3 rounded-full bg-brand-accent"></div>
            <p class="text-sm text-gray-600">SaaS <span class="font-bold font-accent">55%</span></p>
        </div>
        <div class="flex items-center gap-2">
            <div class="w-3 h-3 rounded-full bg-brand-warm"></div>
            <p class="text-sm text-gray-600">Consulting <span class="font-bold font-accent">25%</span></p>
        </div>
        <div class="flex items-center gap-2">
            <div class="w-3 h-3 rounded-full bg-gray-400"></div>
            <p class="text-sm text-gray-600">Other <span class="font-bold font-accent">20%</span></p>
        </div>
    </div>
</div>
```

### CSS Progress Bar

```html
<div>
    <div class="flex items-center justify-between mb-1">
        <p class="text-sm font-bold text-brand-dark">プロジェクトA</p>
        <p class="text-sm font-bold text-brand-accent font-accent">72%</p>
    </div>
    <div class="w-full bg-gray-100 rounded-full h-2.5">
        <div class="bg-brand-accent h-2.5 rounded-full" style="width: 72%;"></div>
    </div>
</div>
```

### Pattern 8: HBF + Top-Bottom Split

Same header/footer as Pattern 3. Body splits vertically into two sections:

```html
<div class="flex-1 px-16 py-6 flex flex-col gap-6">
    <!-- Top Section: Content (e.g., 2-column comparison) -->
    <div class="flex-1 flex gap-8">
        <div class="w-1/2 flex flex-col">
            <div class="flex items-center mb-3">
                <span class="bg-gray-200 text-gray-600 text-xs font-bold px-3 py-1 rounded uppercase">Current</span>
                <h2 class="text-lg font-bold text-brand-dark ml-2">現状の課題</h2>
            </div>
            <div class="flex-1 bg-gray-50 rounded-xl p-5 border border-gray-100">
                <!-- Challenge items -->
            </div>
        </div>
        <div class="w-1/2 flex flex-col">
            <div class="flex items-center mb-3">
                <span class="bg-brand-accent text-white text-xs font-bold px-3 py-1 rounded uppercase">Future</span>
                <h2 class="text-lg font-bold text-brand-dark ml-2">解決策</h2>
            </div>
            <div class="flex-1 bg-blue-50 rounded-xl p-5 border border-blue-100">
                <!-- Solution items -->
            </div>
        </div>
    </div>

    <!-- Bottom Section: KPI/Summary Bar (dark) -->
    <div class="bg-brand-dark rounded-xl p-6 flex items-center">
        <div class="w-1/4 border-r border-gray-700 pr-6">
            <h3 class="text-brand-accent font-bold text-sm">Expected Results</h3>
        </div>
        <div class="w-3/4 flex justify-around items-center pl-6">
            <div class="text-center text-white">
                <p class="text-2xl font-black font-accent">40<span class="text-sm font-normal ml-1">%</span></p>
                <p class="text-xs opacity-70">コスト削減</p>
            </div>
            <!-- More KPI items -->
        </div>
    </div>
</div>
```

### Pattern 9: HBF + Timeline/Roadmap

Same header/footer as Pattern 3. Body has a horizontal timeline bar with phase cards:

```html
<div class="flex-1 px-16 py-6 flex flex-col">
    <!-- Timeline Bar -->
    <div class="relative mb-8">
        <div class="absolute top-5 left-0 right-0 h-1 bg-gray-200 rounded-full"></div>
        <div class="grid grid-cols-4 relative z-10">
            <div class="flex flex-col items-center">
                <div class="w-10 h-10 rounded-full bg-brand-accent flex items-center justify-center text-white text-sm font-bold shadow-lg border-4 border-white">
                    <i class="fas fa-flag"></i>
                </div>
                <p class="text-xs text-brand-accent font-bold mt-2 font-accent">Q1</p>
            </div>
            <!-- Q2, Q3, Q4 with different colors -->
        </div>
    </div>

    <!-- Phase Cards -->
    <div class="grid grid-cols-4 gap-4 flex-1">
        <div class="bg-white rounded-xl border border-gray-100 shadow-sm p-5 flex flex-col border-t-4 border-brand-accent">
            <div class="flex items-center space-x-3 mb-3">
                <div class="w-10 h-10 bg-brand-accent rounded-full flex items-center justify-center flex-shrink-0">
                    <i class="fas fa-flask text-white text-sm"></i>
                </div>
                <div>
                    <span class="inline-block text-xs font-bold px-2 py-0.5 rounded bg-green-50 text-brand-accent">Phase 1</span>
                    <p class="text-xs text-gray-400 mt-0.5">Month 1-3</p>
                </div>
            </div>
            <h3 class="text-base font-bold text-brand-dark mb-2">MVP開発</h3>
            <ul class="space-y-2 flex-1">
                <li class="flex items-start space-x-2">
                    <i class="fas fa-check-circle text-brand-accent text-xs mt-1 flex-shrink-0"></i>
                    <span class="text-xs text-gray-600">プロトタイプ開発</span>
                </li>
            </ul>
            <div class="mt-3 pt-3 border-t border-gray-100">
                <p class="text-xs text-gray-400"><i class="fas fa-users mr-1"></i>4名体制</p>
            </div>
        </div>
        <!-- More phase cards -->
    </div>
</div>
```

### Pattern 10: HBF + KPI Dashboard

Same header/footer as Pattern 3. Body has KPI cards grid at top + visualization area below:

```html
<div class="flex-1 px-16 py-6 flex flex-col gap-5">
    <!-- KPI Cards Row -->
    <div class="grid grid-cols-4 gap-4">
        <div class="bg-white border border-gray-200 rounded-xl p-5 shadow-sm relative overflow-hidden">
            <div class="absolute left-0 top-0 bottom-0 w-1 bg-brand-accent"></div>
            <div class="flex items-center justify-between">
                <div>
                    <p class="text-xs text-gray-500 font-bold uppercase tracking-wider mb-1">Revenue</p>
                    <p class="text-2xl font-black text-brand-dark font-accent">¥48<span class="text-sm font-normal text-gray-500 ml-1">億</span></p>
                </div>
                <div class="w-10 h-10 rounded-full bg-blue-50 flex items-center justify-center text-brand-accent">
                    <i class="fas fa-chart-bar"></i>
                </div>
            </div>
            <p class="text-xs text-green-500 mt-2"><i class="fas fa-arrow-up mr-1"></i>+24% YoY</p>
        </div>
        <!-- 3 more KPI cards -->
    </div>

    <!-- Full-Width Chart/Progress Area -->
    <div class="flex-1 flex gap-6">
        <div class="w-1/2 bg-white border border-gray-200 rounded-xl p-5 shadow-sm">
            <h3 class="text-sm font-bold text-brand-accent uppercase font-accent mb-4">Phase Progress</h3>
            <div class="space-y-4">
                <div>
                    <div class="flex items-center justify-between mb-1">
                        <p class="text-sm font-bold text-brand-dark">Phase 1: 要件定義</p>
                        <p class="text-sm font-bold text-brand-accent font-accent">100%</p>
                    </div>
                    <div class="w-full bg-gray-100 rounded-full h-2.5">
                        <div class="bg-brand-accent h-2.5 rounded-full" style="width: 100%;"></div>
                    </div>
                </div>
            </div>
        </div>
        <div class="w-1/2 bg-white border border-gray-200 rounded-xl p-5 shadow-sm">
            <!-- Effect Metrics / Bar Charts -->
        </div>
    </div>
</div>
```

### Pattern 11: HBF + Grid Table

Same header/footer as Pattern 3. Body is a flex-based comparison table:

```html
<div class="flex-1 px-16 py-6">
    <div class="bg-white rounded-xl overflow-hidden shadow-sm border border-gray-200 h-full flex flex-col">
        <!-- Header Row -->
        <div class="flex bg-brand-dark text-white text-xs font-bold">
            <div class="w-1/5 px-5 py-3">項目</div>
            <div class="w-1/5 px-5 py-3 text-center">自社</div>
            <div class="w-1/5 px-5 py-3 text-center">競合A</div>
            <div class="w-1/5 px-5 py-3 text-center">競合B</div>
            <div class="w-1/5 px-5 py-3 text-center">競合C</div>
        </div>
        <!-- Data Rows -->
        <div class="flex text-sm border-b border-gray-100">
            <div class="w-1/5 px-5 py-3 font-bold text-brand-dark bg-gray-50">
                <div class="flex items-center">
                    <i class="fas fa-tags mr-2 text-brand-accent opacity-70"></i>
                    <p>価格</p>
                </div>
            </div>
            <div class="w-1/5 px-5 py-3 text-center bg-blue-50 bg-opacity-30 border-l-4 border-brand-accent">
                <i class="fas fa-check-circle text-brand-accent text-lg mb-1"></i>
                <p class="font-bold text-brand-dark text-xs">コスト効率</p>
            </div>
            <div class="w-1/5 px-5 py-3 text-center">
                <i class="far fa-circle text-gray-400 text-lg mb-1"></i>
                <p class="text-gray-600 text-xs">標準的</p>
            </div>
            <div class="w-1/5 px-5 py-3 text-center">
                <i class="fas fa-times-circle text-red-400 text-lg mb-1"></i>
                <p class="text-gray-600 text-xs">高コスト</p>
            </div>
            <div class="w-1/5 px-5 py-3 text-center">
                <i class="far fa-circle text-gray-400 text-lg mb-1"></i>
                <p class="text-gray-600 text-xs">標準的</p>
            </div>
        </div>
        <!-- More rows -->
    </div>
</div>
```

### Pattern 12: HBF + Funnel

Same header/footer as Pattern 3. Body has progressively narrowing bars:

```html
<div class="flex-1 px-16 py-6 flex flex-col items-center gap-2">
    <!-- Level 1 (100% width) -->
    <div class="flex items-center w-full" style="max-width: 900px;">
        <div class="bg-brand-accent rounded-lg py-3 px-6 flex items-center justify-between" style="width: 100%;">
            <div class="flex items-center gap-4">
                <i class="fas fa-bullhorn text-white text-lg"></i>
                <div>
                    <p class="text-sm font-bold text-white">認知</p>
                    <p class="text-xs text-white opacity-80">コンテンツ / SNS / 広告</p>
                </div>
            </div>
            <p class="text-lg font-bold text-white font-accent">10,000 <span class="text-xs font-normal">PV/月</span></p>
        </div>
        <p class="ml-4 text-xs text-gray-500 font-accent w-16 text-right">100%</p>
    </div>

    <!-- Arrow + Conversion Rate -->
    <div class="flex items-center justify-center">
        <i class="fas fa-chevron-down text-gray-500 text-xs"></i>
        <span class="text-xs text-gray-500 font-accent ml-2">12.0%</span>
    </div>

    <!-- Level 2 (78% width) -->
    <div class="flex items-center w-full" style="max-width: 900px;">
        <div class="flex justify-center" style="width: 100%;">
            <div class="bg-brand-warm rounded-lg py-3 px-6 flex items-center justify-between" style="width: 78%;">
                <!-- Content -->
            </div>
        </div>
        <p class="ml-4 text-xs text-gray-500 font-accent w-16 text-right">60%</p>
    </div>

    <!-- Continue: 56%, 36%, 20% widths with different colors -->
</div>
```

### Pattern 13: HBF + Vertical Stack (Architecture/Layers)

Same header/footer as Pattern 3. Body stacks full-width layer cards with separators:

```html
<div class="flex-1 px-16 py-6 flex flex-col gap-2">
    <!-- Layer 1: Frontend -->
    <div class="bg-gray-50 rounded-xl border border-gray-200 p-4 flex items-center">
        <div class="w-28 flex-shrink-0">
            <p class="text-xs text-brand-accent font-bold font-accent uppercase tracking-wider">Frontend</p>
            <p class="text-sm font-bold text-brand-dark">クライアント層</p>
        </div>
        <div class="flex-1 flex gap-3">
            <div class="flex-1 bg-white rounded-lg p-3 flex items-center border border-gray-100">
                <i class="fas fa-desktop text-brand-accent mr-3"></i>
                <div>
                    <p class="text-xs font-bold text-brand-dark">Web Dashboard</p>
                    <p class="text-xs text-gray-500">React / Next.js</p>
                </div>
            </div>
            <!-- More tech items -->
        </div>
    </div>

    <!-- Chevron Separator -->
    <div class="flex items-center justify-center">
        <i class="fas fa-chevron-down text-gray-300"></i>
    </div>

    <!-- Layer 2: API (highlighted with accent border) -->
    <div class="bg-blue-50 rounded-xl border-2 border-brand-accent p-4 flex items-center">
        <div class="w-28 flex-shrink-0">
            <p class="text-xs text-brand-accent font-bold font-accent uppercase tracking-wider">Core</p>
            <p class="text-sm font-bold text-brand-dark">サービス層</p>
        </div>
        <div class="flex-1 grid grid-cols-4 gap-3">
            <div class="bg-white rounded-lg p-3 text-center border border-gray-100">
                <i class="fas fa-database text-brand-accent text-lg mb-1"></i>
                <p class="text-xs font-bold text-brand-dark">Data Hub</p>
            </div>
            <!-- More services -->
        </div>
    </div>

    <!-- Chevron Separator -->
    <div class="flex items-center justify-center">
        <i class="fas fa-chevron-down text-gray-300"></i>
    </div>

    <!-- Layer 3: Infrastructure -->
    <div class="bg-gray-50 rounded-xl border border-gray-200 p-4 flex items-center">
        <!-- Similar structure -->
    </div>
</div>
```

### Pattern 14: HBF + 2x2 Grid

Same header/footer as Pattern 3. Body has 2-row x 2-column card grid:

```html
<div class="flex-1 px-16 py-6 flex flex-col gap-5">
    <!-- 2x2 Grid -->
    <div class="flex-1 grid grid-cols-2 gap-4">
        <!-- Card 1 (e.g., Risk: Market) -->
        <div class="bg-white rounded-xl border border-gray-200 p-5 shadow-sm flex flex-col relative">
            <div class="absolute top-3 right-3">
                <span class="text-xs font-bold text-red-600 bg-red-100 px-2 py-0.5 rounded">高</span>
            </div>
            <div class="flex items-center gap-3 mb-3">
                <div class="w-10 h-10 rounded-lg bg-red-50 flex items-center justify-center flex-shrink-0">
                    <i class="fas fa-exclamation-triangle text-red-500"></i>
                </div>
                <div>
                    <p class="text-sm font-bold text-brand-dark">市場リスク</p>
                    <p class="text-xs text-gray-500">市場成長鈍化・規制変更</p>
                </div>
            </div>
            <div class="bg-green-50 rounded-lg p-3 mt-auto">
                <div class="flex items-center space-x-1.5 mb-1">
                    <i class="fas fa-shield-alt text-brand-accent text-xs"></i>
                    <p class="text-xs font-bold text-brand-accent">軽減策</p>
                </div>
                <p class="text-xs text-gray-600">マルチ業界対応、柔軟な設計</p>
            </div>
        </div>
        <!-- Cards 2, 3, 4 with different colors (yellow/blue/green badges) -->
    </div>

    <!-- Optional: Summary Bar -->
    <div class="bg-gray-50 rounded-xl p-4 flex items-center justify-between border border-gray-200">
        <div class="flex items-center gap-3">
            <i class="fas fa-shield-alt text-brand-accent"></i>
            <p class="text-xs font-bold text-brand-dark">総合リスク評価</p>
        </div>
        <div class="flex items-center gap-4">
            <div class="flex items-center gap-2">
                <span class="w-2 h-2 rounded-full bg-red-500"></span>
                <span class="text-xs text-gray-500">高: 1</span>
            </div>
            <div class="flex items-center gap-2">
                <span class="w-2 h-2 rounded-full bg-yellow-500"></span>
                <span class="text-xs text-gray-500">中: 2</span>
            </div>
            <div class="flex items-center gap-2">
                <span class="w-2 h-2 rounded-full bg-green-500"></span>
                <span class="text-xs text-gray-500">低: 1</span>
            </div>
        </div>
    </div>
</div>
```

### Pattern 15: HBF + Stacked Cards (Q&A / Numbered List)

Same header/footer as Pattern 3. Body has vertically stacked full-width cards with numbered badges:

```html
<div class="flex-1 px-16 py-6 flex flex-col gap-4">
    <!-- Card 1 -->
    <div class="bg-white rounded-xl p-5 border border-gray-200 shadow-sm">
        <div class="flex items-start">
            <div class="flex-shrink-0 w-8 h-8 rounded-full bg-brand-accent flex items-center justify-center text-white text-xs font-bold mr-4 font-accent">Q1</div>
            <div class="flex-1">
                <p class="text-sm font-bold text-brand-dark mb-2">質問やポイントのタイトル</p>
                <p class="text-xs text-gray-500 leading-relaxed">回答や説明テキスト。複数行にわたる場合も leading-relaxed で読みやすく保つ。</p>
            </div>
        </div>
    </div>

    <!-- Card 2 -->
    <div class="bg-white rounded-xl p-5 border border-gray-200 shadow-sm">
        <div class="flex items-start">
            <div class="flex-shrink-0 w-8 h-8 rounded-full bg-brand-accent flex items-center justify-center text-white text-xs font-bold mr-4 font-accent">Q2</div>
            <div class="flex-1">
                <p class="text-sm font-bold text-brand-dark mb-2">2番目の質問やポイント</p>
                <p class="text-xs text-gray-500 leading-relaxed">回答テキスト</p>
            </div>
        </div>
    </div>

    <!-- Card 3 (alternate badge color for variety) -->
    <div class="bg-white rounded-xl p-5 border border-gray-200 shadow-sm">
        <div class="flex items-start">
            <div class="flex-shrink-0 w-8 h-8 rounded-full bg-brand-sub flex items-center justify-center text-white text-xs font-bold mr-4 font-accent">Q3</div>
            <div class="flex-1">
                <p class="text-sm font-bold text-brand-dark mb-2">3番目の質問やポイント</p>
                <p class="text-xs text-gray-500 leading-relaxed">回答テキスト</p>
            </div>
        </div>
    </div>

    <!-- Card 4 -->
    <div class="bg-white rounded-xl p-5 border border-gray-200 shadow-sm">
        <div class="flex items-start">
            <div class="flex-shrink-0 w-8 h-8 rounded-full bg-brand-sub flex items-center justify-center text-white text-xs font-bold mr-4 font-accent">Q4</div>
            <div class="flex-1">
                <p class="text-sm font-bold text-brand-dark mb-2">4番目の質問やポイント</p>
                <p class="text-xs text-gray-500 leading-relaxed">回答テキスト</p>
            </div>
        </div>
    </div>
</div>
```

Badge variations: Use `Q1`/`Q2` for Q&A, `01`/`02` for numbered points, or icons (`<i class="fas fa-lightbulb">`) for key insights. Cards 1-2 use `bg-brand-accent`, cards 3-4 use `bg-brand-sub` for visual rhythm. Recommended 4-5 cards max.

### Pattern 16: HBF + TAM/SAM/SOM (Market Size)

**Variant 16a**: Left descriptions + right nested circles.

Same header/footer as Pattern 3. Body:

```html
<div class="flex-1 px-16 py-8 flex gap-8">
    <!-- Left: Descriptions -->
    <div class="w-1/3 flex flex-col justify-between">
        <div class="space-y-6">
            <div class="relative pl-6 border-l-4 border-gray-200">
                <h3 class="text-sm font-bold text-gray-500 uppercase mb-1 font-accent">TAM (Total Addressable Market)</h3>
                <p class="text-2xl font-bold text-gray-800 font-accent">250 <span class="text-sm font-normal text-gray-500">億円/年</span></p>
                <p class="text-xs text-gray-600 mt-2 leading-relaxed">市場全体の理論上の最大規模。</p>
            </div>
            <div class="relative pl-6 border-l-4 border-brand-accent">
                <h3 class="text-sm font-bold text-gray-500 uppercase mb-1 font-accent">SAM (Serviceable Available Market)</h3>
                <p class="text-2xl font-bold text-gray-800 font-accent">50 <span class="text-sm font-normal text-gray-500">億円/年</span></p>
                <p class="text-xs text-gray-600 mt-2 leading-relaxed">ターゲットに提供可能な市場。</p>
            </div>
            <div class="relative pl-6 border-l-4 border-brand-warm">
                <h3 class="text-sm font-bold text-brand-accent uppercase mb-1 font-accent">SOM (Serviceable Obtainable Market)</h3>
                <p class="text-3xl font-bold text-brand-accent font-accent">2.5 <span class="text-lg font-normal text-brand-warm">億円/年</span></p>
                <p class="text-xs text-gray-600 mt-2 leading-relaxed">現実的に獲得可能な市場規模。</p>
            </div>
        </div>
        <div class="bg-white p-4 rounded-xl border border-gray-200 shadow-sm mt-6">
            <div class="flex items-center justify-between mb-2">
                <span class="text-xs font-bold text-gray-500 uppercase font-accent">Market Trend</span>
                <span class="bg-green-100 text-brand-accent text-xs font-bold px-2 py-0.5 rounded-full flex items-center">
                    <i class="fas fa-arrow-trend-up mr-1"></i> CAGR +5%
                </span>
            </div>
            <p class="text-xs text-gray-600 leading-snug">トレンド説明テキスト。</p>
        </div>
    </div>
    <!-- Right: Nested Circles -->
    <div class="w-2/3 flex items-center justify-center relative">
        <div class="relative flex items-center justify-center" style="width: 400px; height: 400px;">
            <div class="absolute rounded-full flex flex-col items-center justify-center bg-blue-100 border-2 border-blue-200" style="width: 360px; height: 360px; padding-top: 30px;">
                <p class="text-brand-dark font-bold">TAM</p>
                <p class="text-2xl font-black text-brand-dark font-accent">¥250億</p>
            </div>
            <div class="absolute rounded-full flex flex-col items-center justify-center bg-blue-300 border-2 border-blue-400 shadow-md" style="width: 240px; height: 240px; padding-top: 30px;">
                <p class="text-brand-dark font-bold text-sm">SAM</p>
                <p class="text-xl font-black text-brand-dark font-accent">¥50億</p>
            </div>
            <div class="absolute rounded-full flex flex-col items-center justify-center bg-brand-accent text-white shadow-lg" style="width: 120px; height: 120px;">
                <p class="text-xs font-bold opacity-90">SOM</p>
                <p class="text-lg font-black font-accent">¥2.5億</p>
            </div>
        </div>
    </div>
</div>
```

**Variant 16b**: Left descriptions + right horizontal bars.

Same header/footer as Pattern 3. Body:

```html
<div class="flex-1 px-16 py-8 flex gap-8">
    <div class="w-1/2 flex flex-col justify-between">
        <div class="space-y-6">
            <div class="relative pl-6 border-l-4 border-gray-200">
                <h3 class="text-sm font-bold text-gray-500 uppercase mb-1 font-accent">TAM</h3>
                <p class="text-2xl font-bold text-gray-800 font-accent">500 <span class="text-sm font-normal text-gray-500">億円/年</span></p>
                <p class="text-xs text-gray-600 mt-2 leading-relaxed">説明テキスト。</p>
            </div>
            <div class="relative pl-6 border-l-4 border-brand-accent">
                <h3 class="text-sm font-bold text-gray-500 uppercase mb-1 font-accent">SAM</h3>
                <p class="text-2xl font-bold text-gray-800 font-accent">80 <span class="text-sm font-normal text-gray-500">億円/年</span></p>
                <p class="text-xs text-gray-600 mt-2 leading-relaxed">説明テキスト。</p>
            </div>
            <div class="relative pl-6 border-l-4 border-brand-warm">
                <h3 class="text-sm font-bold text-brand-accent uppercase mb-1 font-accent">SOM (TARGET)</h3>
                <p class="text-3xl font-bold text-brand-accent font-accent">5 <span class="text-lg font-normal text-brand-warm">億円/年</span></p>
                <p class="text-xs text-gray-600 mt-2 leading-relaxed">説明テキスト。</p>
            </div>
        </div>
        <div class="bg-white p-4 rounded-xl border border-gray-200 shadow-sm mt-6">
            <div class="flex items-center justify-between mb-2">
                <span class="text-xs font-bold text-gray-500 uppercase font-accent">Market Trend</span>
                <span class="bg-green-100 text-brand-accent text-xs font-bold px-2 py-0.5 rounded-full flex items-center">
                    <i class="fas fa-arrow-trend-up mr-1"></i> CAGR +15%
                </span>
            </div>
            <p class="text-xs text-gray-600 leading-snug">トレンド説明テキスト。</p>
        </div>
    </div>
    <div class="w-1/2 flex flex-col justify-center gap-8">
        <div>
            <div class="flex items-end justify-between mb-2">
                <p class="text-sm font-bold text-gray-800 font-accent">TAM</p>
                <p class="text-2xl font-black text-gray-800 font-accent">¥500億</p>
            </div>
            <div class="w-full h-12 rounded-lg flex items-center justify-center bg-gray-100">
                <p class="text-sm font-bold text-brand-dark">ラベル</p>
            </div>
        </div>
        <div>
            <div class="flex items-end justify-between mb-2">
                <p class="text-sm font-bold text-gray-800 font-accent">SAM</p>
                <p class="text-2xl font-black text-gray-800 font-accent">¥80億</p>
            </div>
            <div class="w-3/5 h-12 rounded-lg flex items-center justify-center bg-brand-accent text-white">
                <p class="text-sm font-bold">ラベル</p>
            </div>
        </div>
        <div>
            <div class="flex items-end justify-between mb-2">
                <p class="text-sm font-bold text-brand-accent font-accent">SOM (TARGET)</p>
                <p class="text-3xl font-black text-brand-accent font-accent">¥5億</p>
            </div>
            <div class="w-1/4 h-12 rounded-lg flex items-center justify-center bg-brand-dark text-white shadow-lg">
                <p class="text-sm font-bold">ラベル</p>
            </div>
        </div>
    </div>
</div>
```

### Pattern 17: Chapter Divider

Full-height split (no HBF). Left = chapter number (25%), right = title + content (75%).

```html
<div class="slide flex relative overflow-hidden">
    <div class="w-1/4 h-full bg-brand-dark flex flex-col justify-center items-center relative overflow-hidden">
        <div class="absolute inset-0 opacity-10">
            <svg height="100%" width="100%" xmlns="http://www.w3.org/2000/svg">
                <defs>
                    <pattern id="grid" width="40" height="40" patternUnits="userSpaceOnUse">
                        <path d="M 40 0 L 0 0 0 40" fill="none" stroke="white" stroke-width="1"></path>
                    </pattern>
                </defs>
                <rect width="100%" height="100%" fill="url(#grid)"></rect>
            </svg>
        </div>
        <div class="z-10 flex flex-col items-center">
            <p class="text-brand-accent text-lg font-accent uppercase mb-2" style="letter-spacing: 0.3em;">Chapter</p>
            <h2 class="font-accent font-light text-white leading-none" style="font-size: 8rem; line-height: 1;">01</h2>
            <div class="w-12 h-1 bg-brand-accent mt-6"></div>
        </div>
        <div class="absolute bottom-12 left-1/2 transform -translate-x-1/2 text-gray-500 text-xs tracking-widest uppercase opacity-50 font-accent">
            Section Label
        </div>
    </div>
    <div class="w-3/4 h-full flex flex-col justify-center px-24 relative">
        <div class="absolute top-12 right-16 flex items-center space-x-2 text-gray-400 text-sm tracking-wide">
            <span class="uppercase font-accent">Project</span>
            <i class="fas fa-chevron-right text-xs"></i>
            <span class="text-brand-accent font-bold uppercase">Section 01</span>
        </div>
        <div class="mb-12">
            <h1 class="text-6xl font-bold text-brand-dark tracking-tight leading-tight mb-6">セクションタイトル</h1>
            <p class="text-3xl text-brand-dark font-light border-l-4 border-brand-accent pl-6 py-1">サブタイトル</p>
        </div>
        <div class="max-w-3xl">
            <div class="flex items-start">
                <i class="fas fa-quote-left text-3xl text-gray-200 mr-4 -mt-2"></i>
                <p class="text-xl text-gray-600 leading-relaxed font-medium">
                    本文テキスト。セクションの概要やリード文を配置。
                </p>
            </div>
        </div>
        <div class="absolute bottom-12 right-16 text-right">
            <div class="flex items-center justify-end space-x-3 text-brand-dark mb-1">
                <i class="fas fa-{brand-icon} text-lg text-brand-accent"></i>
                <p class="text-xs font-bold tracking-widest uppercase font-accent">BRAND</p>
            </div>
        </div>
    </div>
</div>
```

### Pattern 18: HBF + Contact

Same header as Pattern 3 (no footer). Body: left = message + CTA, right = contact card.

```html
<div class="slide flex flex-col">
    <!-- Header (same as Pattern 3) -->
    <div class="px-16 pt-10 pb-4 flex justify-between items-end border-b border-gray-200 mx-16">
        <div class="flex items-center space-x-4">
            <div class="w-1.5 h-10 bg-brand-accent"></div>
            <div>
                <p class="text-xs text-gray-400 font-accent tracking-widest uppercase mb-1">Contact</p>
                <h1 class="text-3xl font-bold text-brand-dark tracking-tight">お問い合わせ</h1>
            </div>
        </div>
        <div class="flex items-center space-x-2 text-brand-dark opacity-50">
            <i class="fas fa-{brand-icon} text-lg"></i>
            <p class="text-xs font-bold tracking-widest uppercase font-accent">BRAND</p>
        </div>
    </div>

    <!-- Body -->
    <div class="flex-1 px-16 py-8 flex gap-12 items-center">
        <!-- Left: Message + CTA -->
        <div class="w-1/2 flex flex-col justify-center">
            <h2 class="text-4xl font-bold text-brand-dark leading-tight mb-6">ご興味をお持ちいただき<br />ありがとうございます</h2>
            <p class="text-lg text-gray-600 leading-relaxed mb-8">
                ご不明な点やご要望がございましたら、お気軽にお問い合わせください。担当者が丁寧にご対応いたします。
            </p>
            <div class="flex gap-4">
                <div class="bg-brand-accent text-white px-8 py-3 rounded-lg font-bold text-sm flex items-center shadow-md">
                    <i class="fas fa-envelope mr-2"></i> メールで問い合わせ
                </div>
                <div class="border-2 border-brand-dark text-brand-dark px-8 py-3 rounded-lg font-bold text-sm flex items-center">
                    <i class="fas fa-calendar-alt mr-2"></i> 打ち合わせ予約
                </div>
            </div>
        </div>

        <!-- Right: Contact Card -->
        <div class="w-1/2 flex justify-center">
            <div class="bg-white rounded-2xl shadow-lg border border-gray-200 p-8 w-full max-w-md">
                <div class="flex items-center mb-6">
                    <div class="w-16 h-16 rounded-full bg-brand-dark flex items-center justify-center text-white text-2xl font-bold mr-4">
                        <i class="fas fa-user"></i>
                    </div>
                    <div>
                        <p class="text-lg font-bold text-brand-dark">担当者名</p>
                        <p class="text-sm text-gray-500">役職・部署名</p>
                    </div>
                </div>
                <div class="space-y-4 border-t border-gray-100 pt-6">
                    <div class="flex items-center">
                        <div class="w-10 h-10 rounded-lg bg-blue-50 flex items-center justify-center text-brand-accent mr-4 flex-shrink-0">
                            <i class="fas fa-envelope"></i>
                        </div>
                        <div>
                            <p class="text-xs text-gray-400 uppercase font-accent">Email</p>
                            <p class="text-sm text-brand-dark font-bold">contact@example.com</p>
                        </div>
                    </div>
                    <div class="flex items-center">
                        <div class="w-10 h-10 rounded-lg bg-blue-50 flex items-center justify-center text-brand-accent mr-4 flex-shrink-0">
                            <i class="fas fa-phone"></i>
                        </div>
                        <div>
                            <p class="text-xs text-gray-400 uppercase font-accent">Phone</p>
                            <p class="text-sm text-brand-dark font-bold">03-XXXX-XXXX</p>
                        </div>
                    </div>
                    <div class="flex items-center">
                        <div class="w-10 h-10 rounded-lg bg-blue-50 flex items-center justify-center text-brand-accent mr-4 flex-shrink-0">
                            <i class="fas fa-globe"></i>
                        </div>
                        <div>
                            <p class="text-xs text-gray-400 uppercase font-accent">Website</p>
                            <p class="text-sm text-brand-dark font-bold">https://example.com</p>
                        </div>
                    </div>
                    <div class="flex items-center">
                        <div class="w-10 h-10 rounded-lg bg-blue-50 flex items-center justify-center text-brand-accent mr-4 flex-shrink-0">
                            <i class="fas fa-map-marker-alt"></i>
                        </div>
                        <div>
                            <p class="text-xs text-gray-400 uppercase font-accent">Address</p>
                            <p class="text-sm text-brand-dark font-bold">東京都渋谷区XXX</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>
```

### Pattern 19: HBF + 5-Column Process Flow

Same header/footer as Pattern 3. Body uses `grid-cols-5`.

```html
<div class="flex-1 px-16 py-6 flex flex-col justify-center bg-gray-50">
    <div class="mb-6 mx-4">
        <p class="text-gray-600 text-sm">導入プロセスの説明テキスト。各フェーズの役割分担など。</p>
    </div>
    <div class="grid grid-cols-5 gap-4 relative">
        <!-- Step 1 -->
        <div class="bg-white rounded-lg shadow-sm border-t-4 border-brand-accent p-5 flex flex-col h-full relative">
            <div class="flex justify-between items-start mb-3">
                <span class="text-4xl font-accent font-bold text-gray-100 absolute top-2 right-4 -z-10">01</span>
                <div class="w-10 h-10 rounded-full bg-{light} flex items-center justify-center text-brand-accent mb-2">
                    <i class="fas fa-search"></i>
                </div>
                <span class="text-xs font-bold bg-brand-accent text-white px-2 py-0.5 rounded">Week 1-2</span>
            </div>
            <h3 class="text-lg font-bold text-brand-dark mb-2">現状調査</h3>
            <p class="text-xs text-gray-500 mb-4 flex-grow leading-relaxed">現状の業務フロー分析とヒアリングを実施</p>
            <div class="mt-auto pt-3 border-t border-gray-100">
                <p class="text-xs text-gray-400"><i class="fas fa-users mr-1"></i>2名体制</p>
            </div>
            <div class="absolute top-1/2 -right-3 transform -translate-y-1/2 z-20 text-gray-300 text-xl">
                <i class="fas fa-chevron-right"></i>
            </div>
        </div>
        <!-- Step 2 -->
        <div class="bg-white rounded-lg shadow-sm border-t-4 border-brand-accent p-5 flex flex-col h-full relative">
            <div class="flex justify-between items-start mb-3">
                <span class="text-4xl font-accent font-bold text-gray-100 absolute top-2 right-4 -z-10">02</span>
                <div class="w-10 h-10 rounded-full bg-{light} flex items-center justify-center text-brand-accent mb-2">
                    <i class="fas fa-pencil-alt"></i>
                </div>
                <span class="text-xs font-bold bg-brand-accent text-white px-2 py-0.5 rounded">Week 3-4</span>
            </div>
            <h3 class="text-lg font-bold text-brand-dark mb-2">要件定義</h3>
            <p class="text-xs text-gray-500 mb-4 flex-grow leading-relaxed">要件の整理と優先順位付けを実施</p>
            <div class="mt-auto pt-3 border-t border-gray-100">
                <p class="text-xs text-gray-400"><i class="fas fa-users mr-1"></i>3名体制</p>
            </div>
            <div class="absolute top-1/2 -right-3 transform -translate-y-1/2 z-20 text-gray-300 text-xl">
                <i class="fas fa-chevron-right"></i>
            </div>
        </div>
        <!-- Step 3 -->
        <div class="bg-white rounded-lg shadow-sm border-t-4 border-brand-dark p-5 flex flex-col h-full relative">
            <div class="flex justify-between items-start mb-3">
                <span class="text-4xl font-accent font-bold text-gray-100 absolute top-2 right-4 -z-10">03</span>
                <div class="w-10 h-10 rounded-full bg-{light} flex items-center justify-center text-brand-dark mb-2">
                    <i class="fas fa-cogs"></i>
                </div>
                <span class="text-xs font-bold bg-brand-dark text-white px-2 py-0.5 rounded">Week 5-8</span>
            </div>
            <h3 class="text-lg font-bold text-brand-dark mb-2">開発・構築</h3>
            <p class="text-xs text-gray-500 mb-4 flex-grow leading-relaxed">システム開発と環境構築を実施</p>
            <div class="mt-auto pt-3 border-t border-gray-100">
                <p class="text-xs text-gray-400"><i class="fas fa-users mr-1"></i>5名体制</p>
            </div>
            <div class="absolute top-1/2 -right-3 transform -translate-y-1/2 z-20 text-gray-300 text-xl">
                <i class="fas fa-chevron-right"></i>
            </div>
        </div>
        <!-- Step 4 -->
        <div class="bg-white rounded-lg shadow-sm border-t-4 border-brand-dark p-5 flex flex-col h-full relative">
            <div class="flex justify-between items-start mb-3">
                <span class="text-4xl font-accent font-bold text-gray-100 absolute top-2 right-4 -z-10">04</span>
                <div class="w-10 h-10 rounded-full bg-{light} flex items-center justify-center text-brand-dark mb-2">
                    <i class="fas fa-vial"></i>
                </div>
                <span class="text-xs font-bold bg-brand-dark text-white px-2 py-0.5 rounded">Week 9-10</span>
            </div>
            <h3 class="text-lg font-bold text-brand-dark mb-2">テスト・検証</h3>
            <p class="text-xs text-gray-500 mb-4 flex-grow leading-relaxed">動作検証とユーザー受入テストを実施</p>
            <div class="mt-auto pt-3 border-t border-gray-100">
                <p class="text-xs text-gray-400"><i class="fas fa-users mr-1"></i>4名体制</p>
            </div>
            <div class="absolute top-1/2 -right-3 transform -translate-y-1/2 z-20 text-gray-300 text-xl">
                <i class="fas fa-chevron-right"></i>
            </div>
        </div>
        <!-- Step 5 (last, no arrow) -->
        <div class="bg-white rounded-lg shadow-sm border-t-4 border-brand-warm p-5 flex flex-col h-full relative ring-2 ring-yellow-50">
            <div class="flex justify-between items-start mb-3">
                <span class="text-4xl font-accent font-bold text-gray-100 absolute top-2 right-4 -z-10">05</span>
                <div class="w-10 h-10 rounded-full bg-{light} flex items-center justify-center text-brand-warm mb-2">
                    <i class="fas fa-rocket"></i>
                </div>
                <span class="text-xs font-bold bg-brand-warm text-white px-2 py-0.5 rounded">Week 11-12</span>
            </div>
            <h3 class="text-lg font-bold text-brand-dark mb-2">本番展開</h3>
            <p class="text-xs text-gray-500 mb-4 flex-grow leading-relaxed">本番環境へのデプロイと運用開始</p>
            <div class="mt-auto pt-3 border-t border-gray-100">
                <p class="text-xs text-gray-400"><i class="fas fa-users mr-1"></i>3名体制</p>
            </div>
        </div>
    </div>
</div>
```

**Optional: RACI Box** (attach below process flow for responsibility matrix):

```html
<div class="mt-6 mx-4 bg-white rounded-xl overflow-hidden shadow-sm border border-gray-200">
    <div class="flex bg-brand-dark text-white text-xs font-bold">
        <div class="w-1/5 px-4 py-2">タスク</div>
        <div class="w-1/5 px-4 py-2 text-center">PM</div>
        <div class="w-1/5 px-4 py-2 text-center">開発</div>
        <div class="w-1/5 px-4 py-2 text-center">デザイン</div>
        <div class="w-1/5 px-4 py-2 text-center">クライアント</div>
    </div>
    <div class="flex text-xs border-b border-gray-100">
        <div class="w-1/5 px-4 py-2 font-bold text-brand-dark bg-gray-50">要件定義</div>
        <div class="w-1/5 px-4 py-2 text-center"><span class="inline-block bg-brand-accent text-white font-bold px-2 py-0.5 rounded">R</span></div>
        <div class="w-1/5 px-4 py-2 text-center"><span class="inline-block bg-blue-100 text-brand-dark font-bold px-2 py-0.5 rounded">C</span></div>
        <div class="w-1/5 px-4 py-2 text-center"><span class="inline-block bg-blue-100 text-brand-dark font-bold px-2 py-0.5 rounded">C</span></div>
        <div class="w-1/5 px-4 py-2 text-center"><span class="inline-block bg-green-100 text-green-700 font-bold px-2 py-0.5 rounded">A</span></div>
    </div>
    <div class="flex text-xs border-b border-gray-100">
        <div class="w-1/5 px-4 py-2 font-bold text-brand-dark bg-gray-50">開発</div>
        <div class="w-1/5 px-4 py-2 text-center"><span class="inline-block bg-green-100 text-green-700 font-bold px-2 py-0.5 rounded">A</span></div>
        <div class="w-1/5 px-4 py-2 text-center"><span class="inline-block bg-brand-accent text-white font-bold px-2 py-0.5 rounded">R</span></div>
        <div class="w-1/5 px-4 py-2 text-center"><span class="inline-block bg-yellow-100 text-yellow-700 font-bold px-2 py-0.5 rounded">I</span></div>
        <div class="w-1/5 px-4 py-2 text-center"><span class="inline-block bg-yellow-100 text-yellow-700 font-bold px-2 py-0.5 rounded">I</span></div>
    </div>
</div>
```

RACI legend: **R** = Responsible (bg-brand-accent), **A** = Accountable (bg-green-100), **C** = Consulted (bg-blue-100), **I** = Informed (bg-yellow-100).

### Pattern 20: HBF + VS Competitor Comparison

Same header/footer as Pattern 3. Body: 2 columns with central VS badge.

```html
<div class="flex-1 px-16 py-6 flex gap-6 items-stretch relative">
    <!-- Left: Our Product (highlighted) -->
    <div class="flex-1 bg-white rounded-2xl border-2 border-brand-accent shadow-lg p-6 flex flex-col relative" style="flex: 1.5; transform: scale(1.02);">
        <div class="absolute -top-3 left-6">
            <span class="bg-brand-accent text-white text-xs font-bold px-3 py-1 rounded-full uppercase tracking-wider shadow-md">Recommended</span>
        </div>
        <div class="flex items-center mb-4 mt-2">
            <div class="w-12 h-12 rounded-xl bg-brand-accent flex items-center justify-center text-white mr-3">
                <i class="fas fa-{brand-icon} text-xl"></i>
            </div>
            <div>
                <h3 class="text-xl font-bold text-brand-dark">自社サービス</h3>
                <p class="text-xs text-gray-500">サービスカテゴリ</p>
            </div>
        </div>
        <div class="space-y-3 flex-1">
            <div class="flex items-start">
                <i class="fas fa-check-circle text-brand-accent mt-0.5 mr-2 flex-shrink-0"></i>
                <div>
                    <p class="text-sm font-bold text-brand-dark">強みポイント1</p>
                    <p class="text-xs text-gray-500">説明テキスト</p>
                </div>
            </div>
            <div class="flex items-start">
                <i class="fas fa-check-circle text-brand-accent mt-0.5 mr-2 flex-shrink-0"></i>
                <div>
                    <p class="text-sm font-bold text-brand-dark">強みポイント2</p>
                    <p class="text-xs text-gray-500">説明テキスト</p>
                </div>
            </div>
            <div class="flex items-start">
                <i class="fas fa-check-circle text-brand-accent mt-0.5 mr-2 flex-shrink-0"></i>
                <div>
                    <p class="text-sm font-bold text-brand-dark">強みポイント3</p>
                    <p class="text-xs text-gray-500">説明テキスト</p>
                </div>
            </div>
        </div>
        <div class="mt-4 pt-4 border-t border-gray-100">
            <div class="flex items-baseline justify-between">
                <p class="text-xs text-gray-400 uppercase font-accent">Price</p>
                <p class="text-2xl font-black text-brand-dark font-accent">¥50,000<span class="text-sm font-normal text-gray-500">/月</span></p>
            </div>
        </div>
    </div>

    <!-- VS Badge (center) -->
    <div class="flex items-center justify-center z-10" style="margin: 0 -20px;">
        <div class="w-14 h-14 rounded-full bg-brand-warm flex items-center justify-center text-white font-black text-lg shadow-lg border-4 border-white">
            VS
        </div>
    </div>

    <!-- Right: Competitor -->
    <div class="flex-1 bg-white rounded-2xl border border-gray-200 shadow-sm p-6 flex flex-col opacity-80">
        <div class="flex items-center mb-4 mt-2">
            <div class="w-12 h-12 rounded-xl bg-gray-200 flex items-center justify-center text-gray-500 mr-3">
                <i class="fas fa-building text-xl"></i>
            </div>
            <div>
                <h3 class="text-xl font-bold text-gray-700">競合サービス</h3>
                <p class="text-xs text-gray-500">サービスカテゴリ</p>
            </div>
        </div>
        <div class="space-y-3 flex-1">
            <div class="flex items-start">
                <i class="fas fa-minus-circle text-gray-400 mt-0.5 mr-2 flex-shrink-0"></i>
                <div>
                    <p class="text-sm font-bold text-gray-700">比較ポイント1</p>
                    <p class="text-xs text-gray-500">説明テキスト</p>
                </div>
            </div>
            <div class="flex items-start">
                <i class="fas fa-times-circle text-red-400 mt-0.5 mr-2 flex-shrink-0"></i>
                <div>
                    <p class="text-sm font-bold text-gray-700">弱みポイント2</p>
                    <p class="text-xs text-gray-500">説明テキスト</p>
                </div>
            </div>
            <div class="flex items-start">
                <i class="fas fa-minus-circle text-gray-400 mt-0.5 mr-2 flex-shrink-0"></i>
                <div>
                    <p class="text-sm font-bold text-gray-700">比較ポイント3</p>
                    <p class="text-xs text-gray-500">説明テキスト</p>
                </div>
            </div>
        </div>
        <div class="mt-4 pt-4 border-t border-gray-100">
            <div class="flex items-baseline justify-between">
                <p class="text-xs text-gray-400 uppercase font-accent">Price</p>
                <p class="text-2xl font-black text-gray-700 font-accent">¥80,000<span class="text-sm font-normal text-gray-500">/月</span></p>
            </div>
        </div>
    </div>
</div>
```

---

## Additional Component Snippets

### Severity/Status Badge

```html
<span class="inline-block text-xs font-bold px-2 py-0.5 rounded bg-red-100 text-red-600">高</span>
<span class="inline-block text-xs font-bold px-2 py-0.5 rounded bg-yellow-100 text-yellow-700">中</span>
<span class="inline-block text-xs font-bold px-2 py-0.5 rounded bg-green-100 text-green-700">低</span>
```

### Timeline Item (Vertical)

```html
<div class="flex items-start">
    <div class="flex-shrink-0 w-9 h-9 rounded-full bg-brand-accent flex items-center justify-center text-white text-xs font-bold mr-3 font-accent">Q1</div>
    <div class="border-l-2 border-brand-accent pl-4 pb-4">
        <p class="text-sm font-bold text-brand-dark">タスク名</p>
        <p class="text-xs text-gray-500">説明テキスト</p>
    </div>
</div>
```

### Metric Pill Badge

```html
<span class="inline-flex items-center gap-1 bg-brand-accent bg-opacity-20 text-brand-accent text-xs font-bold px-3 py-1 rounded-full">
    <i class="fas fa-chart-pie"></i> 市場
</span>
```

### Pricing Card (Highlighted Center)

```html
<!-- Standard -->
<div class="bg-white rounded-xl border border-gray-200 p-6 flex flex-col shadow-sm">
    <h3 class="text-lg font-bold text-gray-500 mb-2">Basic</h3>
    <div class="flex items-baseline text-brand-dark mb-4">
        <span class="text-4xl font-black font-accent">¥10,000</span>
        <span class="text-gray-400 ml-2">/ 月</span>
    </div>
    <ul class="space-y-3 text-sm text-gray-600 flex-1">
        <li class="flex items-start">
            <i class="fas fa-check text-green-500 mt-1 mr-3"></i>
            <p>ユーザー数：最大5名</p>
        </li>
    </ul>
</div>

<!-- Highlighted (center card) -->
<div class="bg-white rounded-xl border-2 border-brand-accent p-6 flex flex-col shadow-lg transform scale-105 z-10 relative">
    <div class="absolute top-0 left-1/2 transform -translate-x-1/2 -translate-y-1/2 bg-brand-accent text-white px-4 py-1 rounded-full text-xs font-bold uppercase tracking-wider">
        Recommended
    </div>
    <!-- Same structure as above -->
</div>
```

### Revenue Flow (Horizontal Math)

```html
<div class="flex items-stretch gap-4">
    <div class="flex-1 bg-white rounded-xl p-5 border-l-4 border-brand-accent shadow-sm">
        <p class="text-xs text-gray-500 uppercase mb-1">Revenue</p>
        <p class="text-3xl font-black text-brand-dark">¥70,000</p>
    </div>
    <div class="flex items-center text-3xl text-gray-300">
        <i class="fas fa-minus"></i>
    </div>
    <div class="flex-1 bg-white rounded-xl p-5 border-l-4 border-red-400 shadow-sm">
        <p class="text-xs text-gray-500 uppercase mb-1">Cost</p>
        <p class="text-3xl font-black text-brand-dark">¥20,000</p>
    </div>
    <div class="flex items-center text-3xl text-gray-300">
        <i class="fas fa-equals"></i>
    </div>
    <div class="flex-1 bg-green-50 rounded-xl p-5 border-l-4 border-green-500 shadow-sm">
        <p class="text-xs text-gray-500 uppercase mb-1">Profit</p>
        <p class="text-3xl font-black text-green-700">¥50,000</p>
    </div>
</div>
```

### Positioning Map (2D Plot)

```html
<div class="flex-1 bg-gray-50 rounded-xl border border-gray-100 p-6 relative">
    <!-- Axes -->
    <div class="absolute left-6 top-6 bottom-14 w-px bg-gray-300"></div>
    <div class="absolute left-6 bottom-14 right-6 h-px bg-gray-300"></div>
    <!-- Axis Labels -->
    <div class="absolute left-1 top-1/2 transform -translate-y-1/2 -rotate-90">
        <p class="text-xs text-gray-400">Y軸ラベル →</p>
    </div>
    <div class="absolute bottom-6 left-1/2 transform -translate-x-1/2">
        <p class="text-xs text-gray-400">X軸ラベル →</p>
    </div>
    <!-- Positioned Items -->
    <div class="absolute" style="right: 60px; bottom: 100px;">
        <div class="w-12 h-12 bg-brand-accent rounded-full flex items-center justify-center text-white text-xs font-bold shadow-md">A</div>
    </div>
    <div class="absolute" style="right: 150px; bottom: 60px;">
        <div class="w-12 h-12 bg-gray-300 rounded-full flex items-center justify-center text-white text-xs font-bold">B</div>
    </div>
</div>
```

### Q&A Card (Stacked)

```html
<div class="bg-white rounded-xl p-5 border border-gray-200 shadow-sm">
    <div class="flex items-start">
        <div class="flex-shrink-0 w-8 h-8 rounded-full bg-brand-accent flex items-center justify-center text-white text-xs font-bold mr-4 font-accent">Q1</div>
        <div class="flex-1">
            <p class="text-sm font-bold text-brand-dark mb-2">質問テキスト</p>
            <p class="text-xs text-gray-500 leading-relaxed">回答テキスト</p>
        </div>
    </div>
</div>
```

---

## Chart.js Patterns (When Enabled)

The following patterns are used automatically when data visualizations require Chart.js (determined in Phase 3 slide map). Include `<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>` in `<head>` only on slides that use these patterns. Place the `<script>` for chart initialization just before `</body>`.

### Chart.js: Doughnut / Pie

```html
<!-- Inside the slide body -->
<div class="flex items-center justify-center gap-8">
    <div class="relative" style="width: 280px; height: 280px;">
        <canvas id="doughnutChart"></canvas>
    </div>
    <!-- Legend (manual for better styling control) -->
    <div class="space-y-3">
        <div class="flex items-center gap-2">
            <div class="w-3 h-3 rounded-full" style="background: #2563eb;"></div>
            <p class="text-sm text-gray-600">セグメントA <span class="font-bold font-accent">40%</span></p>
        </div>
        <div class="flex items-center gap-2">
            <div class="w-3 h-3 rounded-full" style="background: #60a5fa;"></div>
            <p class="text-sm text-gray-600">セグメントB <span class="font-bold font-accent">30%</span></p>
        </div>
        <div class="flex items-center gap-2">
            <div class="w-3 h-3 rounded-full" style="background: #93c5fd;"></div>
            <p class="text-sm text-gray-600">セグメントC <span class="font-bold font-accent">20%</span></p>
        </div>
        <div class="flex items-center gap-2">
            <div class="w-3 h-3 rounded-full" style="background: #e5e7eb;"></div>
            <p class="text-sm text-gray-600">その他 <span class="font-bold font-accent">10%</span></p>
        </div>
    </div>
</div>

<!-- Just before </body> -->
<script>
new Chart(document.getElementById('doughnutChart'), {
    type: 'doughnut',
    data: {
        labels: ['セグメントA', 'セグメントB', 'セグメントC', 'その他'],
        datasets: [{
            data: [40, 30, 20, 10],
            backgroundColor: ['#2563eb', '#60a5fa', '#93c5fd', '#e5e7eb'],
            borderWidth: 0
        }]
    },
    options: {
        responsive: true,
        maintainAspectRatio: true,
        cutout: '60%',
        plugins: {
            legend: { display: false },
            tooltip: {
                backgroundColor: 'rgba(15, 23, 42, 0.9)',
                padding: 12,
                titleFont: { family: "'Noto Sans JP', sans-serif", size: 13 },
                bodyFont: { family: "'Noto Sans JP', sans-serif", size: 12 },
                callbacks: {
                    label: function(ctx) { return ' ' + ctx.label + ': ' + ctx.raw + '%'; }
                }
            }
        }
    }
});
</script>
```

**Notes:**
- Set `legend: { display: false }` and build legend manually in HTML for PPTX-convertible labels
- Use brand palette colors from Phase 2
- `cutout: '60%'` for doughnut; omit for solid pie
- For `type: 'pie'`, remove the `cutout` option

### Chart.js: Vertical Bar

```html
<!-- Inside the slide body -->
<div style="width: 100%; height: 300px;">
    <canvas id="barChart"></canvas>
</div>

<!-- Just before </body> -->
<script>
new Chart(document.getElementById('barChart'), {
    type: 'bar',
    data: {
        labels: ['Q1', 'Q2', 'Q3', 'Q4'],
        datasets: [{
            label: '売上（億円）',
            data: [32, 48, 55, 65],
            backgroundColor: '#2563eb',
            borderRadius: 6,
            barPercentage: 0.6
        }]
    },
    options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
            y: {
                beginAtZero: true,
                grid: { color: '#f3f4f6' },
                ticks: {
                    font: { family: "'Noto Sans JP', sans-serif", size: 11 },
                    callback: function(v) { return '¥' + v + '億'; }
                }
            },
            x: {
                grid: { display: false },
                ticks: { font: { family: "'Noto Sans JP', sans-serif", size: 12 } }
            }
        },
        plugins: {
            legend: { display: false },
            tooltip: {
                backgroundColor: 'rgba(15, 23, 42, 0.9)',
                padding: 12,
                titleFont: { family: "'Noto Sans JP', sans-serif", size: 13 },
                bodyFont: { family: "'Noto Sans JP', sans-serif", size: 12 }
            }
        }
    }
});
</script>
```

**Notes:**
- `borderRadius: 6` for rounded bar tops
- Use `barPercentage: 0.6` to avoid overly wide bars
- For grouped bars, add multiple datasets with different colors
- Wrap canvas in a fixed-height container and use `maintainAspectRatio: false`

### Chart.js: Stacked Bar

```html
<!-- Inside the slide body -->
<div style="width: 100%; height: 300px;">
    <canvas id="stackedBarChart"></canvas>
</div>

<!-- Just before </body> -->
<script>
new Chart(document.getElementById('stackedBarChart'), {
    type: 'bar',
    data: {
        labels: ['2023', '2024', '2025', '2026E'],
        datasets: [
            { label: 'SaaS', data: [20, 28, 35, 42], backgroundColor: '#2563eb', borderRadius: 4 },
            { label: 'Consulting', data: [15, 18, 20, 22], backgroundColor: '#60a5fa', borderRadius: 4 },
            { label: 'Other', data: [5, 6, 7, 8], backgroundColor: '#93c5fd', borderRadius: 4 }
        ]
    },
    options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
            x: { stacked: true, grid: { display: false } },
            y: { stacked: true, beginAtZero: true, grid: { color: '#f3f4f6' } }
        },
        plugins: {
            legend: {
                position: 'bottom',
                labels: {
                    font: { family: "'Noto Sans JP', sans-serif", size: 11 },
                    usePointStyle: true,
                    pointStyle: 'circle',
                    padding: 20
                }
            }
        }
    }
});
</script>
```

### Chart.js: Line

```html
<!-- Inside the slide body -->
<div style="width: 100%; height: 300px;">
    <canvas id="lineChart"></canvas>
</div>

<!-- Just before </body> -->
<script>
new Chart(document.getElementById('lineChart'), {
    type: 'line',
    data: {
        labels: ['1月', '2月', '3月', '4月', '5月', '6月'],
        datasets: [{
            label: 'ユーザー数',
            data: [1200, 1900, 3000, 5000, 4200, 6100],
            borderColor: '#2563eb',
            backgroundColor: 'rgba(37, 99, 235, 0.1)',
            borderWidth: 3,
            fill: true,
            tension: 0.4,
            pointBackgroundColor: '#2563eb',
            pointRadius: 5,
            pointHoverRadius: 7
        }]
    },
    options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
            y: {
                beginAtZero: true,
                grid: { color: '#f3f4f6' },
                ticks: { font: { family: "'Noto Sans JP', sans-serif", size: 11 } }
            },
            x: {
                grid: { display: false },
                ticks: { font: { family: "'Noto Sans JP', sans-serif", size: 12 } }
            }
        },
        plugins: {
            legend: { display: false },
            tooltip: {
                backgroundColor: 'rgba(15, 23, 42, 0.9)',
                padding: 12,
                titleFont: { family: "'Noto Sans JP', sans-serif", size: 13 },
                bodyFont: { family: "'Noto Sans JP', sans-serif", size: 12 }
            }
        }
    }
});
</script>
```

**Notes:**
- `tension: 0.4` for smooth curves; `tension: 0` for straight lines
- `fill: true` with semi-transparent `backgroundColor` for area chart effect
- For multi-line, add multiple datasets with distinct `borderColor` values

### Chart.js: Radar

```html
<!-- Inside the slide body -->
<div class="flex items-center justify-center" style="width: 400px; height: 350px; margin: 0 auto;">
    <canvas id="radarChart"></canvas>
</div>

<!-- Just before </body> -->
<script>
new Chart(document.getElementById('radarChart'), {
    type: 'radar',
    data: {
        labels: ['技術力', '価格', 'サポート', 'UI/UX', 'セキュリティ', '拡張性'],
        datasets: [
            {
                label: '自社',
                data: [90, 70, 85, 80, 95, 75],
                borderColor: '#2563eb',
                backgroundColor: 'rgba(37, 99, 235, 0.15)',
                borderWidth: 2,
                pointBackgroundColor: '#2563eb',
                pointRadius: 4
            },
            {
                label: '競合A',
                data: [70, 85, 60, 75, 70, 80],
                borderColor: '#f97316',
                backgroundColor: 'rgba(249, 115, 22, 0.1)',
                borderWidth: 2,
                pointBackgroundColor: '#f97316',
                pointRadius: 4
            }
        ]
    },
    options: {
        responsive: true,
        maintainAspectRatio: true,
        scales: {
            r: {
                beginAtZero: true,
                max: 100,
                ticks: { stepSize: 20, font: { size: 10 }, backdropColor: 'transparent' },
                pointLabels: { font: { family: "'Noto Sans JP', sans-serif", size: 12 } },
                grid: { color: '#e5e7eb' }
            }
        },
        plugins: {
            legend: {
                position: 'bottom',
                labels: {
                    font: { family: "'Noto Sans JP', sans-serif", size: 11 },
                    usePointStyle: true,
                    pointStyle: 'circle',
                    padding: 20
                }
            }
        }
    }
});
</script>
```

**Notes:**
- Best for competitive comparison or capability assessment
- Max 2-3 datasets to avoid visual clutter on a slide
- Keep labels concise (2-4 characters ideal for Japanese)

### Chart.js: Horizontal Bar

```html
<!-- Inside the slide body -->
<div style="width: 100%; height: 300px;">
    <canvas id="hBarChart"></canvas>
</div>

<!-- Just before </body> -->
<script>
new Chart(document.getElementById('hBarChart'), {
    type: 'bar',
    data: {
        labels: ['マーケティング自動化', 'CDP', '計測/分析', 'CRM', 'その他'],
        datasets: [{
            data: [40, 30, 15, 10, 5],
            backgroundColor: ['#2563eb', '#60a5fa', '#93c5fd', '#bfdbfe', '#e5e7eb'],
            borderRadius: 6,
            barPercentage: 0.7
        }]
    },
    options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        scales: {
            x: { beginAtZero: true, max: 50, grid: { color: '#f3f4f6' } },
            y: {
                grid: { display: false },
                ticks: { font: { family: "'Noto Sans JP', sans-serif", size: 12 } }
            }
        },
        plugins: { legend: { display: false } }
    }
});
</script>
```

### Chart.js Style Guidelines

| Setting | Recommended Value | Reason |
|---------|------------------|--------|
| Font family | `'Noto Sans JP', sans-serif` | Match slide typography |
| Tooltip background | `rgba(15, 23, 42, 0.9)` | Consistent dark tooltip |
| Grid color | `#f3f4f6` (gray-100) | Subtle, non-distracting |
| Border width | `0` (doughnut/pie), `2-3` (line) | Clean appearance |
| Border radius | `4-6` (bars) | Match slide card radius |
| Colors | Use Phase 2 brand palette | Visual consistency |
| Legend | `display: false` when manual HTML legend is used | Better PPTX convertibility |
| Canvas container | Fixed height with `maintainAspectRatio: false` | Predictable sizing within slide |

### When to Use Chart.js vs CSS-Only

| Scenario | Recommendation |
|----------|---------------|
| Simple donut with 2-4 segments | CSS `conic-gradient` (simpler, PPTX-friendly) |
| Complex donut with 5+ segments | Chart.js doughnut |
| Horizontal progress bars | CSS width-based bars (simpler, PPTX-friendly) |
| Vertical bar chart with axes | Chart.js bar |
| Line chart / area chart | Chart.js line (no CSS equivalent) |
| Radar / spider chart | Chart.js radar (no CSS equivalent) |
| Stacked bar chart | Chart.js stacked bar (CSS alternative is fragile) |
| Simple percentage indicators | CSS progress bar (simpler, PPTX-friendly) |

---

## DOM Nesting Depth Guidelines

| Slide type | Max depth (body -> text) |
|------------|-------------------------|
| Cover (Pattern 1) | 3-4 levels |
| Content (Pattern 3-6) | 4-5 levels |
| Complex cards (Pattern 8-20) | 5-6 levels max |

Avoid wrapper divs that serve no layout or styling purpose.
