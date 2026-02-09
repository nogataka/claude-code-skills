# Claude Code Skills

Custom skills for [Claude Code](https://docs.anthropic.com/en/docs/claude-code).

## Skills

### create-slide-template

Generate HTML slide presentation templates (1 slide = 1 HTML file, 1280x720px) using Tailwind CSS, Font Awesome, and Google Fonts. Supports 5 styles (creative, elegant, modern, professional, minimalist) and 5 themes (marketing, portfolio, business, technology, education). Includes 15 layout patterns and PPTX conversion compatibility.

```bash
claude install-skill https://github.com/nogataka/claude-code-skills/tree/main/skills/create-slide-template
```

### pdf-to-html

Convert PDF presentations to high-quality HTML slide files using a visual reproduction approach. Pipeline: PDF → slide screenshots (pdftoppm) → Claude writes HTML matching each screenshot using the create-slide-template design system.

```bash
claude install-skill https://github.com/nogataka/claude-code-skills/tree/main/skills/pdf-to-html
```

#### Dependencies

```bash
brew install poppler
```

## License

MIT
