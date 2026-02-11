# Custom Templates

Place your own HTML slide files here as style references.

When generating a new deck, Claude will detect these files in Phase 0 and ask whether to use them. If selected, the visual design (colors, fonts, header/footer, decorative elements) is extracted and used as the primary style guide.

## How to use

### Single template set

1. Copy your HTML slide files directly into this directory
2. Run `/slidekit-create` as usual
3. Claude will ask if you want to use the template

### Multiple template sets

1. Create subdirectories for each template set (e.g., `navy-gold/`, `modern-tech/`)
2. Place HTML slide files inside each subdirectory
3. Run `/slidekit-create` as usual
4. Claude will list the available template sets and ask which one to use

```
templates/
├── navy-gold/
│   ├── 001.html
│   ├── 002.html
│   └── 003.html
├── modern-tech/
│   ├── 001.html
│   └── 002.html
└── README.md
```

## Rules

- **HTML files only** (`.html`) — other file types are ignored
- **Max 5 files per template set** — if more than 5 exist, only the first 5 (alphabetical) are read
- Files should follow the 1280x720px slide format for best results
- Text content in templates is ignored — only the visual style is extracted
- All Mandatory Constraints from SKILL.md still apply to generated output
