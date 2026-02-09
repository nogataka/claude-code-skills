# Custom Templates

Place your own HTML slide files here as style references.

When generating a new deck, Claude will read these files and extract the visual design
(colors, fonts, header/footer, decorative elements) to use as the primary style guide.

## How to use

1. Copy your HTML slide files into this directory
2. Run `/create-slide-template` as usual
3. Claude will automatically detect and reference the templates

## Rules

- **HTML files only** (`.html`) — other file types are ignored
- **Max 5 files** — if more than 5 exist, only the first 5 (alphabetical) are read
- Files should follow the 1280x720px slide format for best results
- Text content in templates is ignored — only the visual style is extracted
- All Mandatory Constraints from SKILL.md still apply to generated output
