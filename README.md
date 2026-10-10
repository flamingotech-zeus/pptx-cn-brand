# Chinese Business Deck Generator

[English](README.md) | [简体中文](README.zh-CN.md)

> Generate editable Chinese business and brand-style presentations with python-pptx, driven by a built-in design system.

## What this is

> **Note on de-identification.** This README is written generically. The skill file itself may still reference internal domains, host names or local absolute paths. Run a scan pass (see `opensource-preflight-audit`) and replace them before publishing.

This is not an image generator that produces slide-shaped pictures. That approach is common and it has a hard ceiling: the output cannot be edited.

Here the deliverable is a real `.pptx` file. Text boxes are text boxes, shapes are shapes. You can open it and keep working on it, which is the difference between a draft and a handoff.

## Why we open-sourced it

We make decks constantly and kept rebuilding the same layout scaffolding at the start of every one. Pulling it into a reusable system removed a repetitive half hour from each.

Design systems are portable by nature. Once the brand-specific colour names are swapped out there is nothing of ours left in it.

## What it does

- **Three built-in palettes** covering different tones, each with a defined set of text and accent colours rather than ad-hoc picking.
- **A header, footer and page-number system** applied consistently, instead of being set per slide.
- **Layout components** for cards, comparisons, data blocks and process flows, so common slide shapes do not get rebuilt by hand.
- **Real editable output.** The generated file behaves like any other `.pptx`.
- **Programmatic control.** Because it is code, generating thirty slides with consistent styling is the same effort as generating three.

## Requirements

Python 3 · python-pptx

## Installation

```bash
git clone https://github.com/<your-username>/pptx-cn-brand.git
```

Then place the folder into your agent skills directory:

| OS | Path |
|---|---|
| Windows | `C:\Users\<you>\.workbuddy\skills\pptx-cn-brand\` |
| macOS / Linux | `~/.workbuddy/skills/pptx-cn-brand/` |

There is no build step. Restart the session and it takes effect.

## How to use

Describe the deck and let it build:

- "Make me a business deck about X, ten pages"
- "Use the brand palette throughout"
- "Generate the slides from this outline"

The skill produces a `.pptx` you can open and edit directly.

## Files

```
├── SKILL.md
└── assets/
    └── pptx_kit.py
```

## Notes and limits

1. **One palette carries a brand name.** Replace it before publishing.
2. **Layout is a starting point.** Generated slides are meant to be adjusted, not submitted unchanged.
3. **Font availability matters.** Chinese decks depend on the fonts present at render time; a missing font falls back, which can shift line breaks.
4. **python-pptx writes the file format, not the rendering.** If you need to see the result as an image, open it in an office suite or render it separately.

## License

MIT
