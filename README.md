# Book Club Guide

The source of the club's published Guide/Wiki and its Edit Log.

**Status:** v0.10, 2026-11-24. Public on GitHub; the site is published by GitHub Pages from `docs/`. Moving to the Scrub-Jay Collective organization.

## What lives where

| Path | What it is | Public? |
|---|---|---|
| `docs/` | The Guide, one page per menu item. GitHub Pages builds the site from this folder. | ✅ yes |
| `data/books.csv` | Book suggestions B01–B12. | ✅ yes |
| `data/revision-history.csv` | One row per guide version. | ✅ yes |
| `private/` | Organizer-only data: the Edit Log (every suggested change and its decision), task owners and dates (Actions), meeting records and recap drafts. Git ignores this folder, so it never reaches GitHub. | 🔒 never |
| `tools/` | Scripts that build the Word file and the Edit Log workbook. | ✅ yes |
| `exports/` | Built files (`Book-Club-Guide-*.docx`, `Book-Club-Edit-Log-*.xlsx`). Rebuilt any time, so not stored. | — |

## Suggesting a change

1. Open the page under `docs/` on GitHub and click the pencil to edit it. GitHub makes a copy and opens a pull request.
2. Say in the pull request what the change does and why. An organizer logs it in the Edit Log.
3. Other founders review the pull request. Merged means Accepted.

For now, suggestions and reviews stay in the founders' shared Google Doc (Suggesting mode); the steps above are for later (E18).

## Privacy rules

- No phone numbers, emails or home addresses anywhere in this repository.
- Recaps go in `private/` first. One moves to `docs/recaps/` only after someone has reviewed it for personal information, and leaving personal details out entirely is best.

## Building the files

Needs [Bun](https://bun.sh). From `tools/`: `bun install` once, then `bun run build`. It writes `exports/Book-Club-Guide-v<version>.docx` and `exports/Book-Club-Edit-Log-v<version>.xlsx`. `bun preview.ts` writes a local preview of the site to `exports/preview/` (open `index.html`).
