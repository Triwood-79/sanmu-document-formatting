---
name: official-document-formatting
description: Create or reformat Chinese official-document DOCX files with a configurable general preset. Use for generating Word documents, standardizing an existing .docx, or viewing and changing the persistent formatting profile; do not use for legacy .doc/.docm files or substantive fact creation.
---

# 通用公文排版

Use this skill to create or reformat `.docx` files without Microsoft Word. Use the deterministic scripts for document changes; do not reproduce their OOXML logic manually.

## Route the request

- New document: accept plain text or Markdown, then run `scripts/docx_engine.py create`.
- Existing document: run `scripts/inspect_docx.py`, show the classification and risk summary, obtain confirmation or corrections, then run `scripts/docx_engine.py format-existing --confirmed`.
- `查看当前排版格式`: run `scripts/profile_manager.py view`.
- `修改排版格式`: run `profile_manager.py begin`, ask only the relevant parameter groups, apply answers with `set`, show `diff`, and save only after the user says `确认保存`.
- `重新录入排版格式`: use `begin --fresh`, then follow the same confirmation flow.
- `取消修改`: run `profile_manager.py cancel`.
- `恢复通用公文预设格式`: explain that the saved customization will be disabled, obtain confirmation, then run `profile_manager.py restore`.

Never treat an ordinary per-document instruction as a persistent profile change. Precedence is: current-request overrides, confirmed user profile, built-in preset.

## Required workflow

1. Read [references/format-fields.md](references/format-fields.md) when creating, reformatting, or changing parameters.
2. Run `scripts/preflight.ps1`. When the Codex workspace dependency locator is available, use its Python path for subsequent scripts.
3. If `python-docx` or `lxml` is missing, explain the missing capability and request authorization before installation. Never install software or fonts silently.
4. For existing files, refuse `.doc` and `.docm`. Stop on corrupt, encrypted, commented, or tracked-change documents and request a clean `.docx` copy.
5. Never overwrite an input file. The engine creates `<stem>_排版后.docx`, adding a numeric suffix when needed.
6. Preserve tables, pictures, text boxes, and unusual landscape sections. Report them; do not normalize their contents.
7. Run `scripts/privacy_scrub.py` and `scripts/validate_docx.py` on every output.
8. Use `scripts/render_docx.py` only when a supported renderer is detected. Without rendering, report structural validation only, never visual approval.

## Profile questions

Ask in small groups: page setup and print mode; global bold; fonts and fallbacks; title hierarchy; body/reference/description/colophon; page numbers. Keep a draft until explicit confirmation and show a current-to-proposed diff first.

Do not store names, organizations, dates, content, or document paths in the profile. Read [references/privacy-and-release.md](references/privacy-and-release.md) before packaging or publishing the skill.

## Content boundary

Formatting does not authorize inventing facts, names, organizations, dates, signatures, or colophon text. Optional document content belongs only to the current file.
