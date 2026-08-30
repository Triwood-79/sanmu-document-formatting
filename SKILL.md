---
name: official-document-formatting
description: Create or reformat Chinese official-document DOCX files with a configurable general preset. Use for generating Word documents, standardizing an existing .docx, or viewing and changing the persistent formatting profile; do not use for legacy .doc/.docm files or substantive fact creation.
---

# 通用公文排版

Use this skill to create or reformat `.docx` files without Microsoft Word. Use the deterministic scripts for document changes; do not reproduce their OOXML logic manually.

## Route the request

- New document: accept plain text or Markdown, then run `scripts/docx_engine.py create`.
- Existing document: run `scripts/inspect_docx.py`, show the paragraph classification, table plan, and risk summary, obtain confirmation or corrections, then run `scripts/docx_engine.py format-existing --confirmed`. The table plan may set `title_paragraph_index` and `header_rows` to `0`, `1`, or `2` for each table.
- `查看当前排版格式`: run `scripts/profile_manager.py view`.
- `修改排版格式`: run `profile_manager.py view`, show the current values for the requested groups, run `begin`, ask only those groups, apply answers with `set`, show `diff`, and save only after the user says `确认保存`.
- `重新录入排版格式`: run `profile_manager.py view` first and display the complete current profile in grouped, user-readable form. Then use `begin --fresh`, ask every V1-open field with its current value shown beside the question, and follow the same diff and confirmation flow. If the user chooses to keep a current value, explicitly write that value into the fresh draft so it is not reset to the built-in default.
- `取消修改`: run `profile_manager.py cancel`.
- `恢复通用公文预设格式`: explain that the saved customization will be disabled, obtain confirmation, then run `profile_manager.py restore`.

Never treat an ordinary per-document instruction as a persistent profile change. Precedence is: current-request overrides, confirmed user profile, built-in preset.

## Required workflow

1. Read [references/format-fields.md](references/format-fields.md) when creating, reformatting, or changing parameters.
2. Run `scripts/preflight.ps1`. When the Codex workspace dependency locator is available, use its Python path for subsequent scripts.
3. If `python-docx` or `lxml` is missing, explain the missing capability and request authorization before installation. Never install software or fonts silently.
4. For existing files, refuse `.doc` and `.docm`. Stop on corrupt, encrypted, commented, or tracked-change documents and request a clean `.docx` copy.
5. Never overwrite an input file. The engine creates `<stem>_排版后.docx`, adding a numeric suffix when needed.
6. Normalize all processed text to black. For tables, normalize font families, text color, and the confirmed global bold setting; preserve font sizes, alignment, spacing, borders, shading, row heights, and column widths. Preserve pictures, text boxes, and unusual landscape sections without internal normalization.
7. V1 does not create or rebuild complete colophon structures. Classify an existing issuance-office/date line as `colophon`, normalize only its Chinese/Latin font families to the body fonts and its text color to black, and preserve its size, bold setting, alignment, spacing, separators, and placement. Tell the user that complex colophon layout still requires manual completion.
8. Run `scripts/privacy_scrub.py` and `scripts/validate_docx.py` on every output.
9. Use `scripts/render_docx.py` only when a supported renderer is detected. Without rendering, report structural validation only, never visual approval.
10. Keep the classification JSON after formatting unless the user explicitly asks to delete it. Before deletion, run `scripts/cleanup_classification.py` without `--confirmed`, show the eligibility result, and obtain confirmation. After confirmation, rerun it with `--confirmed --expected-sha256 <sha256>`. If validation or the hash check fails, keep the file and report the refusal.

## Table formatting

- If a table has a confirmed title paragraph, use the main-title Chinese and Latin font families, black text, and the confirmed global bold setting; preserve its size and other paragraph geometry.
- Use the heading-1 font families, black text, and the global bold setting for the first header row.
- If a second header row is confirmed, use the heading-2 font families, black text, and the global bold setting for that row.
- Use the body font families, black text, and the global bold setting for all remaining table rows.
- Inspection proposes one header row by default and proposes two only when the first row has merged cells or the second row is marked as a repeating header. Always show the proposal before formatting so the user can correct it.

## Profile questions

Ask in small groups: margins and print mode; global bold; fonts and fallbacks; title hierarchy; body/reference/description; page-number font, size, and bold. Show the current value whenever asking for a replacement value.

Fixed line spacing is an explicit V1 question, not an inferred value. Ask separately for the main title and for every other paragraph role: heading 1, heading 2, body, reference note, and description. The user may answer with one shared value for several roles; expand it to each affected field. Keep a draft until explicit confirmation and show a current-to-proposed diff first. Do not offer fields that are fixed in V1.

Do not store names, organizations, dates, content, or document paths in the profile. Read [references/privacy-and-release.md](references/privacy-and-release.md) before packaging or publishing the skill.

## Content boundary

Formatting does not authorize inventing facts, names, organizations, dates, signatures, or other document content. Optional document content belongs only to the current file.
