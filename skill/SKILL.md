---
name: combine-church-sunday-ppt
description: Assemble a Chinese church Sunday-service PowerPoint from a template PPTX, scripture DOCX, four-song worship PPTX, optional sermon PPTX, bulletin DOCX, and announcement QR codes, posters, images, or PDFs. Use for weekly Sunday PPT combination, native worship/sermon copying, Traditional Chinese scripture, announcement pagination and asset matching, and reference-deck validation.
---

# Combine Church Sunday PPT

Use the bundled scripts. Preserve worship, sermon, and response slides as native PowerPoint slides; never reconstruct or flatten them.

## Requirements

- macOS or Windows with desktop Microsoft PowerPoint
- Python 3.10+
- Network access to `springbible.fhl.net` the first time a scripture chapter is used; later runs use the local cache
- Optional `pdftoppm` and Pillow only when an announcement PDF is present

For setup details, read `references/platforms.md` only when installation or platform troubleshooting is needed.

## Inputs

Keep weekly source files in one folder and never modify them:

- template PPTX with `template` in its filename, or an explicit `--template /private/path/template.pptx`
- DOCX containing `宣召經文` and `讀經經文` or `講道經文`
- worship PPTX with four song-title slides
- optional sermon PPTX; when absent, insert one blank placeholder
- bulletin DOCX containing `報告事項`
- optional PNG, JPG, JPEG, or PDF announcement assets

## Fast Path

Run one build command. Do not run a separate manifest inspection or dry-run unless automatic discovery is ambiguous.

```bash
python3 scripts/combine_sunday_ppt.py /weekly/folder \
  --output /weekly/folder/SundayPPT_YYYYMMDD_AM.pptx \
  --json /tmp/SundayPPT-build-result.json
```

The script automatically selects AppleScript on macOS and PowerShell COM on Windows. It reports discovery, scripture, PowerPoint, repair, validation, and total elapsed time.

When discovery selects the wrong file, inspect the manifest once and rerun with explicit overrides:

```bash
python3 scripts/input_manifest.py /weekly/folder
python3 scripts/combine_sunday_ppt.py /weekly/folder \
  --worship /path/worship.pptx \
  --sermon /path/sermon.pptx \
  --scripture-docx /path/scripture.docx \
  --bulletin-docx /path/bulletin.docx \
  --output /path/output.pptx
```

Use `--dry-run` only to diagnose an unusual week. Use `--engine macos` or `--engine windows` only to override platform detection.

## Required Order

1. Template cover and generated call to worship
2. Worship songs 1-3, copied unchanged
3. Welcome and generated scripture pages
4. Sermon, copied unchanged; or one blank placeholder
5. Worship song 4 as response, copied unchanged
6. Offering, welcome, announcements, same-page QR codes, and separate poster slides
7. Doxology and hidden parking-notice slide

## Rules

- Fetch Traditional Chinese Union Version scripture with `scripts/spring_bible.py`.
- Keep scripture in editable template placeholders. Use at most four verses per reading slide, explicit verse numbers, a 32 pt title, and 40 pt body text.
- Keep announcement text at 36 pt. Split text longer than 220 characters at sentence boundaries and repeat the item number.
- Put each QR code on the same slide as its matched announcement and preserve its aspect ratio.
- Put a poster, flyer, map, or PDF on a separate slide immediately after its matched announcement.
- Keep the first three worship songs and the response song in source order and Traditional Chinese.
- Preserve native designs, media, transitions, timing, WDP effects, doxology background, and hidden parking state.
- Never target an arbitrary active presentation. Use exact presentation paths and unique output names.
- Do not leave automation scripts, render caches, or working files in the weekly folder.
- Never bundle or commit a real template, worship deck, sermon deck, bulletin, or finished service deck with this skill.

## Validation

Every normal build automatically compares native copied slides by text, shape count, picture count, media hashes, transitions, and timing. Any mismatch fails the build.

When an accepted finished deck exists, also run:

```bash
python3 scripts/validate_against_reference.py \
  /weekly/folder /tmp/generated.pptx /weekly/folder/finished-reference.pptx
```

This permits a different sermon span but requires matching scripture pages, announcement text, QR placement, worship/response sequence, doxology background, and hidden parking slide. Inspect representative PowerPoint-rendered pages only after these deterministic checks pass.

## Extending

Read `references/customization.md` when adding a naming pattern, asset type, layout rule, validator, or platform backend. Keep platform-specific commands behind the shared deck API and add a regression test.

## Utilities

- `scripts/pptx_inventory.py inspect deck.pptx`: full structural JSON
- `scripts/pptx_inventory.py summary deck.pptx`: slide order, visibility, and text
- `scripts/pptx_inventory.py match target.pptx candidate.pptx`: candidate comparison
- `scripts/spring_bible.py '太14:34-15:9'`: fetch or read cached scripture
