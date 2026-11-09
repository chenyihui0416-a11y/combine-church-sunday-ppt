# Customization Map

Keep additions small and attach them to the narrowest existing extension point.

## New Input Naming Pattern

Edit `_select_pptx` in `scripts/input_manifest.py`. Preserve explicit command-line overrides so unusual weeks can bypass automatic discovery.

## New Announcement Asset Type

Edit `asset_kind` and `pair_assets` in `scripts/input_manifest.py`, then handle the new kind in `prepare_assets` and the announcement loop in `scripts/combine_sunday_ppt.py`.

## New Layout Rule

Add a named geometry or pagination function in `scripts/combine_sunday_ppt.py`. Keep all platform-specific commands inside `AppleScriptDeck` and `PowerShellDeck`; the assembly workflow should call the same method on either backend.

## New Validation Rule

Add deterministic package checks to `scripts/validate_against_reference.py`. Do not rely only on screenshots when native slide objects, media hashes, visibility, or text can be checked directly.

## New Platform

Implement the same deck API as `AppleScriptDeck` and `PowerShellDeck`, add it to `select_engine`, and keep native-copy validation enabled. A platform is not supported until a real PowerPoint end-to-end fixture passes.

## Regression Tests

Put portable logic tests in the repository `tests/` directory. Do not commit weekly bulletin, worship, sermon, QR, or finished-service files; use private fixtures for end-to-end testing.
