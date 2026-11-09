# Platform Notes

## macOS

- Requires desktop Microsoft PowerPoint.
- The first run may request Automation permission for the terminal, Codex, or Claude Code to control PowerPoint.
- Uses AppleScript and then verifies copied slide objects inside the generated PPTX.

## Windows

- Requires desktop Microsoft PowerPoint and Windows PowerShell or PowerShell 7.
- Uses PowerPoint COM through a generated temporary PowerShell script.
- The script inserts worship and sermon slides with `Slides.InsertFromFile`, then runs the same package-level validation used on macOS.

## Optional PDF Support

- Install Poppler so `pdftoppm` is available on `PATH`.
- Install Pillow with `python -m pip install Pillow` for the established centered PDF crop.
- PNG and JPEG dimensions are read directly by the script and require no image library.

## Unsupported

LibreOffice-only and Linux environments are not supported for final native copying. They may inspect manifests and run unit tests, but they cannot guarantee PowerPoint animation, theme, transition, WDP effect, or background fidelity.

