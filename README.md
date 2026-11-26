# Presentation Viewer for Orca

Open PowerPoint presentations as read-only slides in Orca's editor: `.pptx`, slide shows
(`.ppsx`), templates (`.potx`) and their macro-enabled forms (`.pptm`, `.ppsm`, `.potm`; macros
never run). The preview renders text,
shapes, pictures, tables, charts, SmartArt, gradients and backgrounds. Slides keep their
original colours; the controls follow Orca's light or dark theme.

- **Navigate** with the ‹ › buttons, ←/→, Page Up/Page Down, Home/End, or by scrolling. Type a
  number into the slide field to jump straight to that slide.
- **Thumbnails**: the sidebar button shows a strip of slide previews; click one or move through
  them with the arrow keys. Hidden slides are dimmed and marked.
- **Find** (Ctrl/Cmd+F) searches the text of every slide, including tables. Enter and Shift+Enter
  step through the matches, and the shape or table holding each one is outlined.
- **Speaker notes**: the notes button shows the current slide's notes below the slides.
- **Zoom** with − / +, Ctrl/Cmd + / −, or Ctrl+wheel / trackpad pinch. **Fit** fits the slide's
  width (Ctrl/Cmd+0); **Page** fits the whole slide and keeps fitting as the editor resizes.
- **Links**: clicking an HTTP(S) link shows its address to copy. Orca's plugin API cannot open
  URLs yet, and the editor never navigates away.

Orca supplies the file's bytes for local, SSH and paired-host files and sends a new version when
the file changes on disk. The viewer reads the new version while the old one stays on screen,
then swaps it in on the same slide; if the new version cannot be read (for example while a save
is still being written), the previous one stays and the footer says so. It never writes to the
file. Requires Orca 1.4.246 or later.

## Limits

- Legacy `.ppt`, OpenDocument `.odp`, Keynote and password-protected files are unsupported.
- Files up to 64 MiB, with ZIP entry and size limits checked before rendering.
- Find outlines the whole shape or table that contains a match, not the individual words.
- Layout can differ from PowerPoint, notably with fonts that are not installed. Animations,
  transitions, audio and video are not played. Some chart details (such as category labels) may
  be missing. Pictures linked from the web are blocked; only embedded media is shown.
- Rendering runs in Orca's sandboxed editor frame: no network, no scripts from the file.

## Develop

Settings → Plugins → **Development folders…** → add this folder. To rebuild the editor after
changing `build/src`, see [build/README.md](build/README.md).

## Licence

MIT. The bundled renderer is [@aiden0z/pptx-renderer](https://github.com/aiden0z/pptx-renderer)
(Apache-2.0) with JSZip, ECharts and their dependencies; see
[THIRD_PARTY_NOTICES.txt](THIRD_PARTY_NOTICES.txt).
