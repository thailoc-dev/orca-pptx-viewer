# Presentation Viewer for Orca

Open PowerPoint presentations as read-only slides in Orca's editor: `.pptx`, slide shows
(`.ppsx`), templates (`.potx`) and their macro-enabled forms (`.pptm`, `.ppsm`, `.potm`; macros
never run). The preview renders text,
shapes, pictures, tables, charts, SmartArt, gradients and backgrounds. Slides keep their
original colours; the controls follow Orca's light or dark theme.

- **Navigate** with the ‹ › buttons, ←/→, Page Up/Page Down, Home/End, or by scrolling.
- **Zoom** with − / + or Ctrl+wheel / trackpad pinch; **Fit** returns to fit-to-width.
- **Links**: clicking an HTTP(S) link shows its address to copy. Orca's plugin API cannot open
  URLs yet, and the editor never navigates away.

Orca supplies the file's bytes for local, SSH and paired-host files and sends a new version when
the file changes on disk; the viewer reloads and stays on the same slide. It never writes to the
file. Requires Orca 1.4.245 or later.

## Limits

- Legacy `.ppt`, OpenDocument `.odp`, Keynote and password-protected files are unsupported.
- Files up to 64 MiB, with ZIP entry and size limits checked before rendering.
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
