# Babylon.js PPTX Viewer

A PowerPoint (.pptx) viewer that loads files via drag and drop and renders slides on Babylon.js GUI.

## Overview

- Input: .pptx (OOXML ZIP)
- Output: 2D slide rendering on Babylon.js GUI, with thumbnails and notes
- Highlights: layout/master inheritance, background image/effect handling, shapes/text/images, and partial chart/SmartArt support

## How To Run

1. Open this folder in VS Code.
2. Start index.html with Live Server.
3. Drag and drop a .pptx file onto the screen.

Requirements:

- A browser with WebGL2 support
- libs/jszip.min.js available in the repository
- Access to Babylon CDN (https://cdn.babylonjs.com) for local development only

Production note:

- The public Babylon CDN is not recommended for production.
- For production deployments, host Babylon packages on your own origin/CDN and update script URLs accordingly.

Example (self-hosted):

```html
<script src="/vendor/babylon/babylon.js"></script>
<script src="/vendor/babylon/gui/babylon.gui.min.js"></script>
```

You can use either of these approaches:

1. Pin a fixed Babylon version and mirror that exact directory structure.
2. Download Babylon release assets and serve them from your own CDN/server.

## Directory Structure

```text
babylon-pptx-viewer/
|- index.html
|- index.js
|- constants.js
|- color-utils.js
|- zip-helpers.js
|- background.js
|- text-parser.js
|- style-inheritance.js
|- shape-parsers.js
|- slide-parser.js
|- pptx-parser.js
|- slide-renderer.js
|- scene-setup.js
|- gui-frame.js
|- libs/
|  `- jszip.min.js
|- test-data/
|  |- fixtures-manifest.json
|  `- assets/
`- tools/
   `- generate_test_pptx.py
```

## Architecture

```text
.pptx (ZIP)
  -> zip-helpers.js (rels / media)
  -> color-utils.js (theme / color resolution)
  -> style-inheritance.js (layout/master inheritance)
  -> background.js (background image/effects)
  -> shape-parsers.js + text-parser.js (element extraction)
  -> pptx-parser.js (main orchestration)
  -> slide-renderer.js (Babylon GUI rendering)
```

Slide geometry is normalized as x, y, w, h in the range 0.0 to 1.0.

## Support Matrix

### OOXML Elements

| Element | Status | Notes |
|---|---|---|
| p:sp (shape) | Supported | rect/ellipse/roundRect/line/pie/chevron and more |
| p:pic (image) | Supported | srcRect crop is applied |
| p:cxnSp (connector) | Supported | flipH/flipV handled |
| p:grpSp (group) | Supported | recursive expansion |
| p:graphicFrame (table) | Partial | basic table rendering |
| p:graphicFrame (chart) | Partial | basic bar/line/pie/area data rendering |
| p:graphicFrame (diagram/SmartArt) | Partial | selected layouts only |

### Text

| Feature | Status | Notes |
|---|---|---|
| Font size / bold / italic | Supported | scaled for Babylon GUI rendering |
| Color resolution (solidFill/schemeClr/srgb/scrgb) | Supported | tint/shade/lum modifiers applied |
| Paragraph alignment (left/center/right) | Supported |  |
| Bullets / auto numbering | Supported | buChar / buAutoNum |
| Explicit line breaks (<a:br/>) | Supported | split into separate render lines |
| CJK wrapping | Supported | zero-width space assistance |
| Layout/master inheritance | Supported | anchor/cap/fontRef color and more |

### Background

| Feature | Status | Notes |
|---|---|---|
| Solid background | Supported |  |
| Background image | Supported | slide -> layout -> master inheritance |
| Theme background via bgRef | Supported | theme style matrix lookup |
| duotone/art effect/tint/alpha | Partial | visual approximation overlays |

## Current Behavior Notes

- GIF files can be loaded and displayed, but are currently rendered as static images (no animation playback).
- For title/ctrTitle on background-image slides, text color may be adjusted for readability.
- If that adjustment results in white, theme tx1 may be preferred when it is a valid distinct color.

## Logs And Debugging

By default, verbose parser/render debug logs are suppressed to keep console output lightweight.

To enable full debug logs in development, run this in DevTools Console before loading a PPTX:

```js
window.__PPTX_DEBUG__ = true;
```

To disable again:

```js
window.__PPTX_DEBUG__ = false;
```

Common log prefixes:

| Prefix | Purpose |
|---|---|
| [INIT] | startup sequence |
| [PPTX] | parse orchestration |
| [BG] / [BLIP] | background and effect handling |
| [LAYOUT] / [MASTER] | inheritance resolution |
| [TREE] / [SP] / [PIC] / [GF] | shape parsing details |
| [RENDER] | rendering phase |

Typical successful startup sequence:

```text
[INIT] boot sequence start
[INIT/ENGINE] done
[INIT/SCENE] done
[INIT/UI] done
[INIT/INPUT] done
[INIT] boot sequence complete
```

## Known Limitations

1. Pixel-perfect parity with PowerPoint is not guaranteed, especially for background effects.
2. SmartArt is only partially supported.
3. Charts are partially supported and do not cover all visual styles.
4. Timeline-based media playback (including animated GIF playback) is not implemented.
5. Embedded custom fonts are not supported; system fallback fonts are used.
6. The current index.html references Babylon public CDN for development convenience; production should use self-hosted package URLs.

## Troubleshooting

### File library (JSZip) error

- Confirm libs/jszip.min.js exists.
- Clear browser cache and reload.

### Graphics engine initialization failed

- Enable hardware acceleration in the browser.
- Update GPU drivers.
- Retry on Chrome or Edge.

### PPTX parse error

- Open and resave the file in PowerPoint, then retry.
- Check whether the issue reproduces with another .pptx.
- For very large files, check memory usage during parsing.
