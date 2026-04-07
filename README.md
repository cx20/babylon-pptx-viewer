# PPTX Viewer - Babylon.js PowerPoint Simulator

A viewer that loads PowerPoint files (.pptx) via drag-and-drop onto a 3D monitor and renders slides using Babylon.js GUI.

## Demo

Runs in Babylon.js Playground v2.
Playground v2 supports multi-file projects and ES modules, so you can add each .js file in separate tabs and run directly.

## File Structure

```text
pptx-viewer/
|- README.md
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
|- scene-setup.js
|- gui-frame.js
`- slide-renderer.js
```

## Architecture

### Data Flow

```text
.pptx file (ZIP)
    |
    v
+-------------+    +-------------+    +-------------------+
| zip-helpers | -> | color-utils | -> | style-inheritance |
| (ZIP/rels)  |    | (theme)     |    | (layout/master)   |
+-------------+    +-------------+    +-------------------+
    |                                      |
    v                                      v
+-------------+    +-------------+    +-------------------+
| background  |    | text-parser | -> | shape-parsers     |
| (bg/effects)|    | (paragraphs)|    | (sp/pic/grp/gf)   |
+-------------+    +-------------+    +-------------------+
    |                                      |
    v                                      v
+---------------------------------------------------------+
|                      pptx-parser.js                     |
|        (orchestrates all parsing and builds slides[])  |
+---------------------------------------------------------+
    |
    v
+---------------------------------------------------------+
|                    slide-renderer.js                    |
|          (renders slides to Babylon.js GUI canvas)      |
+---------------------------------------------------------+
```

### Shared State (app object)

Created in index.js and passed to renderer-related modules:

```javascript
var app = {
    scene: scene,
    gui: { ... },
    slides: [],
    currentSlide: 0,
    thumbRects: []
};
```

### Slide Data Model

Each slide has the following structure:

```javascript
{
    bg: "#FFFFFF",
    bgImage: "data:image/...", // or null
    bgTint: {
        type: "artEffect",      // "duotone" | "artEffect" | "tint" | "alpha"
        color: "#0E5580"
    },
    elements: [
        { type: "text", text: "...", x: 0.1, y: 0.2, w: 0.8, fontSize: 24, color: "#FFF", ... },
        { type: "shape", shape: "rect", x: 0.0, y: 0.0, w: 0.5, h: 0.5, fillColor: "#ACD433", ... },
        { type: "image", dataUrl: "data:...", x: 0.5, y: 0.0, w: 0.5, h: 1.0, ... }
    ],
    notes: "Slide 1"
}
```

Coordinate system: x, y, w, h are normalized ratios in the range [0.0, 1.0] relative to slide size.

### Element Schema (Issue 04)

All elements are normalized by normalizeElement() and follow this schema.

#### Common Properties (all elements)

| Property | Type | Default | Description |
|---|---|---|---|
| type | string | "shape" | "text" | "shape" | "image" | "table" |
| x | number | 0 | Left position (0.0-1.0) |
| y | number | 0 | Top position (0.0-1.0) |
| rotation | number | 0 | Rotation in degrees |

#### Text Element (type: "text")

| Property | Type | Default | Description |
|---|---|---|---|
| text | string | "" | Text content |
| w | number | 1 | Width (0.0-1.0) |
| fontSize | number | 12 | Point size |
| color | string | "#000000" | Hex color |
| fontWeight | string | "normal" | "normal" or "bold" |
| fontStyle | string | "normal" | "normal" or "italic" |
| fontFamily | string | "Calibri" | Font family |
| align | string | "left" | "left" | "center" | "right" |

#### Shape Element (type: "shape")

| Property | Type | Default | Description |
|---|---|---|---|
| shape | string | "rect" | "rect" | "ellipse" | "line" | "circle" etc. |
| w | number | 1 | Width (0.0-1.0) |
| h | number | 1 | Height (0.0-1.0) |
| fillColor | string | "#FFFFFF" | Fill color |
| strokeColor | string | "#000000" | Stroke color |
| thickness | number | 1 | Stroke width (pixels) |
| x1, y1, x2, y2 | number | - | Line endpoints (line shape only) |

#### Image Element (type: "image")

| Property | Type | Default | Description |
|---|---|---|---|
| dataUrl | string | null | Image data URL |
| w | number | 1 | Width (0.0-1.0) |
| h | number | 1 | Height (0.0-1.0) |
| crop | object | - | {l, t, r, b} crop ratios (0.0-1.0) |

#### Table Element (type: "table")

| Property | Type | Default | Description |
|---|---|---|---|
| rows | number | 0 | Row count |
| cols | number | 0 | Column count |
| tableData | array | [] | Cell data |

Benefits of normalization:
- Every element has guaranteed properties, reducing undefined checks in render paths.
- Parser output stays consistent across shape types.
- New element types/properties can be added by updating a single schema.

## Feature Support

### OOXML Elements

| Element | Status | Notes |
|---|---|---|
| p:sp (shape) | Yes | rect, ellipse, roundRect, etc. |
| p:pic (image) | Yes | srcRect crop supported |
| p:cxnSp (connector) | Yes | flipH/flipV supported |
| p:grpSp (group) | Yes | recursive nested groups supported |
| p:graphicFrame (table) | Partial | basic table rendering |
| p:graphicFrame (chart) | Partial | placeholder rendering |
| p:graphicFrame (diagram) | Partial | placeholder rendering |
| SmartArt | No | not supported yet |
| Animation | No | not supported yet |

### Text

| Feature | Status | Notes |
|---|---|---|
| Font size | Yes | rendered with 0.75 scale |
| Bold/italic | Yes |  |
| Color (solidFill, schemeClr) | Yes | shade/tint/lumMod modifiers supported |
| Alignment | Yes | left/center/right |
| Bullets (buChar) | Yes |  |
| Capitalization (cap="all") | Yes | layout inheritance supported |
| CJK wrapping | Yes | zero-width spaces inserted |
| Line spacing | Partial | approximate |

### Background

| Feature | Status | Notes |
|---|---|---|
| Solid background | Yes |  |
| Image background | Yes |  |
| Gradient | Partial | approximated by first gradient stop |
| Duotone | Yes | grayscale detection + dk2 tint |
| Art effect | Partial | approximated by dk2 overlay |
| Background inheritance (slide->layout->master) | Yes |  |

### Style Inheritance

| Layer | Status | Notes |
|---|---|---|
| Slide-local style | Yes |  |
| Layout placeholders | Yes | cap, anchor, fontSize, color, fontRef |
| Master txStyles | Yes | titleStyle, bodyStyle (non-bg-image slides) |
| Master placeholder fontRef | Yes |  |
| Theme colors (dk1-accent6) | Yes |  |

## Development Guide

### Editing Modules

Each file has a single responsibility. Typical extension points:
- Add a new shape type: extend parseShapeTree in shape-parsers.js.
- Add text formatting support: extend parseParagraphs in text-parser.js.
- Add inheritance rule: extend style-inheritance.js.
- Add chart/SmartArt support: extend parseGraphicFrame in shape-parsers.js.
- Improve rendering: adjust renderSlide in slide-renderer.js.
- Update UI frame: edit gui-frame.js.

### Using Babylon.js Playground v2

Playground v2 supports:
- Multi-file tabs (VS Code-like layout)
- Native ES Modules import/export handling
- NPM package integration
- TypeScript IntelliSense (auto type acquisition)
- Chrome DevTools debugging
- Separate shader files (.wgsl/.glsl)

Add each .js file as a tab and run index.js as the entry point.
For local development, open index.html with Live Server.

### Debugging

Use console log prefixes to locate issues quickly:

| Prefix | File | Purpose |
|---|---|---|
| [PPTX] | pptx-parser.js | parsing orchestration |
| [BG] | background.js | background extraction |
| [BLIP] | background.js | duotone/art-effect handling |
| [TREE] | shape-parsers.js | shape tree traversal |
| [SP] | shape-parsers.js | shape parsing |
| [PIC] | shape-parsers.js | image parsing |
| [GF] | shape-parsers.js | graphicFrame parsing |
| [LAYOUT] | style-inheritance.js | layout style inheritance |
| [MASTER] | style-inheritance.js | master style inheritance |
| [RENDER] | slide-renderer.js | rendering process |

## Known Limitations

1. Background art effects: full parity with PowerPoint pixel processing is difficult in WebGL. The current implementation approximates with a dk2 color overlay.
2. Charts: OOXML chart rendering is not implemented yet; placeholder boxes are shown.
3. SmartArt: dsp:drawing diagram parsing is not implemented yet.
4. Animation: slide animations and transitions are not implemented.
5. Font size: Babylon.js GUI constraints require 75% font scaling.
6. Embedded fonts: custom embedded fonts are not supported; fallbacks such as Segoe UI/Calibri are used.

## Troubleshooting

### Live Server startup failures

#### Error: "renderCanvas not found"

Cause:
- The canvas element is referenced before DOM is fully ready.

Fix:
1. Hard reload the page (Ctrl+Shift+R).
2. Check Live Server delayed-load settings.
3. Ensure script type="module" src="index.js" is placed at the end of body.

#### Error: "Failed to load file library"

Cause:
- libs/jszip.min.js is missing or corrupted.

Fix:
1. Verify libs/jszip.min.js exists.
2. Check browser console for jszip load errors.
3. If corrupted, redownload jszip 3.10.1 from:
   https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js

#### Error: "Failed to initialize graphics engine"

Cause:
- WebGL is disabled or GPU/driver support is unavailable.

Fix:
1. Enable hardware acceleration in browser settings.
2. Update GPU drivers.
3. Try another browser (Chrome/Firefox/Edge).
4. Try private/incognito mode to rule out extension conflicts.

### PPTX loading failure

#### Error: "PPTX parse error"

Cause:
- File is corrupted or unsupported content is present.

Fix:
1. Verify file is a valid .pptx (ZIP-based package).
2. Open in PowerPoint and save again to repair.
3. Test with another .pptx.
4. For very large files, check memory pressure during data URL generation.

### Check phased initialization logs in console

Example successful sequence:

```text
[INIT] boot sequence start
[INIT/ENGINE] starting
[INIT/ENGINE] done
[INIT/SCENE] starting
[INIT/SCENE] done
[INIT/UI] starting
[INIT/UI] done
[INIT/INPUT] starting
[INIT/INPUT] done
[INIT] boot sequence complete
```

On failure:
- [INIT] failed appears
- error code shows category
- dev message includes detailed diagnostics

### Environment-specific known issues

| Environment | Symptom | Cause | Mitigation |
|---|---|---|---|
| Windows Safari | graphics not displayed | no WebGL2 support | use Chrome or Edge |
| iPad/iPhone | touch shape interactions are limited | mobile touch behavior not fully implemented | use desktop browser for testing |
| VPN/proxy | Babylon.js CDN load fails | whitelist not configured | request whitelist for cdn.babylonjs.com |
