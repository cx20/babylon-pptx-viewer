// ============================================================================
// constants.js - Shared constants, namespaces, and default data
// ============================================================================

export var SLIDE_EMU_W = 9144000;
export var SLIDE_EMU_H = 6858000;
export var CANVAS_W = 580;
export var CANVAS_H = 326;
export var TEX_W = 1024;
export var TEX_H = 640;
export var PP = "#D04423"; // PowerPoint orange
export var FONT_SCALE = 0.75;

// XML namespaces
export var A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
export var P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
export var R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// Geometry classification
export var ELLIPSE_SHAPES = ["ellipse", "oval", "circle", "pie", "arc", "chord", "donut"];
export var ROUND_RECT_SHAPES = ["roundRect", "snipRoundRect", "snip1Rect", "snip2SameRect", "round1Rect", "round2SameRect"];

// Default slide (shown before .pptx is loaded)
export function getDefaultSlides() {
    return [{
        bg: "#FFFFFF", bgImage: null, bgTint: null,
        elements: [
            { type: "text", text: "Drag & Drop a .pptx file", x: 0.0, y: 0.30, w: 1.0, fontSize: 24, color: PP, fontWeight: "bold", align: "center" },
            { type: "text", text: "onto this screen to load it", x: 0.0, y: 0.50, w: 1.0, fontSize: 16, color: "#666666", fontWeight: "normal", align: "center" },
            { type: "shape", shape: "line", x1: 0.15, y1: 0.42, x2: 0.85, y2: 0.42, color: PP, thickness: 2 }
        ],
        notes: "Drop a .pptx file to begin."
    }];
}
