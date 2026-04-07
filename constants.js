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
            normalizeElement({ type: "text", text: "Drag & Drop a .pptx file", x: 0.0, y: 0.30, w: 1.0, fontSize: 24, color: PP, fontWeight: "bold", align: "center" }),
            normalizeElement({ type: "text", text: "onto this screen to load it", x: 0.0, y: 0.50, w: 1.0, fontSize: 16, color: "#666666", fontWeight: "normal", align: "center" }),
            normalizeElement({ type: "shape", shape: "line", x1: 0.15, y1: 0.42, x2: 0.85, y2: 0.42, color: PP, thickness: 2 })
        ],
        notes: "Drop a .pptx file to begin."
    }];
}

// Element schema normalization
// Ensures all parsed elements conform to a consistent schema with default values
var DEFAULT_ELEMENT_SCHEMA = {
    // Common properties
    type: "shape",           // "text" | "shape" | "image" | "table"
    x: 0, y: 0, w: 1, h: 1, // Normalized [0,1] coordinates
    rotation: 0,             // Degrees, 0-360

    // Text properties
    text: "",                // Text content
    fontSize: 12,            // Points
    color: "#000000",        // Hex color
    fontWeight: "normal",    // "normal" | "bold"
    fontStyle: "normal",     // "normal" | "italic"
    fontFamily: "Calibri",   // Font name
    align: "left",           // "left" | "center" | "right"

    // Shape properties
    shape: "rect",           // "rect" | "ellipse" | "line" | "circle" etc.
    fillColor: "#FFFFFF",
    strokeColor: "#000000",
    thickness: 1,            // Stroke width

    // Line-specific properties
    x1: 0, y1: 0, x2: 1, y2: 1,

    // Image properties
    dataUrl: null,           // Data URL for images

    // Table properties
    rows: 0, cols: 0, tableData: []
};

export function normalizeElement(el) {
    if (!el || typeof el !== "object") return null;

    var normalized = {};
    
    // Copy all existing properties
    for (var key in el) {
        if (el.hasOwnProperty(key)) {
            normalized[key] = el[key];
        }
    }

    // Apply defaults for missing critical properties
    if (normalized.type === undefined) normalized.type = DEFAULT_ELEMENT_SCHEMA.type;
    if (normalized.x === undefined || normalized.x === null) normalized.x = DEFAULT_ELEMENT_SCHEMA.x;
    if (normalized.y === undefined || normalized.y === null) normalized.y = DEFAULT_ELEMENT_SCHEMA.y;
    if (normalized.rotation === undefined) normalized.rotation = DEFAULT_ELEMENT_SCHEMA.rotation;

    // Type-specific defaults
    if (normalized.type === "text") {
        if (!normalized.text) normalized.text = "";
        if (!normalized.fontSize) normalized.fontSize = DEFAULT_ELEMENT_SCHEMA.fontSize;
        if (!normalized.color) normalized.color = DEFAULT_ELEMENT_SCHEMA.color;
        if (!normalized.fontWeight) normalized.fontWeight = DEFAULT_ELEMENT_SCHEMA.fontWeight;
        if (!normalized.fontStyle) normalized.fontStyle = DEFAULT_ELEMENT_SCHEMA.fontStyle;
        if (!normalized.align) normalized.align = DEFAULT_ELEMENT_SCHEMA.align;
        if (normalized.w === undefined || normalized.w === null) normalized.w = DEFAULT_ELEMENT_SCHEMA.w;
    }
    else if (normalized.type === "shape") {
        if (!normalized.shape) normalized.shape = DEFAULT_ELEMENT_SCHEMA.shape;
        if (normalized.w === undefined || normalized.w === null) normalized.w = DEFAULT_ELEMENT_SCHEMA.w;
        if (normalized.h === undefined || normalized.h === null) normalized.h = DEFAULT_ELEMENT_SCHEMA.h;
        if (!normalized.fillColor) normalized.fillColor = DEFAULT_ELEMENT_SCHEMA.fillColor;
        if (!normalized.strokeColor) normalized.strokeColor = DEFAULT_ELEMENT_SCHEMA.strokeColor;
        if (normalized.thickness === undefined) normalized.thickness = DEFAULT_ELEMENT_SCHEMA.thickness;
        
        // Line-specific
        if (normalized.shape === "line") {
            if (normalized.x1 === undefined) normalized.x1 = DEFAULT_ELEMENT_SCHEMA.x1;
            if (normalized.y1 === undefined) normalized.y1 = DEFAULT_ELEMENT_SCHEMA.y1;
            if (normalized.x2 === undefined) normalized.x2 = DEFAULT_ELEMENT_SCHEMA.x2;
            if (normalized.y2 === undefined) normalized.y2 = DEFAULT_ELEMENT_SCHEMA.y2;
        }
    }
    else if (normalized.type === "image") {
        if (normalized.w === undefined || normalized.w === null) normalized.w = DEFAULT_ELEMENT_SCHEMA.w;
        if (normalized.h === undefined || normalized.h === null) normalized.h = DEFAULT_ELEMENT_SCHEMA.h;
    }
    else if (normalized.type === "table") {
        if (normalized.rows === undefined) normalized.rows = DEFAULT_ELEMENT_SCHEMA.rows;
        if (normalized.cols === undefined) normalized.cols = DEFAULT_ELEMENT_SCHEMA.cols;
        if (!normalized.tableData) normalized.tableData = DEFAULT_ELEMENT_SCHEMA.tableData;
    }

    return normalized;
}
