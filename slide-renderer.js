// ============================================================================
// slide-renderer.js - Slide rendering (main canvas + thumbnails)
// ============================================================================

import { CANVAS_W, CANVAS_H, PP, FONT_SCALE, ELLIPSE_SHAPES, ROUND_RECT_SHAPES } from "./constants.js";

// Apply background tint overlay to a container
function applyTintOverlay(container, bgTint, name) {
    if (!bgTint) return;
    var tintRect = new BABYLON.GUI.Rectangle(name);
    tintRect.width = "100%"; tintRect.height = "100%"; tintRect.thickness = 0;
    if (bgTint.type === "duotone") { tintRect.background = bgTint.dark; tintRect.alpha = 0.55; }
    else if (bgTint.type === "artEffect") { tintRect.background = bgTint.color; tintRect.alpha = 0.40; }
    else if (bgTint.type === "tint") { tintRect.background = bgTint.color; tintRect.alpha = 0.5; }
    else if (bgTint.type === "alpha") { tintRect.background = "#000000"; tintRect.alpha = 1.0 - bgTint.amt; }
    else return;
    container.addControl(tintRect);
}

// Render a single text element into a GUI container
function renderTextElement(el, container, canvasW, canvasH, fontScale) {
    var tb = new BABYLON.GUI.TextBlock();
    // Insert zero-width spaces for CJK word wrapping
    var displayText = el.text.replace(/([\u3000-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF])/g, "\u200B$1");
    var renderFS = Math.round(el.fontSize * fontScale);
    tb.text = displayText; tb.fontSize = renderFS;
    tb.fontWeight = el.fontWeight || "normal"; tb.fontStyle = el.fontStyle || "normal";
    tb.color = el.color; tb.fontFamily = "Segoe UI, Calibri, sans-serif";
    tb.textWrapping = true; tb.resizeToFit = false;
    if (el.w && el.w > 0) {
        var containerW = el.w * canvasW;
        tb.width = containerW + "px";
        var hasCJK = /[\u3000-\u9FFF\uF900-\uFAFF]/.test(el.text);
        var charW = hasCJK ? 1.0 : 0.55;
        var estTextW = el.text.length * renderFS * charW;
        var fs = renderFS;
        if (fs >= 12 && estTextW > containerW * 2) {
            fs = Math.max(8, Math.floor(containerW * 2 / (el.text.length * charW)));
            tb.fontSize = fs;
        }
        var estLines = Math.ceil(el.text.length * fs * charW / containerW) || 1;
        var estHeight = Math.max(estLines * fs * 1.4, fs * 1.5);
        tb.height = estHeight + "px";
    }
    if (el.align === "center") tb.textHorizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_CENTER;
    else if (el.align === "right") tb.textHorizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_RIGHT;
    else tb.textHorizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
    tb.left = (el.x * canvasW) + "px"; tb.top = (el.y * canvasH) + "px";
    tb.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
    tb.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
    if (el.rotation) tb.rotation = el.rotation * Math.PI / 180;
    container.addControl(tb);
}

// Render the current slide onto the main canvas
export function renderSlide(app) {
    if (app.scene.isDisposed) return;
    var sLayer = app.gui.sLayer;
    sLayer.clearControls();
    var slide = app.slides[app.currentSlide];
    app.gui.sCanvas.background = slide.bg;
    console.log("[RENDER] === Rendering slide " + (app.currentSlide + 1) + ": " + slide.elements.length + " elements, bg=" + slide.bg + ", bgImage=" + (slide.bgImage ? "yes" : "no") + " ===");

    // Background image
    if (slide.bgImage) {
        console.log("[RENDER] Background image: " + slide.bgImage.substring(0, 50) + "... (length=" + slide.bgImage.length + ")");
        var bgI = new BABYLON.GUI.Image("sBg", slide.bgImage);
        bgI.stretch = BABYLON.GUI.Image.STRETCH_FILL;
        bgI.width = "100%"; bgI.height = "100%"; sLayer.addControl(bgI);
        if (slide.bgTint) {
            applyTintOverlay(sLayer, slide.bgTint, "bgTint");
            console.log("[RENDER] Applied bgTint overlay: " + JSON.stringify(slide.bgTint));
        }
    }

    slide.elements.forEach(function (el) {
        if (el.type === "shape") {
            if (el.shape === "line") {
                var x1 = (el.x1 || 0) * CANVAS_W, y1 = (el.y1 || 0) * CANVAS_H;
                var x2 = (el.x2 || 0) * CANVAS_W, y2 = (el.y2 || 0) * CANVAS_H;
                var dx = x2 - x1, dy = y2 - y1;
                var len = Math.sqrt(dx * dx + dy * dy), ang = Math.atan2(dy, dx);
                var le = new BABYLON.GUI.Rectangle();
                le.width = len + "px"; le.height = (el.thickness || 2) + "px";
                le.background = el.color || "#000"; le.thickness = 0;
                le.left = ((x1 + x2) / 2 - CANVAS_W / 2) + "px";
                le.top = ((y1 + y2) / 2 - CANVAS_H / 2) + "px";
                le.rotation = ang; sLayer.addControl(le);
            } else if (ELLIPSE_SHAPES.indexOf(el.shape) >= 0) {
                var ell = new BABYLON.GUI.Ellipse();
                ell.left = (el.x * CANVAS_W) + "px"; ell.top = (el.y * CANVAS_H) + "px";
                ell.width = (el.w * CANVAS_W) + "px"; ell.height = (el.h * CANVAS_H) + "px";
                ell.background = el.fillColor || "transparent";
                ell.thickness = el.borderWidth || 0; ell.color = el.borderColor || "transparent";
                ell.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                ell.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                if (el.rotation) ell.rotation = el.rotation * Math.PI / 180;
                sLayer.addControl(ell);
            } else {
                var rect = new BABYLON.GUI.Rectangle();
                rect.left = (el.x * CANVAS_W) + "px"; rect.top = (el.y * CANVAS_H) + "px";
                rect.width = (el.w * CANVAS_W) + "px"; rect.height = (el.h * CANVAS_H) + "px";
                rect.background = el.fillColor || "transparent";
                rect.thickness = el.borderWidth || 0; rect.color = el.borderColor || "transparent";
                if (ROUND_RECT_SHAPES.indexOf(el.shape) >= 0) {
                    rect.cornerRadius = Math.min(el.w * CANVAS_W, el.h * CANVAS_H) * 0.15;
                }
                rect.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                rect.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                if (el.rotation) rect.rotation = el.rotation * Math.PI / 180;
                sLayer.addControl(rect);
            }
        } else if (el.type === "image") {
            var ir = new BABYLON.GUI.Rectangle();
            ir.left = (el.x * CANVAS_W) + "px"; ir.top = (el.y * CANVAS_H) + "px";
            ir.width = (el.w * CANVAS_W) + "px"; ir.height = (el.h * CANVAS_H) + "px";
            ir.thickness = 0;
            ir.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
            ir.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
            var img = new BABYLON.GUI.Image("i_" + Math.random(), el.dataUrl);
            img.stretch = BABYLON.GUI.Image.STRETCH_FILL;
            ir.addControl(img); sLayer.addControl(ir);
        } else if (el.type === "text") {
            console.log("[RENDER] text: '" + el.text.substring(0, 30) + "' left=" + (el.x * CANVAS_W).toFixed(1) + " top=" + (el.y * CANVAS_H).toFixed(1) + " w=" + (el.w ? el.w * CANVAS_W : "-").toString() + " fs=" + el.fontSize + " color=" + el.color);
            renderTextElement(el, sLayer, CANVAS_W, CANVAS_H, FONT_SCALE);
        }
    });
}

// Build slide thumbnails in the sidebar
export function buildThumbnails(app) {
    if (app.scene.isDisposed) return;
    var thumbC = app.gui.thumbC;
    thumbC.clearControls();
    app.thumbRects = [];
    var TW = 96, TH = 54;
    app.slides.forEach(function (slide, idx) {
        var row = new BABYLON.GUI.StackPanel(); row.isVertical = false;
        row.height = (TH + 4) + "px"; row.width = "120px"; row.paddingBottom = "4px";
        var nt = new BABYLON.GUI.TextBlock(); nt.text = (idx + 1).toString();
        nt.color = "#888"; nt.fontSize = 9; nt.fontFamily = "Segoe UI,sans-serif";
        nt.width = "18px"; nt.height = TH + "px"; row.addControl(nt);

        var th = new BABYLON.GUI.Rectangle(); th.width = TW + "px"; th.height = TH + "px";
        th.background = slide.bg; th.thickness = idx === app.currentSlide ? 2 : 1;
        th.color = idx === app.currentSlide ? PP : "#CCC";
        th.cornerRadius = 1; th.shadowColor = "rgba(0,0,0,0.1)"; th.shadowBlur = 3;
        th.clipChildren = true;
        app.thumbRects.push(th);

        if (slide.bgImage) {
            var ti = new BABYLON.GUI.Image("tbg_" + idx, slide.bgImage);
            ti.stretch = BABYLON.GUI.Image.STRETCH_FILL; th.addControl(ti);
            applyTintOverlay(th, slide.bgTint, "tTint_" + idx);
        }

        // Render mini elements
        slide.elements.forEach(function (el) {
            if (el.type === "shape" && el.shape !== "line" && el.fillColor && el.fillColor !== "transparent") {
                var sr = new BABYLON.GUI.Rectangle();
                sr.width = (el.w * TW) + "px"; sr.height = (el.h * TH) + "px";
                sr.left = (el.x * TW) + "px"; sr.top = (el.y * TH) + "px";
                sr.background = el.fillColor; sr.thickness = 0;
                sr.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                sr.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                th.addControl(sr);
            }
            if (el.type === "image" && el.dataUrl) {
                var ic = new BABYLON.GUI.Rectangle();
                ic.width = (el.w * TW) + "px"; ic.height = (el.h * TH) + "px";
                ic.left = (el.x * TW) + "px"; ic.top = (el.y * TH) + "px";
                ic.thickness = 0;
                ic.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                ic.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                var im = new BABYLON.GUI.Image("ti_" + idx + "_" + Math.random().toString(36).substr(2, 4), el.dataUrl);
                im.stretch = BABYLON.GUI.Image.STRETCH_FILL; ic.addControl(im);
                th.addControl(ic);
            }
            if (el.type === "text" && el.text.trim()) {
                var fs = Math.max(3, Math.round(el.fontSize * TW / CANVAS_W));
                var mi = new BABYLON.GUI.TextBlock();
                mi.text = el.text; mi.fontSize = fs;
                mi.fontWeight = el.fontWeight || "normal";
                mi.color = el.color; mi.fontFamily = "Segoe UI,sans-serif";
                mi.left = (el.x * TW) + "px"; mi.top = (el.y * TH) + "px";
                if (el.w) mi.width = (el.w * TW) + "px";
                mi.textHorizontalAlignment = el.align === "center" ? BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_CENTER :
                    el.align === "right" ? BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_RIGHT :
                    BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                mi.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                mi.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                mi.textWrapping = true;
                th.addControl(mi);
            }
        });

        th.isPointerBlocker = true;
        (function (j) {
            th.onPointerClickObservable.add(function () {
                if (app.scene.isDisposed) return;
                app.currentSlide = j;
                renderSlide(app); updateThumbs(app); updateNotes(app); updateStatus(app);
            });
        })(idx);
        row.addControl(th); thumbC.addControl(row);
    });
}

export function updateThumbs(app) {
    app.thumbRects.forEach(function (t, i) {
        t.thickness = i === app.currentSlide ? 2 : 1;
        t.color = i === app.currentSlide ? PP : "#CCC";
    });
}

export function updateNotes(app) {
    var s = app.slides[app.currentSlide];
    app.gui.notesText.text = s.notes || "";
    app.gui.notesText.color = s.notes ? "#555" : "#BBB";
}

export function updateStatus(app) {
    app.gui.stLeft.text = "Slide " + (app.currentSlide + 1) + " of " + app.slides.length;
}
