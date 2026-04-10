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
    else if (bgTint.type === "artEffect") {
        // artEffect approximation can over-blue many templates; keep it subtle.
        tintRect.background = bgTint.color;
        tintRect.alpha = 0.04;
    }
    else if (bgTint.type === "tint") { tintRect.background = bgTint.color; tintRect.alpha = 0.5; }
    else if (bgTint.type === "alpha") { tintRect.background = "#000000"; tintRect.alpha = 1.0 - bgTint.amt; }
    else return;
    container.addControl(tintRect);
}

function createChevronDataUrl(widthPx, heightPx, fillColor, strokeColor, strokeWidth) {
    var w = Math.max(8, Math.round(widthPx));
    var h = Math.max(8, Math.round(heightPx));
    var canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    var ctx = canvas.getContext("2d");
    if (!ctx) return null;

    // Make arrow tip blunter (closer to ~90deg) by reducing tip depth.
    var tipBaseInset = Math.max(8, Math.round(w * 0.20));
    var notchInset = Math.max(6, Math.round(w * 0.16));
    var sidePad = Math.max(3, Math.round(w * 0.14));
    var xL = sidePad;
    var xR = w - sidePad;
    var y1 = Math.round(h * 0.12);
    var y2 = Math.round(h * 0.88);
    var mid = Math.round(h * 0.5);

    ctx.beginPath();
    ctx.moveTo(xL, y1);
    ctx.lineTo(xR - tipBaseInset, y1);
    ctx.lineTo(xR, mid);
    ctx.lineTo(xR - tipBaseInset, y2);
    ctx.lineTo(xL, y2);
    ctx.lineTo(xL + notchInset, mid);
    ctx.closePath();

    var hasFill = fillColor && fillColor !== "transparent";
    if (hasFill) {
        ctx.fillStyle = fillColor;
        ctx.fill();
    }

    var hasStroke = strokeColor && strokeColor !== "transparent";
    if (hasStroke && strokeWidth > 0) {
        ctx.lineWidth = strokeWidth;
        ctx.strokeStyle = strokeColor;
        ctx.stroke();
    }

    return canvas.toDataURL("image/png");
}

function createRightArrowDataUrl(widthPx, heightPx, fillColor, strokeColor, strokeWidth) {
    var w = Math.max(8, Math.round(widthPx));
    var h = Math.max(8, Math.round(heightPx));
    var canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    var ctx = canvas.getContext("2d");
    if (!ctx) return null;

    // Approximate PPT rightArrow preset with a rectangular tail and triangular head.
    var headW = Math.max(4, Math.round(w * 0.38));
    var shaftH = Math.max(4, Math.round(h * 0.52));
    var yTop = Math.round((h - shaftH) / 2);
    var yBottom = yTop + shaftH;
    var tailRight = Math.max(2, w - headW);
    var midY = Math.round(h / 2);

    ctx.beginPath();
    ctx.moveTo(0, yTop);
    ctx.lineTo(tailRight, yTop);
    ctx.lineTo(tailRight, 0);
    ctx.lineTo(w, midY);
    ctx.lineTo(tailRight, h);
    ctx.lineTo(tailRight, yBottom);
    ctx.lineTo(0, yBottom);
    ctx.closePath();

    var hasFill = fillColor && fillColor !== "transparent";
    if (hasFill) {
        ctx.fillStyle = fillColor;
        ctx.fill();
    }

    var hasStroke = strokeColor && strokeColor !== "transparent";
    if (hasStroke && strokeWidth > 0) {
        ctx.lineWidth = strokeWidth;
        ctx.strokeStyle = strokeColor;
        ctx.stroke();
    }

    return canvas.toDataURL("image/png");
}

function createWedgeRoundRectCalloutDataUrl(widthPx, heightPx, fillColor, strokeColor, strokeWidth, pointLeft) {
    var w = Math.max(16, Math.round(widthPx));
    var h = Math.max(16, Math.round(heightPx));
    var canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    var ctx = canvas.getContext("2d");
    if (!ctx) return null;

    var radius = Math.max(3, Math.round(Math.min(w, h) * 0.06));
    var pad = Math.max(1, Math.round((strokeWidth || 1) * 0.5));
    var x = pad;
    var y = pad;
    var rw = w - pad * 2;
    var rh = h - pad * 2;

    // Keep the tail inside bounds; previous implementation drew the tip outside and got clipped.
    var tailH = Math.max(6, Math.round(h * 0.18));
    var bodyH = Math.max(8, rh - tailH);
    var baseY = y + bodyH;
    var tailBaseCenter = pointLeft ? (x + rw * 0.30) : (x + rw * 0.70);
    var tailBaseHalf = Math.max(4, Math.round(Math.min(w, h) * 0.07));
    var tipX = pointLeft ? (tailBaseCenter - tailBaseHalf * 1.8) : (tailBaseCenter + tailBaseHalf * 1.8);
    tipX = Math.max(x + radius + 1, Math.min(x + rw - radius - 1, tipX));
    var tipY = y + rh;

    ctx.beginPath();
    ctx.moveTo(x + radius, y);
    ctx.lineTo(x + rw - radius, y);
    ctx.quadraticCurveTo(x + rw, y, x + rw, y + radius);
    ctx.lineTo(x + rw, baseY - radius);
    ctx.quadraticCurveTo(x + rw, baseY, x + rw - radius, baseY);

    // Bottom edge with integrated wedge tail
    var rightBaseX = tailBaseCenter + tailBaseHalf;
    var leftBaseX = tailBaseCenter - tailBaseHalf;
    ctx.lineTo(rightBaseX, baseY);
    ctx.lineTo(tipX, tipY);
    ctx.lineTo(leftBaseX, baseY);

    ctx.lineTo(x + radius, baseY);
    ctx.quadraticCurveTo(x, baseY, x, baseY - radius);
    ctx.lineTo(x, y + radius);
    ctx.quadraticCurveTo(x, y, x + radius, y);
    ctx.closePath();

    var hasFill = fillColor && fillColor !== "transparent";
    if (hasFill) {
        ctx.fillStyle = fillColor;
        ctx.fill();
    }

    var hasStroke = strokeColor && strokeColor !== "transparent";
    if (hasStroke && strokeWidth > 0) {
        ctx.lineWidth = strokeWidth;
        ctx.strokeStyle = strokeColor;
        ctx.stroke();
    }

    return canvas.toDataURL("image/png");
}

function normalizeDeg(value) {
    var d = Number(value);
    if (!Number.isFinite(d)) return 0;
    while (d < 0) d += 360;
    while (d >= 360) d -= 360;
    return d;
}

function createPieDataUrl(widthPx, heightPx, fillColor, startDeg, endDeg) {
    var w = Math.max(8, Math.round(widthPx));
    var h = Math.max(8, Math.round(heightPx));
    var canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    var ctx = canvas.getContext("2d");
    if (!ctx) return null;

    var cx = w / 2;
    var cy = h / 2;
    var rx = w / 2;
    var ry = h / 2;
    var s = normalizeDeg(startDeg || 0);
    var e = normalizeDeg(endDeg || 360);
    if (e <= s) e += 360;

    ctx.beginPath();
    ctx.moveTo(cx, cy);
    // OOXML pie angles use 0° at the right, increasing clockwise.
    ctx.ellipse(cx, cy, rx, ry, 0, s * Math.PI / 180, e * Math.PI / 180);
    ctx.closePath();
    ctx.fillStyle = fillColor || "#5B7FC5";
    ctx.fill();

    return canvas.toDataURL("image/png");
}

function wrapTextByCharWidth(text, maxChars) {
    if (!text || !Number.isFinite(maxChars) || maxChars < 1) return text || "";
    var lines = (text || "").split("\n");
    var out = [];
    lines.forEach(function (line) {
        if (line.length <= maxChars) {
            out.push(line);
            return;
        }
        var i = 0;
        while (i < line.length) {
            out.push(line.slice(i, i + maxChars));
            i += maxChars;
        }
    });
    return out.join("\n");
}

// Render a single text element into a GUI container
function renderTextElement(el, container, canvasW, canvasH, fontScale) {
    var tb = new BABYLON.GUI.TextBlock();
    var forceCharWrap = el.wrapMode === "char";
    var forceVCenter = el.valign === "center";
    // Only force CJK break opportunities when there are no natural separators.
    var hasCJK = /[\u3000-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF]/.test(el.text);
    var hasNaturalBreak = /[\s\-\/]/.test(el.text);
    var displayText = ((hasCJK && !hasNaturalBreak) || forceCharWrap)
        ? el.text.replace(/([\u3000-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF])/g, "\u200B$1")
        : el.text;
    var renderFS = Math.round(el.fontSize * fontScale);
    tb.text = displayText; tb.fontSize = renderFS;
    tb.fontWeight = el.fontWeight || "normal"; tb.fontStyle = el.fontStyle || "normal";
    tb.color = el.color;
    // Prefer Japanese Office-like fonts first to keep CJK line breaks closer to PowerPoint.
    tb.fontFamily = "Meiryo UI, Meiryo, Yu Gothic UI, MS PGothic, Segoe UI, Calibri, sans-serif";
    tb.textWrapping = true; tb.resizeToFit = false;
    if (el.w && el.w > 0) {
        var containerW = el.w * canvasW;
        tb.width = containerW + "px";
        var charW = hasCJK ? 1.0 : 0.55;
        if (forceCharWrap && hasCJK) {
            var maxChars = Math.max(1, Math.floor(containerW / Math.max(1, renderFS * charW)));
            displayText = wrapTextByCharWidth(el.text, maxChars);
            tb.text = displayText;
        }
        var estTextW = el.text.length * renderFS * charW;
        var fs = renderFS;
        if (fs >= 12 && estTextW > containerW * 2) {
            fs = Math.max(8, Math.floor(containerW * 2 / (el.text.length * charW)));
            tb.fontSize = fs;
        }
        if (el.h && el.h > 0) {
            tb.height = (el.h * canvasH) + "px";
        } else {
            var lineCount = (tb.text || "").split("\n").length;
            var estLines = Math.max(lineCount, Math.ceil(el.text.length * fs * charW / containerW) || 1);
            var estHeight = Math.max(estLines * fs * 1.4, fs * 1.5);
            tb.height = estHeight + "px";
        }
    }
    if (el.align === "center") tb.textHorizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_CENTER;
    else if (el.align === "right") tb.textHorizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_RIGHT;
    else tb.textHorizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
    tb.textVerticalAlignment = forceVCenter ? BABYLON.GUI.Control.VERTICAL_ALIGNMENT_CENTER : BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
    tb.left = (el.x * canvasW) + "px"; tb.top = (el.y * canvasH) + "px";
    tb.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
    tb.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
    if (el.rotation) tb.rotation = el.rotation * Math.PI / 180;
    container.addControl(tb);
}

// Flip an image using canvas (horizontal or vertical flip)
function getFlippedImageDataUrl(sourceDataUrl, flipH, flipV) {
    return new Promise(function(resolve, reject) {
        var img = new Image();
        img.onload = function() {
            var canvas = document.createElement("canvas");
            canvas.width = img.width;
            canvas.height = img.height;
            var ctx = canvas.getContext("2d");
            
            // Apply flip transforms
            if (flipH || flipV) {
                ctx.translate(flipH ? canvas.width : 0, flipV ? canvas.height : 0);
                ctx.scale(flipH ? -1 : 1, flipV ? -1 : 1);
            }
            ctx.drawImage(img, 0, 0);
            resolve(canvas.toDataURL("image/jpeg", 0.9));
        };
        img.onerror = function() {
            reject(new Error("Failed to load image for flipping"));
        };
        img.src = sourceDataUrl;
    });
}

function getEffectiveImageAlpha(el, slideHasBgImage) {
    var alpha = (typeof el.alpha === "number") ? Math.max(0, Math.min(1, el.alpha)) : 1;
    // Readability guard: full-slide decorative overlays can wash out text too much
    // when rendered with plain alpha blending.
    var isNearFullSlide = el.x <= 0.01 && el.y <= 0.01 && el.w >= 0.99 && el.h >= 0.99;
    if (slideHasBgImage && isNearFullSlide && alpha < 1) {
        alpha = Math.min(alpha, 0.18);
    }
    return alpha;
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
                le.background = el.color || el.strokeColor || "#000"; le.thickness = 0;
                le.left = ((x1 + x2) / 2 - CANVAS_W / 2) + "px";
                le.top = ((y1 + y2) / 2 - CANVAS_H / 2) + "px";
                le.rotation = ang; sLayer.addControl(le);
            } else if (ELLIPSE_SHAPES.indexOf(el.shape) >= 0) {
                if (el.shape === "pie" && Number.isFinite(el.pieStart) && Number.isFinite(el.pieEnd)) {
                    var pie = new BABYLON.GUI.Rectangle();
                    pie.left = (el.x * CANVAS_W) + "px"; pie.top = (el.y * CANVAS_H) + "px";
                    pie.width = (el.w * CANVAS_W) + "px"; pie.height = (el.h * CANVAS_H) + "px";
                    pie.thickness = 0; pie.background = "transparent";
                    pie.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                    pie.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                    if (el.rotation) pie.rotation = el.rotation * Math.PI / 180;
                    var pieData = createPieDataUrl(el.w * CANVAS_W, el.h * CANVAS_H, el.fillColor || "#5B7FC5", el.pieStart, el.pieEnd);
                    if (pieData) {
                        var pieImg = new BABYLON.GUI.Image("pie_" + Math.random(), pieData);
                        pieImg.stretch = BABYLON.GUI.Image.STRETCH_FILL;
                        pie.addControl(pieImg);
                    }
                    sLayer.addControl(pie);
                } else {
                    var ellStrokeW = el.thickness !== undefined ? el.thickness : (el.borderWidth || 0);
                    var ellStrokeColor = el.strokeColor || el.borderColor || "transparent";
                    var ell = new BABYLON.GUI.Ellipse();
                    ell.left = (el.x * CANVAS_W) + "px"; ell.top = (el.y * CANVAS_H) + "px";
                    ell.width = (el.w * CANVAS_W) + "px"; ell.height = (el.h * CANVAS_H) + "px";
                    ell.background = el.fillColor || "transparent";
                    ell.thickness = ellStrokeW; ell.color = ellStrokeColor;
                    ell.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                    ell.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                    if (el.rotation) ell.rotation = el.rotation * Math.PI / 180;
                    sLayer.addControl(ell);
                }
            } else if (el.shape === "chevron") {
                var cheStrokeW = el.thickness !== undefined ? el.thickness : (el.borderWidth || 0);
                var cheStrokeColor = el.strokeColor || el.borderColor || "transparent";
                var che = new BABYLON.GUI.Rectangle();
                che.left = (el.x * CANVAS_W) + "px"; che.top = (el.y * CANVAS_H) + "px";
                che.width = (el.w * CANVAS_W) + "px"; che.height = (el.h * CANVAS_H) + "px";
                che.thickness = 0; che.background = "transparent";
                che.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                che.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                if (el.rotation) che.rotation = el.rotation * Math.PI / 180;

                var cheData = createChevronDataUrl(el.w * CANVAS_W, el.h * CANVAS_H, el.fillColor || "transparent", cheStrokeColor, cheStrokeW);
                if (cheData) {
                    var cheImg = new BABYLON.GUI.Image("chev_" + Math.random(), cheData);
                    cheImg.stretch = BABYLON.GUI.Image.STRETCH_FILL;
                    che.addControl(cheImg);
                }
                sLayer.addControl(che);
            } else if (el.shape === "rightArrow") {
                var arrStrokeW = el.thickness !== undefined ? el.thickness : (el.borderWidth || 0);
                var arrStrokeColor = el.strokeColor || el.borderColor || "transparent";
                var arr = new BABYLON.GUI.Rectangle();
                arr.left = (el.x * CANVAS_W) + "px"; arr.top = (el.y * CANVAS_H) + "px";
                arr.width = (el.w * CANVAS_W) + "px"; arr.height = (el.h * CANVAS_H) + "px";
                arr.thickness = 0; arr.background = "transparent";
                arr.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                arr.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                if (el.rotation) arr.rotation = el.rotation * Math.PI / 180;

                var arrData = createRightArrowDataUrl(el.w * CANVAS_W, el.h * CANVAS_H, el.fillColor || "transparent", arrStrokeColor, arrStrokeW);
                if (arrData) {
                    var arrImg = new BABYLON.GUI.Image("arr_" + Math.random(), arrData);
                    arrImg.stretch = BABYLON.GUI.Image.STRETCH_FILL;
                    arr.addControl(arrImg);
                }
                sLayer.addControl(arr);
            } else if (el.shape === "wedgeRoundRectCallout") {
                var callStrokeW = el.thickness !== undefined ? el.thickness : (el.borderWidth || 0);
                var callStrokeColor = el.strokeColor || el.borderColor || "#888";
                if (!callStrokeW || callStrokeW <= 0) callStrokeW = 1;
                var call = new BABYLON.GUI.Rectangle();
                call.left = (el.x * CANVAS_W) + "px"; call.top = (el.y * CANVAS_H) + "px";
                call.width = (el.w * CANVAS_W) + "px"; call.height = (el.h * CANVAS_H) + "px";
                call.thickness = 0; call.background = "transparent";
                call.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                call.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                if (el.rotation) call.rotation = el.rotation * Math.PI / 180;

                var pointLeft = (el.x || 0) > 0.5;
                var callData = createWedgeRoundRectCalloutDataUrl(
                    el.w * CANVAS_W,
                    el.h * CANVAS_H,
                    el.fillColor || "#FFFFFF",
                    callStrokeColor,
                    callStrokeW,
                    pointLeft
                );
                if (callData) {
                    var callImg = new BABYLON.GUI.Image("call_" + Math.random(), callData);
                    callImg.stretch = BABYLON.GUI.Image.STRETCH_FILL;
                    call.addControl(callImg);
                }
                sLayer.addControl(call);
            } else {
                var rectStrokeW = el.thickness !== undefined ? el.thickness : (el.borderWidth || 0);
                var rectStrokeColor = el.strokeColor || el.borderColor || "transparent";
                var rect = new BABYLON.GUI.Rectangle();
                rect.left = (el.x * CANVAS_W) + "px"; rect.top = (el.y * CANVAS_H) + "px";
                rect.width = (el.w * CANVAS_W) + "px"; rect.height = (el.h * CANVAS_H) + "px";
                rect.background = el.fillColor || "transparent";
                rect.thickness = rectStrokeW; rect.color = rectStrokeColor;
                if (ROUND_RECT_SHAPES.indexOf(el.shape) >= 0) {
                    rect.cornerRadius = Math.min(el.w * CANVAS_W, el.h * CANVAS_H) * 0.15;
                }
                rect.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                rect.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                if (el.rotation) rect.rotation = el.rotation * Math.PI / 180;
                sLayer.addControl(rect);
            }
        } else if (el.type === "image") {
            console.log("[RENDER] image: pos=(" + el.x.toFixed(3) + "," + el.y.toFixed(3) + ") size=(" + el.w.toFixed(3) + "," + el.h.toFixed(3) + ") rotation=" + (el.rotation || 0) + " flipH=" + (el.flipH ? "yes" : "no") + " flipV=" + (el.flipV ? "yes" : "no") + " alpha=" + (el.alpha || 1.0));
            // Use flipped image data URL if flip is requested
            var imgUrl = el.dataUrl;
            if (el.flipH || el.flipV) {
                getFlippedImageDataUrl(el.dataUrl, el.flipH, el.flipV).then(function(flippedUrl) {
                    imgUrl = flippedUrl;
                    renderImageElement(el, imgUrl, sLayer, !!slide.bgImage);
                }).catch(function(err) {
                    console.log("[RENDER] flip failed, using unflipped: " + err.message);
                    renderImageElement(el, el.dataUrl, sLayer, !!slide.bgImage);
                });
            } else {
                renderImageElement(el, imgUrl, sLayer, !!slide.bgImage);
            }
        } else if (el.type === "text") {
            console.log("[RENDER] text: '" + el.text.substring(0, 30) + "' left=" + (el.x * CANVAS_W).toFixed(1) + " top=" + (el.y * CANVAS_H).toFixed(1) + " w=" + (el.w ? el.w * CANVAS_W : "-").toString() + " fs=" + el.fontSize + " color=" + el.color);
            renderTextElement(el, sLayer, CANVAS_W, CANVAS_H, FONT_SCALE);
        }
    });
}

function renderImageElement(el, imgUrl, container, slideHasBgImage) {
    var ir = new BABYLON.GUI.Rectangle();
    ir.left = (el.x * CANVAS_W) + "px"; ir.top = (el.y * CANVAS_H) + "px";
    ir.width = (el.w * CANVAS_W) + "px"; ir.height = (el.h * CANVAS_H) + "px";
    ir.thickness = 0;
    ir.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
    ir.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
    if (el.rotation) ir.rotation = el.rotation * Math.PI / 180;
    var img = new BABYLON.GUI.Image("i_" + Math.random(), imgUrl);
    img.stretch = BABYLON.GUI.Image.STRETCH_FILL;
    img.alpha = getEffectiveImageAlpha(el, !!slideHasBgImage);
    ir.addControl(img); container.addControl(ir);
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
                if (ELLIPSE_SHAPES.indexOf(el.shape) >= 0) {
                    if (el.shape === "pie" && Number.isFinite(el.pieStart) && Number.isFinite(el.pieEnd)) {
                        var pieRect = new BABYLON.GUI.Rectangle();
                        pieRect.width = (el.w * TW) + "px"; pieRect.height = (el.h * TH) + "px";
                        pieRect.left = (el.x * TW) + "px"; pieRect.top = (el.y * TH) + "px";
                        pieRect.thickness = 0; pieRect.background = "transparent";
                        pieRect.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                        pieRect.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                        if (el.rotation) pieRect.rotation = el.rotation * Math.PI / 180;
                        var pieDataThumb = createPieDataUrl(el.w * TW, el.h * TH, el.fillColor || "#5B7FC5", el.pieStart, el.pieEnd);
                        if (pieDataThumb) {
                            var pieImgThumb = new BABYLON.GUI.Image("th_pie_" + idx + "_" + Math.random().toString(36).substr(2, 4), pieDataThumb);
                            pieImgThumb.stretch = BABYLON.GUI.Image.STRETCH_FILL;
                            pieRect.addControl(pieImgThumb);
                        }
                        th.addControl(pieRect);
                    } else {
                        var se = new BABYLON.GUI.Ellipse();
                        se.width = (el.w * TW) + "px"; se.height = (el.h * TH) + "px";
                        se.left = (el.x * TW) + "px"; se.top = (el.y * TH) + "px";
                        se.background = el.fillColor;
                        se.thickness = 0;
                        se.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                        se.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                        if (el.rotation) se.rotation = el.rotation * Math.PI / 180;
                        th.addControl(se);
                    }
                } else if (el.shape === "chevron") {
                    var cheDataThumb = createChevronDataUrl(el.w * TW, el.h * TH, el.fillColor || "transparent", "transparent", 0);
                    if (cheDataThumb) {
                        var sc = new BABYLON.GUI.Rectangle();
                        sc.width = (el.w * TW) + "px"; sc.height = (el.h * TH) + "px";
                        sc.left = (el.x * TW) + "px"; sc.top = (el.y * TH) + "px";
                        sc.thickness = 0; sc.background = "transparent";
                        sc.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                        sc.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                        if (el.rotation) sc.rotation = el.rotation * Math.PI / 180;
                        var simg = new BABYLON.GUI.Image("th_chev_" + idx + "_" + Math.random().toString(36).substr(2, 4), cheDataThumb);
                        simg.stretch = BABYLON.GUI.Image.STRETCH_FILL;
                        sc.addControl(simg);
                        th.addControl(sc);
                    }
                } else if (el.shape === "rightArrow") {
                    var arrDataThumb = createRightArrowDataUrl(el.w * TW, el.h * TH, el.fillColor || "transparent", "transparent", 0);
                    if (arrDataThumb) {
                        var sa = new BABYLON.GUI.Rectangle();
                        sa.width = (el.w * TW) + "px"; sa.height = (el.h * TH) + "px";
                        sa.left = (el.x * TW) + "px"; sa.top = (el.y * TH) + "px";
                        sa.thickness = 0; sa.background = "transparent";
                        sa.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                        sa.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                        if (el.rotation) sa.rotation = el.rotation * Math.PI / 180;
                        var sarr = new BABYLON.GUI.Image("th_arr_" + idx + "_" + Math.random().toString(36).substr(2, 4), arrDataThumb);
                        sarr.stretch = BABYLON.GUI.Image.STRETCH_FILL;
                        sa.addControl(sarr);
                        th.addControl(sa);
                    }
                } else if (el.shape === "wedgeRoundRectCallout") {
                    var pointLeftThumb = (el.x || 0) > 0.5;
                    var callDataThumb = createWedgeRoundRectCalloutDataUrl(
                        el.w * TW,
                        el.h * TH,
                        el.fillColor || "#FFFFFF",
                        "#888",
                        1,
                        pointLeftThumb
                    );
                    if (callDataThumb) {
                        var sb = new BABYLON.GUI.Rectangle();
                        sb.width = (el.w * TW) + "px"; sb.height = (el.h * TH) + "px";
                        sb.left = (el.x * TW) + "px"; sb.top = (el.y * TH) + "px";
                        sb.thickness = 0; sb.background = "transparent";
                        sb.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                        sb.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                        if (el.rotation) sb.rotation = el.rotation * Math.PI / 180;
                        var bimg = new BABYLON.GUI.Image("th_call_" + idx + "_" + Math.random().toString(36).substr(2, 4), callDataThumb);
                        bimg.stretch = BABYLON.GUI.Image.STRETCH_FILL;
                        sb.addControl(bimg);
                        th.addControl(sb);
                    }
                } else {
                    var sr = new BABYLON.GUI.Rectangle();
                    sr.width = (el.w * TW) + "px"; sr.height = (el.h * TH) + "px";
                    sr.left = (el.x * TW) + "px"; sr.top = (el.y * TH) + "px";
                    sr.background = el.fillColor; sr.thickness = 0;
                    sr.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                    sr.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                    if (el.rotation) sr.rotation = el.rotation * Math.PI / 180;
                    th.addControl(sr);
                }
            }
            if (el.type === "image" && el.dataUrl) {
                var thumbImgUrl = el.dataUrl;
                var renderThumbImage = function(url) {
                    var ic = new BABYLON.GUI.Rectangle();
                    ic.width = (el.w * TW) + "px"; ic.height = (el.h * TH) + "px";
                    ic.left = (el.x * TW) + "px"; ic.top = (el.y * TH) + "px";
                    ic.thickness = 0;
                    ic.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                    ic.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                    if (el.rotation) ic.rotation = el.rotation * Math.PI / 180;
                    var im = new BABYLON.GUI.Image("ti_" + idx + "_" + Math.random().toString(36).substr(2, 4), url);
                    im.stretch = BABYLON.GUI.Image.STRETCH_FILL;
                    if (typeof el.alpha === "number") im.alpha = Math.max(0, Math.min(1, el.alpha));
                    ic.addControl(im);
                    th.addControl(ic);
                };
                if (el.flipH || el.flipV) {
                    getFlippedImageDataUrl(el.dataUrl, el.flipH, el.flipV).then(function(flippedUrl) {
                        renderThumbImage(flippedUrl);
                    }).catch(function(err) {
                        console.log("[THUMB] flip failed, using unflipped: " + err.message);
                        renderThumbImage(el.dataUrl);
                    });
                } else {
                    renderThumbImage(thumbImgUrl);
                }
            }
            if (el.type === "text" && el.text.trim()) {
                // Use the same text layout path as main slide to avoid thumbnail-only drift.
                var thumbFontScale = FONT_SCALE * (TW / CANVAS_W);
                renderTextElement(el, th, TW, TH, thumbFontScale);
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
