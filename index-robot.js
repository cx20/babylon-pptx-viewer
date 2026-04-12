// ============================================================================
// index-robot.js - Entry point for Robot PowerPoint Viewer
// Robot model flips through slides with animation
// ============================================================================

import { getDefaultSlides, CANVAS_W, CANVAS_H, TEX_W, TEX_H, PP, FONT_SCALE } from "./constants.js";
import { setupRobotScene, animatePageFlip } from "./robot-scene-setup.js";
import { parsePptx } from "./pptx-parser.js";

// Suppress verbose parser/render debug logs in production.
(function configureDebugLogging() {
    if (typeof window === "undefined" || !window.console || typeof window.console.log !== "function") return;
    if (window.__PPTX_DEBUG__) return;

    var rawLog = window.console.log.bind(window.console);
    var debugPrefixes = /^(\[PPTX\]|\[RENDER\]|\[SP\]|\[TREE\]|\[PIC\]|\[BG\]|\[BLIP\]|\[LAYOUT\]|\[MASTER\]|\[GRPSP\]|\[GF\]|\[SmartArt\]|\[TBL\]|\[CXN\]|\[INIT\/ENGINE\]|\[INIT\/SCENE\]|\[INIT\/UI\]|\[INIT\/INPUT\]|\[INIT\/JSZIP\]|\[INIT\] boot sequence)/;

    window.console.log = function () {
        if (arguments.length > 0 && typeof arguments[0] === "string" && debugPrefixes.test(arguments[0])) {
            return;
        }
        return rawLog.apply(window.console, arguments);
    };
})();

var canvas = document.getElementById("renderCanvas");

// ============================================================================
// Error handling
// ============================================================================

var AppError = function(code, userMsg, devMsg) {
    this.code = code;
    this.userMsg = userMsg;
    this.devMsg = devMsg;
};
AppError.prototype = Object.create(Error.prototype);
AppError.prototype.constructor = AppError;

function resolveJsZipConstructor() {
    if (typeof globalThis === "undefined") return null;
    var candidate = globalThis.JSZip;
    if (typeof candidate === "function") return candidate;
    if (candidate && typeof candidate.default === "function") return candidate.default;
    if (candidate && typeof candidate.JSZip === "function") return candidate.JSZip;
    return null;
}

function ensureJsZipLoaded() {
    var JsZipCtor = resolveJsZipConstructor();
    if (!JsZipCtor) {
        throw new AppError(
            'JSZIP_LOAD_FAIL',
            'File library (JSZip) not available.',
            'JSZip constructor could not be resolved'
        );
    }
    globalThis.JSZip = JsZipCtor;

    var hasLoadAsync = typeof JsZipCtor.loadAsync === "function";
    if (!hasLoadAsync) {
        var testZip = new JsZipCtor();
        hasLoadAsync = typeof testZip.loadAsync === "function";
    }
    if (!hasLoadAsync) {
        throw new AppError(
            'JSZIP_LOAD_FAIL',
            'File library loaded but is incompatible.',
            'JSZip does not expose loadAsync'
        );
    }
    console.log("[INIT/JSZIP] JSZip validated OK");
}

// ============================================================================
// Slide texture rendering
// ============================================================================

/**
 * Preloads an image from a src (URL or Data URL).
 * Returns a Promise that resolves to the loaded Image or null on error.
 */
function preloadImage(src) {
    return new Promise(function(resolve) {
        if (!src) {
            console.log("[ROBOT-DEBUG] preloadImage: src is empty/null");
            resolve(null);
            return;
        }
        console.log("[ROBOT-DEBUG] preloadImage: loading src type=" + typeof src + ", length=" + src.length + ", prefix=" + src.substring(0, 50));
        var img = new Image();
        img.crossOrigin = "anonymous";
        img.onload = function() {
            console.log("[ROBOT-DEBUG] preloadImage: SUCCESS loaded " + img.width + "x" + img.height);
            resolve(img);
        };
        img.onerror = function(e) {
            console.warn("[ROBOT-DEBUG] preloadImage: FAILED to load image:", src.substring(0, 100), e);
            resolve(null);
        };
        img.src = src;
    });
}

/**
 * Creates a DynamicTexture from PPTX slide data and applies it to the slide mesh.
 * Supports:
 * - Solid and gradient backgrounds
 * - Background images
 * - Text elements with font styling
 * - Shape elements (rectangle, ellipse, line)
 * - Image elements
 */
async function createSlideTexture(slide, scene, slideMesh) {
    console.log("[ROBOT-DEBUG] createSlideTexture called");
    console.log("[ROBOT-DEBUG] slide object:", JSON.stringify({
        bg: slide.bg,
        bgImage: slide.bgImage ? (typeof slide.bgImage === "string" ? "string(len=" + slide.bgImage.length + ")" : typeof slide.bgImage) : null,
        elementCount: slide.elements ? slide.elements.length : 0
    }));
    console.log("[ROBOT-DEBUG] slideMesh:", slideMesh ? slideMesh.name : "null");
    console.log("[ROBOT-DEBUG] slideMesh.material:", slideMesh && slideMesh.material ? slideMesh.material.name : "null");

    var texWidth = 1024;
    var texHeight = 768;
    
    var dynTex = new BABYLON.DynamicTexture("slideTex", { width: texWidth, height: texHeight }, scene, true);
    var ctx = dynTex.getContext();
    
    // Clear and draw background
    ctx.fillStyle = slide.bg || "#FFFFFF";
    ctx.fillRect(0, 0, texWidth, texHeight);
    console.log("[ROBOT-DEBUG] Background filled with:", slide.bg || "#FFFFFF");
    
    // Draw background image if present
    // bgImage can be a string (Data URL) or an object with .src/.image property
    var bgSrc = null;
    if (slide.bgImage) {
        console.log("[ROBOT-DEBUG] slide.bgImage exists, type=" + typeof slide.bgImage);
        if (typeof slide.bgImage === "string") {
            bgSrc = slide.bgImage;
            console.log("[ROBOT-DEBUG] bgImage is string, length=" + bgSrc.length);
        } else if (slide.bgImage.src) {
            bgSrc = slide.bgImage.src;
            console.log("[ROBOT-DEBUG] bgImage.src found, length=" + bgSrc.length);
        } else if (slide.bgImage.image) {
            bgSrc = slide.bgImage.image;
            console.log("[ROBOT-DEBUG] bgImage.image found, length=" + bgSrc.length);
        } else {
            console.log("[ROBOT-DEBUG] bgImage has unknown structure:", Object.keys(slide.bgImage));
        }
    } else {
        console.log("[ROBOT-DEBUG] No bgImage on slide");
    }
    
    if (bgSrc) {
        console.log("[ROBOT-DEBUG] Loading background image...");
        var bgImg = await preloadImage(bgSrc);
        if (bgImg) {
            ctx.drawImage(bgImg, 0, 0, texWidth, texHeight);
            console.log("[ROBOT-DEBUG] Background image drawn to canvas");
        } else {
            console.log("[ROBOT-DEBUG] Background image failed to load");
        }
    }
    
    // Helper to get image source (parser uses dataUrl, some may use src)
    function getImageSrc(el) {
        return el.dataUrl || el.src || null;
    }
    
    // Collect all image elements that need preloading
    var imageElements = [];
    if (slide.elements && Array.isArray(slide.elements)) {
        slide.elements.forEach(function(el) {
            if (el.type === "image") {
                var imgSrc = getImageSrc(el);
                if (imgSrc) {
                    imageElements.push({ el: el, src: imgSrc });
                } else {
                    console.log("[ROBOT-DEBUG] Image element has no src/dataUrl:", JSON.stringify({x: el.x, y: el.y, w: el.w, h: el.h, rId: el.rId}));
                }
            }
        });
    }
    console.log("[ROBOT-DEBUG] Image elements to load:", imageElements.length);
    
    // Log all element types  
    if (slide.elements && Array.isArray(slide.elements)) {
        var typeCounts = {};
        slide.elements.forEach(function(el) {
            typeCounts[el.type] = (typeCounts[el.type] || 0) + 1;
        });
        console.log("[ROBOT-DEBUG] Element types:", JSON.stringify(typeCounts));
    }
    
    // Preload all images in parallel
    var loadedImages = {};
    if (imageElements.length > 0) {
        var imagePromises = imageElements.map(function(item) {
            return preloadImage(item.src).then(function(img) {
                if (img) {
                    loadedImages[item.src] = img;
                }
            });
        });
        await Promise.all(imagePromises);
        console.log("[ROBOT-DEBUG] Loaded images count:", Object.keys(loadedImages).length);
    }
    
    // Separate elements by type for proper layering
    // Draw order: shapes (background) -> images -> text (foreground)
    var shapeElements = [];
    var imageElementsArray = [];
    var textElements = [];
    
    if (slide.elements && Array.isArray(slide.elements)) {
        slide.elements.forEach(function(el, idx) {
            el._origIdx = idx; // Preserve original order within same type
            if (el.type === "shape") {
                shapeElements.push(el);
            } else if (el.type === "image") {
                imageElementsArray.push(el);
            } else if (el.type === "text") {
                textElements.push(el);
            }
        });
    }
    
    console.log("[ROBOT-DEBUG] Elements: shapes=" + shapeElements.length + " images=" + imageElementsArray.length + " text=" + textElements.length);
    
    // Draw images first (background layer)
    imageElementsArray.forEach(function(el) {
        var imgSrc = getImageSrc(el);
        if (imgSrc) {
            var cachedImg = loadedImages[imgSrc];
            if (cachedImg) {
                renderImageToCanvasWithImg(el, ctx, texWidth, texHeight, cachedImg);
            }
        }
    });
    
    // Draw shapes on top of images
    shapeElements.forEach(function(el) {
        renderShapeToCanvas(el, ctx, texWidth, texHeight);
    });
    
    // Draw text last (foreground) - filter out duplicate/overlapping text and adjust Y for overlap
    var renderedTextRegions = [];
    textElements.forEach(function(el, idx) {
        // Skip text that would overlap with already drawn text at same position (duplicate detection)
        var isDuplicate = renderedTextRegions.some(function(prev) {
            return Math.abs(prev.x - el.x) < 0.005 && 
                   Math.abs(prev.origY - el.y) < 0.005 &&
                   Math.abs(prev.w - (el.w || 0)) < 0.01;
        });
        
        if (isDuplicate) {
            console.log("[ROBOT-DEBUG] Skipping duplicate text at (" + el.x.toFixed(3) + "," + el.y.toFixed(3) + "): '" + (el.text||"").substring(0,30) + "'");
            return;
        }
        
        // Calculate yOffset to avoid overlapping with previously rendered text
        var yOffset = 0;
        var elX = el.x;
        var elY = el.y;
        var elW = el.w || 0;
        
        renderedTextRegions.forEach(function(prev) {
            // Check if in the same horizontal region (overlapping X range)
            var xOverlap = !(prev.x + prev.w < elX || elX + elW < prev.x);
            // Only shift text that was meant to appear below the previous text (based on original Y)
            var isSequentialText = elY >= prev.origY;
            if (xOverlap && isSequentialText) {
                // Check if this text's Y position would overlap with previous text's rendered area
                // If current text starts before previous text ends, we need to shift down
                if (elY + yOffset < prev.endY) {
                    // Shift down to below previously rendered text
                    var newOffset = prev.endY - elY + 0.01; // Add small gap
                    if (newOffset > yOffset) {
                        yOffset = newOffset;
                    }
                }
            }
        });
        
        var result = renderTextToCanvas(el, ctx, texWidth, texHeight, yOffset);
        if (result) {
            renderedTextRegions.push({
                x: result.x,
                origY: el.y,
                startY: result.startY,
                endY: result.endY,
                w: result.w,
                height: result.height
            });
        }
    });
    
    dynTex.update();
    console.log("[ROBOT-DEBUG] DynamicTexture updated");
    
    // Apply to slide mesh
    if (slideMesh && slideMesh.material) {
        slideMesh.material.diffuseTexture = dynTex;
        console.log("[ROBOT-DEBUG] Texture applied to slideMesh.material.diffuseTexture");
    } else {
        console.log("[ROBOT-DEBUG] WARNING: Cannot apply texture - slideMesh:", !!slideMesh, "material:", !!(slideMesh && slideMesh.material));
    }
    
    return dynTex;
}

/**
 * Wraps text to fit within a given width using Canvas 2D measureText.
 * Handles both CJK characters (character-by-character wrap) and
 * Western text (word-by-word wrap).
 */
function wrapTextToWidth(ctx, text, maxWidth) {
    if (!text || maxWidth <= 0) return [text || ""];
    
    var lines = [];
    var paragraphs = text.split("\n");
    
    paragraphs.forEach(function(paragraph) {
        if (!paragraph) {
            lines.push("");
            return;
        }
        
        // Check if text contains CJK characters
        var hasCJK = /[\u3000-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF]/.test(paragraph);
        
        if (hasCJK) {
            // For CJK text, wrap character by character
            var currentLine = "";
            for (var i = 0; i < paragraph.length; i++) {
                var char = paragraph[i];
                var testLine = currentLine + char;
                var metrics = ctx.measureText(testLine);
                
                if (metrics.width > maxWidth && currentLine.length > 0) {
                    lines.push(currentLine);
                    currentLine = char;
                } else {
                    currentLine = testLine;
                }
            }
            if (currentLine) {
                lines.push(currentLine);
            }
        } else {
            // For Western text, wrap word by word
            var words = paragraph.split(/(\s+)/); // Keep whitespace
            var currentLine = "";
            
            for (var j = 0; j < words.length; j++) {
                var word = words[j];
                var testLine = currentLine + word;
                var metrics = ctx.measureText(testLine);
                
                if (metrics.width > maxWidth && currentLine.length > 0) {
                    lines.push(currentLine.trim());
                    currentLine = word;
                } else {
                    currentLine = testLine;
                }
            }
            if (currentLine) {
                lines.push(currentLine.trim());
            }
        }
    });
    
    return lines.length > 0 ? lines : [""];
}

/**
 * Renders text to canvas and returns the actual height used.
 * This allows subsequent text elements to be positioned correctly.
 */
function renderTextToCanvas(el, ctx, canvasW, canvasH, yOffset) {
    yOffset = yOffset || 0;
    var x = (el.x || 0) * canvasW;
    var y = ((el.y || 0) + yOffset) * canvasH; // yOffset is in normalized coordinates
    var w = (el.w || 1) * canvasW;
    var h = (el.h || 0) * canvasH;
    var fontSize = Math.round((el.fontSize || 16) * 1.5);
    
    // Set font before measuring
    var fontStyle = (el.fontWeight === "bold" ? "bold " : "") + 
                    (el.fontStyle === "italic" ? "italic " : "") +
                    fontSize + "px " + (el.fontFamily || "Meiryo UI, Meiryo, Arial, sans-serif");
    ctx.font = fontStyle;
    ctx.fillStyle = el.color || "#000000";
    ctx.textBaseline = "top";
    
    // Wrap text to fit width
    var text = el.text || "";
    var wrappedLines = wrapTextToWidth(ctx, text, w);
    var lineHeight = fontSize * 1.3;
    
    // Calculate total text height for vertical alignment
    var totalTextHeight = wrappedLines.length * lineHeight;
    var startY = y;
    
    // Vertical alignment
    if (el.valign === "center" && h > 0) {
        startY = y + (h - totalTextHeight) / 2;
    } else if (el.valign === "bottom" && h > 0) {
        startY = y + h - totalTextHeight;
    }
    
    // Draw each wrapped line
    wrappedLines.forEach(function(line, i) {
        var drawX = x;
        
        // Horizontal alignment
        if (el.align === "center") {
            var lineWidth = ctx.measureText(line).width;
            drawX = x + (w - lineWidth) / 2;
        } else if (el.align === "right") {
            var lineWidth = ctx.measureText(line).width;
            drawX = x + w - lineWidth;
        }
        
        ctx.fillText(line, drawX, startY + i * lineHeight);
    });
    
    // Return info about where text was drawn in normalized coordinates (for overlap detection)
    return {
        startY: startY / canvasH,
        endY: (startY + totalTextHeight) / canvasH,
        height: totalTextHeight / canvasH,
        x: (el.x || 0),
        w: (el.w || 1)
    };
}

// Ellipse shapes list (from constants.js)
var ELLIPSE_SHAPES = ["ellipse", "oval", "circle", "pie", "arc", "chord", "donut"];
var ROUND_RECT_SHAPES = ["roundRect", "snipRoundRect", "snip1Rect", "snip2SameRect", "round1Rect", "round2SameRect"];

// Draw chevron shape path
function drawChevronPath(ctx, x, y, w, h) {
    var tipBaseInset = Math.max(8, Math.round(w * 0.20));
    var notchInset = Math.max(6, Math.round(w * 0.16));
    var sidePad = Math.max(3, Math.round(w * 0.14));
    var xL = x + sidePad;
    var xR = x + w - sidePad;
    var y1 = y + Math.round(h * 0.12);
    var y2 = y + Math.round(h * 0.88);
    var mid = y + Math.round(h * 0.5);

    ctx.beginPath();
    ctx.moveTo(xL, y1);
    ctx.lineTo(xR - tipBaseInset, y1);
    ctx.lineTo(xR, mid);
    ctx.lineTo(xR - tipBaseInset, y2);
    ctx.lineTo(xL, y2);
    ctx.lineTo(xL + notchInset, mid);
    ctx.closePath();
}

// Draw right arrow shape path
function drawRightArrowPath(ctx, x, y, w, h) {
    var headW = Math.max(4, Math.round(w * 0.38));
    var shaftH = Math.max(4, Math.round(h * 0.52));
    var yTop = y + Math.round((h - shaftH) / 2);
    var yBottom = yTop + shaftH;
    var tailRight = x + Math.max(2, w - headW);
    var midY = y + Math.round(h / 2);

    ctx.beginPath();
    ctx.moveTo(x, yTop);
    ctx.lineTo(tailRight, yTop);
    ctx.lineTo(tailRight, y);
    ctx.lineTo(x + w, midY);
    ctx.lineTo(tailRight, y + h);
    ctx.lineTo(tailRight, yBottom);
    ctx.lineTo(x, yBottom);
    ctx.closePath();
}

// Draw left arrow shape path
function drawLeftArrowPath(ctx, x, y, w, h) {
    var headW = Math.max(4, Math.round(w * 0.38));
    var shaftH = Math.max(4, Math.round(h * 0.52));
    var yTop = y + Math.round((h - shaftH) / 2);
    var yBottom = yTop + shaftH;
    var tailLeft = x + headW;
    var midY = y + Math.round(h / 2);

    ctx.beginPath();
    ctx.moveTo(x + w, yTop);
    ctx.lineTo(tailLeft, yTop);
    ctx.lineTo(tailLeft, y);
    ctx.lineTo(x, midY);
    ctx.lineTo(tailLeft, y + h);
    ctx.lineTo(tailLeft, yBottom);
    ctx.lineTo(x + w, yBottom);
    ctx.closePath();
}

// Draw up arrow shape path
function drawUpArrowPath(ctx, x, y, w, h) {
    var headH = Math.max(4, Math.round(h * 0.38));
    var shaftW = Math.max(4, Math.round(w * 0.52));
    var xLeft = x + Math.round((w - shaftW) / 2);
    var xRight = xLeft + shaftW;
    var tailTop = y + headH;
    var midX = x + Math.round(w / 2);

    ctx.beginPath();
    ctx.moveTo(xLeft, y + h);
    ctx.lineTo(xLeft, tailTop);
    ctx.lineTo(x, tailTop);
    ctx.lineTo(midX, y);
    ctx.lineTo(x + w, tailTop);
    ctx.lineTo(xRight, tailTop);
    ctx.lineTo(xRight, y + h);
    ctx.closePath();
}

// Draw down arrow shape path
function drawDownArrowPath(ctx, x, y, w, h) {
    var headH = Math.max(4, Math.round(h * 0.38));
    var shaftW = Math.max(4, Math.round(w * 0.52));
    var xLeft = x + Math.round((w - shaftW) / 2);
    var xRight = xLeft + shaftW;
    var tailBottom = y + h - headH;
    var midX = x + Math.round(w / 2);

    ctx.beginPath();
    ctx.moveTo(xLeft, y);
    ctx.lineTo(xLeft, tailBottom);
    ctx.lineTo(x, tailBottom);
    ctx.lineTo(midX, y + h);
    ctx.lineTo(x + w, tailBottom);
    ctx.lineTo(xRight, tailBottom);
    ctx.lineTo(xRight, y);
    ctx.closePath();
}

// Draw pentagon shape path
function drawPentagonPath(ctx, x, y, w, h) {
    var cx = x + w / 2;
    ctx.beginPath();
    ctx.moveTo(cx, y);
    ctx.lineTo(x + w, y + h * 0.38);
    ctx.lineTo(x + w * 0.82, y + h);
    ctx.lineTo(x + w * 0.18, y + h);
    ctx.lineTo(x, y + h * 0.38);
    ctx.closePath();
}

// Draw hexagon shape path
function drawHexagonPath(ctx, x, y, w, h) {
    var inset = w * 0.25;
    ctx.beginPath();
    ctx.moveTo(x + inset, y);
    ctx.lineTo(x + w - inset, y);
    ctx.lineTo(x + w, y + h / 2);
    ctx.lineTo(x + w - inset, y + h);
    ctx.lineTo(x + inset, y + h);
    ctx.lineTo(x, y + h / 2);
    ctx.closePath();
}

// Draw triangle shape path
function drawTrianglePath(ctx, x, y, w, h) {
    ctx.beginPath();
    ctx.moveTo(x + w / 2, y);
    ctx.lineTo(x + w, y + h);
    ctx.lineTo(x, y + h);
    ctx.closePath();
}

// Draw diamond shape path
function drawDiamondPath(ctx, x, y, w, h) {
    ctx.beginPath();
    ctx.moveTo(x + w / 2, y);
    ctx.lineTo(x + w, y + h / 2);
    ctx.lineTo(x + w / 2, y + h);
    ctx.lineTo(x, y + h / 2);
    ctx.closePath();
}

// Draw parallelogram shape path
function drawParallelogramPath(ctx, x, y, w, h) {
    var skew = w * 0.2;
    ctx.beginPath();
    ctx.moveTo(x + skew, y);
    ctx.lineTo(x + w, y);
    ctx.lineTo(x + w - skew, y + h);
    ctx.lineTo(x, y + h);
    ctx.closePath();
}

// Draw trapezoid shape path
function drawTrapezoidPath(ctx, x, y, w, h) {
    var inset = w * 0.2;
    ctx.beginPath();
    ctx.moveTo(x + inset, y);
    ctx.lineTo(x + w - inset, y);
    ctx.lineTo(x + w, y + h);
    ctx.lineTo(x, y + h);
    ctx.closePath();
}

// Draw plus/cross shape path
function drawPlusPath(ctx, x, y, w, h) {
    var armW = w / 3;
    var armH = h / 3;
    ctx.beginPath();
    ctx.moveTo(x + armW, y);
    ctx.lineTo(x + 2 * armW, y);
    ctx.lineTo(x + 2 * armW, y + armH);
    ctx.lineTo(x + w, y + armH);
    ctx.lineTo(x + w, y + 2 * armH);
    ctx.lineTo(x + 2 * armW, y + 2 * armH);
    ctx.lineTo(x + 2 * armW, y + h);
    ctx.lineTo(x + armW, y + h);
    ctx.lineTo(x + armW, y + 2 * armH);
    ctx.lineTo(x, y + 2 * armH);
    ctx.lineTo(x, y + armH);
    ctx.lineTo(x + armW, y + armH);
    ctx.closePath();
}

// Draw star shape path (5-pointed)
function drawStarPath(ctx, x, y, w, h) {
    var cx = x + w / 2;
    var cy = y + h / 2;
    var outerR = Math.min(w, h) / 2;
    var innerR = outerR * 0.38;
    var points = 5;
    
    ctx.beginPath();
    for (var i = 0; i < points * 2; i++) {
        var r = i % 2 === 0 ? outerR : innerR;
        var angle = (i * Math.PI / points) - Math.PI / 2;
        var px = cx + r * Math.cos(angle);
        var py = cy + r * Math.sin(angle);
        if (i === 0) ctx.moveTo(px, py);
        else ctx.lineTo(px, py);
    }
    ctx.closePath();
}

// Draw rounded rectangle path
function drawRoundRectPath(ctx, x, y, w, h, radius) {
    radius = Math.min(radius || 8, w / 2, h / 2);
    ctx.beginPath();
    ctx.moveTo(x + radius, y);
    ctx.lineTo(x + w - radius, y);
    ctx.quadraticCurveTo(x + w, y, x + w, y + radius);
    ctx.lineTo(x + w, y + h - radius);
    ctx.quadraticCurveTo(x + w, y + h, x + w - radius, y + h);
    ctx.lineTo(x + radius, y + h);
    ctx.quadraticCurveTo(x, y + h, x, y + h - radius);
    ctx.lineTo(x, y + radius);
    ctx.quadraticCurveTo(x, y, x + radius, y);
    ctx.closePath();
}

// Draw callout (wedgeRoundRectCallout) shape path
function drawCalloutPath(ctx, x, y, w, h, pointLeft) {
    var radius = Math.max(3, Math.round(Math.min(w, h) * 0.06));
    var tailH = Math.max(6, Math.round(h * 0.18));
    var bodyH = Math.max(8, h - tailH);
    var baseY = y + bodyH;
    var tailBaseCenter = pointLeft ? (x + w * 0.30) : (x + w * 0.70);
    var tailBaseHalf = Math.max(4, Math.round(Math.min(w, h) * 0.07));
    var tipX = pointLeft ? (tailBaseCenter - tailBaseHalf * 1.8) : (tailBaseCenter + tailBaseHalf * 1.8);
    tipX = Math.max(x + radius + 1, Math.min(x + w - radius - 1, tipX));
    var tipY = y + h;

    ctx.beginPath();
    ctx.moveTo(x + radius, y);
    ctx.lineTo(x + w - radius, y);
    ctx.quadraticCurveTo(x + w, y, x + w, y + radius);
    ctx.lineTo(x + w, baseY - radius);
    ctx.quadraticCurveTo(x + w, baseY, x + w - radius, baseY);

    // Bottom edge with tail
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
}

// Draw pie/arc shape
function drawPiePath(ctx, x, y, w, h, startDeg, endDeg) {
    var cx = x + w / 2;
    var cy = y + h / 2;
    var rx = w / 2;
    var ry = h / 2;
    
    // Normalize angles
    var s = startDeg || 0;
    var e = endDeg || 360;
    while (s < 0) s += 360;
    while (s >= 360) s -= 360;
    while (e < 0) e += 360;
    if (e <= s) e += 360;

    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.ellipse(cx, cy, rx, ry, 0, s * Math.PI / 180, e * Math.PI / 180);
    ctx.closePath();
}

function renderShapeToCanvas(el, ctx, canvasW, canvasH) {
    var x = (el.x || 0) * canvasW;
    var y = (el.y || 0) * canvasH;
    var w = (el.w || 0.1) * canvasW;
    var h = (el.h || 0.1) * canvasH;
    
    // Handle different property naming conventions
    var fillColor = el.fillColor || el.fill || el.color || "transparent";
    var strokeColor = el.strokeColor || el.borderColor || el.stroke || "transparent";
    var strokeWidth = el.thickness !== undefined ? el.thickness : (el.borderWidth || 1);
    
    ctx.fillStyle = fillColor;
    ctx.strokeStyle = strokeColor;
    ctx.lineWidth = strokeWidth;
    
    var shape = (el.shape || "rect").toLowerCase();
    
    // For shapes with rotation, we need to transform coordinates
    var hasRotation = el.rotation && el.rotation !== 0;
    var drawX = x;
    var drawY = y;
    
    if (hasRotation) {
        ctx.save();
        // Translate to center of shape, rotate, then draw from -w/2, -h/2
        ctx.translate(x + w / 2, y + h / 2);
        ctx.rotate(el.rotation * Math.PI / 180);
        // After this transform, we draw with origin at center
        drawX = -w / 2;
        drawY = -h / 2;
    }
    
    if (shape === "line") {
        var x1 = (el.x1 || el.x || 0) * canvasW;
        var y1 = (el.y1 || el.y || 0) * canvasH;
        var x2 = (el.x2 || el.x + el.w || 0) * canvasW;
        var y2 = (el.y2 || el.y + el.h || 0) * canvasH;
        
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();
    } else if (shape.indexOf("bentconnector") >= 0 || shape.indexOf("curvedconnector") >= 0) {
        // Draw bent/curved connector as a curved line between opposite corners
        var startX = el.flipH ? x + w : x;
        var startY = el.flipV ? y + h : y;
        var endX = el.flipH ? x : x + w;
        var endY = el.flipV ? y : y + h;
        
        // Calculate control points for a smooth S-curve
        var midX = (startX + endX) / 2;
        var midY = (startY + endY) / 2;
        
        ctx.beginPath();
        ctx.moveTo(startX, startY);
        
        if (shape.indexOf("curved") >= 0) {
            // Curved connector - use quadratic bezier through midpoint
            var cx1 = startX;
            var cy1 = midY;
            var cx2 = endX;
            var cy2 = midY;
            ctx.bezierCurveTo(cx1, cy1, cx2, cy2, endX, endY);
        } else {
            // Bent connector - use two line segments with a bend
            if (Math.abs(endX - startX) > Math.abs(endY - startY)) {
                // Horizontal dominant - bend at midX
                ctx.lineTo(midX, startY);
                ctx.lineTo(midX, endY);
            } else {
                // Vertical dominant - bend at midY
                ctx.lineTo(startX, midY);
                ctx.lineTo(endX, midY);
            }
            ctx.lineTo(endX, endY);
        }
        
        ctx.lineWidth = strokeWidth > 0 ? strokeWidth : 2;
        ctx.strokeStyle = strokeColor !== "transparent" ? strokeColor : fillColor;
        ctx.stroke();
        
        // Draw arrowhead at end
        var arrowSize = Math.max(8, strokeWidth * 3);
        var angle = Math.atan2(endY - midY, endX - midX);
        ctx.beginPath();
        ctx.moveTo(endX, endY);
        ctx.lineTo(endX - arrowSize * Math.cos(angle - Math.PI/6), endY - arrowSize * Math.sin(angle - Math.PI/6));
        ctx.lineTo(endX - arrowSize * Math.cos(angle + Math.PI/6), endY - arrowSize * Math.sin(angle + Math.PI/6));
        ctx.closePath();
        ctx.fillStyle = strokeColor !== "transparent" ? strokeColor : fillColor;
        ctx.fill();
    } else if (shape === "chevron") {
        drawChevronPath(ctx, drawX, drawY, w, h);
        ctx.fillStyle = fillColor;
        ctx.strokeStyle = strokeColor;
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (shape === "rightarrow") {
        // For rotated arrows, we need to save context, translate to center, rotate, and draw
        if (hasRotation) {
            // Already in rotated context (centered), draw at drawX, drawY
            drawRightArrowPath(ctx, drawX, drawY, w, h);
        } else {
            drawRightArrowPath(ctx, x, y, w, h);
        }
        ctx.fillStyle = fillColor;
        ctx.strokeStyle = strokeColor;
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (shape === "leftarrow") {
        drawLeftArrowPath(ctx, drawX, drawY, w, h);
        ctx.fillStyle = fillColor;
        ctx.strokeStyle = strokeColor;
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (shape === "uparrow") {
        drawUpArrowPath(ctx, drawX, drawY, w, h);
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (shape === "downarrow") {
        drawDownArrowPath(ctx, drawX, drawY, w, h);
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (shape === "pentagon") {
        drawPentagonPath(ctx, drawX, drawY, w, h);
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (shape === "hexagon") {
        drawHexagonPath(ctx, drawX, drawY, w, h);
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (shape === "triangle" || shape === "isoctriangle") {
        drawTrianglePath(ctx, drawX, drawY, w, h);
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (shape === "diamond") {
        drawDiamondPath(ctx, drawX, drawY, w, h);
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (shape === "parallelogram") {
        drawParallelogramPath(ctx, drawX, drawY, w, h);
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (shape === "trapezoid") {
        drawTrapezoidPath(ctx, drawX, drawY, w, h);
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (shape === "plus" || shape === "cross" || shape === "mathplus") {
        drawPlusPath(ctx, drawX, drawY, w, h);
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (shape === "star5" || shape === "star") {
        drawStarPath(ctx, drawX, drawY, w, h);
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (shape === "pie" && Number.isFinite(el.pieStart) && Number.isFinite(el.pieEnd)) {
        drawPiePath(ctx, drawX, drawY, w, h, el.pieStart, el.pieEnd);
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (ELLIPSE_SHAPES.indexOf(shape) >= 0) {
        ctx.beginPath();
        ctx.ellipse(drawX + w/2, drawY + h/2, w/2, h/2, 0, 0, Math.PI * 2);
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (shape === "wedgeroundrectcallout" || shape === "callout") {
        var pointLeft = (el.x || 0) > 0.5;
        drawCalloutPath(ctx, drawX, drawY, w, h, pointLeft);
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else if (ROUND_RECT_SHAPES.indexOf(shape) >= 0 || shape === "roundrect") {
        var radius = Math.min(w, h) * 0.15;
        drawRoundRectPath(ctx, drawX, drawY, w, h, radius);
        if (fillColor !== "transparent") ctx.fill();
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.stroke();
    } else {
        // Default: rectangle
        if (fillColor !== "transparent") ctx.fillRect(drawX, drawY, w, h);
        if (strokeColor !== "transparent" && strokeWidth > 0) ctx.strokeRect(drawX, drawY, w, h);
    }
    
    // Restore context if we rotated
    if (hasRotation) {
        ctx.restore();
    }
}

function renderImageToCanvasWithImg(el, ctx, canvasW, canvasH, img) {
    var x = (el.x || 0) * canvasW;
    var y = (el.y || 0) * canvasH;
    var w = (el.w || 0.5) * canvasW;
    var h = (el.h || 0.5) * canvasH;
    
    try {
        ctx.drawImage(img, x, y, w, h);
    } catch (e) {
        console.warn("Failed to draw image:", e);
    }
}

// ============================================================================
// GUI setup (controller panel at bottom)
// ============================================================================

function buildRobotGui(scene) {
    var ui = BABYLON.GUI.AdvancedDynamicTexture.CreateFullscreenUI("UI");

    var panel = new BABYLON.GUI.Rectangle();
    panel.width = "520px";
    panel.height = "130px";
    panel.cornerRadius = 12;
    panel.color = "white";
    panel.thickness = 1;
    panel.background = "rgba(0,0,0,0.55)";
    panel.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_CENTER;
    panel.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_BOTTOM;
    panel.top = "-20px";
    ui.addControl(panel);

    var stack = new BABYLON.GUI.StackPanel();
    stack.width = 0.95;
    stack.isVertical = true;
    panel.addControl(stack);

    var title = new BABYLON.GUI.TextBlock();
    title.height = "28px";
    title.color = "white";
    title.fontSize = 18;
    title.text = "Robot Presentation";
    title.fontFamily = "Segoe UI, sans-serif";
    stack.addControl(title);

    var pageText = new BABYLON.GUI.TextBlock();
    pageText.height = "24px";
    pageText.color = "white";
    pageText.fontSize = 16;
    pageText.text = "Slide 1 / 1";
    pageText.fontFamily = "Segoe UI, sans-serif";
    stack.addControl(pageText);

    var slider = new BABYLON.GUI.Slider();
    slider.minimum = 1;
    slider.maximum = 1;
    slider.value = 1;
    slider.step = 1;
    slider.height = "20px";
    slider.width = "440px";
    slider.color = PP;
    slider.background = "#555555";
    slider.borderColor = "white";
    stack.addControl(slider);

    var buttonRow = new BABYLON.GUI.StackPanel();
    buttonRow.isVertical = false;
    buttonRow.height = "36px";
    stack.addControl(buttonRow);

    function makeButton(text) {
        var button = BABYLON.GUI.Button.CreateSimpleButton(text, text);
        button.width = "140px";
        button.height = "32px";
        button.color = "white";
        button.cornerRadius = 8;
        button.thickness = 1;
        button.background = PP;
        button.paddingLeft = "8px";
        button.paddingRight = "8px";
        button.fontFamily = "Segoe UI, sans-serif";
        return button;
    }

    var prevButton = makeButton("◀ Prev");
    var nextButton = makeButton("Next ▶");
    buttonRow.addControl(prevButton);
    buttonRow.addControl(nextButton);

    var fpsText = new BABYLON.GUI.TextBlock();
    fpsText.width = "120px";
    fpsText.height = "32px";
    fpsText.color = "white";
    fpsText.fontSize = 14;
    fpsText.fontFamily = "Segoe UI, sans-serif";
    fpsText.textHorizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_CENTER;
    fpsText.text = "0 fps";
    buttonRow.addControl(fpsText);

    // Drop hint
    var dropHint = new BABYLON.GUI.TextBlock();
    dropHint.height = "20px";
    dropHint.color = "#aaaaaa";
    dropHint.fontSize = 12;
    dropHint.text = "Drop .pptx file to load";
    dropHint.fontFamily = "Segoe UI, sans-serif";
    stack.addControl(dropHint);

    return {
        ui: ui,
        titleText: title,
        pageText: pageText,
        slider: slider,
        prevButton: prevButton,
        nextButton: nextButton,
        fpsText: fpsText,
        dropHint: dropHint
    };
}

// ============================================================================
// Main application
// ============================================================================

var engine = null;
var sceneToRender = null;

var createDefaultEngine = function() {
    return new BABYLON.Engine(canvas, true, { 
        preserveDrawingBuffer: true, 
        stencil: true, 
        disableWebGL2Support: false 
    });
};

var startRenderLoop = function(engine) {
    engine.runRenderLoop(function() {
        if (sceneToRender && sceneToRender.activeCamera) {
            sceneToRender.render();
        }
    });
};

async function runApp() {
    console.log("[INIT] Robot PowerPoint Viewer starting...");

    // Validate JSZip
    ensureJsZipLoaded();

    // Create engine
    engine = createDefaultEngine();
    if (!engine) {
        throw new AppError('ENGINE_NULL', 'Failed to create engine', 'engine is null');
    }
    window.engine = engine;
    startRenderLoop(engine);

    // Create scene
    var scene = new BABYLON.Scene(engine);
    scene.clearColor = new BABYLON.Color4(0.02, 0.02, 0.05, 1.0);
    sceneToRender = scene;
    window.scene = scene;

    // Setup robot scene
    var sceneObjs = await setupRobotScene(scene, canvas);
    var slideMesh = sceneObjs.slideMesh;
    var importedMeshes = sceneObjs.importedMeshes;
    
    console.log("[ROBOT-DEBUG] Scene setup complete");
    console.log("[ROBOT-DEBUG] slideMesh:", slideMesh ? slideMesh.name : "NOT FOUND");
    console.log("[ROBOT-DEBUG] slideMesh.material:", slideMesh && slideMesh.material ? slideMesh.material.name : "NOT FOUND");
    console.log("[ROBOT-DEBUG] importedMeshes count:", importedMeshes.length);
    if (slideMesh) {
        console.log("[ROBOT-DEBUG] slideMesh position:", slideMesh.position.toString());
        console.log("[ROBOT-DEBUG] slideMesh isVisible:", slideMesh.isVisible);
    }

    // Build GUI
    var gui = buildRobotGui(scene);

    // Application state
    var app = {
        scene: scene,
        gui: gui,
        slides: getDefaultSlides(),
        currentSlide: 0,
        slideMesh: slideMesh,
        importedMeshes: importedMeshes,
        isAnimating: false,
        textureCache: {}
    };

    // Render current slide
    async function renderCurrentSlide() {
        console.log("[ROBOT-DEBUG] renderCurrentSlide called, currentSlide:", app.currentSlide);
        if (!app.slides || app.slides.length === 0) {
            console.log("[ROBOT-DEBUG] No slides to render");
            return;
        }
        var slide = app.slides[app.currentSlide];
        console.log("[ROBOT-DEBUG] Rendering slide:", app.currentSlide, "slideMesh:", slideMesh ? slideMesh.name : "null");
        await createSlideTexture(slide, scene, slideMesh);
        gui.pageText.text = "Slide " + (app.currentSlide + 1) + " / " + app.slides.length;
        console.log("[ROBOT-DEBUG] renderCurrentSlide complete");
    }

    // Update slider
    function updateSlider() {
        gui.slider.maximum = Math.max(1, app.slides.length);
        gui.slider.value = app.currentSlide + 1;
    }

    // Change slide with animation
    function changeSlide(newIndex, useAnimation) {
        newIndex = Math.max(0, Math.min(app.slides.length - 1, newIndex));
        if (newIndex === app.currentSlide) return;
        if (app.isAnimating) return;

        if (!useAnimation || importedMeshes.length === 0) {
            app.currentSlide = newIndex;
            renderCurrentSlide();
            updateSlider();
            return;
        }

        app.isAnimating = true;
        animatePageFlip(scene, importedMeshes, function() {
            app.currentSlide = newIndex;
            renderCurrentSlide();
            updateSlider();
            setTimeout(function() {
                app.isAnimating = false;
            }, 600);
        });
    }

    // Navigation buttons
    gui.prevButton.onPointerClickObservable.add(function() {
        changeSlide(app.currentSlide - 1, true);
    });

    gui.nextButton.onPointerClickObservable.add(function() {
        changeSlide(app.currentSlide + 1, true);
    });

    // Slider
    var syncingSlider = false;
    gui.slider.onValueChangedObservable.add(function(value) {
        if (syncingSlider) return;
        var newIndex = Math.round(value) - 1;
        changeSlide(newIndex, false);
    });

    // Keyboard navigation
    scene.onKeyboardObservable.add(function(kb) {
        if (kb.type !== BABYLON.KeyboardEventTypes.KEYDOWN) return;
        var k = kb.event.key;
        var changed = false;
        
        if (k === "ArrowRight" || k === "ArrowDown" || k === "PageDown" || k === " ") {
            if (app.currentSlide < app.slides.length - 1) {
                changeSlide(app.currentSlide + 1, true);
                changed = true;
            }
        } else if (k === "ArrowLeft" || k === "ArrowUp" || k === "PageUp") {
            if (app.currentSlide > 0) {
                changeSlide(app.currentSlide - 1, true);
                changed = true;
            }
        } else if (k === "Home") {
            changeSlide(0, false);
            changed = true;
        } else if (k === "End") {
            changeSlide(app.slides.length - 1, false);
            changed = true;
        }
        
        if (changed) {
            kb.event.preventDefault();
        }
    });

    // Drag & drop PPTX
    var dropOverlay = document.createElement("div");
    dropOverlay.style.cssText = "position:fixed;top:0;left:0;width:100%;height:100%;" +
        "background:rgba(208,68,35,0.3);z-index:9999;pointer-events:none;" +
        "font-size:32px;color:white;display:none;align-items:center;justify-content:center;" +
        "font-family:Segoe UI,sans-serif;";
    dropOverlay.textContent = "Drop .pptx here";
    document.body.appendChild(dropOverlay);

    var dragCounter = 0;
    
    document.addEventListener("dragenter", function(e) {
        e.preventDefault();
        dragCounter++;
        dropOverlay.style.display = "flex";
    });
    
    document.addEventListener("dragleave", function(e) {
        e.preventDefault();
        dragCounter--;
        if (dragCounter <= 0) {
            dragCounter = 0;
            dropOverlay.style.display = "none";
        }
    });
    
    document.addEventListener("dragover", function(e) {
        e.preventDefault();
    });
    
    document.addEventListener("drop", async function(e) {
        e.preventDefault();
        dragCounter = 0;
        dropOverlay.style.display = "none";
        
        var files = e.dataTransfer.files;
        if (files.length === 0) return;
        
        var file = files[0];
        if (!file.name.toLowerCase().endsWith(".pptx")) {
            alert("Please drop a .pptx file");
            return;
        }

        gui.titleText.text = "Loading: " + file.name + "...";
        gui.dropHint.text = "";
        
        try {
            var ab = await file.arrayBuffer();
            
            var ns = await parsePptx(ab,
                function onStructureReady(partialSlides) {
                    if (partialSlides.length === 0) return;
                    app.slides = partialSlides;
                    app.currentSlide = 0;
                    renderCurrentSlide();
                    updateSlider();
                    gui.titleText.text = file.name.replace(".pptx", "") + " - Loading images...";
                },
                function onSlideImagesReady(idx) {
                    if (idx === app.currentSlide) {
                        renderCurrentSlide();
                    }
                },
                function onAllImagesReady() {
                    if (!app.slides || app.slides.length === 0) return;
                    renderCurrentSlide();
                    updateSlider();
                    gui.titleText.text = file.name.replace(".pptx", "");
                }
            );
            
            if (ns.length > 0) {
                app.slides = ns;
                app.currentSlide = 0;
                renderCurrentSlide();
                updateSlider();
            }
        } catch (err) {
            console.error("PPTX parse error:", err);
            gui.titleText.text = "Error: " + err.message;
        }
    });

    // FPS counter
    scene.onBeforeRenderObservable.add(function() {
        gui.fpsText.text = engine.getFps().toFixed() + " fps";
    });

    // Initial render
    renderCurrentSlide();
    updateSlider();

    console.log("[INIT] Robot PowerPoint Viewer ready");
    return scene;
}

// Resize handler
window.addEventListener("resize", function() {
    if (engine) engine.resize();
});

// Start application
runApp().catch(function(err) {
    console.error("[INIT] Failed to start Robot PowerPoint Viewer:", err);
});

export default runApp;
