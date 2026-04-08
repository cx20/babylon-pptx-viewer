// ============================================================================
// shape-parsers.js - Shape tree parsing (sp, pic, cxnSp, grpSp, graphicFrame)
// ============================================================================

import { A_NS, P_NS, R_NS, CANVAS_H, normalizeElement } from "./constants.js";
import { resolveColor, themeColors } from "./color-utils.js";
import { parseParagraphs, parseOutline, getPresetGeometry, getShapeFill } from "./text-parser.js";

function colorLuma(hex) {
    if (!hex || hex.charAt(0) !== "#" || hex.length !== 7) return 0;
    var r = parseInt(hex.substr(1, 2), 16);
    var g = parseInt(hex.substr(3, 2), 16);
    var b = parseInt(hex.substr(5, 2), 16);
    if (!Number.isFinite(r) || !Number.isFinite(g) || !Number.isFinite(b)) return 0;
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
}

function txBodyHasExplicitColor(txBody) {
    if (!txBody) return false;
    var rPrs = txBody.getElementsByTagNameNS(A_NS, "rPr");
    for (var i = 0; i < rPrs.length; i++) {
        if (rPrs[i].getElementsByTagNameNS(A_NS, "solidFill").length > 0) return true;
    }
    var endR = txBody.getElementsByTagNameNS(A_NS, "endParaRPr");
    for (var j = 0; j < endR.length; j++) {
        if (endR[j].getElementsByTagNameNS(A_NS, "solidFill").length > 0) return true;
    }
    var defR = txBody.getElementsByTagNameNS(A_NS, "defRPr");
    for (var k = 0; k < defR.length; k++) {
        if (defR[k].getElementsByTagNameNS(A_NS, "solidFill").length > 0) return true;
    }
    return false;
}

function parseShapeTree(spTreeNode, slideW, slideH, images, relsAll, opts) {
    opts = opts || {};
    var elements = [];
    var skipPlaceholders = opts.skipPlaceholders || false;
    var hasBgImage = opts.hasBgImage || false;
    var layoutStyles = opts.layoutStyles || {};
    var chartDataMap = opts.chartDataMap || {};
    var diagramDataMap = opts.diagramDataMap || {};
    var sourceLayer = opts.sourceLayer || "slide";
    // Dark background slides generally require light fallback text for readability.
    var defaultTextColor = hasBgImage ? (themeColors.lt1 || "#FFFFFF") : (themeColors.tx1 || "#333");
    // Group transform: convert child coords to slide fraction coords
    var gOffX = opts.gOffX || 0, gOffY = opts.gOffY || 0;
    var gScaleX = opts.gScaleX || 1, gScaleY = opts.gScaleY || 1;

    function finiteOr(value, fallback) {
        return Number.isFinite(value) ? value : fallback;
    }

    function toFracX(emu) { return finiteOr((gOffX + emu * gScaleX) / slideW, 0); }
    function toFracY(emu) { return finiteOr((gOffY + emu * gScaleY) / slideH, 0); }
    function toFracW(emu) { return finiteOr(emu * gScaleX / slideW, 0); }
    function toFracH(emu) { return finiteOr(emu * gScaleY / slideH, 0); }

    // Iterate direct children
    var childCount = {sp:0, pic:0, grpSp:0, cxnSp:0, graphicFrame:0, other:0};
    for (var ci = 0; ci < spTreeNode.childNodes.length; ci++) {
        var child = spTreeNode.childNodes[ci];
        if (child.nodeType !== 1) continue;
        var ln = child.localName;
        if (ln === "sp") childCount.sp++;
        else if (ln === "pic") childCount.pic++;
        else if (ln === "grpSp") childCount.grpSp++;
        else if (ln === "cxnSp") childCount.cxnSp++;
        else if (ln === "graphicFrame") childCount.graphicFrame++;
        else childCount.other++;
    }
    console.log("[TREE] children: sp=" + childCount.sp + " pic=" + childCount.pic + " grpSp=" + childCount.grpSp + " cxnSp=" + childCount.cxnSp + " graphicFrame=" + childCount.graphicFrame + " other=" + childCount.other + " skipPH=" + skipPlaceholders);

    for (var ci = 0; ci < spTreeNode.childNodes.length; ci++) {
        var child = spTreeNode.childNodes[ci];
        if (child.nodeType !== 1) continue; // element nodes only
        var localName = child.localName;

        // --- sp (shape) ---
        if (localName === "sp") {
            parseSp(child, elements, slideW, slideH, skipPlaceholders, defaultTextColor, toFracX, toFracY, toFracW, toFracH, layoutStyles, hasBgImage, sourceLayer);
        }
        // --- pic (picture) ---
        else if (localName === "pic") {
            parsePic(child, elements, slideW, slideH, images, relsAll, hasBgImage, toFracX, toFracY, toFracW, toFracH);
        }
        // --- cxnSp (connector) ---
        else if (localName === "cxnSp") {
            parseCxnSp(child, elements, slideW, slideH, toFracX, toFracY);
        }
        // --- grpSp (group shape) - RECURSIVE ---
        else if (localName === "grpSp") {
            parseGrpSp(child, elements, slideW, slideH, images, relsAll, opts);
        }
        // --- graphicFrame (chart / table / diagram) ---
        else if (localName === "graphicFrame") {
            parseGraphicFrame(child, elements, slideW, slideH, images, relsAll, chartDataMap, diagramDataMap, defaultTextColor, toFracX, toFracY, toFracW, toFracH);
        }
    }
    return elements;
}

// --- Parse sp (shape with optional text) ---
function parseSp(sp, elements, slideW, slideH, skipPH, defTextColor, fx, fy, fw, fh, layoutStyles, hasBgImage, sourceLayer) {
    layoutStyles = layoutStyles || {};
    hasBgImage = hasBgImage || false;
    sourceLayer = sourceLayer || "slide";
    // Placeholder detection
    var phType = "", phIdx = -1;
    var nvSpPr = sp.getElementsByTagNameNS(P_NS, "nvSpPr")[0];
    if (nvSpPr) {
        var nvPr = nvSpPr.getElementsByTagNameNS(P_NS, "nvPr")[0];
        if (nvPr) {
            var ph = nvPr.getElementsByTagNameNS(P_NS, "ph")[0];
            if (ph) { phType = ph.getAttribute("type") || "body"; phIdx = parseInt(ph.getAttribute("idx")) || 0; }
        }
    }
    if (skipPH && phType) {
        console.log("[SP] SKIP placeholder type='" + phType + "'");
        return;
    }

    // Transform
    var xfrm = sp.getElementsByTagNameNS(A_NS, "xfrm")[0];
    var ox = 0, oy = 0, cx = 0, cy = 0, rot = 0;
    if (xfrm) {
        var off = xfrm.getElementsByTagNameNS(A_NS, "off")[0];
        var ext = xfrm.getElementsByTagNameNS(A_NS, "ext")[0];
        if (off) { ox = parseInt(off.getAttribute("x")) || 0; oy = parseInt(off.getAttribute("y")) || 0; }
        if (ext) { cx = parseInt(ext.getAttribute("cx")) || 0; cy = parseInt(ext.getAttribute("cy")) || 0; }
        rot = parseInt(xfrm.getAttribute("rot")) || 0;
    }
    var fracX = fx(ox), fracY = fy(oy), fracW = fw(cx), fracH = fh(cy);
    var rotDeg = rot / 60000;

    // Shape visual properties
    // p:spPr is in presentation namespace in normal PPTX shape nodes.
    var spPr = sp.getElementsByTagNameNS(P_NS, "spPr")[0];
    if (!spPr) spPr = sp.getElementsByTagNameNS(A_NS, "spPr")[0];
    var geom = getPresetGeometry(spPr);
    var outline = parseOutline(spPr);
    var fill = getShapeFill(spPr);
    var hasGradFill = !!(spPr && spPr.getElementsByTagNameNS(A_NS, "gradFill")[0]);

    // Style-based fill/color
    var styleFontColor = null;
    var style = sp.getElementsByTagNameNS(P_NS, "style")[0];
    if (style) {
        if (!fill && !outline) {
            var fillRef = style.getElementsByTagNameNS(A_NS, "fillRef")[0];
            if (fillRef) { var fc = resolveColor(fillRef); if (fc) fill = fc; }
        }
        var fontRef = style.getElementsByTagNameNS(A_NS, "fontRef")[0];
        if (fontRef) styleFontColor = resolveColor(fontRef);
    }

    // Master backgrounds sometimes encode glow spots as gradient ellipses.
    // Rendering them as flat fills causes large white circle artifacts.
    var isEllipseGeom = geom === "ellipse" || geom === "oval" || geom === "circle" || geom === "donut" || geom === "pie" || geom === "arc" || geom === "chord";
    if (sourceLayer === "master" && hasGradFill && isEllipseGeom && (!outline || outline.width <= 0)) {
        fill = null;
    }

    // Emit shape rectangle/ellipse
    if ((fill || outline) && cx > 0 && cy > 0) {
        elements.push(normalizeElement({
            type: "shape", shape: geom, x: fracX, y: fracY, w: fracW, h: fracH,
            fillColor: fill || "transparent",
            strokeColor: outline ? outline.color : "transparent",
            thickness: outline ? outline.width : 0,
            rotation: rotDeg
        }));
    }

    // Placeholder font defaults
    var phFS = 14, phFC = defTextColor;
    var phLayout = layoutStyles[phType];
    if (phType === "title" || phType === "ctrTitle") phFS = 32;
    else if (phType === "subTitle") {
        phFS = 20;
        // PowerPoint convention: subtitle uses accent1 color on dark background slides
        if (hasBgImage && themeColors.accent1) phFC = themeColors.accent1;
    }
    else if (phType === "body" || phType === "obj") phFS = 18;
    else if (!phType && cy > 0) {
        // Non-placeholder text boxes in PPT usually default to ~18pt.
        // Height-based estimation made bullet text too large in fixtures.
        phFS = 18;
    }
    // Apply layout fontRef color as default, then slide's own style overrides
    if (phLayout && phLayout.fontRefColor) phFC = phLayout.fontRefColor;
    if (styleFontColor) phFC = styleFontColor;

    // Text body
    var txBody = sp.getElementsByTagNameNS(P_NS, "txBody")[0];
    if (!txBody) txBody = sp.getElementsByTagNameNS(A_NS, "txBody")[0];
    if (!txBody) {
        console.log("[SP] geom=" + geom + " ph='" + phType + "' fill=" + fill + " pos=(" + fracX.toFixed(3) + "," + fracY.toFixed(3) + ") NO txBody");
        return;
    }

    // For bg-image slides, title placeholders are usually light text unless explicitly styled.
    if (hasBgImage && (phType === "title" || phType === "ctrTitle") && !txBodyHasExplicitColor(txBody)) {
        if (colorLuma(phFC) < 0.55) phFC = themeColors.lt1 || "#FFFFFF";
    }

    console.log("[SP] geom=" + geom + " ph='" + phType + "' fill=" + fill + " phFS=" + phFS + " phFC=" + phFC + " pos=(" + fracX.toFixed(3) + "," + fracY.toFixed(3) + ") size=(" + fracW.toFixed(3) + "," + fracH.toFixed(3) + ")");

    // Body properties (anchor, insets)
    var bodyPr = txBody.getElementsByTagNameNS(A_NS, "bodyPr")[0];
    // Default anchor based on placeholder type
    // ctrTitle typically has bottom-aligned text in PowerPoint
    var defaultAnchor = "t";
    if (phType === "ctrTitle") defaultAnchor = "b";
    else if (phType === "title") defaultAnchor = "b";
    // Override with layout placeholder anchor
    if (phLayout && phLayout.anchor) defaultAnchor = phLayout.anchor;
    
    var anchor = defaultAnchor;
    var iL = 91440 / slideW, iT = 45720 / slideH, iR = 91440 / slideW, iB = 45720 / slideH;
    if (bodyPr) {
        var explicitAnchor = bodyPr.getAttribute("anchor");
        if (explicitAnchor) anchor = explicitAnchor;
        var lI = bodyPr.getAttribute("lIns"), tI = bodyPr.getAttribute("tIns");
        var rI = bodyPr.getAttribute("rIns"), bI = bodyPr.getAttribute("bIns");
        if (lI !== null) iL = parseInt(lI) / slideW;
        if (tI !== null) iT = parseInt(tI) / slideH;
        if (rI !== null) iR = parseInt(rI) / slideW;
        if (bI !== null) iB = parseInt(bI) / slideH;
    }
    console.log("[SP]   anchor=" + anchor + " (explicit=" + (bodyPr && bodyPr.getAttribute("anchor") || "none") + " default=" + defaultAnchor + ")");

    var layoutCap = (phLayout && phLayout.cap) ? phLayout.cap : "";
    var paras = parseParagraphs(txBody, phFS, phFC, layoutCap);
    paras.forEach(function(p, pi) {
        if (!p.isEmpty) console.log("[SP]   para[" + pi + "] '" + p.text.substring(0,30) + "' fs=" + p.fontSize + " color=" + p.color + " align=" + p.align);
    });

    // Calculate vertical positioning based on anchor
    var totalH = 0, paraH = [];
    paras.forEach(function (p) {
        var h = p.isEmpty ? p.fontSize * 0.6 : p.fontSize * p.lineSpacing;
        h += (p.spaceBefore || 0) + (p.spaceAfter || 0);
        paraH.push(h); totalH += h;
    });
    var areaTop = fracY + iT, areaH = fracH - iT - iB;
    var thFrac = totalH / CANVAS_H;
    var startY = areaTop;
    if (anchor === "ctr" || anchor === "mid") startY = areaTop + (areaH - thFrac) / 2;
    else if (anchor === "b") startY = areaTop + areaH - thFrac;

    var curY = startY;
    // Approximate PPT default list indent (0.375 inch per level).
    var indentPerLevel = 342900 / slideW;
    paras.forEach(function (p, pi) {
        if (!p.isEmpty) {
            var level = p.level || 0;
            var indentX = level * indentPerLevel;
            var textX = fracX + iL + indentX;
            var textW = Math.max(0.01, fracW - iL - iR - indentX);
            elements.push(normalizeElement({
                type: "text", text: p.text,
                x: textX, y: curY, w: textW,
                fontSize: p.fontSize, color: p.color,
                fontWeight: p.fontWeight, fontStyle: p.italic ? "italic" : "normal",
                align: p.align, rotation: rotDeg
            }));
        }
        curY += paraH[pi] / CANVAS_H;
    });
}

// --- Parse pic (picture) ---
function parsePic(pic, elements, slideW, slideH, images, relsAll, hasBgImage, fx, fy, fw, fh) {
    var xfrm = pic.getElementsByTagNameNS(A_NS, "xfrm")[0]; if (!xfrm) return;
    var off = xfrm.getElementsByTagNameNS(A_NS, "off")[0];
    var ext = xfrm.getElementsByTagNameNS(A_NS, "ext")[0];
    if (!off || !ext) return;
    var ox = parseInt(off.getAttribute("x")) || 0, oy = parseInt(off.getAttribute("y")) || 0;
    var cx = parseInt(ext.getAttribute("cx")) || 0, cy = parseInt(ext.getAttribute("cy")) || 0;

    // Skip full-bleed background images (already handled)
    if (hasBgImage && cx > slideW * 0.7 && cy > slideH * 0.7) return;

    var blip = pic.getElementsByTagNameNS(A_NS, "blip")[0];
    if (!blip) return;
    var rId = blip.getAttribute("r:embed") || blip.getAttributeNS(R_NS, "embed");
    if (!rId || !images[rId]) return;

    // srcRect (crop) - approximate by adjusting position/size
    var fracX = fx(ox), fracY = fy(oy), fracW = fw(cx), fracH = fh(cy);
    var blipFill = pic.getElementsByTagNameNS(A_NS, "blipFill")[0];
    var cropL = 0, cropT = 0, cropR = 0, cropB = 0;
    if (blipFill) {
        var srcRect = blipFill.getElementsByTagNameNS(A_NS, "srcRect")[0];
        if (srcRect) {
            cropL = (parseInt(srcRect.getAttribute("l")) || 0) / 100000;
            cropT = (parseInt(srcRect.getAttribute("t")) || 0) / 100000;
            cropR = (parseInt(srcRect.getAttribute("r")) || 0) / 100000;
            cropB = (parseInt(srcRect.getAttribute("b")) || 0) / 100000;
        }
    }

    elements.push(normalizeElement({
        type: "image", dataUrl: images[rId],
        x: fracX, y: fracY, w: fracW, h: fracH,
        crop: { l: cropL, t: cropT, r: cropR, b: cropB }
    }));
    console.log("[PIC] image at (" + fracX.toFixed(3) + "," + fracY.toFixed(3) + ") size=(" + fracW.toFixed(3) + "," + fracH.toFixed(3) + ") crop=L" + (cropL*100).toFixed(0) + "%,T" + (cropT*100).toFixed(0) + "%,R" + (cropR*100).toFixed(0) + "%,B" + (cropB*100).toFixed(0) + "%");
}

// --- Parse cxnSp (connector line) ---
function parseCxnSp(cxn, elements, slideW, slideH, fx, fy) {
    var xfrm = cxn.getElementsByTagNameNS(A_NS, "xfrm")[0]; if (!xfrm) return;
    var off = xfrm.getElementsByTagNameNS(A_NS, "off")[0];
    var ext = xfrm.getElementsByTagNameNS(A_NS, "ext")[0];
    if (!off || !ext) return;
    var x1 = parseInt(off.getAttribute("x")) || 0, y1 = parseInt(off.getAttribute("y")) || 0;
    var w = parseInt(ext.getAttribute("cx")) || 0, h = parseInt(ext.getAttribute("cy")) || 0;
    var flipH = xfrm.getAttribute("flipH") === "1", flipV = xfrm.getAttribute("flipV") === "1";
    // p:spPr is typically used for connector style/outline.
    var spPr = cxn.getElementsByTagNameNS(P_NS, "spPr")[0];
    if (!spPr) spPr = cxn.getElementsByTagNameNS(A_NS, "spPr")[0];
    var ol = parseOutline(spPr);
    console.log("[CXN] connector line color=" + (ol?ol.color:"#000") + " flipH=" + flipH + " flipV=" + flipV);
    elements.push(normalizeElement({
        type: "shape", shape: "line",
        x1: fx(flipH ? x1 + w : x1), y1: fy(flipV ? y1 + h : y1),
        x2: fx(flipH ? x1 : x1 + w), y2: fy(flipV ? y1 : y1 + h),
        color: ol ? ol.color : "#000", thickness: ol ? ol.width : 1
    }));
}

// --- Parse grpSp (group shape) - RECURSIVE ---
function parseGrpSp(grpSp, elements, slideW, slideH, images, relsAll, parentOpts) {
    console.log("[GRPSP] Parsing group shape");
    var grpSpPr = grpSp.getElementsByTagNameNS(A_NS, "grpSpPr")[0];
    if (!grpSpPr) grpSpPr = grpSp.getElementsByTagNameNS(P_NS, "grpSpPr")[0];

    // Group has two coordinate spaces:
    // off/ext = position and size on parent
    // chOff/chExt = child coordinate space
    var offX = 0, offY = 0, extW = 1, extH = 1, chOffX = 0, chOffY = 0, chExtW = 1, chExtH = 1;
    if (grpSpPr) {
        var xfrm = grpSpPr.getElementsByTagNameNS(A_NS, "xfrm")[0];
        if (xfrm) {
            var off = xfrm.getElementsByTagNameNS(A_NS, "off")[0];
            var ext = xfrm.getElementsByTagNameNS(A_NS, "ext")[0];
            var chOff = xfrm.getElementsByTagNameNS(A_NS, "chOff")[0];
            var chExt = xfrm.getElementsByTagNameNS(A_NS, "chExt")[0];
            if (off) { offX = parseInt(off.getAttribute("x")) || 0; offY = parseInt(off.getAttribute("y")) || 0; }
            if (ext) { extW = parseInt(ext.getAttribute("cx")) || 1; extH = parseInt(ext.getAttribute("cy")) || 1; }
            if (chOff) { chOffX = parseInt(chOff.getAttribute("x")) || 0; chOffY = parseInt(chOff.getAttribute("y")) || 0; }
            if (chExt) { chExtW = parseInt(chExt.getAttribute("cx")) || 1; chExtH = parseInt(chExt.getAttribute("cy")) || 1; }
        }
    }

    // Calculate transform: child EMU → slide EMU
    var pGOffX = (parentOpts && parentOpts.gOffX) || 0;
    var pGOffY = (parentOpts && parentOpts.gOffY) || 0;
    var pGScaleX = (parentOpts && parentOpts.gScaleX) || 1;
    var pGScaleY = (parentOpts && parentOpts.gScaleY) || 1;

    var ratioX = chExtW !== 0 ? (extW / chExtW) : 1;
    var ratioY = chExtH !== 0 ? (extH / chExtH) : 1;
    if (chExtW === 0 || chExtH === 0) {
        console.warn("[GRPSP] invalid chExt detected, fallback ratio=1", { chExtW: chExtW, chExtH: chExtH });
    }
    if (!Number.isFinite(ratioX)) ratioX = 1;
    if (!Number.isFinite(ratioY)) ratioY = 1;

    var newGOffX = pGOffX + (offX - chOffX * ratioX) * pGScaleX;
    var newGOffY = pGOffY + (offY - chOffY * ratioY) * pGScaleY;
    var newGScaleX = pGScaleX * ratioX;
    var newGScaleY = pGScaleY * ratioY;

    var childOpts = {
        skipPlaceholders: parentOpts ? parentOpts.skipPlaceholders : false,
        hasBgImage: parentOpts ? parentOpts.hasBgImage : false,
        layoutStyles: parentOpts ? (parentOpts.layoutStyles || {}) : {},
        chartDataMap: parentOpts ? (parentOpts.chartDataMap || {}) : {},
        diagramDataMap: parentOpts ? (parentOpts.diagramDataMap || {}) : {},
        sourceLayer: parentOpts ? (parentOpts.sourceLayer || "slide") : "slide",
        gOffX: newGOffX, gOffY: newGOffY,
        gScaleX: newGScaleX, gScaleY: newGScaleY
    };
    console.log("[GRPSP]   off=(" + offX + "," + offY + ") ext=(" + extW + "," + extH + ") chOff=(" + chOffX + "," + chOffY + ") chExt=(" + chExtW + "," + chExtH + ") scale=(" + newGScaleX.toFixed(3) + "," + newGScaleY.toFixed(3) + ")");

    var childElements = parseShapeTree(grpSp, slideW, slideH, images, relsAll, childOpts);
    childElements.forEach(function (el) { elements.push(el); });
}

// --- Parse graphicFrame (chart / table / diagram) ---
function parseGraphicFrame(gf, elements, slideW, slideH, images, relsAll, chartDataMap, diagramDataMap, defTextColor, fx, fy, fw, fh) {
    // graphicFrame uses p:xfrm, not a:xfrm
    var xfrm = gf.getElementsByTagNameNS(P_NS, "xfrm")[0];
    if (!xfrm) xfrm = gf.getElementsByTagNameNS(A_NS, "xfrm")[0];
    if (!xfrm) { console.log("[GF] No xfrm found in graphicFrame"); return; }
    // off/ext may be in a: or p: namespace depending on xfrm parent
    var off = xfrm.getElementsByTagNameNS(A_NS, "off")[0] || xfrm.getElementsByTagNameNS(P_NS, "off")[0];
    var ext = xfrm.getElementsByTagNameNS(A_NS, "ext")[0] || xfrm.getElementsByTagNameNS(P_NS, "ext")[0];
    // Also try without namespace (some serializers omit prefix for children)
    if (!off || !ext) {
        for (var ci = 0; ci < xfrm.childNodes.length; ci++) {
            var cn = xfrm.childNodes[ci];
            if (cn.localName === "off" && !off) off = cn;
            if (cn.localName === "ext" && !ext) ext = cn;
        }
    }
    if (!off || !ext) { console.log("[GF] No off/ext in graphicFrame xfrm"); return; }
    var ox = parseInt(off.getAttribute("x")) || 0, oy = parseInt(off.getAttribute("y")) || 0;
    var cx = parseInt(ext.getAttribute("cx")) || 0, cy = parseInt(ext.getAttribute("cy")) || 0;
    var fracX = fx(ox), fracY = fy(oy), fracW = fw(cx), fracH = fh(cy);

    // Check graphic data namespace to determine type
    // Structure: p:graphicFrame > a:graphic > a:graphicData
    var graphic = gf.getElementsByTagNameNS(A_NS, "graphic")[0];
    var graphicData = null;
    if (graphic) {
        graphicData = graphic.getElementsByTagNameNS(A_NS, "graphicData")[0];
    }
    if (!graphicData) {
        // Fallback: search without namespace
        graphicData = gf.getElementsByTagName("a:graphicData")[0];
    }
    if (!graphicData) {
        // Broader fallback
        graphicData = gf.getElementsByTagNameNS(A_NS, "graphicData")[0];
    }
    var uri = graphicData ? (graphicData.getAttribute("uri") || "") : "";
    console.log("[GF] graphicFrame at (" + fracX.toFixed(3) + "," + fracY.toFixed(3) + ") size=(" + fracW.toFixed(3) + "," + fracH.toFixed(3) + ") uri=" + uri + " hasGraphicData=" + !!graphicData);

    // Chart: render clustered bars from pre-parsed chart cache when available
    if (uri.indexOf("chart") !== -1 && graphicData) {
        var chartNode = null;
        for (var cni = 0; cni < graphicData.childNodes.length; cni++) {
            var child = graphicData.childNodes[cni];
            if (child.nodeType === 1 && child.localName === "chart") { chartNode = child; break; }
        }
        var chartRid = chartNode ? (chartNode.getAttribute("r:id") || chartNode.getAttributeNS(R_NS, "id") || "") : "";
        var chartData = chartRid ? chartDataMap[chartRid] : null;
        if (chartData && chartData.type === "barChart") {
            renderBarChart(chartData, elements, fracX, fracY, fracW, fracH, defTextColor);
            return;
        }
    }

    if (uri.indexOf("diagram") !== -1 && graphicData) {
        var relIdsNode = null;
        for (var ri = 0; ri < graphicData.childNodes.length; ri++) {
            var rn = graphicData.childNodes[ri];
            if (rn.nodeType === 1 && rn.localName === "relIds") { relIdsNode = rn; break; }
        }
        var dmRid = "";
        if (relIdsNode) {
            dmRid = relIdsNode.getAttribute("r:dm") || relIdsNode.getAttributeNS(R_NS, "dm") || relIdsNode.getAttribute("dm") || "";
        }
        var dgm = dmRid ? diagramDataMap[dmRid] : null;
        if (!dgm) {
            for (var key in diagramDataMap) {
                dgm = diagramDataMap[key];
                if (dgm) break;
            }
        }
        if (renderSmartArtDiagram(dgm, elements, fracX, fracY, fracW, fracH, defTextColor)) return;
    }

    // Table (a:tbl)
    if (uri.indexOf("table") !== -1 || uri.indexOf("dgm") !== -1) {
        parseTable(graphicData, elements, fracX, fracY, fracW, fracH, defTextColor);
        return;
    }

    // Chart or other - render as placeholder box
    var label = "Chart";
    if (uri.indexOf("chart") !== -1) label = "📊 Chart";
    else if (uri.indexOf("diagram") !== -1) label = "📐 Diagram";
    else if (uri.indexOf("ole") !== -1) label = "📎 OLE Object";
    else label = "📋 Object";

    elements.push(normalizeElement({
        type: "shape", shape: "rect", x: fracX, y: fracY, w: fracW, h: fracH,
        fillColor: "rgba(200,200,200,0.3)", strokeColor: "#999", thickness: 1, rotation: 0
    }));
    elements.push(normalizeElement({
        type: "text", text: label,
        x: fracX, y: fracY + fracH * 0.35, w: fracW,
        fontSize: 12, color: "#666", fontWeight: "normal", fontStyle: "normal", align: "center"
    }));
}

function parseGdValue(gdNode) {
    if (!gdNode) return 0;
    var fmla = gdNode.getAttribute("fmla") || "";
    var m = fmla.match(/val\s+(-?\d+)/i);
    return m ? (parseInt(m[1], 10) || 0) : 0;
}

function readTextFromTxBody(txBody) {
    if (!txBody) return "";
    var ts = txBody.getElementsByTagNameNS(A_NS, "t");
    var parts = [];
    for (var i = 0; i < ts.length; i++) {
        var tx = (ts[i].textContent || "").trim();
        if (tx) parts.push(tx);
    }
    return parts.join(" ").trim();
}

function readTextColorFromTxBody(txBody, fallbackColor) {
    if (!txBody) return fallbackColor;
    var rPr = txBody.getElementsByTagNameNS(A_NS, "rPr")[0] || txBody.getElementsByTagNameNS(A_NS, "endParaRPr")[0];
    if (!rPr) return fallbackColor;
    var sf = rPr.getElementsByTagNameNS(A_NS, "solidFill")[0];
    if (!sf) return fallbackColor;
    var c = resolveColor(sf);
    return c || fallbackColor;
}

function parseSimpleXfrm(node) {
    if (!node) return null;
    var xfrm = null;
    if (node.localName === "xfrm" || node.localName === "txXfrm") xfrm = node;
    if (!xfrm) xfrm = node.getElementsByTagNameNS(A_NS, "xfrm")[0];
    if (!xfrm) return null;

    var off = xfrm.getElementsByTagNameNS(A_NS, "off")[0] || xfrm.getElementsByTagNameNS("*", "off")[0];
    var ext = xfrm.getElementsByTagNameNS(A_NS, "ext")[0] || xfrm.getElementsByTagNameNS("*", "ext")[0];
    if (!off || !ext) return null;
    return {
        x: parseInt(off.getAttribute("x"), 10) || 0,
        y: parseInt(off.getAttribute("y"), 10) || 0,
        w: parseInt(ext.getAttribute("cx"), 10) || 0,
        h: parseInt(ext.getAttribute("cy"), 10) || 0
    };
}

function renderSmartArtDiagram(diagramEntry, elements, fracX, fracY, fracW, fracH, defTextColor) {
    if (!diagramEntry) return false;
    var loTypeId = getDiagramLayoutType(diagramEntry);
    var isPieLayout = /chart3|cycle|pie|wedge/i.test(loTypeId || "");

    if (isPieLayout) {
        if (renderSmartArtFromDrawing(diagramEntry, elements, fracX, fracY, fracW, fracH, defTextColor)) return true;
        return renderSmartArtFallbackFromData(diagramEntry, elements, fracX, fracY, fracW, fracH, defTextColor);
    }

    return renderSmartArtFromDrawingGeneric(diagramEntry, elements, fracX, fracY, fracW, fracH, defTextColor);
}

function getDiagramLayoutType(diagramEntry) {
    var dataDoc = diagramEntry ? diagramEntry.dataDoc : null;
    if (!dataDoc) return "";
    var ptNodes = dataDoc.getElementsByTagNameNS("*", "pt");
    for (var i = 0; i < ptNodes.length; i++) {
        var pt = ptNodes[i];
        if ((pt.getAttribute("type") || "") !== "doc") continue;
        var prSet = pt.getElementsByTagNameNS("*", "prSet")[0];
        if (!prSet) continue;
        var loTypeId = prSet.getAttribute("loTypeId") || "";
        if (loTypeId) return loTypeId;
    }
    return "";
}

function getTextAlignFromTxBody(txBody) {
    if (!txBody) return "left";
    var pPr = txBody.getElementsByTagNameNS(A_NS, "pPr")[0];
    if (!pPr) return "left";
    var algn = (pPr.getAttribute("algn") || "l").toLowerCase();
    if (algn === "ctr" || algn === "center") return "center";
    if (algn === "r" || algn === "right") return "right";
    return "left";
}

function getTextFontSizeFromTxBody(txBody, fallback) {
    if (!txBody) return fallback;
    var rPr = txBody.getElementsByTagNameNS(A_NS, "rPr")[0] || txBody.getElementsByTagNameNS(A_NS, "endParaRPr")[0];
    if (!rPr) return fallback;
    var sz = parseInt(rPr.getAttribute("sz"), 10);
    if (!Number.isFinite(sz) || sz <= 0) return fallback;
    return Math.max(10, Math.round(sz / 100));
}

function pickReadableTextColor(baseColor, bgFill, fallbackColor) {
    if (baseColor) return baseColor;
    var c = (bgFill || "").trim();
    if (c.charAt(0) !== "#" || (c.length !== 7 && c.length !== 4)) {
        return fallbackColor || "#000000";
    }
    if (c.length === 4) {
        c = "#" + c.charAt(1) + c.charAt(1) + c.charAt(2) + c.charAt(2) + c.charAt(3) + c.charAt(3);
    }
    var r = parseInt(c.substr(1, 2), 16);
    var g = parseInt(c.substr(3, 2), 16);
    var b = parseInt(c.substr(5, 2), 16);
    if (!Number.isFinite(r) || !Number.isFinite(g) || !Number.isFinite(b)) return fallbackColor || "#000000";
    var luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return luminance < 0.55 ? "#FFFFFF" : "#000000";
}

function getSmartArtStyleColorRefs(spNode) {
    if (!spNode) return { fillColor: null, fontColor: null };
    var styleNode = spNode.getElementsByTagNameNS("*", "style")[0];
    if (!styleNode) return { fillColor: null, fontColor: null };
    var fillRef = styleNode.getElementsByTagNameNS(A_NS, "fillRef")[0];
    var fontRef = styleNode.getElementsByTagNameNS(A_NS, "fontRef")[0];
    return {
        fillColor: fillRef ? resolveColor(fillRef) : null,
        fontColor: fontRef ? resolveColor(fontRef) : null
    };
}

function getBlipEmbedRid(spPr) {
    if (!spPr) return "";
    var blip = spPr.getElementsByTagNameNS(A_NS, "blip")[0];
    if (blip) {
        var rid = blip.getAttribute("r:embed") || blip.getAttributeNS(R_NS, "embed") || "";
        if (rid) return rid;
    }
    var all = spPr.getElementsByTagNameNS("*", "svgBlip");
    if (all && all.length > 0) {
        return all[0].getAttribute("r:embed") || all[0].getAttributeNS(R_NS, "embed") || "";
    }
    return "";
}

function renderSmartArtFromDrawingGeneric(diagramEntry, elements, fracX, fracY, fracW, fracH, defTextColor) {
    var drawingDoc = diagramEntry.drawingDoc;
    if (!drawingDoc) return false;
    var spNodes = drawingDoc.getElementsByTagName("dsp:sp");
    if (!spNodes || spNodes.length === 0) spNodes = drawingDoc.getElementsByTagNameNS("*", "sp");
    if (!spNodes || spNodes.length === 0) return false;

    var bounds = [];
    var minX = Number.POSITIVE_INFINITY, minY = Number.POSITIVE_INFINITY;
    var maxX = Number.NEGATIVE_INFINITY, maxY = Number.NEGATIVE_INFINITY;
    for (var i = 0; i < spNodes.length; i++) {
        var spPr0 = spNodes[i].getElementsByTagNameNS("*", "spPr")[0];
        var x0 = parseSimpleXfrm(spPr0);
        if (!x0 || x0.w <= 0 || x0.h <= 0) continue;
        bounds.push({ sp: spNodes[i], spPr: spPr0, xfrm: x0 });
        minX = Math.min(minX, x0.x);
        minY = Math.min(minY, x0.y);
        maxX = Math.max(maxX, x0.x + x0.w);
        maxY = Math.max(maxY, x0.y + x0.h);
    }
    if (bounds.length === 0 || maxX <= minX || maxY <= minY) return false;

    var bw = maxX - minX;
    var bh = maxY - minY;
    function mapX(v) { return fracX + ((v - minX) / bw) * fracW; }
    function mapY(v) { return fracY + ((v - minY) / bh) * fracH; }
    function mapW(v) { return (v / bw) * fracW; }
    function mapH(v) { return (v / bh) * fracH; }

    var textEls = [];
    for (var bi = 0; bi < bounds.length; bi++) {
        var item = bounds[bi];
        var sp = item.sp;
        var spPr = item.spPr;
        var x = item.xfrm;

        var sx = mapX(x.x), sy = mapY(x.y), sw = mapW(x.w), sh = mapH(x.h);
        var geom = getPresetGeometry(spPr) || "rect";
        var fill = getShapeFill(spPr);
        var outline = parseOutline(spPr);
        var hasBlip = !!spPr.getElementsByTagNameNS(A_NS, "blipFill")[0];
        var styleRefs = getSmartArtStyleColorRefs(sp);
        if (styleRefs.fillColor) fill = styleRefs.fillColor;

        if (hasBlip) {
            var blipRid = getBlipEmbedRid(spPr);
            var imgMap = diagramEntry.drawingImageMap || {};
            if (blipRid && imgMap[blipRid]) {
                elements.push(normalizeElement({
                    type: "image",
                    dataUrl: imgMap[blipRid],
                    x: sx, y: sy, w: sw, h: sh
                }));
            } else {
                elements.push(normalizeElement({
                    type: "shape", shape: "rect", x: sx, y: sy, w: sw, h: sh,
                    fillColor: "rgba(122,209,243,0.20)",
                    strokeColor: "#66C7EA",
                    thickness: 1
                }));
            }
        } else if (fill || outline) {
            elements.push(normalizeElement({
                type: "shape", shape: geom, x: sx, y: sy, w: sw, h: sh,
                fillColor: fill || "transparent",
                strokeColor: outline ? outline.color : "transparent",
                thickness: outline ? outline.width : 0
            }));
        }

        var txBody = sp.getElementsByTagNameNS("*", "txBody")[0];
        var text = readTextFromTxBody(txBody);
        if (!text) continue;
        var txXfrm = parseSimpleXfrm(sp.getElementsByTagNameNS("*", "txXfrm")[0]);

        var tx = sx + sw * 0.08;
        var ty = sy + sh * 0.36;
        var tw = sw * 0.84;
        if (txXfrm && txXfrm.w > 0 && txXfrm.h > 0) {
            tx = mapX(txXfrm.x);
            ty = mapY(txXfrm.y);
            tw = mapW(txXfrm.w);
        }
        var rawTextColor = readTextColorFromTxBody(txBody, styleRefs.fontColor || null);
        var finalTextColor = pickReadableTextColor(rawTextColor, fill, defTextColor || "#000000");

        textEls.push(normalizeElement({
            type: "text", text: text,
            x: tx, y: ty, w: tw,
            fontSize: getTextFontSizeFromTxBody(txBody, 16),
            color: finalTextColor,
            align: getTextAlignFromTxBody(txBody)
        }));
    }

    for (var ti = 0; ti < textEls.length; ti++) elements.push(textEls[ti]);
    return elements.length > 0;
}

function renderSmartArtFromDrawing(diagramEntry, elements, fracX, fracY, fracW, fracH, defTextColor) {
    var drawingDoc = diagramEntry.drawingDoc;
    if (!drawingDoc) return false;
    var spNodes = drawingDoc.getElementsByTagName("dsp:sp");
    if (!spNodes || spNodes.length === 0) {
        spNodes = drawingDoc.getElementsByTagNameNS("*", "sp");
    }
    if (!spNodes || spNodes.length === 0) return false;

    var pieParts = [];
    var textEls = [];
    var minX = Number.POSITIVE_INFINITY, minY = Number.POSITIVE_INFINITY;
    var maxX = Number.NEGATIVE_INFINITY, maxY = Number.NEGATIVE_INFINITY;

    for (var i = 0; i < spNodes.length; i++) {
        var sp = spNodes[i];
        var spPr = sp.getElementsByTagNameNS("*", "spPr")[0];
        if (!spPr) continue;
        var geom = spPr.getElementsByTagNameNS(A_NS, "prstGeom")[0];
        if (!geom || geom.getAttribute("prst") !== "pie") continue;

        var x = parseSimpleXfrm(spPr);
        if (!x || x.w <= 0 || x.h <= 0) continue;
        minX = Math.min(minX, x.x);
        minY = Math.min(minY, x.y);
        maxX = Math.max(maxX, x.x + x.w);
        maxY = Math.max(maxY, x.y + x.h);

        var gdNodes = geom.getElementsByTagNameNS(A_NS, "gd");
        var adj1 = 0, adj2 = 21600000;
        for (var g = 0; g < gdNodes.length; g++) {
            var name = gdNodes[g].getAttribute("name") || "";
            if (name === "adj1") adj1 = parseGdValue(gdNodes[g]);
            if (name === "adj2") adj2 = parseGdValue(gdNodes[g]);
        }

        var txBody = sp.getElementsByTagNameNS("*", "txBody")[0];
        var txXfrmNode = sp.getElementsByTagNameNS("*", "txXfrm")[0];
        var txXfrm = parseSimpleXfrm(txXfrmNode);

        pieParts.push({
            xfrm: x,
            txXfrm: txXfrm,
            text: readTextFromTxBody(txBody),
            startDeg: adj1 / 60000,
            endDeg: adj2 / 60000,
            fillColor: getShapeFill(spPr) || "#5B7FC5",
            textColor: readTextColorFromTxBody(txBody, defTextColor || "#FFFFFF")
        });
    }

    if (pieParts.length === 0 || !Number.isFinite(minX) || !Number.isFinite(maxX) || maxX <= minX || maxY <= minY) {
        return false;
    }

    var bw = maxX - minX;
    var bh = maxY - minY;
    function mapX(v) { return fracX + ((v - minX) / bw) * fracW; }
    function mapY(v) { return fracY + ((v - minY) / bh) * fracH; }
    function mapW(v) { return (v / bw) * fracW; }
    function mapH(v) { return (v / bh) * fracH; }

    for (var p = 0; p < pieParts.length; p++) {
        var part = pieParts[p];
        var sx = mapX(part.xfrm.x);
        var sy = mapY(part.xfrm.y);
        var sw = mapW(part.xfrm.w);
        var sh = mapH(part.xfrm.h);
        elements.push(normalizeElement({
            type: "shape", shape: "pie",
            x: sx, y: sy, w: sw, h: sh,
            fillColor: part.fillColor,
            strokeColor: "transparent",
            thickness: 0,
            pieStart: part.startDeg,
            pieEnd: part.endDeg
        }));

        var tx = sx + sw * 0.25;
        var ty = sy + sh * 0.42;
        var tw = sw * 0.5;
        var th = sh * 0.12;
        if (part.txXfrm && part.txXfrm.w > 0 && part.txXfrm.h > 0) {
            // Drawing part already stores text box geometry for each wedge.
            tx = mapX(part.txXfrm.x);
            ty = mapY(part.txXfrm.y);
            tw = mapW(part.txXfrm.w);
            th = mapH(part.txXfrm.h);
        } else {
            // Fallback when text transform is missing.
            var midDeg = (part.startDeg + part.endDeg) * 0.5;
            var rad = midDeg * Math.PI / 180;
            var r = Math.min(sw, sh) * 0.28;
            tx = sx + sw * 0.5 + Math.cos(rad) * r - sw * 0.18;
            ty = sy + sh * 0.5 + Math.sin(rad) * r - sh * 0.06;
            tw = sw * 0.36;
            th = sh * 0.12;
        }
        if (part.text) {
            textEls.push(normalizeElement({
                type: "text", text: part.text,
                x: tx, y: ty, w: tw,
                fontSize: Math.max(12, Math.round((th || sh * 0.12) * CANVAS_H * 0.65)),
                color: part.textColor,
                align: "center",
                fontWeight: "normal"
            }));
        }
    }
    for (var ti = 0; ti < textEls.length; ti++) elements.push(textEls[ti]);
    return true;
}

function renderSmartArtFallbackFromData(diagramEntry, elements, fracX, fracY, fracW, fracH, defTextColor) {
    var dataDoc = diagramEntry.dataDoc;
    if (!dataDoc) return false;
    var ptNodes = dataDoc.getElementsByTagNameNS("*", "pt");
    var labels = [];
    for (var i = 0; i < ptNodes.length; i++) {
        var pt = ptNodes[i];
        var t = pt.getAttribute("type") || "";
        if (t === "doc" || t === "pres" || t === "parTrans" || t === "sibTrans") continue;
        var tx = readTextFromTxBody(pt);
        if (tx) labels.push(tx);
    }
    labels = labels.slice(0, 7);
    if (labels.length === 0) return false;

    var palette = ["#466CB4", "#6E88C8", "#95A6D2", "#B5C1DF", "#7A95CF", "#5D7FC0", "#8AA0CF"];
    var textEls = [];
    var n = labels.length;
    for (var li = 0; li < n; li++) {
        var start = (li * 360) / n;
        var end = ((li + 1) * 360) / n;
        var offset = li === 0 ? 0.015 : 0;
        var midRad = ((start + end) * 0.5) * Math.PI / 180;
        var ox = Math.cos(midRad) * offset;
        var oy = Math.sin(midRad) * offset;
        elements.push(normalizeElement({
            type: "shape", shape: "pie",
            x: fracX + ox, y: fracY + oy, w: fracW, h: fracH,
            fillColor: palette[li % palette.length],
            strokeColor: "transparent", thickness: 0,
            pieStart: start, pieEnd: end
        }));
        textEls.push(normalizeElement({
            type: "text", text: labels[li],
            x: fracX + fracW * (0.5 + Math.cos(midRad) * 0.28) - fracW * 0.12 + ox,
            y: fracY + fracH * (0.5 + Math.sin(midRad) * 0.28) - fracH * 0.05 + oy,
            w: fracW * 0.24,
            fontSize: 20, color: defTextColor || "#FFFFFF", align: "center"
        }));
    }
    for (var tj = 0; tj < textEls.length; tj++) elements.push(textEls[tj]);
    return true;
}

function renderBarChart(chartData, elements, fracX, fracY, fracW, fracH, defTextColor) {
    var categories = chartData.categories || [];
    var series = chartData.series || [];
    if (categories.length === 0 || series.length === 0) return;

    var topPad = fracH * 0.06;
    var bottomPad = fracH * 0.20;
    var leftPad = fracW * 0.08;
    var rightPad = fracW * 0.04;
    var plotX = fracX + leftPad;
    var plotY = fracY + topPad;
    var plotW = Math.max(0.01, fracW - leftPad - rightPad);
    var plotH = Math.max(0.01, fracH - topPad - bottomPad);

    // Prefer chart XML area fills when available.
    var chartAreaFill = chartData.chartAreaFill;
    var plotAreaFill = chartData.plotAreaFill;

    if (chartAreaFill && chartAreaFill !== "transparent") {
        elements.push(normalizeElement({
            type: "shape", shape: "rect",
            x: fracX, y: fracY, w: fracW, h: fracH,
            fillColor: chartAreaFill,
            strokeColor: "transparent",
            thickness: 0
        }));
    }

    if (plotAreaFill && plotAreaFill !== "transparent") {
        elements.push(normalizeElement({
            type: "shape", shape: "rect",
            x: plotX, y: plotY, w: plotW, h: plotH,
            fillColor: plotAreaFill,
            strokeColor: "transparent",
            thickness: 0
        }));
    }

    var maxV = Math.max(1, chartData.maxValue || 1);
    var axisMax = Math.ceil(maxV / 10) * 10;
    if (axisMax <= 0) axisMax = 10;

    var labelColor = defTextColor || "#666666";
    var lightText = colorLuma(labelColor) > 0.7;
    var gridColor = lightText ? "#8CB1C8" : "#CCCCCC";

    // Horizontal grid lines + Y labels
    for (var gi = 0; gi <= 6; gi++) {
        var t = gi / 6;
        var gy = plotY + plotH * t;
        var value = Math.round(axisMax * (1 - t));
        elements.push(normalizeElement({
            type: "shape", shape: "line",
            x1: plotX, y1: gy,
            x2: plotX + plotW, y2: gy,
            color: gridColor, thickness: 1
        }));
        elements.push(normalizeElement({
            type: "text", text: String(value),
            x: plotX - fracW * 0.05, y: gy - fracH * 0.015, w: fracW * 0.04,
            fontSize: 10, color: labelColor, align: "right"
        }));
    }

    // Bars
    var groupW = plotW / categories.length;
    var innerPad = groupW * 0.12;
    var barGap = groupW * 0.06;
    var barsAreaW = groupW - innerPad * 2;
    var barW = Math.max(groupW * 0.08, (barsAreaW - barGap * (series.length - 1)) / series.length);

    var defaultSeriesColors = ["#AAD232", "#DCBE32", "#ED7D31", "#5B9BD5", "#70AD47"];

    for (var ci = 0; ci < categories.length; ci++) {
        var gx = plotX + ci * groupW + innerPad;
        for (var si = 0; si < series.length; si++) {
            var s = series[si];
            var val = (s.values && ci < s.values.length) ? s.values[ci] : 0;
            val = Number.isFinite(val) ? val : 0;
            var bh = plotH * (Math.max(0, val) / axisMax);
            var bx = gx + si * (barW + barGap);
            var by = plotY + plotH - bh;
            elements.push(normalizeElement({
                type: "shape", shape: "rect",
                x: bx, y: by, w: barW, h: Math.max(fracH * 0.003, bh),
                fillColor: defaultSeriesColors[si % defaultSeriesColors.length],
                strokeColor: "transparent", thickness: 0
            }));
        }

        // Category labels
        elements.push(normalizeElement({
            type: "text", text: categories[ci],
            x: plotX + ci * groupW, y: plotY + plotH + fracH * 0.02, w: groupW,
            fontSize: 11, color: labelColor, align: "center"
        }));
    }

    // Legend at bottom-center
    var legY = fracY + fracH - fracH * 0.06;
    var blockSize = Math.min(fracW * 0.03, fracH * 0.03);
    var gapW = fracW * 0.03;
    var totalW = series.length * (blockSize + gapW + fracW * 0.12);
    var curX = fracX + (fracW - totalW) / 2;
    for (var li = 0; li < series.length; li++) {
        var c = defaultSeriesColors[li % defaultSeriesColors.length];
        elements.push(normalizeElement({
            type: "shape", shape: "rect",
            x: curX, y: legY, w: blockSize, h: blockSize,
            fillColor: c, strokeColor: "transparent", thickness: 0
        }));
        elements.push(normalizeElement({
            type: "text", text: series[li].name || ("Series " + (li + 1)),
            x: curX + blockSize + fracW * 0.01, y: legY - fracH * 0.008, w: fracW * 0.12,
            fontSize: 10, color: labelColor, align: "left"
        }));
        curX += blockSize + gapW + fracW * 0.12;
    }
}

// --- Parse a:tbl (table) from graphicData ---
function parseTable(graphicData, elements, fracX, fracY, fracW, fracH, defTextColor) {
    if (!graphicData) return;
    var tbl = graphicData.getElementsByTagNameNS(A_NS, "tbl")[0];
    if (!tbl) return;
    var tblGrid = tbl.getElementsByTagNameNS(A_NS, "tblGrid")[0];
    var rows = tbl.getElementsByTagNameNS(A_NS, "tr");
    if (!rows || rows.length === 0) return;
    console.log("[TBL] Table: " + rows.length + " rows, cols=" + (tblGrid ? tblGrid.getElementsByTagNameNS(A_NS, "gridCol").length : "?"));

    // Calculate column widths from tblGrid
    var colWidths = [];
    var totalW = 0;
    if (tblGrid) {
        var gridCols = tblGrid.getElementsByTagNameNS(A_NS, "gridCol");
        for (var i = 0; i < gridCols.length; i++) {
            var w = parseInt(gridCols[i].getAttribute("w")) || 100000;
            colWidths.push(w); totalW += w;
        }
    }
    if (totalW === 0) totalW = 1;

    // Calculate row heights
    var rowHeights = [];
    var totalH = 0;
    for (var r = 0; r < rows.length; r++) {
        var h = parseInt(rows[r].getAttribute("h")) || 300000;
        rowHeights.push(h); totalH += h;
    }
    if (totalH === 0) totalH = 1;

    // Render cells
    var curY = fracY;
    for (var r = 0; r < rows.length; r++) {
        var rh = fracH * (rowHeights[r] / totalH);
        var cells = rows[r].getElementsByTagNameNS(A_NS, "tc");
        var curX = fracX;
        for (var c = 0; c < cells.length && c < colWidths.length; c++) {
            var cw = fracW * (colWidths[c] / totalW);

            // Cell background
            var tcPr = cells[c].getElementsByTagNameNS(A_NS, "tcPr")[0];
            var cellFill = null;
            if (tcPr) {
                var sf = tcPr.getElementsByTagNameNS(A_NS, "solidFill")[0];
                if (sf) cellFill = resolveColor(sf);
            }

            // Cell border
            elements.push(normalizeElement({
                type: "shape", shape: "rect", x: curX, y: curY, w: cw, h: rh,
                fillColor: cellFill || "transparent", strokeColor: "#AAA", thickness: 1, rotation: 0
            }));

            // Cell text
            var txBody = cells[c].getElementsByTagNameNS(A_NS, "txBody")[0];
            if (txBody) {
                var cellParas = parseParagraphs(txBody, 10, defTextColor);
                var textY = curY + 0.005;
                cellParas.forEach(function (p) {
                    if (!p.isEmpty) {
                        elements.push(normalizeElement({
                            type: "text", text: p.text, x: curX + 0.005, y: textY, w: cw - 0.01,
                            fontSize: Math.min(p.fontSize, 14), color: p.color,
                            fontWeight: p.fontWeight, fontStyle: p.italic ? "italic" : "normal", align: p.align
                        }));
                    }
                    textY += p.fontSize * 1.2 / CANVAS_H;
                });
            }
            curX += cw;
        }
        curY += rh;
    }
}

// ========================================================================

export { parseShapeTree };
