// ============================================================================
// shape-parsers.js - Shape tree parsing (sp, pic, cxnSp, grpSp, graphicFrame)
// ============================================================================

import { A_NS, P_NS, R_NS, CANVAS_H, normalizeElement } from "./constants.js";
import { resolveColor, themeColors } from "./color-utils.js";
import { parseParagraphs, parseOutline, getPresetGeometry, getShapeFill } from "./text-parser.js";

function parseShapeTree(spTreeNode, slideW, slideH, images, relsAll, opts) {
    opts = opts || {};
    var elements = [];
    var skipPlaceholders = opts.skipPlaceholders || false;
    var hasBgImage = opts.hasBgImage || false;
    var layoutStyles = opts.layoutStyles || {};
    var chartDataMap = opts.chartDataMap || {};
    // Prefer theme text color regardless of background image presence.
    // Per-template placeholder/style inheritance can still override this later.
    var defaultTextColor = themeColors.tx1 || "#333";
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
            parseSp(child, elements, slideW, slideH, skipPlaceholders, defaultTextColor, toFracX, toFracY, toFracW, toFracH, layoutStyles, hasBgImage);
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
            parseGraphicFrame(child, elements, slideW, slideH, images, relsAll, chartDataMap, defaultTextColor, toFracX, toFracY, toFracW, toFracH);
        }
    }
    return elements;
}

// --- Parse sp (shape with optional text) ---
function parseSp(sp, elements, slideW, slideH, skipPH, defTextColor, fx, fy, fw, fh, layoutStyles, hasBgImage) {
    layoutStyles = layoutStyles || {};
    hasBgImage = hasBgImage || false;
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
        gOffX: newGOffX, gOffY: newGOffY,
        gScaleX: newGScaleX, gScaleY: newGScaleY
    };
    console.log("[GRPSP]   off=(" + offX + "," + offY + ") ext=(" + extW + "," + extH + ") chOff=(" + chOffX + "," + chOffY + ") chExt=(" + chExtW + "," + chExtH + ") scale=(" + newGScaleX.toFixed(3) + "," + newGScaleY.toFixed(3) + ")");

    var childElements = parseShapeTree(grpSp, slideW, slideH, images, relsAll, childOpts);
    childElements.forEach(function (el) { elements.push(el); });
}

// --- Parse graphicFrame (chart / table / diagram) ---
function parseGraphicFrame(gf, elements, slideW, slideH, images, relsAll, chartDataMap, defTextColor, fx, fy, fw, fh) {
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

    // PowerPoint-like chart area panel.
    elements.push(normalizeElement({
        type: "shape", shape: "rect",
        x: fracX, y: fracY, w: fracW, h: fracH,
        fillColor: "#ECECEC", strokeColor: "#D4D4D4", thickness: 1
    }));

    // Plot area is slightly brighter than chart area.
    elements.push(normalizeElement({
        type: "shape", shape: "rect",
        x: plotX, y: plotY, w: plotW, h: plotH,
        fillColor: "#F8F8F8", strokeColor: "#CFCFCF", thickness: 1
    }));

    var maxV = Math.max(1, chartData.maxValue || 1);
    var axisMax = Math.ceil(maxV / 10) * 10;
    if (axisMax <= 0) axisMax = 10;

    // Horizontal grid lines + Y labels
    for (var gi = 0; gi <= 6; gi++) {
        var t = gi / 6;
        var gy = plotY + plotH * t;
        var value = Math.round(axisMax * (1 - t));
        elements.push(normalizeElement({
            type: "shape", shape: "line",
            x1: plotX, y1: gy,
            x2: plotX + plotW, y2: gy,
            color: "#CCCCCC", thickness: 1
        }));
        elements.push(normalizeElement({
            type: "text", text: String(value),
            x: plotX - fracW * 0.05, y: gy - fracH * 0.015, w: fracW * 0.04,
            fontSize: 10, color: "#666666", align: "right"
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
            fontSize: 11, color: "#666666", align: "center"
        }));
    }

    // Legend at bottom-center
    var legY = fracY + fracH - fracH * 0.06;
    var blockW = fracW * 0.06;
    var gapW = fracW * 0.03;
    var totalW = series.length * (blockW + gapW + fracW * 0.12);
    var curX = fracX + (fracW - totalW) / 2;
    for (var li = 0; li < series.length; li++) {
        var c = defaultSeriesColors[li % defaultSeriesColors.length];
        elements.push(normalizeElement({
            type: "shape", shape: "rect",
            x: curX, y: legY, w: blockW, h: fracH * 0.02,
            fillColor: c, strokeColor: "transparent", thickness: 0
        }));
        elements.push(normalizeElement({
            type: "text", text: series[li].name || ("Series " + (li + 1)),
            x: curX + blockW + fracW * 0.01, y: legY - fracH * 0.008, w: fracW * 0.12,
            fontSize: 10, color: "#666666", align: "left"
        }));
        curX += blockW + gapW + fracW * 0.12;
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
