// ============================================================================
// pptx-parser.js - Main PPTX orchestrator (coordinates all parsing)
// ============================================================================

import { SLIDE_EMU_W, SLIDE_EMU_H, A_NS } from "./constants.js";
import { parseThemeXml, resolveColor } from "./color-utils.js";
import { parseRelsFile, buildImageMap, loadImageAsDataUrl, clearImageCache } from "./zip-helpers.js";
import { extractBackground, extractBlipEffects } from "./background.js";
import { extractPlaceholderStyles, extractMasterTxStyles } from "./style-inheritance.js";
import { parseSlideXml } from "./slide-parser.js";

var C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";

function perfNow() {
    return (typeof performance !== "undefined" && performance.now) ? performance.now() : Date.now();
}

function pushPerfStep(list, name, startedAt) {
    list.push({ name: name, ms: perfNow() - startedAt });
}

function getChartText(node) {
    if (!node) return "";
    var v = node.getElementsByTagNameNS(C_NS, "v")[0];
    if (v && v.textContent) return v.textContent;
    var t = node.getElementsByTagNameNS(C_NS, "t")[0];
    if (t && t.textContent) return t.textContent;
    return "";
}

function extractSeriesName(serNode) {
    if (!serNode) return "Series";
    var tx = serNode.getElementsByTagNameNS(C_NS, "tx")[0];
    if (!tx) return "Series";
    var strRef = tx.getElementsByTagNameNS(C_NS, "strRef")[0];
    if (strRef) {
        var strCache = strRef.getElementsByTagNameNS(C_NS, "strCache")[0];
        if (strCache) {
            var pt = strCache.getElementsByTagNameNS(C_NS, "pt")[0];
            if (pt) {
                var pv = getChartText(pt);
                if (pv) return pv;
            }
        }
    }
    var tv = tx.getElementsByTagNameNS(C_NS, "v")[0];
    if (tv && tv.textContent) return tv.textContent;
    return "Series";
}

function extractCategoryValues(serNode) {
    var values = [];
    if (!serNode) return values;
    var cat = serNode.getElementsByTagNameNS(C_NS, "cat")[0];
    if (!cat) return values;
    var strRef = cat.getElementsByTagNameNS(C_NS, "strRef")[0];
    if (strRef) {
        var strCache = strRef.getElementsByTagNameNS(C_NS, "strCache")[0];
        if (strCache) {
            var pts = strCache.getElementsByTagNameNS(C_NS, "pt");
            for (var i = 0; i < pts.length; i++) {
                values.push(getChartText(pts[i]));
            }
            return values;
        }
    }
    var numRef = cat.getElementsByTagNameNS(C_NS, "numRef")[0];
    if (numRef) {
        var numCache = numRef.getElementsByTagNameNS(C_NS, "numCache")[0];
        if (numCache) {
            var npts = numCache.getElementsByTagNameNS(C_NS, "pt");
            for (var j = 0; j < npts.length; j++) {
                values.push(getChartText(npts[j]));
            }
        }
    }
    return values;
}

function extractNumValues(serNode) {
    var values = [];
    if (!serNode) return values;
    var val = serNode.getElementsByTagNameNS(C_NS, "val")[0];
    if (!val) return values;
    var numRef = val.getElementsByTagNameNS(C_NS, "numRef")[0];
    if (numRef) {
        var numCache = numRef.getElementsByTagNameNS(C_NS, "numCache")[0];
        if (numCache) {
            var pts = numCache.getElementsByTagNameNS(C_NS, "pt");
            for (var i = 0; i < pts.length; i++) {
                var t = getChartText(pts[i]);
                var n = parseFloat(t);
                values.push(Number.isFinite(n) ? n : 0);
            }
            return values;
        }
    }
    var numLit = val.getElementsByTagNameNS(C_NS, "numLit")[0];
    if (numLit) {
        var lpts = numLit.getElementsByTagNameNS(C_NS, "pt");
        for (var j = 0; j < lpts.length; j++) {
            var lt = getChartText(lpts[j]);
            var ln = parseFloat(lt);
            values.push(Number.isFinite(ln) ? ln : 0);
        }
    }
    return values;
}

function extractFillFromSpPr(spPr) {
    if (!spPr) return null;
    var noFill = spPr.getElementsByTagNameNS(A_NS, "noFill")[0];
    if (noFill) return "transparent";

    var solidFill = spPr.getElementsByTagNameNS(A_NS, "solidFill")[0];
    if (solidFill) {
        var c = resolveColor(solidFill);
        if (c) return c;
    }

    var gradFill = spPr.getElementsByTagNameNS(A_NS, "gradFill")[0];
    if (gradFill) {
        var gs = gradFill.getElementsByTagNameNS(A_NS, "gs");
        if (gs.length > 0) {
            var gc = resolveColor(gs[0]);
            if (gc) return gc;
        }
    }
    return null;
}

function getDirectChildByTagNameNS(parent, namespaceUri, localName) {
    if (!parent) return null;
    for (var i = 0; i < parent.childNodes.length; i++) {
        var child = parent.childNodes[i];
        if (child.nodeType !== 1) continue;
        if (child.namespaceURI === namespaceUri && child.localName === localName) return child;
    }
    return null;
}

async function buildChartDataMap(zip, slideBasePath, relsAll) {
    var map = {};
    if (!relsAll) return map;
    for (var rId in relsAll) {
        var target = relsAll[rId] || "";
        if (target.indexOf("chart") === -1 || target.indexOf(".xml") === -1) continue;
        var fullPath = (slideBasePath + target).replace(/[^/]+\/\.\.\//g, "");
        var cf = zip.file(fullPath);
        if (!cf) continue;
        try {
            var cdoc = new DOMParser().parseFromString(await cf.async("string"), "application/xml");
            var chart = cdoc.getElementsByTagNameNS(C_NS, "chart")[0];
            if (!chart) continue;
            var plotArea = chart.getElementsByTagNameNS(C_NS, "plotArea")[0];
            if (!plotArea) continue;

            // Chart/plot area fills (c:spPr with a:* fill children).
            var chartAreaFill = null;
            var plotAreaFill = null;
            var chartSpPr = getDirectChildByTagNameNS(chart, C_NS, "spPr");
            if (chartSpPr) chartAreaFill = extractFillFromSpPr(chartSpPr);
            var plotSpPr = getDirectChildByTagNameNS(plotArea, C_NS, "spPr");
            if (plotSpPr) plotAreaFill = extractFillFromSpPr(plotSpPr);

            var chartTypeNode =
                plotArea.getElementsByTagNameNS(C_NS, "barChart")[0] ||
                plotArea.getElementsByTagNameNS(C_NS, "lineChart")[0] ||
                plotArea.getElementsByTagNameNS(C_NS, "pieChart")[0] ||
                plotArea.getElementsByTagNameNS(C_NS, "areaChart")[0] ||
                null;
            if (!chartTypeNode) continue;

            var serNodes = chartTypeNode.getElementsByTagNameNS(C_NS, "ser");
            if (!serNodes || serNodes.length === 0) continue;

            var categories = [];
            var series = [];
            var maxValue = 0;
            for (var i = 0; i < serNodes.length; i++) {
                var s = serNodes[i];
                var cats = extractCategoryValues(s);
                if (categories.length === 0 && cats.length > 0) categories = cats;
                var vals = extractNumValues(s);
                for (var v = 0; v < vals.length; v++) {
                    if (vals[v] > maxValue) maxValue = vals[v];
                }
                series.push({
                    name: extractSeriesName(s),
                    values: vals
                });
            }

            if (categories.length === 0 && series.length > 0) {
                var fallbackLen = series[0].values.length;
                for (var c = 0; c < fallbackLen; c++) categories.push(String(c + 1));
            }

            map[rId] = {
                type: chartTypeNode.localName,
                categories: categories,
                series: series,
                maxValue: maxValue,
                chartAreaFill: chartAreaFill,
                plotAreaFill: plotAreaFill
            };
            console.log("[PPTX] Chart loaded for " + rId + ": type=" + chartTypeNode.localName + " cats=" + categories.length + " series=" + series.length);
        } catch (e) {
            console.warn("[PPTX] Failed to parse chart for " + rId + ":", e);
        }
    }
    return map;
}

async function loadXmlPart(zip, fullPath) {
    var f = zip.file(fullPath);
    if (!f) return null;
    try {
        return new DOMParser().parseFromString(await f.async("string"), "application/xml");
    } catch (e) {
        console.warn("[PPTX] Failed to parse XML part:", fullPath, e);
        return null;
    }
}

function resolvePartPath(basePath, target) {
    return (basePath + target).replace(/[^/]+\/\.\.\//g, "");
}

async function buildDiagramDataMap(zip, slideBasePath, relsAll) {
    var map = {};
    if (!relsAll) return map;

    var globalLayoutPath = null;
    var globalQuickStylePath = null;
    var globalColorsPath = null;
    var globalDrawingPath = null;

    for (var rid0 in relsAll) {
        var t0 = (relsAll[rid0] || "").toLowerCase();
        if (t0.indexOf("/diagrams/") === -1) continue;
        if (t0.indexOf("layout") !== -1 && !globalLayoutPath) globalLayoutPath = resolvePartPath(slideBasePath, relsAll[rid0]);
        else if (t0.indexOf("quickstyle") !== -1 && !globalQuickStylePath) globalQuickStylePath = resolvePartPath(slideBasePath, relsAll[rid0]);
        else if (t0.indexOf("colors") !== -1 && !globalColorsPath) globalColorsPath = resolvePartPath(slideBasePath, relsAll[rid0]);
        else if (t0.indexOf("drawing") !== -1 && !globalDrawingPath) globalDrawingPath = resolvePartPath(slideBasePath, relsAll[rid0]);
    }

    for (var rId in relsAll) {
        var target = relsAll[rId] || "";
        if (target.indexOf("/diagrams/") === -1 || target.toLowerCase().indexOf("data") === -1 || target.toLowerCase().indexOf(".xml") === -1) {
            continue;
        }

        var dataPath = resolvePartPath(slideBasePath, target);
        var dataDoc = await loadXmlPart(zip, dataPath);
        if (!dataDoc) continue;

        var dataDir = dataPath.replace(/[^/]*$/, "");
        var relsPath = dataPath.replace(/([^/]+)$/, "_rels/$1.rels");
        var rels = await parseRelsFile(zip, relsPath);

        var entry = {
            dataDoc: dataDoc,
            drawingDoc: null,
            layoutDoc: null,
            quickStyleDoc: null,
            colorsDoc: null,
            drawingImageMap: {}
        };
        var drawingPath = null;

        // 1) Prefer data part local rels if present
        for (var drId in rels.all) {
            var partTarget = rels.all[drId] || "";
            var fullPath = resolvePartPath(dataDir, partTarget);
            var lower = fullPath.toLowerCase();
            if (lower.indexOf("drawing") !== -1) {
                entry.drawingDoc = await loadXmlPart(zip, fullPath);
                if (entry.drawingDoc) drawingPath = fullPath;
            }
            else if (lower.indexOf("layout") !== -1) entry.layoutDoc = await loadXmlPart(zip, fullPath);
            else if (lower.indexOf("quickstyle") !== -1) entry.quickStyleDoc = await loadXmlPart(zip, fullPath);
            else if (lower.indexOf("colors") !== -1) entry.colorsDoc = await loadXmlPart(zip, fullPath);
        }

        // 2) Fallback: dataModelExt relId often points to drawing part in slide rels
        if (!entry.drawingDoc) {
            var extNodes = dataDoc.getElementsByTagNameNS("*", "dataModelExt");
            if (extNodes && extNodes.length > 0) {
                var drawingRid = extNodes[0].getAttribute("relId") || extNodes[0].getAttribute("r:relId") || "";
                if (drawingRid && relsAll[drawingRid]) {
                    var fullDrawingPath = resolvePartPath(slideBasePath, relsAll[drawingRid]);
                    entry.drawingDoc = await loadXmlPart(zip, fullDrawingPath);
                    if (entry.drawingDoc) drawingPath = fullDrawingPath;
                }
            }
        }

        // 3) Fallback: slide rels often contain layout/quickStyle/colors directly
        if (!entry.layoutDoc && globalLayoutPath) entry.layoutDoc = await loadXmlPart(zip, globalLayoutPath);
        if (!entry.quickStyleDoc && globalQuickStylePath) entry.quickStyleDoc = await loadXmlPart(zip, globalQuickStylePath);
        if (!entry.colorsDoc && globalColorsPath) entry.colorsDoc = await loadXmlPart(zip, globalColorsPath);
        if (!entry.drawingDoc && globalDrawingPath) {
            entry.drawingDoc = await loadXmlPart(zip, globalDrawingPath);
            if (entry.drawingDoc) drawingPath = globalDrawingPath;
        }

        if (drawingPath) {
            var drawingBase = drawingPath.replace(/[^/]*$/, "");
            var drawingRelsPath = drawingPath.replace(/([^/]+)$/, "_rels/$1.rels");
            var drawingRels = await parseRelsFile(zip, drawingRelsPath);
            for (var imgRid in drawingRels.all) {
                var imgTarget = drawingRels.all[imgRid] || "";
                var imgData = await loadImageAsDataUrl(zip, drawingBase, imgTarget);
                if (imgData) entry.drawingImageMap[imgRid] = imgData;
            }
        }

        map[rId] = entry;
        console.log("[PPTX] Diagram loaded for " + rId + ": data=" + !!entry.dataDoc + " drawing=" + !!entry.drawingDoc + " quickStyle=" + !!entry.quickStyleDoc + " colors=" + !!entry.colorsDoc + " drawImages=" + Object.keys(entry.drawingImageMap).length);
    }

    return map;
}

async function parsePptx(arrayBuffer, onStructureReady, onSlideImagesReady, onAllImagesReady) {
    var t0 = perfNow();
    var perfSummary = { totalMs: 0, steps: [], slides: [] };
    console.log("[PPTX] === Starting PPTX parse ===");
    clearImageCache();
    var zip = new JSZip();
    var stepStart = perfNow();
    await zip.loadAsync(arrayBuffer);
    pushPerfStep(perfSummary.steps, "zip.loadAsync", stepStart);
    console.log("[PPTX] ZIP loaded, files: " + Object.keys(zip.files).length);

    // Parse theme colors
    stepStart = perfNow();
    await parseThemeXml(zip);
    pushPerfStep(perfSummary.steps, "parseThemeXml", stepStart);

    // Slide dimensions
    var slideW = SLIDE_EMU_W, slideH = SLIDE_EMU_H;
    stepStart = perfNow();
    var pf = zip.file("ppt/presentation.xml");
    if (pf) {
        var pdoc = new DOMParser().parseFromString(await pf.async("string"), "application/xml");
        var ss = pdoc.getElementsByTagName("p:sldSz")[0];
        if (ss) {
            slideW = parseInt(ss.getAttribute("cx")) || SLIDE_EMU_W;
            slideH = parseInt(ss.getAttribute("cy")) || SLIDE_EMU_H;
        }
    }
    pushPerfStep(perfSummary.steps, "presentation.xml", stepStart);
    console.log("[PPTX] Slide dimensions: " + slideW + " x " + slideH + " EMU");

    // Enumerate slides
    stepStart = perfNow();
    var slideFiles = [];
    zip.forEach(function (path) {
        var m = path.match(/^ppt\/slides\/slide(\d+)\.xml$/);
        if (m) slideFiles.push({ path: path, num: parseInt(m[1]) });
    });
    slideFiles.sort(function (a, b) { return a.num - b.num; });
    pushPerfStep(perfSummary.steps, "enumerateSlides", stepStart);
    console.log("[PPTX] Found " + slideFiles.length + " slides: " + slideFiles.map(function(s){return "slide"+s.num;}).join(", "));

    // Background cache for layout/master
    var bgCache = {};
    var layoutStylesCache = {};
    var masterTxStylesCache = {};
    var relsCache = {};

    async function getRelsCached(relsPath) {
        if (relsCache[relsPath] !== undefined) return relsCache[relsPath];
        relsCache[relsPath] = await parseRelsFile(zip, relsPath);
        return relsCache[relsPath];
    }

    async function getLayerBackground(xmlPath, relsPath, basePath) {
        if (bgCache[xmlPath] !== undefined) return bgCache[xmlPath];
        var f = zip.file(xmlPath);
        if (!f) { bgCache[xmlPath] = { bg: null, masterTarget: null, basePath: basePath }; return bgCache[xmlPath]; }
        var xml = await f.async("string");
        var rels = await parseRelsFile(zip, relsPath);
        var bg = await extractBackground(xml, zip, basePath, rels.all, slideW, slideH);
        bgCache[xmlPath] = { bg: bg, masterTarget: rels.master, basePath: basePath };
        return bgCache[xmlPath];
    }

    // Process each slide
    var newSlides = [];
    for (var i = 0; i < slideFiles.length; i++) {
        var sf = slideFiles[i];
        var slidePerf = { slide: sf.num, steps: [], totalMs: 0 };
        var slideStart = perfNow();
        stepStart = perfNow();
        var xmlStr = await zip.file(sf.path).async("string");
        var slideRels = await parseRelsFile(zip, "ppt/slides/_rels/slide" + sf.num + ".xml.rels");
        pushPerfStep(slidePerf.steps, "slideXml+rels", stepStart);
        console.log("[PPTX] Slide " + sf.num + " rels: images=" + Object.keys(slideRels.images).length + " layout=" + (slideRels.layout||"none"));
        console.log("[PPTX]   all rels: " + JSON.stringify(slideRels.all));

        // Build image map
        stepStart = perfNow();
        // Image loading is deferred to phase 2 (parallel across all slides).
        var images = {};
        pushPerfStep(slidePerf.steps, "buildImageMap", stepStart);
        // Accumulate image-loading context; filled as layout/master paths are resolved.
        var _ctxLayoutBase = null, _ctxLayoutImageRels = {}, _ctxMasterBase = null, _ctxMasterImageRels = {};

        // Build chart data map from related chart parts
        stepStart = perfNow();
        var chartDataMap = await buildChartDataMap(zip, "ppt/slides/", slideRels.all);
        var diagramDataMap = await buildDiagramDataMap(zip, "ppt/slides/", slideRels.all);
        pushPerfStep(slidePerf.steps, "buildChart+DiagramData", stepStart);

        // === Background inheritance: slide → layout → master ===
        stepStart = perfNow();
        var bgResult = await extractBackground(xmlStr, zip, "ppt/slides/", slideRels.all, slideW, slideH);
        pushPerfStep(slidePerf.steps, "extractBackground", stepStart);
        console.log("[PPTX] Slide " + sf.num + " bg chain: slide=" + (bgResult ? (typeof bgResult === "string" ? "image" : JSON.stringify(bgResult)) : "none"));

        if (!bgResult && slideRels.layout) {
            stepStart = perfNow();
            var layoutPath = ("ppt/slides/" + slideRels.layout).replace(/[^/]+\/\.\.\//g, "");
            console.log("[PPTX]   checking layout: " + layoutPath);
            var layoutBase = layoutPath.replace(/[^/]*$/, "");
            var layoutRelsPath = layoutPath.replace(/([^/]+)$/, "_rels/$1.rels");
            var layerData = await getLayerBackground(layoutPath, layoutRelsPath, layoutBase);
            bgResult = layerData.bg;
            console.log("[PPTX]   layout bg=" + (bgResult ? (typeof bgResult === "string" ? "image" : JSON.stringify(bgResult)) : "none"));

            // Try master
            if (!bgResult && layerData.masterTarget) {
                var masterPath = (layoutBase + layerData.masterTarget).replace(/[^/]+\/\.\.\//g, "");
                console.log("[PPTX]   checking master: " + masterPath);
                var masterBase = masterPath.replace(/[^/]*$/, "");
                var masterRelsPath = masterPath.replace(/([^/]+)$/, "_rels/$1.rels");
                var masterData = await getLayerBackground(masterPath, masterRelsPath, masterBase);
                bgResult = masterData.bg;
                console.log("[PPTX]   master bg=" + (bgResult ? (typeof bgResult === "string" ? "image" : JSON.stringify(bgResult)) : "none"));
            }
            pushPerfStep(slidePerf.steps, "backgroundInheritance", stepStart);
        }

        var hasBgImage = !!(bgResult && (typeof bgResult === "string" || (typeof bgResult === "object" && !!bgResult.image)));
        var bgImageRid = (bgResult && typeof bgResult === "object" && bgResult.bgImageRid) ? bgResult.bgImageRid : null;
        console.log("[PPTX] Slide " + sf.num + " hasBgImage=" + hasBgImage + " bgImageRid=" + (bgImageRid || "none"));

        // Parse slide shapes
        console.log("[PPTX] Parsing slide " + sf.num + " shapes...");
        // Extract layout placeholder styles for inheritance
        var layoutStyles = {};
        var masterTxStyles = { titleColor: null, bodyColor: null, otherColor: null };
        if (slideRels.layout) {
            stepStart = perfNow();
            var layoutPathStyles = ("ppt/slides/" + slideRels.layout).replace(/[^/]+\/\.\.\//g, "");
            // Also read master txStyles
            if (layoutStylesCache[layoutPathStyles]) {
                layoutStyles = layoutStylesCache[layoutPathStyles];
            } else {
                layoutStyles = await extractPlaceholderStyles(zip, layoutPathStyles);
                layoutStylesCache[layoutPathStyles] = layoutStyles;
            }
            var layoutRelsPathStyles = layoutPathStyles.replace(/([^/]+)$/, "_rels/$1.rels");
            var layoutRelsForMaster = await getRelsCached(layoutRelsPathStyles);
            if (layoutRelsForMaster.master) {
                var masterPathTx = (layoutPathStyles.replace(/[^/]*$/, "") + layoutRelsForMaster.master).replace(/[^/]+\/\.\.\//g, "");
                if (masterTxStylesCache[masterPathTx]) {
                    masterTxStyles = masterTxStylesCache[masterPathTx];
                } else {
                    masterTxStyles = await extractMasterTxStyles(zip, masterPathTx);
                    masterTxStylesCache[masterPathTx] = masterTxStyles;
                }
            }
            pushPerfStep(slidePerf.steps, "extractLayoutStyles", stepStart);
        }
        // Apply master txStyles as fallback fontRefColor for placeholders.
        // Body/other text should still inherit master colors on bg-image slides.
        if (masterTxStyles.bodyColor) {
            if (!layoutStyles.subTitle) layoutStyles.subTitle = {};
            if (!layoutStyles.subTitle.fontRefColor && !layoutStyles.subTitle.color) layoutStyles.subTitle.fontRefColor = masterTxStyles.bodyColor;
            if (!layoutStyles.body) layoutStyles.body = {};
            if (!layoutStyles.body.fontRefColor && !layoutStyles.body.color) layoutStyles.body.fontRefColor = masterTxStyles.bodyColor;
        }
        if (masterTxStyles.otherColor) {
            if (!layoutStyles[""]) layoutStyles[""] = {};
            if (!layoutStyles[""].fontRefColor) layoutStyles[""].fontRefColor = masterTxStyles.otherColor;
        }
        if (!hasBgImage && masterTxStyles.titleColor) {
            if (!layoutStyles.title) layoutStyles.title = {};
            if (!layoutStyles.title.fontRefColor) layoutStyles.title.fontRefColor = masterTxStyles.titleColor;
            if (!layoutStyles.ctrTitle) layoutStyles.ctrTitle = {};
            if (!layoutStyles.ctrTitle.fontRefColor) layoutStyles.ctrTitle.fontRefColor = masterTxStyles.titleColor;
        }
        // Master placeholder fontRef colors generally apply to all slides.
        // On bg-image slides, title placeholders should not be forced to tx1-like dark colors.
        for (var phKey in masterTxStyles.phFontRef) {
            if (hasBgImage && (phKey === "title" || phKey === "ctrTitle")) continue;
            if (!layoutStyles[phKey]) layoutStyles[phKey] = {};
            if (!layoutStyles[phKey].fontRefColor) layoutStyles[phKey].fontRefColor = masterTxStyles.phFontRef[phKey];
        }
        // Map master 'body' fontRef to 'subTitle' if subTitle doesn't have its own
        if (masterTxStyles.phFontRef.body && (!layoutStyles.subTitle || !layoutStyles.subTitle.fontRefColor)) {
            if (!layoutStyles.subTitle) layoutStyles.subTitle = {};
            if (!layoutStyles.subTitle.color) layoutStyles.subTitle.fontRefColor = masterTxStyles.phFontRef.body;
        }
        stepStart = perfNow();
        var parsed = parseSlideXml(xmlStr, slideW, slideH, images, slideRels.all, hasBgImage, bgImageRid, false, layoutStyles, chartDataMap, diagramDataMap, "slide");
        pushPerfStep(slidePerf.steps, "parseSlideXml(slide)", stepStart);
        console.log("[PPTX] Slide " + sf.num + " own elements: " + parsed.elements.length);
    parsed.elements.forEach(function(el) { if (el.type === "image" && el.rId) el._imgSrc = "slide"; });

        // Parse layout shapes (non-placeholder decorations only)
        if (slideRels.layout) {
            var layoutPath2 = ("ppt/slides/" + slideRels.layout).replace(/[^/]+\/\.\.\//g, "");
            var layoutBase2 = layoutPath2.replace(/[^/]*$/, "");
            var layoutFile = zip.file(layoutPath2);
            if (layoutFile) {
                stepStart = perfNow();
                console.log("[PPTX] Parsing layout shapes for slide " + sf.num + ": " + layoutPath2);
                var layoutRels2 = await getRelsCached(layoutPath2.replace(/([^/]+)$/, "_rels/$1.rels"));
                var layoutParsed = parseSlideXml(
                    await layoutFile.async("string"), slideW, slideH, {}, layoutRels2.all, hasBgImage, null, true, {}, {}, {}, "layout"
                );
                layoutParsed.elements.forEach(function(el) { if (el.type === "image" && el.rId) el._imgSrc = "layout"; });
                console.log("[PPTX]   layout contributed " + layoutParsed.elements.length + " elements");
                parsed.elements = layoutParsed.elements.concat(parsed.elements);
                pushPerfStep(slidePerf.steps, "parseSlideXml(layout)", stepStart);
                            _ctxLayoutBase = layoutBase2;
                            _ctxLayoutImageRels = layoutRels2.images;
            }

            // Parse master shapes (non-placeholder decorations only)
            var layoutRelsForMasterShapes = await getRelsCached(layoutPath2.replace(/([^/]+)$/, "_rels/$1.rels"));
            if (layoutRelsForMasterShapes.master) {
                var masterPath2 = (layoutBase2 + layoutRelsForMasterShapes.master).replace(/[^/]+\/\.\.\//g, "");
                var masterBase2 = masterPath2.replace(/[^/]*$/, "");
                var masterFile = zip.file(masterPath2);
                if (masterFile) {
                    stepStart = perfNow();
                    console.log("[PPTX] Parsing master shapes for slide " + sf.num + ": " + masterPath2);
                    var masterRels2 = await getRelsCached(masterPath2.replace(/([^/]+)$/, "_rels/$1.rels"));
                    var masterParsed = parseSlideXml(
                        await masterFile.async("string"), slideW, slideH, {}, masterRels2.all, hasBgImage, null, true, {}, {}, {}, "master"
                    );
                    masterParsed.elements.forEach(function(el) { if (el.type === "image" && el.rId) el._imgSrc = "master"; });
                    console.log("[PPTX]   master contributed " + masterParsed.elements.length + " elements");
                    parsed.elements = masterParsed.elements.concat(parsed.elements);
                    pushPerfStep(slidePerf.steps, "parseSlideXml(master)", stepStart);
                                    _ctxMasterBase = masterBase2;
                                    _ctxMasterImageRels = masterRels2.images;
                }
            }
        }

        // Build slide object
        var bgTint = null;
        if (hasBgImage) {
            if (bgResult && typeof bgResult === "object" && bgResult.bgTint) {
                bgTint = bgResult.bgTint;
            } else {
                bgTint = extractBlipEffects(xmlStr);
            }
        }
        if (bgTint) console.log("[PPTX] Slide " + sf.num + " bgTint: " + JSON.stringify(bgTint));
        var slide = {
            bg: parsed.bgColor, bgImage: null, bgTint: bgTint,
            elements: parsed.elements, notes: "Slide " + sf.num
        };
        if (bgResult) {
            if (typeof bgResult === "string") slide.bgImage = bgResult;
            else if (bgResult.image) slide.bgImage = bgResult.image;
            else if (bgResult.solidColor) slide.bg = bgResult.solidColor;
        }

        console.log("[PPTX] Slide " + sf.num + " DONE: " + slide.elements.length + " elements, bg=" + slide.bg + ", bgImage=" + (slide.bgImage ? "yes" : "no"));
        var elSummary = {shape:0, text:0, image:0};
        slide.elements.forEach(function(el) { elSummary[el.type] = (elSummary[el.type]||0) + 1; });
        console.log("[PPTX]   breakdown: shapes=" + elSummary.shape + " texts=" + elSummary.text + " images=" + elSummary.image);

        newSlides.push(slide);
                // Store context needed for deferred image loading in phase 2.
                slide._ctx = {
                    imageRels: slideRels.images,
                    imageBasePath: "ppt/slides/",
                    layoutBase: _ctxLayoutBase,
                    layoutImageRels: _ctxLayoutImageRels,
                    masterBase: _ctxMasterBase,
                    masterImageRels: _ctxMasterImageRels
                };
        slidePerf.totalMs = perfNow() - slideStart;
        perfSummary.slides.push(slidePerf);
    }
    // Phase 1 complete — fire callback so UI can display structure without images.
    if (typeof onStructureReady === "function") {
        try { onStructureReady(newSlides); } catch(e) {}
    }

    perfSummary.totalMs = perfNow() - t0;
    if (typeof window !== "undefined") {
        window.__PPTX_PERF__ = window.__PPTX_PERF__ || {};
        window.__PPTX_PERF__.lastParsePhase2ImagesMs = null;
        window.__PPTX_PERF__.lastParse = perfSummary;
    }
    console.info("[PERF] parsePptx total=" + perfSummary.totalMs.toFixed(1) + "ms slides=" + perfSummary.slides.length);
    perfSummary.steps.forEach(function (step) {
        console.info("[PERF] parse step " + step.name + "=" + step.ms.toFixed(1) + "ms");
    });
    perfSummary.slides.forEach(function (sp) {
        console.info("[PERF] slide " + sp.slide + " total=" + sp.totalMs.toFixed(1) + "ms");
        sp.steps.forEach(function (step) {
            console.info("[PERF] slide " + sp.slide + " " + step.name + "=" + step.ms.toFixed(1) + "ms");
        });
    });
    console.log("[PPTX] === Parse phase1 complete: " + newSlides.length + " slides in " + perfSummary.totalMs.toFixed(0) + "ms ===");

    // Phase 2 — load all slide images concurrently across every slide.
    // Start this in background so parsePptx can return immediately after phase 1.
    var phase2Start = perfNow();
    Promise.all(newSlides.map(function(slide, idx) {
        var ctx = slide._ctx;
        if (!ctx) return Promise.resolve();
        delete slide._ctx;
        return Promise.all([
            buildImageMap(zip, ctx.imageBasePath, ctx.imageRels),
            ctx.layoutBase ? buildImageMap(zip, ctx.layoutBase, ctx.layoutImageRels) : Promise.resolve({}),
            ctx.masterBase ? buildImageMap(zip, ctx.masterBase, ctx.masterImageRels) : Promise.resolve({})
        ]).then(function(maps) {
            var slideImgs = maps[0], layoutImgs = maps[1], masterImgs = maps[2];
            slide.elements.forEach(function(el) {
                if (el.type !== "image" || !el.rId) return;
                var map = el._imgSrc === "layout" ? layoutImgs
                        : el._imgSrc === "master"  ? masterImgs
                        : slideImgs;
                el.dataUrl = map[el.rId] || null;
                delete el._imgSrc; delete el.rId;
            });
            if (typeof onSlideImagesReady === "function") {
                try { onSlideImagesReady(idx); } catch(e) {}
            }
        });
    })).then(function () {
        var phase2Ms = perfNow() - phase2Start;
        if (typeof window !== "undefined") {
            window.__PPTX_PERF__ = window.__PPTX_PERF__ || {};
            window.__PPTX_PERF__.lastParsePhase2ImagesMs = phase2Ms;
        }
        console.info("[PERF] phase2 images ms=" + phase2Ms.toFixed(1));
        if (typeof onAllImagesReady === "function") {
            try { onAllImagesReady(); } catch(e) {}
        }
    }).catch(function (e) {
        console.warn("[PPTX] phase2 image loading failed", e);
    });

    return newSlides;
}

export { parsePptx };
