// ============================================================================
// pptx-parser.js - Main PPTX orchestrator (coordinates all parsing)
// ============================================================================

import { SLIDE_EMU_W, SLIDE_EMU_H } from "./constants.js";
import { parseThemeXml } from "./color-utils.js";
import { parseRelsFile, buildImageMap, loadImageAsDataUrl } from "./zip-helpers.js";
import { extractBackground, extractBlipEffects } from "./background.js";
import { extractPlaceholderStyles, extractMasterTxStyles } from "./style-inheritance.js";
import { parseSlideXml } from "./slide-parser.js";

var C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";

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
                maxValue: maxValue
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

async function parsePptx(arrayBuffer) {
    var t0 = performance.now();
    console.log("[PPTX] === Starting PPTX parse ===");
    var zip = new JSZip();
    await zip.loadAsync(arrayBuffer);
    console.log("[PPTX] ZIP loaded, files: " + Object.keys(zip.files).length);

    // Parse theme colors
    await parseThemeXml(zip);

    // Slide dimensions
    var slideW = SLIDE_EMU_W, slideH = SLIDE_EMU_H;
    var pf = zip.file("ppt/presentation.xml");
    if (pf) {
        var pdoc = new DOMParser().parseFromString(await pf.async("string"), "application/xml");
        var ss = pdoc.getElementsByTagName("p:sldSz")[0];
        if (ss) {
            slideW = parseInt(ss.getAttribute("cx")) || SLIDE_EMU_W;
            slideH = parseInt(ss.getAttribute("cy")) || SLIDE_EMU_H;
        }
    }
    console.log("[PPTX] Slide dimensions: " + slideW + " x " + slideH + " EMU");

    // Enumerate slides
    var slideFiles = [];
    zip.forEach(function (path) {
        var m = path.match(/^ppt\/slides\/slide(\d+)\.xml$/);
        if (m) slideFiles.push({ path: path, num: parseInt(m[1]) });
    });
    slideFiles.sort(function (a, b) { return a.num - b.num; });
    console.log("[PPTX] Found " + slideFiles.length + " slides: " + slideFiles.map(function(s){return "slide"+s.num;}).join(", "));

    // Background cache for layout/master
    var bgCache = {};

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
        var xmlStr = await zip.file(sf.path).async("string");
        var slideRels = await parseRelsFile(zip, "ppt/slides/_rels/slide" + sf.num + ".xml.rels");
        console.log("[PPTX] Slide " + sf.num + " rels: images=" + Object.keys(slideRels.images).length + " layout=" + (slideRels.layout||"none"));
        console.log("[PPTX]   all rels: " + JSON.stringify(slideRels.all));

        // Build image map
        var images = await buildImageMap(zip, "ppt/slides/", slideRels.images);
        console.log("[PPTX] Slide " + sf.num + " images loaded: " + Object.keys(images).filter(function(k){return !!images[k];}).length);

        // Build chart data map from related chart parts
        var chartDataMap = await buildChartDataMap(zip, "ppt/slides/", slideRels.all);
        var diagramDataMap = await buildDiagramDataMap(zip, "ppt/slides/", slideRels.all);

        // === Background inheritance: slide → layout → master ===
        var bgResult = await extractBackground(xmlStr, zip, "ppt/slides/", slideRels.all, slideW, slideH);
        console.log("[PPTX] Slide " + sf.num + " bg chain: slide=" + (bgResult ? (typeof bgResult === "string" ? "image" : JSON.stringify(bgResult)) : "none"));

        if (!bgResult && slideRels.layout) {
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
        }

        var hasBgImage = bgResult && typeof bgResult === "string";
        console.log("[PPTX] Slide " + sf.num + " hasBgImage=" + hasBgImage);

        // Parse slide shapes
        console.log("[PPTX] Parsing slide " + sf.num + " shapes...");
        // Extract layout placeholder styles for inheritance
        var layoutStyles = {};
        var masterTxStyles = { titleColor: null, bodyColor: null, otherColor: null };
        if (slideRels.layout) {
            var layoutPathStyles = ("ppt/slides/" + slideRels.layout).replace(/[^/]+\/\.\.\//g, "");
            layoutStyles = await extractPlaceholderStyles(zip, layoutPathStyles);
            // Also read master txStyles
            var layoutRelsForMaster = await parseRelsFile(zip, layoutPathStyles.replace(/([^/]+)$/, "_rels/$1.rels"));
            if (layoutRelsForMaster.master) {
                var masterPathTx = (layoutPathStyles.replace(/[^/]*$/, "") + layoutRelsForMaster.master).replace(/[^/]+\/\.\.\//g, "");
                masterTxStyles = await extractMasterTxStyles(zip, masterPathTx);
            }
        }
        // Apply master txStyles as fallback fontRefColor for placeholders
        // txStyles only for slides WITHOUT background images
        if (!hasBgImage) {
            if (masterTxStyles.bodyColor) {
                if (!layoutStyles.subTitle) layoutStyles.subTitle = {};
                if (!layoutStyles.subTitle.fontRefColor) layoutStyles.subTitle.fontRefColor = masterTxStyles.bodyColor;
                if (!layoutStyles.body) layoutStyles.body = {};
                if (!layoutStyles.body.fontRefColor) layoutStyles.body.fontRefColor = masterTxStyles.bodyColor;
            }
            if (masterTxStyles.titleColor) {
                if (!layoutStyles.title) layoutStyles.title = {};
                if (!layoutStyles.title.fontRefColor) layoutStyles.title.fontRefColor = masterTxStyles.titleColor;
                if (!layoutStyles.ctrTitle) layoutStyles.ctrTitle = {};
                if (!layoutStyles.ctrTitle.fontRefColor) layoutStyles.ctrTitle.fontRefColor = masterTxStyles.titleColor;
            }
        }
        // Master placeholder fontRef colors apply to all slides, including bgImage slides.
        // This keeps template-defined title/body colors instead of forcing white text.
        for (var phKey in masterTxStyles.phFontRef) {
            if (!layoutStyles[phKey]) layoutStyles[phKey] = {};
            if (!layoutStyles[phKey].fontRefColor) layoutStyles[phKey].fontRefColor = masterTxStyles.phFontRef[phKey];
        }
        // Map master 'body' fontRef to 'subTitle' if subTitle doesn't have its own
        if (masterTxStyles.phFontRef.body && (!layoutStyles.subTitle || !layoutStyles.subTitle.fontRefColor)) {
            if (!layoutStyles.subTitle) layoutStyles.subTitle = {};
            layoutStyles.subTitle.fontRefColor = masterTxStyles.phFontRef.body;
        }
        var parsed = parseSlideXml(xmlStr, slideW, slideH, images, slideRels.all, hasBgImage, false, layoutStyles, chartDataMap, diagramDataMap);
        console.log("[PPTX] Slide " + sf.num + " own elements: " + parsed.elements.length);

        // Parse layout shapes (non-placeholder decorations only)
        if (slideRels.layout) {
            var layoutPath2 = ("ppt/slides/" + slideRels.layout).replace(/[^/]+\/\.\.\//g, "");
            var layoutBase2 = layoutPath2.replace(/[^/]*$/, "");
            var layoutFile = zip.file(layoutPath2);
            if (layoutFile) {
                console.log("[PPTX] Parsing layout shapes for slide " + sf.num + ": " + layoutPath2);
                var layoutRels2 = await parseRelsFile(zip, layoutPath2.replace(/([^/]+)$/, "_rels/$1.rels"));
                var layoutImgs = await buildImageMap(zip, layoutBase2, layoutRels2.images);
                var layoutParsed = parseSlideXml(
                    await layoutFile.async("string"), slideW, slideH, layoutImgs, layoutRels2.all, hasBgImage, true, {}, {}, {}
                );
                console.log("[PPTX]   layout contributed " + layoutParsed.elements.length + " elements");
                parsed.elements = layoutParsed.elements.concat(parsed.elements);
            }
        }

        // Build slide object
        var bgTint = hasBgImage ? extractBlipEffects(xmlStr) : null;
        if (bgTint) console.log("[PPTX] Slide " + sf.num + " bgTint: " + JSON.stringify(bgTint));
        var slide = {
            bg: parsed.bgColor, bgImage: null, bgTint: bgTint,
            elements: parsed.elements, notes: "Slide " + sf.num
        };
        if (bgResult) {
            if (typeof bgResult === "string") slide.bgImage = bgResult;
            else if (bgResult.solidColor) slide.bg = bgResult.solidColor;
        }

        console.log("[PPTX] Slide " + sf.num + " DONE: " + slide.elements.length + " elements, bg=" + slide.bg + ", bgImage=" + (slide.bgImage ? "yes" : "no"));
        var elSummary = {shape:0, text:0, image:0};
        slide.elements.forEach(function(el) { elSummary[el.type] = (elSummary[el.type]||0) + 1; });
        console.log("[PPTX]   breakdown: shapes=" + elSummary.shape + " texts=" + elSummary.text + " images=" + elSummary.image);

        newSlides.push(slide);
    }
    console.log("[PPTX] === Parse complete: " + newSlides.length + " slides in " + (performance.now()-t0).toFixed(0) + "ms ===");
    return newSlides;
}

export { parsePptx };
