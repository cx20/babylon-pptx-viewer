// ============================================================================
// pptx-parser.js - Main PPTX orchestrator (coordinates all parsing)
// ============================================================================

import { SLIDE_EMU_W, SLIDE_EMU_H } from "./constants.js";
import { parseThemeXml } from "./color-utils.js";
import { parseRelsFile, buildImageMap } from "./zip-helpers.js";
import { extractBackground, extractBlipEffects } from "./background.js";
import { extractPlaceholderStyles, extractMasterTxStyles } from "./style-inheritance.js";
import { parseSlideXml } from "./slide-parser.js";

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
        var parsed = parseSlideXml(xmlStr, slideW, slideH, images, slideRels.all, hasBgImage, false, layoutStyles);
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
                    await layoutFile.async("string"), slideW, slideH, layoutImgs, layoutRels2.all, hasBgImage, true
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
