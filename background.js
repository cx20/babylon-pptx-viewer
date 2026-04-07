// ============================================================================
// background.js - Background image/color extraction and blip effect detection
// ============================================================================

import { A_NS, P_NS, R_NS } from "./constants.js";
import { resolveColor, themeColors, hexToRgb, applyColorModifiers } from "./color-utils.js";
import { loadImageAsDataUrl, parseRelsFile } from "./zip-helpers.js";

function resolveClrNodeWithPh(clrNode, phClr) {
    if (!clrNode) return null;
    if (clrNode.localName === "srgbClr") {
        return applyColorModifiers("#" + (clrNode.getAttribute("val") || "000000"), clrNode);
    }
    if (clrNode.localName === "schemeClr") {
        var key = clrNode.getAttribute("val") || "";
        var base = key === "phClr" ? (phClr || themeColors.bg1 || "#FFFFFF") : (themeColors[key] || "#333333");
        return applyColorModifiers(base, clrNode);
    }
    if (clrNode.localName === "prstClr") {
        var pv = (clrNode.getAttribute("val") || "").toLowerCase();
        var base2 = pv === "white" ? "#FFFFFF" : pv === "black" ? "#000000" : "#808080";
        return applyColorModifiers(base2, clrNode);
    }
    return null;
}

function resolveFillNodeColor(fillNode, phClr) {
    if (!fillNode) return null;
    if (fillNode.localName === "solidFill") {
        for (var i = 0; i < fillNode.childNodes.length; i++) {
            var cn = fillNode.childNodes[i];
            if (cn.nodeType !== 1) continue;
            var c = resolveClrNodeWithPh(cn, phClr);
            if (c) return c;
        }
    }
    if (fillNode.localName === "gradFill") {
        var gs = fillNode.getElementsByTagNameNS(A_NS, "gs");
        if (gs.length > 0) {
            for (var gi = 0; gi < gs[0].childNodes.length; gi++) {
                var gcn = gs[0].childNodes[gi];
                if (gcn.nodeType !== 1) continue;
                var gc = resolveClrNodeWithPh(gcn, phClr);
                if (gc) return gc;
            }
        }
    }
    return null;
}

async function resolveBgRefFromTheme(zip, bgRef) {
    var idx = parseInt(bgRef.getAttribute("idx") || "1001", 10);
    var phClr = resolveColor(bgRef) || themeColors.bg1 || "#FFFFFF";
    var themeFile = zip.file("ppt/theme/theme1.xml");
    if (!themeFile) return null;

    var tdoc = new DOMParser().parseFromString(await themeFile.async("string"), "application/xml");
    var fmtScheme = tdoc.getElementsByTagNameNS(A_NS, "fmtScheme")[0];
    if (!fmtScheme) return null;
    var bgFillStyleLst = fmtScheme.getElementsByTagNameNS(A_NS, "bgFillStyleLst")[0];
    if (!bgFillStyleLst) return null;

    var styles = [];
    for (var i = 0; i < bgFillStyleLst.childNodes.length; i++) {
        var n = bgFillStyleLst.childNodes[i];
        if (n.nodeType === 1) styles.push(n);
    }
    if (styles.length === 0) return null;

    // OOXML style matrix mapping: background refs commonly start at 1000.
    // 1000 -> first bg style, 1001 -> second, 1002 -> third.
    var styleIdx = idx >= 1000 ? (idx - 1000) : (idx - 1);
    if (styleIdx < 0) styleIdx = 0;
    if (styleIdx >= styles.length) styleIdx = styles.length - 1;
    var styleNode = styles[styleIdx];
    if (!styleNode) return null;

    // Background image style
    if (styleNode.localName === "blipFill") {
        var blip = styleNode.getElementsByTagNameNS(A_NS, "blip")[0];
        if (blip) {
            var rId = blip.getAttribute("r:embed") || blip.getAttributeNS(R_NS, "embed");
            if (rId) {
                var themeRels = await parseRelsFile(zip, "ppt/theme/_rels/theme1.xml.rels");
                var target = themeRels.all[rId];
                if (target) {
                    var img = await loadImageAsDataUrl(zip, "ppt/theme/", target);
                    if (img) return img;
                }
            }
        }
    }

    // Solid/gradient style fallback
    var styleColor = resolveFillNodeColor(styleNode, phClr);
    if (styleColor) return { solidColor: styleColor };
    return null;
}

// Extract background from a slide/layout/master XML
// Returns: dataUrl string (bg image), {solidColor:"#..."}, or null
export async function extractBackground(xmlStr, zip, basePath, relsAll, slideW, slideH) {
    var doc = new DOMParser().parseFromString(xmlStr, "application/xml");
    var cSld = doc.getElementsByTagNameNS(P_NS, "cSld")[0];
    if (!cSld) { console.log("[BG] No cSld found, base=" + basePath); return null; }

    var bg = cSld.getElementsByTagNameNS(P_NS, "bg")[0];
    console.log("[BG] base=" + basePath + " hasBgElement=" + !!bg);
    if (bg) {
        var bgPr = bg.getElementsByTagNameNS(P_NS, "bgPr")[0];
        if (bgPr) {
            var blipFill = bgPr.getElementsByTagNameNS(A_NS, "blipFill")[0];
            console.log("[BG]   bgPr found, hasBlipFill=" + !!blipFill);
            if (blipFill) {
                var blip = blipFill.getElementsByTagNameNS(A_NS, "blip")[0];
                if (blip) {
                    var rId = blip.getAttribute("r:embed") || blip.getAttributeNS(R_NS, "embed");
                    console.log("[BG]   blipFill rId=" + rId + " target=" + (relsAll[rId] || "NOT FOUND"));
                    // Check for alt image in blip extLst
                    var allBlipDesc = blip.getElementsByTagName("*");
                    for (var bi = 0; bi < allBlipDesc.length; bi++) {
                        var altEmbed = allBlipDesc[bi].getAttribute("r:embed") || allBlipDesc[bi].getAttributeNS(R_NS, "embed");
                        if (altEmbed && altEmbed !== rId && relsAll[altEmbed]) {
                            console.log("[BG]   found alt image in blip extLst: " + altEmbed + " → " + relsAll[altEmbed]);
                        }
                    }
                    if (rId && relsAll[rId]) {
                        var img = await loadImageAsDataUrl(zip, basePath, relsAll[rId]);
                        if (img) return img;
                    }
                }
            }
            var sf = bgPr.getElementsByTagNameNS(A_NS, "solidFill")[0];
            if (sf) { var c = resolveColor(sf); if (c) return { solidColor: c }; }
            var gf = bgPr.getElementsByTagNameNS(A_NS, "gradFill")[0];
            if (gf) {
                var gs = gf.getElementsByTagNameNS(A_NS, "gs");
                if (gs.length > 0) { var c = resolveColor(gs[0]); if (c) return { solidColor: c }; }
            }
        }
        var bgRef = bg.getElementsByTagNameNS(P_NS, "bgRef")[0];
        if (!bgRef) bgRef = bg.getElementsByTagNameNS(A_NS, "bgRef")[0];
        console.log("[BG]   hasBgRef=" + !!bgRef);
        if (bgRef) {
            var fromTheme = await resolveBgRefFromTheme(zip, bgRef);
            if (fromTheme) {
                console.log("[BG]   bgRef resolved via theme style: " + (typeof fromTheme === "string" ? "image" : JSON.stringify(fromTheme)));
                return fromTheme;
            }
            var c = resolveColor(bgRef);
            if (c) return { solidColor: c };
        }
    }

    // Check spTree for full-bleed background images
    var spTree = cSld.getElementsByTagNameNS(P_NS, "spTree")[0];
    if (!spTree) { console.log("[BG]   no spTree"); return null; }
    var pics = spTree.getElementsByTagNameNS(P_NS, "pic");
    console.log("[BG]   spTree pics=" + pics.length + " slideW=" + slideW + " slideH=" + slideH);
    for (var i = 0; i < pics.length; i++) {
        var pic = pics[i];
        var xfrm = pic.getElementsByTagNameNS(A_NS, "xfrm")[0]; if (!xfrm) continue;
        var off = xfrm.getElementsByTagNameNS(A_NS, "off")[0];
        var ext = xfrm.getElementsByTagNameNS(A_NS, "ext")[0];
        if (!off || !ext) continue;
        var cx = parseInt(ext.getAttribute("cx")) || 0, cy = parseInt(ext.getAttribute("cy")) || 0;
        if (cx > slideW * 0.7 && cy > slideH * 0.7) {
            var blip = pic.getElementsByTagNameNS(A_NS, "blip")[0];
            if (blip) {
                var rId = blip.getAttribute("r:embed") || blip.getAttributeNS(R_NS, "embed");
                if (rId && relsAll[rId]) {
                    var img = await loadImageAsDataUrl(zip, basePath, relsAll[rId]);
                    if (img) return img;
                }
            }
        }
    }
    return null;
}

// Detect duotone, art effects, tint, etc. on background blip
export function extractBlipEffects(xmlStr) {
    var doc = new DOMParser().parseFromString(xmlStr, "application/xml");
    var cSld = doc.getElementsByTagNameNS(P_NS, "cSld")[0];
    if (!cSld) return null;
    var bg = cSld.getElementsByTagNameNS(P_NS, "bg")[0];
    if (!bg) return null;
    var blip = bg.getElementsByTagNameNS(A_NS, "blip")[0];
    if (!blip) return null;

    // Debug
    var blipKids = [];
    for (var ci = 0; ci < blip.childNodes.length; ci++) {
        if (blip.childNodes[ci].nodeType === 1) blipKids.push(blip.childNodes[ci].localName);
    }
    console.log("[BLIP] blip children: [" + blipKids.join(", ") + "]");

    // Search entire bg subtree for art effect URI
    var allBgEls = bg.getElementsByTagName("*");
    for (var i = 0; i < allBgEls.length; i++) {
        var uri = (allBgEls[i].getAttribute("uri") || "").toUpperCase();
        if (uri.indexOf("BEBA8EAE") !== -1) {
            console.log("[BLIP] Art effect found in bg subtree");
            return { type: "artEffect", color: themeColors.dk2 || "#0E5580" };
        }
    }

    // Duotone
    var duotone = null;
    for (var ci = 0; ci < blip.childNodes.length; ci++) {
        if (blip.childNodes[ci].localName === "duotone") { duotone = blip.childNodes[ci]; break; }
    }
    if (duotone) {
        var colors = [], rawVals = [];
        for (var i = 0; i < duotone.childNodes.length; i++) {
            var cn = duotone.childNodes[i];
            if (cn.nodeType !== 1) continue;
            if (cn.localName === "srgbClr") {
                colors.push(applyColorModifiers("#" + cn.getAttribute("val"), cn));
                rawVals.push("srgb:" + cn.getAttribute("val"));
            } else if (cn.localName === "schemeClr") {
                var val = cn.getAttribute("val") || "";
                colors.push(applyColorModifiers(themeColors[val] || "#000000", cn));
                rawVals.push("scheme:" + val);
            } else if (cn.localName === "prstClr") {
                var pv = cn.getAttribute("val") || "black";
                colors.push(applyColorModifiers(pv === "black" ? "#000000" : pv === "white" ? "#FFFFFF" : "#808080", cn));
                rawVals.push("prst:" + pv);
            }
        }
        console.log("[BLIP] duotone: " + rawVals.join(", ") + " → " + colors.join(", "));
        if (colors.length >= 2) {
            var c1 = hexToRgb(colors[0]), c2 = hexToRgb(colors[1]);
            var gray1 = Math.abs(c1.r - c1.g) < 15 && Math.abs(c1.g - c1.b) < 15;
            var gray2 = Math.abs(c2.r - c2.g) < 15 && Math.abs(c2.g - c2.b) < 15;
            if (gray1 && gray2) {
                console.log("[BLIP] Grayscale duotone detected, applying dk2 tint");
                return { type: "artEffect", color: themeColors.dk2 || "#0E5580" };
            } else {
                return { type: "duotone", dark: colors[0], light: colors[1] };
            }
        }
    }

    var clrChange = blip.getElementsByTagNameNS(A_NS, "clrChange")[0];
    if (clrChange) return { type: "tint", color: themeColors.dk2 || "#0E5580" };
    var alphaModFix = blip.getElementsByTagNameNS(A_NS, "alphaModFix")[0];
    if (alphaModFix) return { type: "alpha", amt: parseInt(alphaModFix.getAttribute("amt") || "100000") / 100000 };

    return null;
}
