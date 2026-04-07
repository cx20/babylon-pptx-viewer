// ============================================================================
// style-inheritance.js - Layout/master placeholder style extraction
// ============================================================================

import { A_NS, P_NS } from "./constants.js";
import { resolveColor } from "./color-utils.js";
import { emuToFontPx } from "./text-parser.js";

// Extract text styles from layout placeholder shapes
export async function extractPlaceholderStyles(zip, layoutPath) {
    var styles = {};
    var f = zip.file(layoutPath);
    if (!f) return styles;
    var xml = await f.async("string");
    var doc = new DOMParser().parseFromString(xml, "application/xml");
    var spTree = doc.getElementsByTagNameNS(P_NS, "spTree")[0];
    if (!spTree) return styles;

    var sps = spTree.getElementsByTagNameNS(P_NS, "sp");
    for (var i = 0; i < sps.length; i++) {
        var sp = sps[i];
        var nvSpPr = sp.getElementsByTagNameNS(P_NS, "nvSpPr")[0];
        if (!nvSpPr) continue;
        var nvPr = nvSpPr.getElementsByTagNameNS(P_NS, "nvPr")[0];
        if (!nvPr) continue;
        var ph = nvPr.getElementsByTagNameNS(P_NS, "ph")[0];
        if (!ph) continue;
        var phType = ph.getAttribute("type") || "body";
        var style = {};

        var txBody = sp.getElementsByTagNameNS(P_NS, "txBody")[0];
        if (!txBody) txBody = sp.getElementsByTagNameNS(A_NS, "txBody")[0];
        if (txBody) {
            var bodyPr = txBody.getElementsByTagNameNS(A_NS, "bodyPr")[0];
            if (bodyPr) {
                var anc = bodyPr.getAttribute("anchor");
                if (anc) style.anchor = anc;
            }
            var lstStyle = txBody.getElementsByTagNameNS(A_NS, "lstStyle")[0];
            if (lstStyle) {
                var pPrs = [];
                for (var lvl = 1; lvl <= 9; lvl++) {
                    var pp = lstStyle.getElementsByTagNameNS(A_NS, "lvl" + lvl + "pPr");
                    if (pp.length > 0) pPrs.push(pp[0]);
                }
                var defPPr = lstStyle.getElementsByTagNameNS(A_NS, "defPPr");
                if (defPPr.length > 0) pPrs.push(defPPr[0]);
                for (var j = 0; j < pPrs.length; j++) {
                    var dr = pPrs[j].getElementsByTagNameNS(A_NS, "defRPr")[0];
                    if (dr) {
                        if (dr.getAttribute("cap")) style.cap = dr.getAttribute("cap");
                        var sz = dr.getAttribute("sz");
                        if (sz) style.fontSize = emuToFontPx(parseInt(sz));
                        if (dr.getAttribute("b") === "1") style.bold = true;
                        var sf = dr.getElementsByTagNameNS(A_NS, "solidFill")[0];
                        if (sf) { var c = resolveColor(sf); if (c) style.color = c; }
                    }
                }
            }
            var paras = txBody.getElementsByTagNameNS(A_NS, "p");
            for (var j = 0; j < paras.length; j++) {
                var pPr = paras[j].getElementsByTagNameNS(A_NS, "pPr")[0];
                if (pPr) {
                    var dr = pPr.getElementsByTagNameNS(A_NS, "defRPr")[0];
                    if (dr) {
                        if (dr.getAttribute("cap") && !style.cap) style.cap = dr.getAttribute("cap");
                        var sz = dr.getAttribute("sz");
                        if (sz && !style.fontSize) style.fontSize = emuToFontPx(parseInt(sz));
                        var sf = dr.getElementsByTagNameNS(A_NS, "solidFill")[0];
                        if (sf && !style.color) { var c = resolveColor(sf); if (c) style.color = c; }
                    }
                }
            }
        }

        if (Object.keys(style).length > 0) {
            styles[phType] = style;
            console.log("[LAYOUT] placeholder '" + phType + "' styles: " + JSON.stringify(style));
        }

        var pStyle = sp.getElementsByTagNameNS(P_NS, "style")[0];
        if (pStyle) {
            var fontRef = pStyle.getElementsByTagNameNS(A_NS, "fontRef")[0];
            if (fontRef) {
                var frc = resolveColor(fontRef);
                if (frc) {
                    if (!styles[phType]) styles[phType] = {};
                    styles[phType].fontRefColor = frc;
                    console.log("[LAYOUT] placeholder '" + phType + "' fontRef color: " + frc);
                }
            }
        }
    }
    return styles;
}

// Extract text styles from slide master's p:txStyles and placeholder fontRef
export async function extractMasterTxStyles(zip, masterPath) {
    var result = { titleColor: null, bodyColor: null, otherColor: null, phFontRef: {} };
    var f = zip.file(masterPath);
    if (!f) return result;
    var xml = await f.async("string");
    var doc = new DOMParser().parseFromString(xml, "application/xml");

    var txStyles = doc.getElementsByTagNameNS(P_NS, "txStyles")[0];
    if (txStyles) {
        function extractStyleColor(styleName) {
            var styleEl = txStyles.getElementsByTagNameNS(P_NS, styleName)[0];
            if (!styleEl) return null;
            var lvl1 = styleEl.getElementsByTagNameNS(A_NS, "lvl1pPr")[0];
            if (lvl1) {
                var dr = lvl1.getElementsByTagNameNS(A_NS, "defRPr")[0];
                if (dr) {
                    var sf = dr.getElementsByTagNameNS(A_NS, "solidFill")[0];
                    if (sf) return resolveColor(sf);
                }
            }
            return null;
        }
        result.titleColor = extractStyleColor("titleStyle");
        result.bodyColor = extractStyleColor("bodyStyle");
        result.otherColor = extractStyleColor("otherStyle");
    }

    var cSld = doc.getElementsByTagNameNS(P_NS, "cSld")[0];
    if (cSld) {
        var spTree = cSld.getElementsByTagNameNS(P_NS, "spTree")[0];
        if (spTree) {
            var sps = spTree.getElementsByTagNameNS(P_NS, "sp");
            for (var i = 0; i < sps.length; i++) {
                var sp = sps[i];
                var nvSpPr = sp.getElementsByTagNameNS(P_NS, "nvSpPr")[0];
                if (!nvSpPr) continue;
                var nvPr = nvSpPr.getElementsByTagNameNS(P_NS, "nvPr")[0];
                if (!nvPr) continue;
                var ph = nvPr.getElementsByTagNameNS(P_NS, "ph")[0];
                if (!ph) continue;
                var phType = ph.getAttribute("type") || "body";
                var pStyle = sp.getElementsByTagNameNS(P_NS, "style")[0];
                if (pStyle) {
                    var fontRef = pStyle.getElementsByTagNameNS(A_NS, "fontRef")[0];
                    if (fontRef) {
                        var frc = resolveColor(fontRef);
                        if (frc) {
                            result.phFontRef[phType] = frc;
                            console.log("[MASTER] placeholder '" + phType + "' fontRef color: " + frc);
                        }
                    }
                }
            }
        }
    }

    console.log("[MASTER] txStyles: title=" + result.titleColor + " body=" + result.bodyColor + " other=" + result.otherColor);
    return result;
}
