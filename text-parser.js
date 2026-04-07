// ============================================================================
// text-parser.js - Text/paragraph parsing and shape geometry helpers
// ============================================================================

import { A_NS } from "./constants.js";
import { resolveColor } from "./color-utils.js";

export function emuToFontPx(hundredthsPt) {
    return Math.round(hundredthsPt / 100 * 0.75);
}

export function parseOutline(spPr) {
    if (!spPr) return null;
    var ln = null;
    for (var i = 0; i < spPr.childNodes.length; i++) {
        if (spPr.childNodes[i].localName === "ln") { ln = spPr.childNodes[i]; break; }
    }
    if (!ln) return null;
    if (ln.getElementsByTagNameNS(A_NS, "noFill")[0]) return null;
    var w = parseInt(ln.getAttribute("w")) || 12700;
    var sf = ln.getElementsByTagNameNS(A_NS, "solidFill")[0];
    return { width: Math.max(1, Math.round(w / 12700)), color: (sf ? resolveColor(sf) : null) || "#000" };
}

export function getPresetGeometry(spPr) {
    if (!spPr) return "rect";
    var pg = spPr.getElementsByTagNameNS(A_NS, "prstGeom")[0];
    return pg ? (pg.getAttribute("prst") || "rect") : "rect";
}

export function getShapeFill(spPr) {
    if (!spPr) return null;
    for (var i = 0; i < spPr.childNodes.length; i++) {
        var n = spPr.childNodes[i];
        if (n.localName === "noFill") return null;
        if (n.localName === "solidFill") return resolveColor(n) || "#CCCCCC";
        if (n.localName === "gradFill") {
            var gs = n.getElementsByTagNameNS(A_NS, "gs");
            if (gs.length > 0) return resolveColor(gs[0]) || "#CCCCCC";
        }
    }
    return null;
}

// Parse paragraphs from a txBody element
export function parseParagraphs(txBody, defaultFS, defaultFC, layoutCap) {
    defaultFS = defaultFS || 14;
    defaultFC = defaultFC || "#333";
    layoutCap = layoutCap || "";
    var result = [], paras = txBody.getElementsByTagNameNS(A_NS, "p");
    for (var p = 0; p < paras.length; p++) {
        var para = paras[p], runs = para.getElementsByTagNameNS(A_NS, "r");
        var fs = defaultFS, fw = "normal", fc = defaultFC, fi = false, cap = layoutCap;

        var pPr = para.getElementsByTagNameNS(A_NS, "pPr")[0];
        if (pPr) {
            var dr = pPr.getElementsByTagNameNS(A_NS, "defRPr")[0];
            if (dr) {
                var sz = dr.getAttribute("sz"); if (sz) fs = emuToFontPx(parseInt(sz));
                if (dr.getAttribute("b") === "1") fw = "bold";
                if (dr.getAttribute("cap")) cap = dr.getAttribute("cap");
                var dsf = dr.getElementsByTagNameNS(A_NS, "solidFill")[0];
                if (dsf) { var dc = resolveColor(dsf); if (dc) fc = dc; }
            }
        }
        var epr = para.getElementsByTagNameNS(A_NS, "endParaRPr")[0];
        if (epr) {
            var esz = epr.getAttribute("sz");
            if (esz && fs === defaultFS) fs = emuToFontPx(parseInt(esz));
            if (epr.getAttribute("b") === "1" && fw === "normal") fw = "bold";
            if (epr.getAttribute("cap") && !cap) cap = epr.getAttribute("cap");
            var ef = epr.getElementsByTagNameNS(A_NS, "solidFill")[0];
            if (ef && fc === defaultFC) { var ec = resolveColor(ef); if (ec) fc = ec; }
        }

        var txt = "";
        for (var r = 0; r < runs.length; r++) {
            var rPr = runs[r].getElementsByTagNameNS(A_NS, "rPr")[0];
            if (rPr) {
                var rsz = rPr.getAttribute("sz");
                if (rsz) { var rfs = emuToFontPx(parseInt(rsz)); if (rfs > fs) fs = rfs; }
                if (rPr.getAttribute("b") === "1") fw = "bold";
                if (rPr.getAttribute("i") === "1") fi = true;
                if (rPr.getAttribute("cap") && !cap) cap = rPr.getAttribute("cap");
                var rsf = rPr.getElementsByTagNameNS(A_NS, "solidFill")[0];
                if (rsf) { var rc = resolveColor(rsf); if (rc) fc = rc; }
            }
            var t = runs[r].getElementsByTagNameNS(A_NS, "t")[0];
            if (t) txt += t.textContent;
        }

        var align = "left";
        var level = 0;
        if (pPr) {
            var al = pPr.getAttribute("algn");
            if (al === "ctr") align = "center"; else if (al === "r") align = "right";
            var lvl = pPr.getAttribute("lvl");
            if (lvl !== null) {
                var parsedLevel = parseInt(lvl, 10);
                if (!isNaN(parsedLevel)) level = Math.max(0, parsedLevel);
            }
        }
        if (pPr) {
            var bc = pPr.getElementsByTagNameNS(A_NS, "buChar")[0];
            if (bc && txt.trim()) txt = (bc.getAttribute("char") || "•") + " " + txt;
        }
        var lnSpc = 1.0;
        if (pPr) {
            var ls = pPr.getElementsByTagNameNS(A_NS, "lnSpc")[0];
            if (ls) {
                var spcPct = ls.getElementsByTagNameNS(A_NS, "spcPct")[0];
                if (spcPct) {
                    var v = parseInt(spcPct.getAttribute("val")) || 100000;
                    lnSpc = Math.max(0.8, Math.min(2.0, v / 100000));
                }
            }
        }

        if (cap === "all" && txt) txt = txt.toUpperCase();

        result.push({
            text: txt, fontSize: Math.min(Math.max(fs, 6), 60),
            fontWeight: fw, color: fc, italic: fi, align: align,
            isEmpty: txt.trim().length === 0, lineSpacing: lnSpc, level: level
        });
    }
    return result;
}
