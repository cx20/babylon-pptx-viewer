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
    var autoNumCounters = {}; // Track per-level auto-numbering
    for (var p = 0; p < paras.length; p++) {
        var para = paras[p];
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

        function normalizeRunText(s) {
            if (!s) return "";
            // U+000B (vertical tab) appears in some PPTX text runs and should render as spacing.
            // Strip other control chars that degrade GUI text rendering.
            return s
                .replace(/\u000B/g, " ")
                .replace(/[\u0000-\u0008\u000C\u000E-\u001F\u007F]/g, "");
        }

        var txt = "";
        var lineBreakPositions = []; // text offsets where <a:br/> elements occurred
        for (var ci = 0; ci < para.childNodes.length; ci++) {
            var child = para.childNodes[ci];
            if (child.nodeType !== 1) continue;
            if (child.localName === "br" && child.namespaceURI === A_NS) {
                // Explicit line break (Shift+Enter in PowerPoint) — record current offset
                lineBreakPositions.push(txt.length);
            } else if (child.localName === "r" && child.namespaceURI === A_NS) {
                var rPr = child.getElementsByTagNameNS(A_NS, "rPr")[0];
                if (rPr) {
                    var rsz = rPr.getAttribute("sz");
                    if (rsz) { var rfs = emuToFontPx(parseInt(rsz)); if (rfs > fs) fs = rfs; }
                    if (rPr.getAttribute("b") === "1") fw = "bold";
                    if (rPr.getAttribute("i") === "1") fi = true;
                    if (rPr.getAttribute("cap") && !cap) cap = rPr.getAttribute("cap");
                    var rsf = rPr.getElementsByTagNameNS(A_NS, "solidFill")[0];
                    if (rsf) { var rc = resolveColor(rsf); if (rc) fc = rc; }
                }
                var t = child.getElementsByTagNameNS(A_NS, "t")[0];
                if (t) txt += normalizeRunText(t.textContent);
            }
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

        // Bullet / numbering (only apply if paragraph has text)
        var hasBullet = false;
        if (pPr && txt.trim()) {
            var buNone = pPr.getElementsByTagNameNS(A_NS, "buNone")[0];
            if (!buNone) {
                var bc = pPr.getElementsByTagNameNS(A_NS, "buChar")[0];
                var ban = pPr.getElementsByTagNameNS(A_NS, "buAutoNum")[0];
                if (bc) {
                    txt = (bc.getAttribute("char") || "\u2022") + " " + txt;
                    hasBullet = true;
                } else if (ban) {
                    // Auto-numbering: increment counter for this level, reset deeper levels
                    autoNumCounters[level] = (autoNumCounters[level] || 0) + 1;
                    for (var lk in autoNumCounters) {
                        if (parseInt(lk, 10) > level) delete autoNumCounters[lk];
                    }
                    var bnType = ban.getAttribute("type") || "arabicPeriod";
                    var startAt = parseInt(ban.getAttribute("startAt"), 10) || 1;
                    var num = autoNumCounters[level] + startAt - 1;
                    var prefix;
                    if (bnType === "romanLcPeriod") {
                        var roms = ["i","ii","iii","iv","v","vi","vii","viii","ix","x"];
                        prefix = (roms[num - 1] || num) + ". ";
                    } else if (bnType === "alphaLcPeriod") {
                        prefix = String.fromCharCode(96 + num) + ". ";
                    } else {
                        prefix = num + ". ";
                    }
                    txt = prefix + txt;
                    hasBullet = true;
                }
            }
        }

        var lnSpc = 1.0;
        if (pPr) {
            var ls = pPr.getElementsByTagNameNS(A_NS, "lnSpc")[0];
            if (ls) {
                var spcPct = ls.getElementsByTagNameNS(A_NS, "spcPct")[0];
                var spcPts = ls.getElementsByTagNameNS(A_NS, "spcPts")[0];
                if (spcPts) {
                    // Absolute point-based line spacing
                    var ptH = parseInt(spcPts.getAttribute("val"), 10) / 100 * 0.75; // hundredths-pt → px
                    lnSpc = Math.max(0.8, ptH / Math.max(1, fs * 0.75)); // ratio relative to fontSize-px
                } else if (spcPct) {
                    var v = parseInt(spcPct.getAttribute("val"), 10) || 100000;
                    lnSpc = Math.max(0.8, Math.min(2.0, v / 100000));
                }
            }
        }

        // Paragraph spacing (spcBef / spcAft) in pixels
        var spaceBefore = 0, spaceAfter = 0;
        if (pPr) {
            var sbb = pPr.getElementsByTagNameNS(A_NS, "spcBef")[0];
            if (sbb) {
                var sbbPts = sbb.getElementsByTagNameNS(A_NS, "spcPts")[0];
                var sbbPct = sbb.getElementsByTagNameNS(A_NS, "spcPct")[0];
                if (sbbPts) spaceBefore = parseInt(sbbPts.getAttribute("val"), 10) / 100 * 0.75;
                else if (sbbPct) spaceBefore = (parseInt(sbbPct.getAttribute("val"), 10) / 100000) * fs * 0.75;
            }
            var sba = pPr.getElementsByTagNameNS(A_NS, "spcAft")[0];
            if (sba) {
                var sbaPts = sba.getElementsByTagNameNS(A_NS, "spcPts")[0];
                var sbaPct = sba.getElementsByTagNameNS(A_NS, "spcPct")[0];
                if (sbaPts) spaceAfter = parseInt(sbaPts.getAttribute("val"), 10) / 100 * 0.75;
                else if (sbaPct) spaceAfter = (parseInt(sbaPct.getAttribute("val"), 10) / 100000) * fs * 0.75;
            }
        }

        txt = normalizeRunText(txt);
        if (cap === "all" && txt) txt = txt.toUpperCase();

        // Split paragraph at <a:br/> positions into separate line entries
        if (lineBreakPositions.length === 0) {
            result.push({
                text: txt, fontSize: Math.min(Math.max(fs, 6), 60),
                fontWeight: fw, color: fc, italic: fi, align: align,
                isEmpty: txt.trim().length === 0, lineSpacing: lnSpc, level: level,
                spaceBefore: spaceBefore, spaceAfter: spaceAfter
            });
        } else {
            var brParts = [];
            var brPrev = 0;
            lineBreakPositions.forEach(function(pos) { brParts.push(txt.slice(brPrev, pos)); brPrev = pos; });
            brParts.push(txt.slice(brPrev));
            brParts.forEach(function(partTxt, si) {
                result.push({
                    text: partTxt, fontSize: Math.min(Math.max(fs, 6), 60),
                    fontWeight: fw, color: fc, italic: fi, align: align,
                    isEmpty: partTxt.trim().length === 0, lineSpacing: lnSpc, level: level,
                    spaceBefore: si === 0 ? spaceBefore : 0,
                    spaceAfter: si === brParts.length - 1 ? spaceAfter : 0
                });
            });
        }
    }
    return result;
}
