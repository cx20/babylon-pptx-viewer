// ============================================================================
// color-utils.js - Color resolution, modifiers, and theme management
// ============================================================================

import { A_NS } from "./constants.js";

// Theme color map - populated by parseThemeXml()
export var themeColors = {
    dk1: "#000000", dk2: "#44546A", lt1: "#FFFFFF", lt2: "#E7E6E6",
    accent1: "#4472C4", accent2: "#ED7D31", accent3: "#A5A5A5",
    accent4: "#FFC000", accent5: "#5B9BD5", accent6: "#70AD47",
    hlink: "#0563C1", folHlink: "#954F72",
    tx1: "#000000", tx2: "#44546A", bg1: "#FFFFFF", bg2: "#E7E6E6"
};

export function hexToRgb(hex) {
    hex = hex.replace("#", "");
    return { r: parseInt(hex.substr(0, 2), 16), g: parseInt(hex.substr(2, 2), 16), b: parseInt(hex.substr(4, 2), 16) };
}

export function rgbToHex(r, g, b) {
    return "#" + [r, g, b].map(function (c) {
        return Math.max(0, Math.min(255, Math.round(c))).toString(16).padStart(2, "0");
    }).join("");
}

export function applyColorModifiers(hex, node) {
    if (!node || !hex) return hex;
    var rgb = hexToRgb(hex);
    for (var i = 0; i < node.childNodes.length; i++) {
        var cn = node.childNodes[i];
        if (cn.nodeType !== 1) continue;
        var v = parseInt(cn.getAttribute("val") || "100000");
        var pct = v / 100000;
        if (cn.localName === "shade") {
            rgb.r = Math.round(rgb.r * pct);
            rgb.g = Math.round(rgb.g * pct);
            rgb.b = Math.round(rgb.b * pct);
        } else if (cn.localName === "tint") {
            rgb.r = Math.round(rgb.r + (255 - rgb.r) * (1 - pct));
            rgb.g = Math.round(rgb.g + (255 - rgb.g) * (1 - pct));
            rgb.b = Math.round(rgb.b + (255 - rgb.b) * (1 - pct));
        } else if (cn.localName === "lumMod") {
            rgb.r = Math.round(rgb.r * pct);
            rgb.g = Math.round(rgb.g * pct);
            rgb.b = Math.round(rgb.b * pct);
        } else if (cn.localName === "lumOff") {
            var off = 255 * pct;
            rgb.r = Math.round(rgb.r + off);
            rgb.g = Math.round(rgb.g + off);
            rgb.b = Math.round(rgb.b + off);
        }
    }
    return rgbToHex(rgb.r, rgb.g, rgb.b);
}

// Resolve color from an XML node containing srgbClr or schemeClr children
export function resolveColor(node) {
    if (!node) return null;
    var srgb = node.getElementsByTagNameNS(A_NS, "srgbClr")[0];
    if (srgb) return applyColorModifiers("#" + srgb.getAttribute("val"), srgb);
    var scheme = node.getElementsByTagNameNS(A_NS, "schemeClr")[0];
    if (scheme) {
        var val = scheme.getAttribute("val") || "";
        var base = themeColors[val] || "#333333";
        return applyColorModifiers(base, scheme);
    }
    return null;
}

// Parse theme1.xml and populate themeColors
export async function parseThemeXml(zip) {
    var tf = zip.file("ppt/theme/theme1.xml");
    if (!tf) return;
    var xml = await tf.async("string");
    var doc = new DOMParser().parseFromString(xml, "application/xml");
    var cs = doc.getElementsByTagNameNS(A_NS, "clrScheme")[0];
    if (!cs) return;
    function extractColor(tagName) {
        var el = cs.getElementsByTagNameNS(A_NS, tagName)[0];
        if (!el) return null;
        var s = el.getElementsByTagNameNS(A_NS, "srgbClr")[0];
        if (s) return "#" + s.getAttribute("val");
        var sys = el.getElementsByTagNameNS(A_NS, "sysClr")[0];
        if (sys) return "#" + (sys.getAttribute("lastClr") || sys.getAttribute("val") || "000000");
        return null;
    }
    ["dk1", "dk2", "lt1", "lt2", "accent1", "accent2", "accent3",
        "accent4", "accent5", "accent6", "hlink", "folHlink"].forEach(function (k) {
            var c = extractColor(k); if (c) themeColors[k] = c;
        });
    themeColors.tx1 = themeColors.dk1;
    themeColors.tx2 = themeColors.dk2;
    themeColors.bg1 = themeColors.lt1;
    themeColors.bg2 = themeColors.lt2;
    console.log("[PPTX] Theme colors loaded:", JSON.stringify(themeColors));
}
