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

// OOXML preset color names → hex (ECMA-376 §20.1.2.3.33)
var PRESET_COLORS = {
    aliceBlue:"#F0F8FF", antiqueWhite:"#FAEBD7", aqua:"#00FFFF",
    aquamarine:"#7FFFD4", azure:"#F0FFFF", beige:"#F5F5DC",
    bisque:"#FFE4C4", black:"#000000", blanchedAlmond:"#FFEBCD",
    blue:"#0000FF", blueViolet:"#8A2BE2", brown:"#A52A2A",
    burlyWood:"#DEB887", cadetBlue:"#5F9EA0", chartreuse:"#7FFF00",
    chocolate:"#D2691E", coral:"#FF7F50", cornflowerBlue:"#6495ED",
    cornsilk:"#FFF8DC", crimson:"#DC143C", cyan:"#00FFFF",
    darkBlue:"#00008B", darkCyan:"#008B8B", darkGoldenrod:"#B8860B",
    darkGray:"#A9A9A9", darkGreen:"#006400", darkKhaki:"#BDB76B",
    darkMagenta:"#8B008B", darkOliveGreen:"#556B2F", darkOrange:"#FF8C00",
    darkOrchid:"#9932CC", darkRed:"#8B0000", darkSalmon:"#E9967A",
    darkSeaGreen:"#8FBC8F", darkSlateBlue:"#483D8B", darkSlateGray:"#2F4F4F",
    darkTurquoise:"#00CED1", darkViolet:"#9400D3", deepPink:"#FF1493",
    deepSkyBlue:"#00BFFF", dimGray:"#696969", dodgerBlue:"#1E90FF",
    firebrick:"#B22222", floralWhite:"#FFFAF0", forestGreen:"#228B22",
    fuchsia:"#FF00FF", gainsboro:"#DCDCDC", ghostWhite:"#F8F8FF",
    gold:"#FFD700", goldenrod:"#DAA520", gray:"#808080",
    green:"#008000", greenYellow:"#ADFF2F", honeydew:"#F0FFF0",
    hotPink:"#FF69B4", indianRed:"#CD5C5C", indigo:"#4B0082",
    ivory:"#FFFFF0", khaki:"#F0E68C", lavender:"#E6E6FA",
    lavenderBlush:"#FFF0F5", lawnGreen:"#7CFC00", lemonChiffon:"#FFFACD",
    lightBlue:"#ADD8E6", lightCoral:"#F08080", lightCyan:"#E0FFFF",
    lightGoldenrodYellow:"#FAFAD2", lightGray:"#D3D3D3", lightGreen:"#90EE90",
    lightPink:"#FFB6C1", lightSalmon:"#FFA07A", lightSeaGreen:"#20B2AA",
    lightSkyBlue:"#87CEFA", lightSlateGray:"#778899", lightSteelBlue:"#B0C4DE",
    lightYellow:"#FFFFE0", lime:"#00FF00", limeGreen:"#32CD32",
    linen:"#FAF0E6", magenta:"#FF00FF", maroon:"#800000",
    medAquamarine:"#66CDAA", medBlue:"#0000CD", medOrchid:"#BA55D3",
    medPurple:"#9370DB", medSeaGreen:"#3CB371", medSlateBlue:"#7B68EE",
    medSpringGreen:"#00FA9A", medTurquoise:"#48D1CC", medVioletRed:"#C71585",
    midnightBlue:"#191970", mintCream:"#F5FFFA", mistyRose:"#FFE4E1",
    moccasin:"#FFE4B5", navajoWhite:"#FFDEAD", navy:"#000080",
    oldLace:"#FDF5E6", olive:"#808000", oliveDrab:"#6B8E23",
    orange:"#FFA500", orangeRed:"#FF4500", orchid:"#DA70D6",
    paleGoldenrod:"#EEE8AA", paleGreen:"#98FB98", paleTurquoise:"#AFEEEE",
    paleVioletRed:"#DB7093", papayaWhip:"#FFEFD5", peachPuff:"#FFDAB9",
    peru:"#CD853F", pink:"#FFC0CB", plum:"#DDA0DD",
    powderBlue:"#B0E0E6", purple:"#800080", red:"#FF0000",
    rosyBrown:"#BC8F8F", royalBlue:"#4169E1", saddleBrown:"#8B4513",
    salmon:"#FA8072", sandyBrown:"#F4A460", seaGreen:"#2E8B57",
    seaShell:"#FFF5EE", sienna:"#A0522D", silver:"#C0C0C0",
    skyBlue:"#87CEEB", slateBlue:"#6A5ACD", slateGray:"#708090",
    snow:"#FFFAFA", springGreen:"#00FF7F", steelBlue:"#4682B4",
    tan:"#D2B48C", teal:"#008080", thistle:"#D8BFD8",
    tomato:"#FF6347", turquoise:"#40E0D0", violet:"#EE82EE",
    wheat:"#F5DEB3", white:"#FFFFFF", whiteSmoke:"#F5F5F5",
    yellow:"#FFFF00", yellowGreen:"#9ACD32"
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

// Convert RGB [0-255] to HLS [0-1] (OOXML / Microsoft HLS model)
export function rgbToHls(r, g, b) {
    r /= 255; g /= 255; b /= 255;
    var max = Math.max(r, g, b), min = Math.min(r, g, b);
    var l = (max + min) / 2, h = 0, s = 0;
    if (max !== min) {
        var d = max - min;
        s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
        else if (max === g) h = ((b - r) / d + 2) / 6;
        else h = ((r - g) / d + 4) / 6;
    }
    return { h: h, l: l, s: s };
}

function _hue2rgb(p, q, t) {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
}

// Convert HLS [0-1] back to RGB [0-255]
export function hlsToRgb(h, l, s) {
    if (s === 0) {
        var v = Math.round(l * 255);
        return { r: v, g: v, b: v };
    }
    var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    var p = 2 * l - q;
    return {
        r: Math.round(_hue2rgb(p, q, h + 1 / 3) * 255),
        g: Math.round(_hue2rgb(p, q, h) * 255),
        b: Math.round(_hue2rgb(p, q, h - 1 / 3) * 255)
    };
}

export function applyColorModifiers(hex, node) {
    if (!node || !hex) return hex;
    var rgb = hexToRgb(hex);
    for (var i = 0; i < node.childNodes.length; i++) {
        var cn = node.childNodes[i];
        if (cn.nodeType !== 1) continue;
        var lname = cn.localName;
        var v = parseInt(cn.getAttribute("val") || "100000");
        var pct = v / 100000;
        if (lname === "shade") {
            // shade: val=0 → black, val=100000 → original (RGB multiply)
            rgb.r = Math.round(rgb.r * pct);
            rgb.g = Math.round(rgb.g * pct);
            rgb.b = Math.round(rgb.b * pct);
        } else if (lname === "tint") {
            // tint: val=0 → no change, val=100000 → white
            rgb.r = Math.round(rgb.r + (255 - rgb.r) * pct);
            rgb.g = Math.round(rgb.g + (255 - rgb.g) * pct);
            rgb.b = Math.round(rgb.b + (255 - rgb.b) * pct);
        } else if (lname === "lumMod" || lname === "lumOff" ||
                   lname === "hueMod" || lname === "hueOff" ||
                   lname === "satMod" || lname === "satOff") {
            // HLS-space modifiers (ECMA-376 §20.1.2.3)
            var hls = rgbToHls(rgb.r, rgb.g, rgb.b);
            if (lname === "lumMod") {
                hls.l = Math.min(1, Math.max(0, hls.l * pct));
            } else if (lname === "lumOff") {
                // lumOff val is in 1000ths of a percent → same scale as l∈[0,1]
                hls.l = Math.min(1, Math.max(0, hls.l + pct));
            } else if (lname === "hueMod") {
                hls.h = ((hls.h * pct) % 1 + 1) % 1;
            } else if (lname === "hueOff") {
                // hueOff val is in 1/60000 of a degree
                hls.h = ((hls.h + v / 21600000) % 1 + 1) % 1;
            } else if (lname === "satMod") {
                hls.s = Math.min(1, Math.max(0, hls.s * pct));
            } else if (lname === "satOff") {
                hls.s = Math.min(1, Math.max(0, hls.s + pct));
            }
            var modified = hlsToRgb(hls.h, hls.l, hls.s);
            rgb.r = modified.r; rgb.g = modified.g; rgb.b = modified.b;
        }
    }
    return rgbToHex(rgb.r, rgb.g, rgb.b);
}

// Resolve color from an XML node containing srgbClr, schemeClr, or prstClr children
export function resolveColor(node) {
    if (!node) return null;
    var srgb = node.getElementsByTagNameNS(A_NS, "srgbClr")[0];
    if (srgb) return applyColorModifiers("#" + srgb.getAttribute("val"), srgb);
    var scrgb = node.getElementsByTagNameNS(A_NS, "scrgbClr")[0];
    if (scrgb) {
        var rr = parseInt(scrgb.getAttribute("r") || "0", 10);
        var gg = parseInt(scrgb.getAttribute("g") || "0", 10);
        var bb = parseInt(scrgb.getAttribute("b") || "0", 10);
        if (!Number.isFinite(rr)) rr = 0;
        if (!Number.isFinite(gg)) gg = 0;
        if (!Number.isFinite(bb)) bb = 0;
        rr = Math.max(0, Math.min(100000, rr));
        gg = Math.max(0, Math.min(100000, gg));
        bb = Math.max(0, Math.min(100000, bb));
        return applyColorModifiers(rgbToHex(rr * 255 / 100000, gg * 255 / 100000, bb * 255 / 100000), scrgb);
    }
    var scheme = node.getElementsByTagNameNS(A_NS, "schemeClr")[0];
    if (scheme) {
        var val = scheme.getAttribute("val") || "";
        var base = themeColors[val] || "#333333";
        return applyColorModifiers(base, scheme);
    }
    var prstClr = node.getElementsByTagNameNS(A_NS, "prstClr")[0];
    if (prstClr) {
        var pname = prstClr.getAttribute("val") || "";
        var pbase = PRESET_COLORS[pname] || "#000000";
        return applyColorModifiers(pbase, prstClr);
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
