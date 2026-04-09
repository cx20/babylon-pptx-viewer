// ============================================================================
// zip-helpers.js - ZIP file operations and relationship parsing
// ============================================================================

import { R_NS } from "./constants.js";

// Parse a .rels file and return structured relationship data
export async function parseRelsFile(zip, relsPath) {
    var result = { images: {}, layout: null, master: null, chart: null, all: {} };
    var f = zip.file(relsPath);
    if (!f) return result;
    var doc = new DOMParser().parseFromString(await f.async("string"), "application/xml");
    var rels = doc.getElementsByTagName("Relationship");
    for (var i = 0; i < rels.length; i++) {
        var r = rels[i];
        var id = r.getAttribute("Id"), type = r.getAttribute("Type") || "", tgt = r.getAttribute("Target") || "";
        result.all[id] = tgt;
        if (type.indexOf("/image") !== -1) result.images[id] = tgt;
        if (type.indexOf("/slideLayout") !== -1) result.layout = tgt;
        if (type.indexOf("/slideMaster") !== -1) result.master = tgt;
        if (type.indexOf("/chart") !== -1) result.chart = tgt;
    }
    return result;
}

// Resolve image from zip given base path and target
// Convert SVG text to a PNG data URL via an off-screen canvas.
// This avoids Babylon.js rendering issues with SVGs that lack explicit width/height attrs.
async function svgToPngDataUrl(svgText) {
    // Extract dimensions from viewBox if explicit attrs are absent
    var width = 96, height = 96;
    var vbMatch = svgText.match(/viewBox\s*=\s*["']([^"']+)["']/);
    if (vbMatch) {
        var vb = vbMatch[1].trim().split(/[\s,]+/);
        if (vb.length >= 4) {
            width  = Math.round(parseFloat(vb[2])) || 96;
            height = Math.round(parseFloat(vb[3])) || 96;
        }
    }
    // Inject width/height so browsers give the SVG intrinsic dimensions
    if (!/\s+width=/.test(svgText)) {
        svgText = svgText.replace(/<svg(\s|\b)/, '<svg width="' + width + '" height="' + height + '" ');
    }

    return new Promise(function (resolve) {
        var SCALE = 2; // Render at 2× for sharper icons
        var canvas = document.createElement("canvas");
        canvas.width  = width  * SCALE;
        canvas.height = height * SCALE;
        var ctx = canvas.getContext("2d");

        var svgBlob = new Blob([svgText], { type: "image/svg+xml" });
        var url = URL.createObjectURL(svgBlob);
        var img = new Image();

        img.onload = function () {
            ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
            URL.revokeObjectURL(url);
            resolve(canvas.toDataURL("image/png"));
        };
        img.onerror = function () {
            URL.revokeObjectURL(url);
            // Fall back: return SVG data URL directly
            var rd = new FileReader();
            rd.onload = function () { resolve(rd.result); };
            rd.onerror = function () { resolve(null); };
            rd.readAsDataURL(svgBlob);
        };
        img.src = url;
    });
}

export async function loadImageAsDataUrl(zip, basePath, target) {
    if (!target) return null;
    var fullPath = (basePath + target).replace(/[^/]+\/\.\.\//g, "");
    var f = zip.file(fullPath);
    if (!f) return null;
    try {
        var ext = fullPath.split(".").pop().toLowerCase();
        if (ext === "svg") {
            var svgText = await f.async("string");
            return await svgToPngDataUrl(svgText);
        }
        var blob = await f.async("blob");
        var mime = (ext === "jpg" || ext === "jpeg") ? "image/jpeg" :
            ext === "gif" ? "image/gif" : "image/png";
        return await new Promise(function (res) {
            var rd = new FileReader();
            rd.onload = function () { res(rd.result); };
            rd.readAsDataURL(new Blob([blob], { type: mime }));
        });
    } catch (e) { return null; }
}

// Build image map {rId: dataUrl} for a set of image relationships
export async function buildImageMap(zip, basePath, imageRels) {
    var map = {};
    for (var rId in imageRels) {
        map[rId] = await loadImageAsDataUrl(zip, basePath, imageRels[rId]);
    }
    return map;
}
