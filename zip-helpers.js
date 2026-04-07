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
export async function loadImageAsDataUrl(zip, basePath, target) {
    if (!target) return null;
    var fullPath = (basePath + target).replace(/[^/]+\/\.\.\//g, "");
    var f = zip.file(fullPath);
    if (!f) return null;
    try {
        var blob = await f.async("blob");
        var ext = fullPath.split(".").pop().toLowerCase();
        var mime = (ext === "jpg" || ext === "jpeg") ? "image/jpeg" :
            ext === "gif" ? "image/gif" : ext === "svg" ? "image/svg+xml" : "image/png";
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
