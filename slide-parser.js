// ============================================================================
// slide-parser.js - Converts a single slide XML string into element array
// ============================================================================

import { A_NS, P_NS } from "./constants.js";
import { resolveColor } from "./color-utils.js";
import { parseShapeTree } from "./shape-parsers.js";

export function parseSlideXml(xmlStr, slideW, slideH, images, relsAll, hasBgImage, skipPlaceholders, layoutStyles, chartDataMap, diagramDataMap, sourceLayer) {
    var doc = new DOMParser().parseFromString(xmlStr, "application/xml");
    var bgColor = "#FFFFFF";

    var cSld = doc.getElementsByTagNameNS(P_NS, "cSld")[0];
    if (cSld) {
        var bg = cSld.getElementsByTagNameNS(P_NS, "bg")[0];
        if (bg) {
            var sf = bg.getElementsByTagNameNS(A_NS, "solidFill")[0];
            if (sf) { var c = resolveColor(sf); if (c) bgColor = c; }
        }
    }

    var spTree = cSld ? cSld.getElementsByTagNameNS(P_NS, "spTree")[0] : null;
    if (!spTree) return { elements: [], bgColor: bgColor };

    var opts = {
        skipPlaceholders: skipPlaceholders || false,
        hasBgImage: hasBgImage || false,
        layoutStyles: layoutStyles || {},
        chartDataMap: chartDataMap || {},
        diagramDataMap: diagramDataMap || {},
        sourceLayer: sourceLayer || "slide",
        gOffX: 0, gOffY: 0, gScaleX: 1, gScaleY: 1
    };
    var elements = parseShapeTree(spTree, slideW, slideH, images, relsAll, opts);
    return { elements: elements, bgColor: bgColor };
}
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        