// ============================================================================
// Babylon.js PPTX Viewer - PowerPoint Simulator with Drag & Drop
// ============================================================================
// Renders .pptx files on a 3D monitor using Babylon.js GUI.
// Supports: background images/colors, theme colors, sp, grpSp, pic, cxnSp,
//           graphicFrame (chart placeholder, basic table), text with font
//           inheritance, slide/layout/master inheritance chain.
// ============================================================================

var createScene = async function () {
    var scene = new BABYLON.Scene(engine);
    scene.clearColor = new BABYLON.Color3(0.08, 0.08, 0.15);

    // --- Load JSZip ---
    await new Promise(function (resolve, reject) {
        if (window.JSZip && (typeof window.JSZip.loadAsync === "function" ||
            (window.JSZip.prototype && typeof window.JSZip.prototype.loadAsync === "function"))) {
            resolve(); return;
        }
        var s = document.createElement("script");
        s.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
        s.onload = resolve; s.onerror = reject;
        document.head.appendChild(s);
    });

    // ========================================================================
    // SECTION 1: Scene Setup - Camera, Lights, PC Model
    // ========================================================================
    var camera = new BABYLON.ArcRotateCamera("cam", -Math.PI / 2, Math.PI / 3, 14,
        new BABYLON.Vector3(0, 2.5, 0), scene);
    camera.attachControl(canvas, true);
    camera.wheelPrecision = 50;
    camera.lowerRadiusLimit = 5; camera.upperRadiusLimit = 30;
    camera.lowerBetaLimit = 0.1; camera.upperBetaLimit = Math.PI / 2 - 0.05;
    camera.keysUp = []; camera.keysDown = []; camera.keysLeft = []; camera.keysRight = [];

    var light = new BABYLON.HemisphericLight("light", new BABYLON.Vector3(0, 1, -0.3), scene);
    light.intensity = 1.0; light.groundColor = new BABYLON.Color3(0.3, 0.3, 0.35);
    var sLight = new BABYLON.PointLight("sLight", new BABYLON.Vector3(0, 3.5, -1.5), scene);
    sLight.intensity = 0.3; sLight.diffuse = new BABYLON.Color3(0.8, 0.9, 1.0);

    // Monitor
    var darkMat = new BABYLON.StandardMaterial("darkMat", scene);
    darkMat.diffuseColor = new BABYLON.Color3(0.15, 0.15, 0.15);
    darkMat.specularColor = new BABYLON.Color3(0.3, 0.3, 0.3);
    var monitorCase = BABYLON.MeshBuilder.CreateBox("mc", { width: 9, height: 5.8, depth: 0.4 }, scene);
    monitorCase.position.y = 3.4; monitorCase.material = darkMat;
    var bezelMat = new BABYLON.StandardMaterial("bzMat", scene);
    bezelMat.diffuseColor = new BABYLON.Color3(0.1, 0.1, 0.1);
    var bezel = BABYLON.MeshBuilder.CreateBox("bz", { width: 9.1, height: 5.9, depth: 0.35 }, scene);
    bezel.position.y = 3.4; bezel.position.z = 0.05; bezel.material = bezelMat;
    var standNeck = BABYLON.MeshBuilder.CreateBox("sn", { width: 0.8, height: 1.8, depth: 0.3 }, scene);
    standNeck.position.y = 0.9; standNeck.position.z = 0.3; standNeck.material = darkMat;
    var standBase = BABYLON.MeshBuilder.CreateBox("sBase", { width: 4, height: 0.15, depth: 2.5 }, scene);
    standBase.position.y = 0; standBase.material = darkMat;
    var deskMat = new BABYLON.StandardMaterial("dskMat", scene);
    deskMat.diffuseColor = new BABYLON.Color3(0.35, 0.25, 0.18);
    var desk = BABYLON.MeshBuilder.CreateBox("desk", { width: 14, height: 0.15, depth: 7 }, scene);
    desk.position.y = -0.15; desk.material = deskMat;
    // Screen plane for GUI texture
    var screenPlane = BABYLON.MeshBuilder.CreatePlane("screen", { width: 8.5, height: 5.3 }, scene);
    screenPlane.parent = monitorCase;
    screenPlane.position.z = -0.21; screenPlane.rotation.y = Math.PI; screenPlane.scaling.x = -1;

    // ========================================================================
    // SECTION 2: Constants & Slide Data Model
    // ========================================================================
    var SLIDE_EMU_W = 9144000;  // Default 10" slide width in EMU
    var SLIDE_EMU_H = 6858000;  // Default 7.5" slide height in EMU
    var CANVAS_W = 580;         // Pixel width of slide canvas in GUI
    var CANVAS_H = 326;         // Pixel height of slide canvas in GUI

    // XML namespaces
    var A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
    var P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
    var R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // Default slide with drop instructions
    var slides = [{
        bg: "#FFFFFF", bgImage: null,
        elements: [
            { type: "text", text: "Drag & Drop a .pptx file", x: 0.0, y: 0.30, w: 1.0, fontSize: 24, color: "#D04423", fontWeight: "bold", align: "center" },
            { type: "text", text: "onto this screen to load it", x: 0.0, y: 0.50, w: 1.0, fontSize: 16, color: "#666666", fontWeight: "normal", align: "center" },
            { type: "shape", shape: "line", x1: 0.15, y1: 0.42, x2: 0.85, y2: 0.42, color: "#D04423", thickness: 2 }
        ],
        notes: "Drop a .pptx file to begin."
    }];
    var currentSlide = 0;

    // ========================================================================
    // SECTION 3: Theme Color System
    // ========================================================================
    // Theme color map - populated from ppt/theme/theme1.xml
    var themeColors = {
        dk1: "#000000", dk2: "#44546A", lt1: "#FFFFFF", lt2: "#E7E6E6",
        accent1: "#4472C4", accent2: "#ED7D31", accent3: "#A5A5A5",
        accent4: "#FFC000", accent5: "#5B9BD5", accent6: "#70AD47",
        hlink: "#0563C1", folHlink: "#954F72",
        tx1: "#000000", tx2: "#44546A", bg1: "#FFFFFF", bg2: "#E7E6E6"
    };

    // Color modifier helpers
    function hexToRgb(hex) {
        hex = hex.replace("#", "");
        return { r: parseInt(hex.substr(0, 2), 16), g: parseInt(hex.substr(2, 2), 16), b: parseInt(hex.substr(4, 2), 16) };
    }
    function rgbToHex(r, g, b) {
        return "#" + [r, g, b].map(function (c) { return Math.max(0, Math.min(255, Math.round(c))).toString(16).padStart(2, "0"); }).join("");
    }
    function applyColorModifiers(hex, node) {
        if (!node || !hex) return hex;
        var rgb = hexToRgb(hex);
        for (var i = 0; i < node.childNodes.length; i++) {
            var cn = node.childNodes[i];
            if (cn.nodeType !== 1) continue;
            var v = parseInt(cn.getAttribute("val") || "100000");
            var pct = v / 100000;
            if (cn.localName === "shade") {
                rgb.r = Math.round(rgb.r * pct); rgb.g = Math.round(rgb.g * pct); rgb.b = Math.round(rgb.b * pct);
            } else if (cn.localName === "tint") {
                rgb.r = Math.round(rgb.r + (255 - rgb.r) * (1 - pct));
                rgb.g = Math.round(rgb.g + (255 - rgb.g) * (1 - pct));
                rgb.b = Math.round(rgb.b + (255 - rgb.b) * (1 - pct));
            } else if (cn.localName === "lumMod") {
                rgb.r = Math.round(rgb.r * pct); rgb.g = Math.round(rgb.g * pct); rgb.b = Math.round(rgb.b * pct);
            } else if (cn.localName === "lumOff") {
                var off = 255 * pct;
                rgb.r = Math.round(rgb.r + off); rgb.g = Math.round(rgb.g + off); rgb.b = Math.round(rgb.b + off);
            }
        }
        return rgbToHex(rgb.r, rgb.g, rgb.b);
    }

    // Resolve color from XML node (srgbClr or schemeClr), with modifiers
    function resolveColor(node) {
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

    // Parse theme XML and populate themeColors
    async function parseThemeXml(zip) {
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
        themeColors.tx1 = themeColors.dk1; themeColors.tx2 = themeColors.dk2;
        themeColors.bg1 = themeColors.lt1; themeColors.bg2 = themeColors.lt2;
        console.log("[PPTX] Theme colors loaded:", JSON.stringify(themeColors));
    }

    // ========================================================================
    // SECTION 4: ZIP / Relationship Helpers
    // ========================================================================
    // Parse .rels file and return structured relationship data
    async function parseRelsFile(zip, relsPath) {
        var result = { images: {}, layout: null, master: null, chart: null, all: {} };
        var f = zip.file(relsPath); if (!f) return result;
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
    async function loadImageAsDataUrl(zip, basePath, target) {
        if (!target) return null;
        var fullPath = (basePath + target).replace(/[^/]+\/\.\.\//g, "");
        var f = zip.file(fullPath); if (!f) return null;
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
    async function buildImageMap(zip, basePath, imageRels) {
        var map = {};
        for (var rId in imageRels) {
            map[rId] = await loadImageAsDataUrl(zip, basePath, imageRels[rId]);
        }
        return map;
    }

    // ========================================================================
    // SECTION 5: Background Extraction (with inheritance chain)
    // ========================================================================
    // Returns: dataUrl string (bg image), {solidColor:"#..."}, or null
    async function extractBackground(xmlStr, zip, basePath, relsAll, slideW, slideH) {
        var doc = new DOMParser().parseFromString(xmlStr, "application/xml");
        var cSld = doc.getElementsByTagNameNS(P_NS, "cSld")[0];
        if (!cSld) { console.log("[BG] No cSld found, base=" + basePath); return null; }

        var bg = cSld.getElementsByTagNameNS(P_NS, "bg")[0];
        console.log("[BG] base=" + basePath + " hasBgElement=" + !!bg);
        if (bg) {
            // bgPr > blipFill (image background)
            var bgPr = bg.getElementsByTagNameNS(P_NS, "bgPr")[0];
            if (bgPr) {
                var blipFill = bgPr.getElementsByTagNameNS(A_NS, "blipFill")[0];
                console.log("[BG]   bgPr found, hasBlipFill=" + !!blipFill);
                if (blipFill) {
                    var blip = blipFill.getElementsByTagNameNS(A_NS, "blip")[0];
                    if (blip) {
                        var rId = blip.getAttribute("r:embed") || blip.getAttributeNS(R_NS, "embed");
                        console.log("[BG]   blipFill rId=" + rId + " target=" + (relsAll[rId]||"NOT FOUND"));
                        
                        // Check if blip has extLst with alternative image reference (pre-rendered art effect)
                        var allBlipDescendants = blip.getElementsByTagName("*");
                        for (var bi = 0; bi < allBlipDescendants.length; bi++) {
                            var bel = allBlipDescendants[bi];
                            var altEmbed = bel.getAttribute("r:embed") || bel.getAttributeNS(R_NS, "embed");
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
                // bgPr > solidFill
                var sf = bgPr.getElementsByTagNameNS(A_NS, "solidFill")[0];
                if (sf) { var c = resolveColor(sf); if (c) return { solidColor: c }; }
                // bgPr > gradFill - approximate with first color
                var gf = bgPr.getElementsByTagNameNS(A_NS, "gradFill")[0];
                if (gf) {
                    var gs = gf.getElementsByTagNameNS(A_NS, "gs");
                    if (gs.length > 0) {
                        var c = resolveColor(gs[0]); if (c) return { solidColor: c };
                    }
                }
            }
            // bgRef (theme background reference)
            var bgRef = bg.getElementsByTagNameNS(P_NS, "bgRef")[0];
            if (!bgRef) bgRef = bg.getElementsByTagNameNS(A_NS, "bgRef")[0];
            console.log("[BG]   hasBgRef=" + !!bgRef);
            if (bgRef) {
                var c = resolveColor(bgRef); if (c) return { solidColor: c };
            }
        }

        // Check spTree for full-bleed background images (common in masters/layouts)
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
            var ox = parseInt(off.getAttribute("x")) || 0, oy = parseInt(off.getAttribute("y")) || 0;
            var cx = parseInt(ext.getAttribute("cx")) || 0, cy = parseInt(ext.getAttribute("cy")) || 0;
            // Image covering ≥70% of slide = background
            console.log("[BG]   pic[" + i + "] pos=(" + ox + "," + oy + ") size=(" + cx + "," + cy + ") covers=" + (cx/slideW*100).toFixed(0) + "%x" + (cy/slideH*100).toFixed(0) + "%");
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

    // ========================================================================
    // SECTION 6: Shape Outline & Geometry Helpers
    // ========================================================================
    function emuToFontPx(hundredthsPt) { return Math.round(hundredthsPt / 100 * 0.75); }

    function parseOutline(spPr) {
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

    function getPresetGeometry(spPr) {
        if (!spPr) return "rect";
        var pg = spPr.getElementsByTagNameNS(A_NS, "prstGeom")[0];
        return pg ? (pg.getAttribute("prst") || "rect") : "rect";
    }

    function getShapeFill(spPr) {
        if (!spPr) return null;
        for (var i = 0; i < spPr.childNodes.length; i++) {
            var n = spPr.childNodes[i];
            if (n.localName === "noFill") return null;
            if (n.localName === "solidFill") return resolveColor(n) || "#CCCCCC";
            if (n.localName === "gradFill") {
                // Approximate gradient with first gs color
                var gs = n.getElementsByTagNameNS(A_NS, "gs");
                if (gs.length > 0) return resolveColor(gs[0]) || "#CCCCCC";
            }
        }
        return null;
    }

    // ========================================================================
    // SECTION 7: Text / Paragraph Parser
    // ========================================================================
    function parseParagraphs(txBody, defaultFS, defaultFC, layoutCap) {
        defaultFS = defaultFS || 14; defaultFC = defaultFC || "#333"; layoutCap = layoutCap || "";
        var result = [], paras = txBody.getElementsByTagNameNS(A_NS, "p");
        for (var p = 0; p < paras.length; p++) {
            var para = paras[p], runs = para.getElementsByTagNameNS(A_NS, "r");
            var fs = defaultFS, fw = "normal", fc = defaultFC, fi = false, cap = layoutCap;

            // defRPr from pPr
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
            // endParaRPr fallback
            var epr = para.getElementsByTagNameNS(A_NS, "endParaRPr")[0];
            if (epr) {
                var esz = epr.getAttribute("sz");
                if (esz && fs === defaultFS) fs = emuToFontPx(parseInt(esz));
                if (epr.getAttribute("b") === "1" && fw === "normal") fw = "bold";
                if (epr.getAttribute("cap") && !cap) cap = epr.getAttribute("cap");
                var ef = epr.getElementsByTagNameNS(A_NS, "solidFill")[0];
                if (ef && fc === defaultFC) { var ec = resolveColor(ef); if (ec) fc = ec; }
            }

            // Process runs for text and run-level formatting
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

            // Alignment
            var align = "left";
            if (pPr) {
                var al = pPr.getAttribute("algn");
                if (al === "ctr") align = "center"; else if (al === "r") align = "right";
            }
            // Bullets
            if (pPr) {
                var bc = pPr.getElementsByTagNameNS(A_NS, "buChar")[0];
                if (bc && txt.trim()) txt = (bc.getAttribute("char") || "•") + " " + txt;
            }
            // Line spacing (approximate)
            var lnSpc = 1.5; // default multiplier
            if (pPr) {
                var ls = pPr.getElementsByTagNameNS(A_NS, "lnSpc")[0];
                if (ls) {
                    var spcPct = ls.getElementsByTagNameNS(A_NS, "spcPct")[0];
                    if (spcPct) { var v = parseInt(spcPct.getAttribute("val")) || 100000; lnSpc = v / 100000 * 1.5; }
                }
            }

            // Apply capitalization (cap="all" → ALL CAPS)
            if (cap === "all" && txt) txt = txt.toUpperCase();

            result.push({
                text: txt, fontSize: Math.min(Math.max(fs, 6), 60),
                fontWeight: fw, color: fc, italic: fi, align: align,
                isEmpty: txt.trim().length === 0, lineSpacing: lnSpc
            });
        }
        return result;
    }

    // ========================================================================
    // SECTION 7b: Layout Placeholder Style Extraction
    // ========================================================================
    // Extract text styles (cap, fontSize, color, anchor) from layout placeholder shapes
    // so slide shapes can inherit them.
    async function extractPlaceholderStyles(zip, layoutPath) {
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

            // Check bodyPr for anchor
            var txBody = sp.getElementsByTagNameNS(P_NS, "txBody")[0];
            if (!txBody) txBody = sp.getElementsByTagNameNS(A_NS, "txBody")[0];
            if (txBody) {
                var bodyPr = txBody.getElementsByTagNameNS(A_NS, "bodyPr")[0];
                if (bodyPr) {
                    var anc = bodyPr.getAttribute("anchor");
                    if (anc) style.anchor = anc;
                }

                // Check lstStyle for defRPr properties
                var lstStyle = txBody.getElementsByTagNameNS(A_NS, "lstStyle")[0];
                if (lstStyle) {
                    // Check all levels (lvl1pPr through lvl9pPr) and defPPr
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

                // Also check direct paragraphs' defRPr in layout placeholder
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

            // Also read p:style > a:fontRef for font color
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

    // Extract text styles from slide master's p:txStyles AND placeholder fontRef colors
    async function extractMasterTxStyles(zip, masterPath) {
        var result = { titleColor: null, bodyColor: null, otherColor: null, phFontRef: {} };
        var f = zip.file(masterPath);
        if (!f) return result;
        var xml = await f.async("string");
        var doc = new DOMParser().parseFromString(xml, "application/xml");

        // 1. Read txStyles
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

        // 2. Read placeholder shapes in master spTree for fontRef colors
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

    // ========================================================================
    // SECTION 7c: Duotone/Tint Detection for Background Images
    // ========================================================================
    function extractBlipEffects(xmlStr) {
        var doc = new DOMParser().parseFromString(xmlStr, "application/xml");
        var cSld = doc.getElementsByTagNameNS(P_NS, "cSld")[0];
        if (!cSld) return null;
        var bg = cSld.getElementsByTagNameNS(P_NS, "bg")[0];
        if (!bg) return null;
        var blip = bg.getElementsByTagNameNS(A_NS, "blip")[0];
        if (!blip) return null;

        // Debug: dump blip children
        var blipKids = [];
        for (var ci = 0; ci < blip.childNodes.length; ci++) {
            if (blip.childNodes[ci].nodeType === 1) blipKids.push(blip.childNodes[ci].localName);
        }
        console.log("[BLIP] blip children: [" + blipKids.join(", ") + "]");

        // Search ENTIRE bg subtree for art effect URI
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
                    console.log("[BLIP] Grayscale duotone detected, applying dk2 tint as art effect approximation");
                    return { type: "artEffect", color: themeColors.dk2 || "#0E5580" };
                } else {
                    return { type: "duotone", dark: colors[0], light: colors[1] };
                }
            }
        }

        // Check for colorChange / clrRepl
        var clrChange = blip.getElementsByTagNameNS(A_NS, "clrChange")[0];
        if (clrChange) return { type: "tint", color: themeColors.dk2 || "#0E5580" };

        var alphaModFix = blip.getElementsByTagNameNS(A_NS, "alphaModFix")[0];
        if (alphaModFix) return { type: "alpha", amt: parseInt(alphaModFix.getAttribute("amt") || "100000") / 100000 };

        return null;
    }

    // ========================================================================
    // SECTION 8: Shape Tree Parser (sp, grpSp, pic, cxnSp, graphicFrame)
    // ========================================================================
    // Parse all children of a spTree/grpSp node.
    // groupOff/groupExt: the group's child coordinate space (for nested grpSp)
    // parentOff: the group's position on the slide
    function parseShapeTree(spTreeNode, slideW, slideH, images, relsAll, opts) {
        opts = opts || {};
        var elements = [];
        var skipPlaceholders = opts.skipPlaceholders || false;
        var hasBgImage = opts.hasBgImage || false;
        var layoutStyles = opts.layoutStyles || {};
        var defaultTextColor = hasBgImage ? (themeColors.lt1 || "#FFF") : (themeColors.tx1 || "#333");
        // Group transform: convert child coords to slide fraction coords
        var gOffX = opts.gOffX || 0, gOffY = opts.gOffY || 0;
        var gScaleX = opts.gScaleX || 1, gScaleY = opts.gScaleY || 1;

        function toFracX(emu) { return (gOffX + emu * gScaleX) / slideW; }
        function toFracY(emu) { return (gOffY + emu * gScaleY) / slideH; }
        function toFracW(emu) { return emu * gScaleX / slideW; }
        function toFracH(emu) { return emu * gScaleY / slideH; }

        // Iterate direct children
        var childCount = {sp:0, pic:0, grpSp:0, cxnSp:0, graphicFrame:0, other:0};
        for (var ci = 0; ci < spTreeNode.childNodes.length; ci++) {
            var child = spTreeNode.childNodes[ci];
            if (child.nodeType !== 1) continue;
            var ln = child.localName;
            if (ln === "sp") childCount.sp++;
            else if (ln === "pic") childCount.pic++;
            else if (ln === "grpSp") childCount.grpSp++;
            else if (ln === "cxnSp") childCount.cxnSp++;
            else if (ln === "graphicFrame") childCount.graphicFrame++;
            else childCount.other++;
        }
        console.log("[TREE] children: sp=" + childCount.sp + " pic=" + childCount.pic + " grpSp=" + childCount.grpSp + " cxnSp=" + childCount.cxnSp + " graphicFrame=" + childCount.graphicFrame + " other=" + childCount.other + " skipPH=" + skipPlaceholders);

        for (var ci = 0; ci < spTreeNode.childNodes.length; ci++) {
            var child = spTreeNode.childNodes[ci];
            if (child.nodeType !== 1) continue; // element nodes only
            var localName = child.localName;

            // --- sp (shape) ---
            if (localName === "sp") {
                parseSp(child, elements, slideW, slideH, skipPlaceholders, defaultTextColor, toFracX, toFracY, toFracW, toFracH, layoutStyles, hasBgImage);
            }
            // --- pic (picture) ---
            else if (localName === "pic") {
                parsePic(child, elements, slideW, slideH, images, relsAll, hasBgImage, toFracX, toFracY, toFracW, toFracH);
            }
            // --- cxnSp (connector) ---
            else if (localName === "cxnSp") {
                parseCxnSp(child, elements, slideW, slideH, toFracX, toFracY);
            }
            // --- grpSp (group shape) - RECURSIVE ---
            else if (localName === "grpSp") {
                parseGrpSp(child, elements, slideW, slideH, images, relsAll, opts);
            }
            // --- graphicFrame (chart / table / diagram) ---
            else if (localName === "graphicFrame") {
                parseGraphicFrame(child, elements, slideW, slideH, images, relsAll, defaultTextColor, toFracX, toFracY, toFracW, toFracH);
            }
        }
        return elements;
    }

    // --- Parse sp (shape with optional text) ---
    function parseSp(sp, elements, slideW, slideH, skipPH, defTextColor, fx, fy, fw, fh, layoutStyles, hasBgImage) {
        layoutStyles = layoutStyles || {};
        hasBgImage = hasBgImage || false;
        // Placeholder detection
        var phType = "", phIdx = -1;
        var nvSpPr = sp.getElementsByTagNameNS(P_NS, "nvSpPr")[0];
        if (nvSpPr) {
            var nvPr = nvSpPr.getElementsByTagNameNS(P_NS, "nvPr")[0];
            if (nvPr) {
                var ph = nvPr.getElementsByTagNameNS(P_NS, "ph")[0];
                if (ph) { phType = ph.getAttribute("type") || "body"; phIdx = parseInt(ph.getAttribute("idx")) || 0; }
            }
        }
        if (skipPH && phType) {
            console.log("[SP] SKIP placeholder type='" + phType + "'");
            return;
        }

        // Transform
        var xfrm = sp.getElementsByTagNameNS(A_NS, "xfrm")[0];
        var ox = 0, oy = 0, cx = 0, cy = 0, rot = 0;
        if (xfrm) {
            var off = xfrm.getElementsByTagNameNS(A_NS, "off")[0];
            var ext = xfrm.getElementsByTagNameNS(A_NS, "ext")[0];
            if (off) { ox = parseInt(off.getAttribute("x")) || 0; oy = parseInt(off.getAttribute("y")) || 0; }
            if (ext) { cx = parseInt(ext.getAttribute("cx")) || 0; cy = parseInt(ext.getAttribute("cy")) || 0; }
            rot = parseInt(xfrm.getAttribute("rot")) || 0;
        }
        var fracX = fx(ox), fracY = fy(oy), fracW = fw(cx), fracH = fh(cy);
        var rotDeg = rot / 60000;

        // Shape visual properties
        var spPr = sp.getElementsByTagNameNS(A_NS, "spPr")[0];
        var geom = getPresetGeometry(spPr);
        var outline = parseOutline(spPr);
        var fill = getShapeFill(spPr);

        // Style-based fill/color
        var styleFontColor = null;
        var style = sp.getElementsByTagNameNS(P_NS, "style")[0];
        if (style) {
            if (!fill && !outline) {
                var fillRef = style.getElementsByTagNameNS(A_NS, "fillRef")[0];
                if (fillRef) { var fc = resolveColor(fillRef); if (fc) fill = fc; }
            }
            var fontRef = style.getElementsByTagNameNS(A_NS, "fontRef")[0];
            if (fontRef) styleFontColor = resolveColor(fontRef);
        }

        // Emit shape rectangle/ellipse
        if ((fill || outline) && cx > 0 && cy > 0) {
            elements.push({
                type: "shape", shape: geom, x: fracX, y: fracY, w: fracW, h: fracH,
                fillColor: fill || "transparent",
                borderColor: outline ? outline.color : "transparent",
                borderWidth: outline ? outline.width : 0,
                rotation: rotDeg
            });
        }

        // Placeholder font defaults
        var phFS = 14, phFC = defTextColor;
        var phLayout = layoutStyles[phType];
        if (phType === "title" || phType === "ctrTitle") phFS = 32;
        else if (phType === "subTitle") {
            phFS = 20;
            // PowerPoint convention: subtitle uses accent1 color on dark background slides
            if (hasBgImage && themeColors.accent1) phFC = themeColors.accent1;
        }
        else if (phType === "body" || phType === "obj") phFS = 18;
        else if (!phType && cy > 0) phFS = Math.min(Math.max(Math.round(fracH * CANVAS_H * 0.4), 10), 36);
        // Apply layout fontRef color as default, then slide's own style overrides
        if (phLayout && phLayout.fontRefColor) phFC = phLayout.fontRefColor;
        if (styleFontColor) phFC = styleFontColor;

        // Text body
        var txBody = sp.getElementsByTagNameNS(P_NS, "txBody")[0];
        if (!txBody) txBody = sp.getElementsByTagNameNS(A_NS, "txBody")[0];
        if (!txBody) {
            console.log("[SP] geom=" + geom + " ph='" + phType + "' fill=" + fill + " pos=(" + fracX.toFixed(3) + "," + fracY.toFixed(3) + ") NO txBody");
            return;
        }

        console.log("[SP] geom=" + geom + " ph='" + phType + "' fill=" + fill + " phFS=" + phFS + " phFC=" + phFC + " pos=(" + fracX.toFixed(3) + "," + fracY.toFixed(3) + ") size=(" + fracW.toFixed(3) + "," + fracH.toFixed(3) + ")");

        // Body properties (anchor, insets)
        var bodyPr = txBody.getElementsByTagNameNS(A_NS, "bodyPr")[0];
        // Default anchor based on placeholder type
        // ctrTitle typically has bottom-aligned text in PowerPoint
        var defaultAnchor = "t";
        if (phType === "ctrTitle") defaultAnchor = "b";
        else if (phType === "title") defaultAnchor = "b";
        // Override with layout placeholder anchor
        if (phLayout && phLayout.anchor) defaultAnchor = phLayout.anchor;
        
        var anchor = defaultAnchor;
        var iL = 91440 / slideW, iT = 45720 / slideH, iR = 91440 / slideW, iB = 45720 / slideH;
        if (bodyPr) {
            var explicitAnchor = bodyPr.getAttribute("anchor");
            if (explicitAnchor) anchor = explicitAnchor;
            var lI = bodyPr.getAttribute("lIns"), tI = bodyPr.getAttribute("tIns");
            var rI = bodyPr.getAttribute("rIns"), bI = bodyPr.getAttribute("bIns");
            if (lI !== null) iL = parseInt(lI) / slideW;
            if (tI !== null) iT = parseInt(tI) / slideH;
            if (rI !== null) iR = parseInt(rI) / slideW;
            if (bI !== null) iB = parseInt(bI) / slideH;
        }
        console.log("[SP]   anchor=" + anchor + " (explicit=" + (bodyPr && bodyPr.getAttribute("anchor") || "none") + " default=" + defaultAnchor + ")");

        var layoutCap = (phLayout && phLayout.cap) ? phLayout.cap : "";
        var paras = parseParagraphs(txBody, phFS, phFC, layoutCap);
        paras.forEach(function(p, pi) {
            if (!p.isEmpty) console.log("[SP]   para[" + pi + "] '" + p.text.substring(0,30) + "' fs=" + p.fontSize + " color=" + p.color + " align=" + p.align);
        });

        // Calculate vertical positioning based on anchor
        var totalH = 0, paraH = [];
        paras.forEach(function (p) {
            var h = p.isEmpty ? p.fontSize * 0.6 : p.fontSize * p.lineSpacing;
            paraH.push(h); totalH += h;
        });
        var areaTop = fracY + iT, areaH = fracH - iT - iB;
        var thFrac = totalH / CANVAS_H;
        var startY = areaTop;
        if (anchor === "ctr" || anchor === "mid") startY = areaTop + (areaH - thFrac) / 2;
        else if (anchor === "b") startY = areaTop + areaH - thFrac;

        var curY = startY;
        paras.forEach(function (p, pi) {
            if (!p.isEmpty) {
                elements.push({
                    type: "text", text: p.text,
                    x: fracX + iL, y: curY, w: fracW - iL - iR,
                    fontSize: p.fontSize, color: p.color,
                    fontWeight: p.fontWeight, fontStyle: p.italic ? "italic" : "normal",
                    align: p.align, rotation: rotDeg
                });
            }
            curY += paraH[pi] / CANVAS_H;
        });
    }

    // --- Parse pic (picture) ---
    function parsePic(pic, elements, slideW, slideH, images, relsAll, hasBgImage, fx, fy, fw, fh) {
        var xfrm = pic.getElementsByTagNameNS(A_NS, "xfrm")[0]; if (!xfrm) return;
        var off = xfrm.getElementsByTagNameNS(A_NS, "off")[0];
        var ext = xfrm.getElementsByTagNameNS(A_NS, "ext")[0];
        if (!off || !ext) return;
        var ox = parseInt(off.getAttribute("x")) || 0, oy = parseInt(off.getAttribute("y")) || 0;
        var cx = parseInt(ext.getAttribute("cx")) || 0, cy = parseInt(ext.getAttribute("cy")) || 0;

        // Skip full-bleed background images (already handled)
        if (hasBgImage && cx > slideW * 0.7 && cy > slideH * 0.7) return;

        var blip = pic.getElementsByTagNameNS(A_NS, "blip")[0];
        if (!blip) return;
        var rId = blip.getAttribute("r:embed") || blip.getAttributeNS(R_NS, "embed");
        if (!rId || !images[rId]) return;

        // srcRect (crop) - approximate by adjusting position/size
        var fracX = fx(ox), fracY = fy(oy), fracW = fw(cx), fracH = fh(cy);
        var blipFill = pic.getElementsByTagNameNS(A_NS, "blipFill")[0];
        var cropL = 0, cropT = 0, cropR = 0, cropB = 0;
        if (blipFill) {
            var srcRect = blipFill.getElementsByTagNameNS(A_NS, "srcRect")[0];
            if (srcRect) {
                cropL = (parseInt(srcRect.getAttribute("l")) || 0) / 100000;
                cropT = (parseInt(srcRect.getAttribute("t")) || 0) / 100000;
                cropR = (parseInt(srcRect.getAttribute("r")) || 0) / 100000;
                cropB = (parseInt(srcRect.getAttribute("b")) || 0) / 100000;
            }
        }

        elements.push({
            type: "image", dataUrl: images[rId],
            x: fracX, y: fracY, w: fracW, h: fracH,
            crop: { l: cropL, t: cropT, r: cropR, b: cropB }
        });
        console.log("[PIC] image at (" + fracX.toFixed(3) + "," + fracY.toFixed(3) + ") size=(" + fracW.toFixed(3) + "," + fracH.toFixed(3) + ") crop=L" + (cropL*100).toFixed(0) + "%,T" + (cropT*100).toFixed(0) + "%,R" + (cropR*100).toFixed(0) + "%,B" + (cropB*100).toFixed(0) + "%");
    }

    // --- Parse cxnSp (connector line) ---
    function parseCxnSp(cxn, elements, slideW, slideH, fx, fy) {
        var xfrm = cxn.getElementsByTagNameNS(A_NS, "xfrm")[0]; if (!xfrm) return;
        var off = xfrm.getElementsByTagNameNS(A_NS, "off")[0];
        var ext = xfrm.getElementsByTagNameNS(A_NS, "ext")[0];
        if (!off || !ext) return;
        var x1 = parseInt(off.getAttribute("x")) || 0, y1 = parseInt(off.getAttribute("y")) || 0;
        var w = parseInt(ext.getAttribute("cx")) || 0, h = parseInt(ext.getAttribute("cy")) || 0;
        var flipH = xfrm.getAttribute("flipH") === "1", flipV = xfrm.getAttribute("flipV") === "1";
        var spPr = cxn.getElementsByTagNameNS(A_NS, "spPr")[0];
        var ol = parseOutline(spPr);
        console.log("[CXN] connector line color=" + (ol?ol.color:"#000") + " flipH=" + flipH + " flipV=" + flipV);
        elements.push({
            type: "shape", shape: "line",
            x1: fx(flipH ? x1 + w : x1), y1: fy(flipV ? y1 + h : y1),
            x2: fx(flipH ? x1 : x1 + w), y2: fy(flipV ? y1 : y1 + h),
            color: ol ? ol.color : "#000", thickness: ol ? ol.width : 1
        });
    }

    // --- Parse grpSp (group shape) - RECURSIVE ---
    function parseGrpSp(grpSp, elements, slideW, slideH, images, relsAll, parentOpts) {
        console.log("[GRPSP] Parsing group shape");
        var grpSpPr = grpSp.getElementsByTagNameNS(A_NS, "grpSpPr")[0];
        if (!grpSpPr) grpSpPr = grpSp.getElementsByTagNameNS(P_NS, "grpSpPr")[0];

        // Group has two coordinate spaces:
        // off/ext = position and size on parent
        // chOff/chExt = child coordinate space
        var offX = 0, offY = 0, extW = 1, extH = 1, chOffX = 0, chOffY = 0, chExtW = 1, chExtH = 1;
        if (grpSpPr) {
            var xfrm = grpSpPr.getElementsByTagNameNS(A_NS, "xfrm")[0];
            if (xfrm) {
                var off = xfrm.getElementsByTagNameNS(A_NS, "off")[0];
                var ext = xfrm.getElementsByTagNameNS(A_NS, "ext")[0];
                var chOff = xfrm.getElementsByTagNameNS(A_NS, "chOff")[0];
                var chExt = xfrm.getElementsByTagNameNS(A_NS, "chExt")[0];
                if (off) { offX = parseInt(off.getAttribute("x")) || 0; offY = parseInt(off.getAttribute("y")) || 0; }
                if (ext) { extW = parseInt(ext.getAttribute("cx")) || 1; extH = parseInt(ext.getAttribute("cy")) || 1; }
                if (chOff) { chOffX = parseInt(chOff.getAttribute("x")) || 0; chOffY = parseInt(chOff.getAttribute("y")) || 0; }
                if (chExt) { chExtW = parseInt(chExt.getAttribute("cx")) || 1; chExtH = parseInt(chExt.getAttribute("cy")) || 1; }
            }
        }

        // Calculate transform: child EMU → slide EMU
        var pGOffX = (parentOpts && parentOpts.gOffX) || 0;
        var pGOffY = (parentOpts && parentOpts.gOffY) || 0;
        var pGScaleX = (parentOpts && parentOpts.gScaleX) || 1;
        var pGScaleY = (parentOpts && parentOpts.gScaleY) || 1;

        var newGOffX = pGOffX + (offX - chOffX * (extW / chExtW)) * pGScaleX;
        var newGOffY = pGOffY + (offY - chOffY * (extH / chExtH)) * pGScaleY;
        var newGScaleX = pGScaleX * (extW / chExtW);
        var newGScaleY = pGScaleY * (extH / chExtH);

        var childOpts = {
            skipPlaceholders: parentOpts ? parentOpts.skipPlaceholders : false,
            hasBgImage: parentOpts ? parentOpts.hasBgImage : false,
            layoutStyles: parentOpts ? (parentOpts.layoutStyles || {}) : {},
            gOffX: newGOffX, gOffY: newGOffY,
            gScaleX: newGScaleX, gScaleY: newGScaleY
        };
        console.log("[GRPSP]   off=(" + offX + "," + offY + ") ext=(" + extW + "," + extH + ") chOff=(" + chOffX + "," + chOffY + ") chExt=(" + chExtW + "," + chExtH + ") scale=(" + newGScaleX.toFixed(3) + "," + newGScaleY.toFixed(3) + ")");

        var childElements = parseShapeTree(grpSp, slideW, slideH, images, relsAll, childOpts);
        childElements.forEach(function (el) { elements.push(el); });
    }

    // --- Parse graphicFrame (chart / table / diagram) ---
    function parseGraphicFrame(gf, elements, slideW, slideH, images, relsAll, defTextColor, fx, fy, fw, fh) {
        // graphicFrame uses p:xfrm, not a:xfrm
        var xfrm = gf.getElementsByTagNameNS(P_NS, "xfrm")[0];
        if (!xfrm) xfrm = gf.getElementsByTagNameNS(A_NS, "xfrm")[0];
        if (!xfrm) { console.log("[GF] No xfrm found in graphicFrame"); return; }
        // off/ext may be in a: or p: namespace depending on xfrm parent
        var off = xfrm.getElementsByTagNameNS(A_NS, "off")[0] || xfrm.getElementsByTagNameNS(P_NS, "off")[0];
        var ext = xfrm.getElementsByTagNameNS(A_NS, "ext")[0] || xfrm.getElementsByTagNameNS(P_NS, "ext")[0];
        // Also try without namespace (some serializers omit prefix for children)
        if (!off || !ext) {
            for (var ci = 0; ci < xfrm.childNodes.length; ci++) {
                var cn = xfrm.childNodes[ci];
                if (cn.localName === "off" && !off) off = cn;
                if (cn.localName === "ext" && !ext) ext = cn;
            }
        }
        if (!off || !ext) { console.log("[GF] No off/ext in graphicFrame xfrm"); return; }
        var ox = parseInt(off.getAttribute("x")) || 0, oy = parseInt(off.getAttribute("y")) || 0;
        var cx = parseInt(ext.getAttribute("cx")) || 0, cy = parseInt(ext.getAttribute("cy")) || 0;
        var fracX = fx(ox), fracY = fy(oy), fracW = fw(cx), fracH = fh(cy);

        // Check graphic data namespace to determine type
        // Structure: p:graphicFrame > a:graphic > a:graphicData
        var graphic = gf.getElementsByTagNameNS(A_NS, "graphic")[0];
        var graphicData = null;
        if (graphic) {
            graphicData = graphic.getElementsByTagNameNS(A_NS, "graphicData")[0];
        }
        if (!graphicData) {
            // Fallback: search without namespace
            graphicData = gf.getElementsByTagName("a:graphicData")[0];
        }
        if (!graphicData) {
            // Broader fallback
            graphicData = gf.getElementsByTagNameNS(A_NS, "graphicData")[0];
        }
        var uri = graphicData ? (graphicData.getAttribute("uri") || "") : "";
        console.log("[GF] graphicFrame at (" + fracX.toFixed(3) + "," + fracY.toFixed(3) + ") size=(" + fracW.toFixed(3) + "," + fracH.toFixed(3) + ") uri=" + uri + " hasGraphicData=" + !!graphicData);

        // Table (a:tbl)
        if (uri.indexOf("table") !== -1 || uri.indexOf("dgm") !== -1) {
            parseTable(graphicData, elements, fracX, fracY, fracW, fracH, defTextColor);
            return;
        }

        // Chart or other - render as placeholder box
        var label = "Chart";
        if (uri.indexOf("chart") !== -1) label = "📊 Chart";
        else if (uri.indexOf("diagram") !== -1) label = "📐 Diagram";
        else if (uri.indexOf("ole") !== -1) label = "📎 OLE Object";
        else label = "📋 Object";

        elements.push({
            type: "shape", shape: "rect", x: fracX, y: fracY, w: fracW, h: fracH,
            fillColor: "rgba(200,200,200,0.3)", borderColor: "#999", borderWidth: 1, rotation: 0
        });
        elements.push({
            type: "text", text: label,
            x: fracX, y: fracY + fracH * 0.35, w: fracW,
            fontSize: 12, color: "#666", fontWeight: "normal", fontStyle: "normal", align: "center"
        });
    }

    // --- Parse a:tbl (table) from graphicData ---
    function parseTable(graphicData, elements, fracX, fracY, fracW, fracH, defTextColor) {
        if (!graphicData) return;
        var tbl = graphicData.getElementsByTagNameNS(A_NS, "tbl")[0];
        if (!tbl) return;
        var tblGrid = tbl.getElementsByTagNameNS(A_NS, "tblGrid")[0];
        var rows = tbl.getElementsByTagNameNS(A_NS, "tr");
        if (!rows || rows.length === 0) return;
        console.log("[TBL] Table: " + rows.length + " rows, cols=" + (tblGrid ? tblGrid.getElementsByTagNameNS(A_NS, "gridCol").length : "?"));

        // Calculate column widths from tblGrid
        var colWidths = [];
        var totalW = 0;
        if (tblGrid) {
            var gridCols = tblGrid.getElementsByTagNameNS(A_NS, "gridCol");
            for (var i = 0; i < gridCols.length; i++) {
                var w = parseInt(gridCols[i].getAttribute("w")) || 100000;
                colWidths.push(w); totalW += w;
            }
        }
        if (totalW === 0) totalW = 1;

        // Calculate row heights
        var rowHeights = [];
        var totalH = 0;
        for (var r = 0; r < rows.length; r++) {
            var h = parseInt(rows[r].getAttribute("h")) || 300000;
            rowHeights.push(h); totalH += h;
        }
        if (totalH === 0) totalH = 1;

        // Render cells
        var curY = fracY;
        for (var r = 0; r < rows.length; r++) {
            var rh = fracH * (rowHeights[r] / totalH);
            var cells = rows[r].getElementsByTagNameNS(A_NS, "tc");
            var curX = fracX;
            for (var c = 0; c < cells.length && c < colWidths.length; c++) {
                var cw = fracW * (colWidths[c] / totalW);

                // Cell background
                var tcPr = cells[c].getElementsByTagNameNS(A_NS, "tcPr")[0];
                var cellFill = null;
                if (tcPr) {
                    var sf = tcPr.getElementsByTagNameNS(A_NS, "solidFill")[0];
                    if (sf) cellFill = resolveColor(sf);
                }

                // Cell border
                elements.push({
                    type: "shape", shape: "rect", x: curX, y: curY, w: cw, h: rh,
                    fillColor: cellFill || "transparent", borderColor: "#AAA", borderWidth: 1, rotation: 0
                });

                // Cell text
                var txBody = cells[c].getElementsByTagNameNS(A_NS, "txBody")[0];
                if (txBody) {
                    var cellParas = parseParagraphs(txBody, 10, defTextColor);
                    var textY = curY + 0.005;
                    cellParas.forEach(function (p) {
                        if (!p.isEmpty) {
                            elements.push({
                                type: "text", text: p.text, x: curX + 0.005, y: textY, w: cw - 0.01,
                                fontSize: Math.min(p.fontSize, 14), color: p.color,
                                fontWeight: p.fontWeight, fontStyle: p.italic ? "italic" : "normal", align: p.align
                            });
                        }
                        textY += p.fontSize * 1.2 / CANVAS_H;
                    });
                }
                curX += cw;
            }
            curY += rh;
        }
    }

    // ========================================================================
    // SECTION 9: Slide XML → Elements (top-level parser)
    // ========================================================================
    function parseSlideXml(xmlStr, slideW, slideH, images, relsAll, hasBgImage, skipPlaceholders, layoutStyles) {
        var doc = new DOMParser().parseFromString(xmlStr, "application/xml");
        var bgColor = "#FFFFFF";

        // Extract background color (solid only; image handled separately)
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
            gOffX: 0, gOffY: 0, gScaleX: 1, gScaleY: 1
        };
        var elements = parseShapeTree(spTree, slideW, slideH, images, relsAll, opts);
        return { elements: elements, bgColor: bgColor };
    }

    // ========================================================================
    // SECTION 10: Main PPTX Parser (orchestrates everything)
    // ========================================================================
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
            // Master placeholder fontRef colors apply to ALL slides (including bgImage)
            // But skip title-type placeholders on bgImage slides (they should stay white)
            for (var phKey in masterTxStyles.phFontRef) {
                if (hasBgImage && (phKey === "title" || phKey === "ctrTitle")) continue;
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

    // ========================================================================
    // SECTION 11: GUI Construction (PowerPoint UI Frame)
    // ========================================================================
    var TEX_W = 1024, TEX_H = 640;
    var advTex = BABYLON.GUI.AdvancedDynamicTexture.CreateForMesh(screenPlane, TEX_W, TEX_H);
    var PP = "#D04423";

    var root = new BABYLON.GUI.Grid(); root.background = "#E0E0E0"; advTex.addControl(root);
    root.addRowDefinition(28, true); root.addRowDefinition(62, true);
    root.addRowDefinition(1.0); root.addRowDefinition(22, true);
    root.addColumnDefinition(1.0);

    // --- Title Bar ---
    var titleBar = new BABYLON.GUI.Rectangle(); titleBar.background = PP; titleBar.thickness = 0;
    root.addControl(titleBar, 0, 0);
    var pIcon = new BABYLON.GUI.TextBlock(); pIcon.text = "P"; pIcon.color = "white"; pIcon.fontSize = 16;
    pIcon.fontWeight = "bold"; pIcon.fontFamily = "Segoe UI,sans-serif";
    pIcon.textHorizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
    pIcon.left = "8px"; pIcon.width = "20px"; titleBar.addControl(pIcon);
    var titleText = new BABYLON.GUI.TextBlock(); titleText.text = "Presentation1 - PowerPoint";
    titleText.color = "white"; titleText.fontSize = 12; titleText.fontFamily = "Segoe UI,sans-serif";
    titleBar.addControl(titleText);
    var wBP = new BABYLON.GUI.StackPanel(); wBP.isVertical = false;
    wBP.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_RIGHT;
    wBP.width = "90px"; wBP.height = "28px"; titleBar.addControl(wBP);
    ["—", "☐", "✕"].forEach(function (sym, idx) {
        var b = BABYLON.GUI.Button.CreateSimpleButton("wb" + idx, sym);
        b.width = "30px"; b.height = "28px"; b.color = "white"; b.fontSize = 11;
        b.thickness = 0; b.background = "transparent";
        b.pointerEnterAnimation = function () { b.background = idx === 2 ? "#E81123" : "rgba(255,255,255,0.15)"; };
        b.pointerOutAnimation = function () { b.background = "transparent"; };
        wBP.addControl(b);
    });

    // --- Ribbon ---
    var ribbonC = new BABYLON.GUI.Rectangle(); ribbonC.background = "#F3F3F3"; ribbonC.thickness = 0;
    root.addControl(ribbonC, 1, 0);
    var tabP = new BABYLON.GUI.StackPanel(); tabP.isVertical = false;
    tabP.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
    tabP.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
    tabP.height = "22px"; tabP.left = "4px"; ribbonC.addControl(tabP);
    ["File", "Home", "Insert", "Design", "Transitions", "Animations", "Slide Show", "Review", "View"].forEach(function (l, i) {
        var tw = (l.length * 7 + 18) + "px";
        if (i === 0) {
            var r = new BABYLON.GUI.Rectangle(); r.width = tw; r.height = "22px"; r.background = PP; r.thickness = 0;
            var t = new BABYLON.GUI.TextBlock(); t.text = l; t.color = "white"; t.fontSize = 11;
            t.fontFamily = "Segoe UI,sans-serif"; r.addControl(t); tabP.addControl(r);
        } else {
            var t = new BABYLON.GUI.TextBlock(); t.text = l; t.color = i === 1 ? PP : "#555";
            t.fontSize = 11; t.fontFamily = "Segoe UI,sans-serif"; t.fontWeight = i === 1 ? "bold" : "normal";
            t.width = tw; t.height = "22px"; tabP.addControl(t);
        }
    });
    // Toolbar
    var toolRow = new BABYLON.GUI.StackPanel(); toolRow.isVertical = false;
    toolRow.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
    toolRow.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_BOTTOM;
    toolRow.height = "36px"; toolRow.left = "6px"; toolRow.top = "-2px"; ribbonC.addControl(toolRow);
    var mkTG = function (items, gl) {
        var g = new BABYLON.GUI.Grid(); g.width = (items.length * 28 + 8) + "px"; g.height = "36px";
        g.addRowDefinition(24, true); g.addRowDefinition(12, true); g.addColumnDefinition(1.0);
        var br = new BABYLON.GUI.StackPanel(); br.isVertical = false;
        br.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
        g.addControl(br, 0, 0);
        items.forEach(function (it) {
            var b = BABYLON.GUI.Button.CreateSimpleButton("t_" + it.l, it.i);
            b.width = "26px"; b.height = "22px"; b.fontSize = it.fs || 12;
            b.color = "#444"; b.thickness = 0; b.background = "transparent";
            b.fontWeight = it.b ? "bold" : "normal";
            b.pointerEnterAnimation = function () { b.background = "#FDE8E0"; };
            b.pointerOutAnimation = function () { b.background = "transparent"; };
            br.addControl(b);
        });
        var lb = new BABYLON.GUI.TextBlock(); lb.text = gl; lb.color = "#999"; lb.fontSize = 8;
        lb.fontFamily = "Segoe UI,sans-serif"; g.addControl(lb, 1, 0); return g;
    };
    var mkSep = function () {
        var s = new BABYLON.GUI.Rectangle(); s.width = "1px"; s.height = "32px";
        s.background = "#D4D4D4"; s.thickness = 0; return s;
    };
    toolRow.addControl(mkTG([{ i: "📋", l: "paste", fs: 14 }, { i: "✂", l: "cut", fs: 13 }, { i: "📄", l: "copy", fs: 11 }], "Clipboard"));
    toolRow.addControl(mkSep());
    toolRow.addControl(mkTG([{ i: "B", l: "bold", b: true }, { i: "I", l: "italic" }, { i: "U", l: "ul" }, { i: "A▼", l: "fc", fs: 10 }], "Font"));
    toolRow.addControl(mkSep());
    toolRow.addControl(mkTG([{ i: "□", l: "shape", fs: 14 }, { i: "⊕", l: "icons", fs: 13 }, { i: "▶", l: "media", fs: 12 }], "Insert"));
    var rBdr = new BABYLON.GUI.Rectangle(); rBdr.height = "1px"; rBdr.background = "#D4D4D4"; rBdr.thickness = 0;
    rBdr.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_BOTTOM; ribbonC.addControl(rBdr);

    // --- Main Content ---
    var mainArea = new BABYLON.GUI.Grid(); mainArea.background = "#E0E0E0";
    root.addControl(mainArea, 2, 0);
    mainArea.addColumnDefinition(130, true); mainArea.addColumnDefinition(1.0);
    mainArea.addRowDefinition(1.0); mainArea.addRowDefinition(70, true);

    // Thumbnail panel
    var slidePanel = new BABYLON.GUI.Rectangle(); slidePanel.background = "#F0F0F0"; slidePanel.thickness = 0;
    mainArea.addControl(slidePanel, 0, 0);
    var pBdr = new BABYLON.GUI.Rectangle(); pBdr.width = "1px"; pBdr.background = "#D0D0D0"; pBdr.thickness = 0;
    pBdr.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_RIGHT; slidePanel.addControl(pBdr);
    var thumbC = new BABYLON.GUI.StackPanel(); thumbC.isVertical = true;
    thumbC.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
    thumbC.paddingTop = "8px"; thumbC.width = "126px"; slidePanel.addControl(thumbC);

    var thumbRects = [];
    var buildThumbnails = function () {
        if (scene.isDisposed) return;
        thumbC.clearControls(); thumbRects = [];
        var TW = 96, TH = 54; // thumb pixel dimensions
        slides.forEach(function (slide, idx) {
            var row = new BABYLON.GUI.StackPanel(); row.isVertical = false;
            row.height = (TH + 4) + "px"; row.width = "120px"; row.paddingBottom = "4px";
            var nt = new BABYLON.GUI.TextBlock(); nt.text = (idx + 1).toString();
            nt.color = "#888"; nt.fontSize = 9; nt.fontFamily = "Segoe UI,sans-serif";
            nt.width = "18px"; nt.height = TH + "px"; row.addControl(nt);

            var th = new BABYLON.GUI.Rectangle(); th.width = TW + "px"; th.height = TH + "px";
            th.background = slide.bg; th.thickness = idx === currentSlide ? 2 : 1;
            th.color = idx === currentSlide ? PP : "#CCC";
            th.cornerRadius = 1; th.shadowColor = "rgba(0,0,0,0.1)"; th.shadowBlur = 3;
            th.clipChildren = true;
            thumbRects.push(th);

            // Background image + tint
            if (slide.bgImage) {
                var ti = new BABYLON.GUI.Image("tbg_" + idx, slide.bgImage);
                ti.stretch = BABYLON.GUI.Image.STRETCH_FILL; th.addControl(ti);
                if (slide.bgTint) {
                    var tTint = new BABYLON.GUI.Rectangle("tTint_" + idx);
                    tTint.width = "100%"; tTint.height = "100%"; tTint.thickness = 0;
                    if (slide.bgTint.type === "duotone") { tTint.background = slide.bgTint.dark; tTint.alpha = 0.55; }
                    else if (slide.bgTint.type === "artEffect") { tTint.background = slide.bgTint.color; tTint.alpha = 0.40; }
                    else if (slide.bgTint.type === "tint") { tTint.background = slide.bgTint.color; tTint.alpha = 0.5; }
                    else if (slide.bgTint.type === "alpha") { tTint.background = "#000000"; tTint.alpha = 1.0 - slide.bgTint.amt; }
                    th.addControl(tTint);
                }
            }

            // Render mini elements
            slide.elements.forEach(function (el) {
                if (el.type === "shape" && el.shape !== "line" && el.fillColor && el.fillColor !== "transparent") {
                    var sr = new BABYLON.GUI.Rectangle();
                    sr.width = (el.w * TW) + "px"; sr.height = (el.h * TH) + "px";
                    sr.left = (el.x * TW) + "px"; sr.top = (el.y * TH) + "px";
                    sr.background = el.fillColor; sr.thickness = 0;
                    sr.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                    sr.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                    th.addControl(sr);
                }
                if (el.type === "image" && el.dataUrl) {
                    var ic = new BABYLON.GUI.Rectangle();
                    ic.width = (el.w * TW) + "px"; ic.height = (el.h * TH) + "px";
                    ic.left = (el.x * TW) + "px"; ic.top = (el.y * TH) + "px";
                    ic.thickness = 0;
                    ic.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                    ic.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                    var im = new BABYLON.GUI.Image("ti_" + idx + "_" + Math.random().toString(36).substr(2,4), el.dataUrl);
                    im.stretch = BABYLON.GUI.Image.STRETCH_FILL; ic.addControl(im);
                    th.addControl(ic);
                }
                if (el.type === "text" && el.text.trim()) {
                    var fs = Math.max(3, Math.round(el.fontSize * TW / CANVAS_W));
                    var mi = new BABYLON.GUI.TextBlock();
                    mi.text = el.text; mi.fontSize = fs;
                    mi.fontWeight = el.fontWeight || "normal";
                    mi.color = el.color; mi.fontFamily = "Segoe UI,sans-serif";
                    mi.left = (el.x * TW) + "px"; mi.top = (el.y * TH) + "px";
                    if (el.w) mi.width = (el.w * TW) + "px";
                    mi.textHorizontalAlignment = el.align === "center" ? BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_CENTER :
                        el.align === "right" ? BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_RIGHT :
                        BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                    mi.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                    mi.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                    mi.textWrapping = true;
                    th.addControl(mi);
                }
            });

            th.isPointerBlocker = true;
            (function (j) {
                th.onPointerClickObservable.add(function () {
                    if (scene.isDisposed) return;
                    currentSlide = j; renderSlide(); updateThumbs(); updateNotes(); updateStatus();
                });
            })(idx);
            row.addControl(th); thumbC.addControl(row);
        });
    };
    var updateThumbs = function () {
        thumbRects.forEach(function (t, i) {
            t.thickness = i === currentSlide ? 2 : 1;
            t.color = i === currentSlide ? PP : "#CCC";
        });
    };

    // --- Slide editor area ---
    var edArea = new BABYLON.GUI.Rectangle(); edArea.background = "#E0E0E0"; edArea.thickness = 0;
    mainArea.addControl(edArea, 0, 1);
    var sCanvas = new BABYLON.GUI.Rectangle(); sCanvas.width = CANVAS_W + "px"; sCanvas.height = CANVAS_H + "px";
    sCanvas.background = "#FFF"; sCanvas.thickness = 1; sCanvas.color = "#CCC";
    sCanvas.shadowColor = "rgba(0,0,0,0.2)"; sCanvas.shadowBlur = 10;
    sCanvas.shadowOffsetX = 2; sCanvas.shadowOffsetY = 3; edArea.addControl(sCanvas);
    var sLayer = new BABYLON.GUI.Rectangle(); sLayer.width = "100%"; sLayer.height = "100%";
    sLayer.thickness = 0; sLayer.background = "transparent"; sLayer.clipChildren = true;
    sCanvas.addControl(sLayer);

    // ========================================================================
    // SECTION 12: Slide Renderer
    // ========================================================================
    var ELLIPSE_SHAPES = ["ellipse", "oval", "circle", "pie", "arc", "chord", "donut"];
    var ROUND_RECT_SHAPES = ["roundRect", "snipRoundRect", "snip1Rect", "snip2SameRect", "round1Rect", "round2SameRect"];

    var renderSlide = function () {
        if (scene.isDisposed) return;
        sLayer.clearControls();
        var slide = slides[currentSlide];
        sCanvas.background = slide.bg;
        console.log("[RENDER] === Rendering slide " + (currentSlide+1) + ": " + slide.elements.length + " elements, bg=" + slide.bg + ", bgImage=" + (slide.bgImage?"yes":"no") + " ===");

        // Background image
        if (slide.bgImage) {
            console.log("[RENDER] Background image: " + slide.bgImage.substring(0, 50) + "... (length=" + slide.bgImage.length + ")");
            var bgI = new BABYLON.GUI.Image("sBg", slide.bgImage);
            bgI.stretch = BABYLON.GUI.Image.STRETCH_FILL;
            bgI.width = "100%"; bgI.height = "100%"; sLayer.addControl(bgI);

            // Apply duotone/tint/artEffect overlay
            if (slide.bgTint) {
                var tintRect = new BABYLON.GUI.Rectangle("bgTint");
                tintRect.width = "100%"; tintRect.height = "100%"; tintRect.thickness = 0;
                var doApply = true;
                if (slide.bgTint.type === "duotone") {
                    tintRect.background = slide.bgTint.dark;
                    tintRect.alpha = 0.55;
                } else if (slide.bgTint.type === "artEffect") {
                    tintRect.background = slide.bgTint.color;
                    tintRect.alpha = 0.40;
                } else if (slide.bgTint.type === "tint") {
                    tintRect.background = slide.bgTint.color;
                    tintRect.alpha = 0.5;
                } else if (slide.bgTint.type === "alpha") {
                    tintRect.background = "#000000";
                    tintRect.alpha = 1.0 - slide.bgTint.amt;
                } else { doApply = false; }
                if (doApply) {
                    sLayer.addControl(tintRect);
                    console.log("[RENDER] Applied bgTint overlay: " + JSON.stringify(slide.bgTint));
                }
            }
        }

        slide.elements.forEach(function (el) {
            if (el.type === "shape") {
                if (el.shape === "line") {
                    var x1 = (el.x1 || 0) * CANVAS_W, y1 = (el.y1 || 0) * CANVAS_H;
                    var x2 = (el.x2 || 0) * CANVAS_W, y2 = (el.y2 || 0) * CANVAS_H;
                    var dx = x2 - x1, dy = y2 - y1;
                    var len = Math.sqrt(dx * dx + dy * dy), ang = Math.atan2(dy, dx);
                    var le = new BABYLON.GUI.Rectangle();
                    le.width = len + "px"; le.height = (el.thickness || 2) + "px";
                    le.background = el.color || "#000"; le.thickness = 0;
                    le.left = ((x1 + x2) / 2 - CANVAS_W / 2) + "px";
                    le.top = ((y1 + y2) / 2 - CANVAS_H / 2) + "px";
                    le.rotation = ang; sLayer.addControl(le);
                } else if (ELLIPSE_SHAPES.indexOf(el.shape) >= 0) {
                    var ell = new BABYLON.GUI.Ellipse();
                    ell.left = (el.x * CANVAS_W) + "px"; ell.top = (el.y * CANVAS_H) + "px";
                    ell.width = (el.w * CANVAS_W) + "px"; ell.height = (el.h * CANVAS_H) + "px";
                    ell.background = el.fillColor || "transparent";
                    ell.thickness = el.borderWidth || 0; ell.color = el.borderColor || "transparent";
                    ell.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                    ell.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                    if (el.rotation) ell.rotation = el.rotation * Math.PI / 180;
                    sLayer.addControl(ell);
                } else {
                    var rect = new BABYLON.GUI.Rectangle();
                    rect.left = (el.x * CANVAS_W) + "px"; rect.top = (el.y * CANVAS_H) + "px";
                    rect.width = (el.w * CANVAS_W) + "px"; rect.height = (el.h * CANVAS_H) + "px";
                    rect.background = el.fillColor || "transparent";
                    rect.thickness = el.borderWidth || 0; rect.color = el.borderColor || "transparent";
                    if (ROUND_RECT_SHAPES.indexOf(el.shape) >= 0)
                        rect.cornerRadius = Math.min(el.w * CANVAS_W, el.h * CANVAS_H) * 0.15;
                    rect.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                    rect.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                    if (el.rotation) rect.rotation = el.rotation * Math.PI / 180;
                    sLayer.addControl(rect);
                }
            } else if (el.type === "image") {
                var ir = new BABYLON.GUI.Rectangle();
                ir.left = (el.x * CANVAS_W) + "px"; ir.top = (el.y * CANVAS_H) + "px";
                ir.width = (el.w * CANVAS_W) + "px"; ir.height = (el.h * CANVAS_H) + "px";
                ir.thickness = 0;
                ir.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                ir.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                var img = new BABYLON.GUI.Image("i_" + Math.random(), el.dataUrl);
                img.stretch = BABYLON.GUI.Image.STRETCH_FILL;
                ir.addControl(img); sLayer.addControl(ir);
            } else if (el.type === "text") {
                console.log("[RENDER] text: '" + el.text.substring(0,30) + "' left=" + (el.x*CANVAS_W).toFixed(1) + " top=" + (el.y*CANVAS_H).toFixed(1) + " w=" + (el.w?el.w*CANVAS_W:"-").toString() + " fs=" + el.fontSize + " color=" + el.color);
                var tb = new BABYLON.GUI.TextBlock();
                // Insert zero-width spaces between CJK characters for word wrapping
                var displayText = el.text.replace(/([\u3000-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF])/g, "\u200B$1");
                var renderFS = Math.round(el.fontSize * 0.75);
                tb.text = displayText; tb.fontSize = renderFS;
                tb.fontWeight = el.fontWeight || "normal"; tb.fontStyle = el.fontStyle || "normal";
                tb.color = el.color; tb.fontFamily = "Segoe UI, Calibri, sans-serif";
                tb.textWrapping = true; tb.resizeToFit = false;
                if (el.w && el.w > 0) {
                    var containerW = el.w * CANVAS_W;
                    tb.width = containerW + "px";
                    // Estimate text width and shrink font if single-line text overflows
                    var hasCJK = /[\u3000-\u9FFF\uF900-\uFAFF]/.test(el.text);
                    var charW = hasCJK ? 1.0 : 0.55;
                    var estTextW = el.text.length * renderFS * charW;
                    var fs = renderFS;
                    if (fs >= 12 && estTextW > containerW * 2) {
                        fs = Math.max(8, Math.floor(containerW * 2 / (el.text.length * charW)));
                        tb.fontSize = fs;
                    } else { fs = renderFS; }
                    var estLines = Math.ceil(el.text.length * fs * charW / containerW) || 1;
                    var estHeight = Math.max(estLines * fs * 1.4, fs * 1.5);
                    tb.height = estHeight + "px";
                }
                if (el.align === "center") tb.textHorizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_CENTER;
                else if (el.align === "right") tb.textHorizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_RIGHT;
                else tb.textHorizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                tb.left = (el.x * CANVAS_W) + "px"; tb.top = (el.y * CANVAS_H) + "px";
                tb.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
                tb.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
                if (el.rotation) tb.rotation = el.rotation * Math.PI / 180;
                sLayer.addControl(tb);
            }
        });
    };

    // --- Notes pane ---
    var nA = new BABYLON.GUI.Rectangle(); nA.background = "#FAFAFA"; nA.thickness = 0; mainArea.addControl(nA, 1, 0);
    var nR = new BABYLON.GUI.Rectangle(); nR.background = "#FAFAFA"; nR.thickness = 0; mainArea.addControl(nR, 1, 1);
    var nBdr = new BABYLON.GUI.Rectangle(); nBdr.height = "1px"; nBdr.background = "#D0D0D0"; nBdr.thickness = 0;
    nBdr.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP; nR.addControl(nBdr);
    var nLbl = new BABYLON.GUI.TextBlock(); nLbl.text = "Notes"; nLbl.fontSize = 9; nLbl.color = "#999";
    nLbl.fontFamily = "Segoe UI,sans-serif"; nLbl.textHorizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
    nLbl.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
    nLbl.top = "4px"; nLbl.left = "10px"; nR.addControl(nLbl);
    var notesText = new BABYLON.GUI.TextBlock("nT"); notesText.text = ""; notesText.fontSize = 10;
    notesText.color = "#555"; notesText.fontFamily = "Segoe UI,sans-serif";
    notesText.textHorizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
    notesText.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
    notesText.textWrapping = true; notesText.top = "18px"; notesText.left = "10px"; notesText.width = "95%";
    nR.addControl(notesText);
    var updateNotes = function () {
        var s = slides[currentSlide];
        notesText.text = s.notes || ""; notesText.color = s.notes ? "#555" : "#BBB";
    };

    // --- Status bar ---
    var stBar = new BABYLON.GUI.Grid(); stBar.background = PP;
    stBar.addRowDefinition(1.0); stBar.addColumnDefinition(0.4); stBar.addColumnDefinition(0.6);
    root.addControl(stBar, 3, 0);
    var stLeft = new BABYLON.GUI.TextBlock("stL"); stLeft.color = "white"; stLeft.fontSize = 10;
    stLeft.fontFamily = "Segoe UI,sans-serif";
    stLeft.textHorizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_LEFT;
    stLeft.paddingLeft = "10px"; stBar.addControl(stLeft, 0, 0);
    var stRight = new BABYLON.GUI.StackPanel(); stRight.isVertical = false;
    stRight.horizontalAlignment = BABYLON.GUI.Control.HORIZONTAL_ALIGNMENT_RIGHT;
    stRight.right = "8px"; stBar.addControl(stRight, 0, 1);
    ["☰", "⊞", "📖", "▶"].forEach(function (ic, i) {
        var v = new BABYLON.GUI.TextBlock(); v.text = ic;
        v.color = i === 0 ? "white" : "rgba(255,255,255,0.5)";
        v.fontSize = i === 3 ? 10 : 11; v.width = "22px"; v.height = "20px"; stRight.addControl(v);
    });
    var zP = new BABYLON.GUI.TextBlock(); zP.text = "  68%"; zP.color = "white"; zP.fontSize = 10;
    zP.fontFamily = "Segoe UI,sans-serif"; zP.width = "40px"; stRight.addControl(zP);
    var updateStatus = function () { stLeft.text = "Slide " + (currentSlide + 1) + " of " + slides.length; };

    // ========================================================================
    // SECTION 13: Keyboard Navigation
    // ========================================================================
    scene.onKeyboardObservable.add(function (kb) {
        if (kb.type !== BABYLON.KeyboardEventTypes.KEYDOWN) return;
        var k = kb.event.key, ch = false;
        if (k === "ArrowRight" || k === "ArrowDown" || k === "PageDown") {
            if (currentSlide < slides.length - 1) { currentSlide++; ch = true; }
        } else if (k === "ArrowLeft" || k === "ArrowUp" || k === "PageUp") {
            if (currentSlide > 0) { currentSlide--; ch = true; }
        } else if (k === "Home") { currentSlide = 0; ch = true; }
        else if (k === "End") { currentSlide = slides.length - 1; ch = true; }
        if (ch) { kb.event.preventDefault(); renderSlide(); updateThumbs(); updateNotes(); updateStatus(); }
    });

    // ========================================================================
    // SECTION 14: Drag & Drop Handler (with Playground re-run safety)
    // ========================================================================
    window.__pptxGen = (window.__pptxGen || 0) + 1;
    var myGen = window.__pptxGen;

    var dropOverlay = document.createElement("div");
    dropOverlay.style.cssText = "position:fixed;top:0;left:0;width:100%;height:100%;" +
        "background:rgba(208,68,35,0.3);z-index:9999;pointer-events:none;" +
        "font-size:32px;color:white;display:none;align-items:center;justify-content:center;" +
        "font-family:Segoe UI,sans-serif;";
    dropOverlay.textContent = "Drop .pptx here";
    document.body.appendChild(dropOverlay);

    var dragCounter = 0;
    var onDragEnter = function (e) {
        e.preventDefault(); if (myGen !== window.__pptxGen) return;
        dragCounter++; dropOverlay.style.display = "flex";
    };
    var onDragLeave = function (e) {
        e.preventDefault(); if (myGen !== window.__pptxGen) return;
        dragCounter--; if (dragCounter <= 0) { dragCounter = 0; dropOverlay.style.display = "none"; }
    };
    var onDragOver = function (e) { e.preventDefault(); };
    var onDrop = async function (e) {
        e.preventDefault(); dragCounter = 0; dropOverlay.style.display = "none";
        if (myGen !== window.__pptxGen || scene.isDisposed) return;
        var files = e.dataTransfer.files; if (files.length === 0) return;
        var file = files[0];
        if (!file.name.toLowerCase().endsWith(".pptx")) { alert("Please drop a .pptx file"); return; }

        titleText.text = "Loading: " + file.name + "...";
        try {
            var ab = await file.arrayBuffer();
            if (scene.isDisposed || myGen !== window.__pptxGen) return;
            var ns = await parsePptx(ab);
            if (scene.isDisposed || myGen !== window.__pptxGen) return;
            if (ns.length > 0) {
                slides = ns; currentSlide = 0;
                buildThumbnails(); renderSlide(); updateThumbs(); updateNotes(); updateStatus();
                titleText.text = file.name.replace(".pptx", "") + " - PowerPoint";
            }
        } catch (err) {
            console.error("PPTX parse error:", err);
            if (!scene.isDisposed) titleText.text = "Error: " + err.message;
        }
    };

    document.addEventListener("dragenter", onDragEnter);
    document.addEventListener("dragleave", onDragLeave);
    document.addEventListener("dragover", onDragOver);
    document.addEventListener("drop", onDrop);

    // Cleanup on scene dispose (prevents stale listeners on Playground re-run)
    scene.onDisposeObservable.add(function () {
        document.removeEventListener("dragenter", onDragEnter);
        document.removeEventListener("dragleave", onDragLeave);
        document.removeEventListener("dragover", onDragOver);
        document.removeEventListener("drop", onDrop);
        if (dropOverlay.parentNode) dropOverlay.parentNode.removeChild(dropOverlay);
    });

    // ========================================================================
    // SECTION 15: Initial Render
    // ========================================================================
    buildThumbnails(); renderSlide(); updateNotes(); updateStatus();

    return scene;
};

export default createScene;
