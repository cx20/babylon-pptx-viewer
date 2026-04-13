// ============================================================================
// gui-frame.js - PowerPoint UI frame (title bar, ribbon, panels, status bar)
// ============================================================================

import { TEX_W, TEX_H, PP, CANVAS_W, CANVAS_H } from "./constants.js";

// Build the full PowerPoint UI and return references to dynamic controls
export function buildGuiFrame(screenPlane) {
    var advTex = BABYLON.GUI.AdvancedDynamicTexture.CreateForMesh(screenPlane, TEX_W, TEX_H);

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
    var thumbScroll = new BABYLON.GUI.ScrollViewer();
    thumbScroll.width = "126px"; thumbScroll.height = "100%";
    thumbScroll.thickness = 0; thumbScroll.background = "transparent";
    thumbScroll.barSize = 8; thumbScroll.wheelPrecision = 30;
    thumbScroll.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
    slidePanel.addControl(thumbScroll);
    var thumbC = new BABYLON.GUI.StackPanel(); thumbC.isVertical = true;
    thumbC.verticalAlignment = BABYLON.GUI.Control.VERTICAL_ALIGNMENT_TOP;
    thumbC.adaptHeightToChildren = true;
    thumbC.paddingTop = "8px"; thumbC.width = "126px"; thumbScroll.addControl(thumbC);

    // Editor area
    var edArea = new BABYLON.GUI.Rectangle(); edArea.background = "#E0E0E0"; edArea.thickness = 0;
    mainArea.addControl(edArea, 0, 1);
    var sCanvas = new BABYLON.GUI.Rectangle(); sCanvas.width = CANVAS_W + "px"; sCanvas.height = CANVAS_H + "px";
    sCanvas.background = "#FFF"; sCanvas.thickness = 1; sCanvas.color = "#CCC";
    sCanvas.shadowColor = "rgba(0,0,0,0.2)"; sCanvas.shadowBlur = 10;
    sCanvas.shadowOffsetX = 2; sCanvas.shadowOffsetY = 3; edArea.addControl(sCanvas);
    var sLayer = new BABYLON.GUI.Rectangle(); sLayer.width = "100%"; sLayer.height = "100%";
    sLayer.thickness = 0; sLayer.background = "transparent"; sLayer.clipChildren = true;
    sCanvas.addControl(sLayer);

    // Notes pane
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

    // Status bar
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

    return {
        titleText: titleText,
        thumbC: thumbC,
        thumbScroll: thumbScroll,
        sCanvas: sCanvas,
        sLayer: sLayer,
        notesText: notesText,
        stLeft: stLeft
    };
}
