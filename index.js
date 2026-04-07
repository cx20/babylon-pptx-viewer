// ============================================================================
// index.js - Entry point: createScene, drag & drop, keyboard navigation
// ============================================================================

import { getDefaultSlides } from "./constants.js";
import { setupScene } from "./scene-setup.js";
import { buildGuiFrame } from "./gui-frame.js";
import { parsePptx } from "./pptx-parser.js";
import { renderSlide, buildThumbnails, updateThumbs, updateNotes, updateStatus } from "./slide-renderer.js";

var canvas = document.getElementById("renderCanvas");

var startRenderLoop = function (engine, canvas) {
  engine.runRenderLoop(function () {
    if (sceneToRender && sceneToRender.activeCamera) {
      sceneToRender.render();
    }
  });
}

var engine = null;
var scene = null;
var sceneToRender = null;
var createDefaultEngine = function () { return new BABYLON.Engine(canvas, true, { preserveDrawingBuffer: true, stencil: true, disableWebGL2Support: false }); };

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

    // --- Build 3D scene ---
    var sceneObjs = setupScene(scene, canvas);

    // --- Build GUI ---
    var gui = buildGuiFrame(sceneObjs.screenPlane);

    // --- App state (shared across all modules) ---
    var app = {
        scene: scene,
        gui: gui,
        slides: getDefaultSlides(),
        currentSlide: 0,
        thumbRects: []
    };

    // --- Keyboard Navigation ---
    scene.onKeyboardObservable.add(function (kb) {
        if (kb.type !== BABYLON.KeyboardEventTypes.KEYDOWN) return;
        var k = kb.event.key, ch = false;
        if (k === "ArrowRight" || k === "ArrowDown" || k === "PageDown") {
            if (app.currentSlide < app.slides.length - 1) { app.currentSlide++; ch = true; }
        } else if (k === "ArrowLeft" || k === "ArrowUp" || k === "PageUp") {
            if (app.currentSlide > 0) { app.currentSlide--; ch = true; }
        } else if (k === "Home") { app.currentSlide = 0; ch = true; }
        else if (k === "End") { app.currentSlide = app.slides.length - 1; ch = true; }
        if (ch) {
            kb.event.preventDefault();
            renderSlide(app); updateThumbs(app); updateNotes(app); updateStatus(app);
        }
    });

    // --- Drag & Drop Handler ---
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

        gui.titleText.text = "Loading: " + file.name + "...";
        try {
            var ab = await file.arrayBuffer();
            if (scene.isDisposed || myGen !== window.__pptxGen) return;
            var ns = await parsePptx(ab);
            if (scene.isDisposed || myGen !== window.__pptxGen) return;
            if (ns.length > 0) {
                app.slides = ns; app.currentSlide = 0;
                buildThumbnails(app); renderSlide(app); updateThumbs(app); updateNotes(app); updateStatus(app);
                gui.titleText.text = file.name.replace(".pptx", "") + " - PowerPoint";
            }
        } catch (err) {
            console.error("PPTX parse error:", err);
            if (!scene.isDisposed) gui.titleText.text = "Error: " + err.message;
        }
    };

    document.addEventListener("dragenter", onDragEnter);
    document.addEventListener("dragleave", onDragLeave);
    document.addEventListener("dragover", onDragOver);
    document.addEventListener("drop", onDrop);

    scene.onDisposeObservable.add(function () {
        document.removeEventListener("dragenter", onDragEnter);
        document.removeEventListener("dragleave", onDragLeave);
        document.removeEventListener("dragover", onDragOver);
        document.removeEventListener("drop", onDrop);
        if (dropOverlay.parentNode) dropOverlay.parentNode.removeChild(dropOverlay);
    });

    // --- Initial Render ---
    buildThumbnails(app); renderSlide(app); updateNotes(app); updateStatus(app);

    return scene;
};

window.initFunction = async function() {

  var asyncEngineCreation = async function() {
    try {
      return createDefaultEngine();
    } catch(e) {
      console.log("the available createEngine function failed. Creating the default engine instead");
      return createDefaultEngine();
    }
  }

  engine = await asyncEngineCreation();
  window.engine = engine;

  const engineOptions = engine.getCreationOptions?.();
  if (!engineOptions || engineOptions.audioEngine !== false) {

  }
  if (!engine) throw 'engine should not be null.';
  startRenderLoop(engine, canvas);
  scene = createScene();
  window.scene = scene;
};

initFunction().then(() => {
  scene.then(returnedScene => { sceneToRender = returnedScene; });
});

// Resize
window.addEventListener("resize", function () {
  engine.resize();
});

export default createScene;
