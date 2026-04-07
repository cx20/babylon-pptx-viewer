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

async function ensureJsZipLoaded() {
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
}

async function createEnginePhase() {
    console.log("[INIT/ENGINE] starting");
    if (!canvas) throw new Error("renderCanvas not found");

    try {
        engine = createDefaultEngine();
    } catch (e) {
        console.warn("[INIT/ENGINE] createDefaultEngine failed once, retrying", e);
        engine = createDefaultEngine();
    }

    if (!engine) throw new Error("engine should not be null");
    window.engine = engine;
    startRenderLoop(engine, canvas);
    console.log("[INIT/ENGINE] done");
    return engine;
}

async function createScenePhase() {
    console.log("[INIT/SCENE] starting");
    await ensureJsZipLoaded();

    var sceneInstance = new BABYLON.Scene(engine);
    sceneInstance.clearColor = new BABYLON.Color3(0.08, 0.08, 0.15);
    var sceneObjs = setupScene(sceneInstance, canvas);

    console.log("[INIT/SCENE] done");
    return { sceneInstance: sceneInstance, sceneObjs: sceneObjs };
}

function createUiPhase(sceneInstance, sceneObjs) {
    console.log("[INIT/UI] starting");
    var gui = buildGuiFrame(sceneObjs.screenPlane);
    var app = {
        scene: sceneInstance,
        gui: gui,
        slides: getDefaultSlides(),
        currentSlide: 0,
        thumbRects: []
    };

    buildThumbnails(app); renderSlide(app); updateNotes(app); updateStatus(app);
    console.log("[INIT/UI] done");
    return app;
}

function registerInputPhase(sceneInstance, app) {
    console.log("[INIT/INPUT] starting");

    sceneInstance.onKeyboardObservable.add(function (kb) {
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
        if (myGen !== window.__pptxGen || sceneInstance.isDisposed) return;
        var files = e.dataTransfer.files; if (files.length === 0) return;
        var file = files[0];
        if (!file.name.toLowerCase().endsWith(".pptx")) { alert("Please drop a .pptx file"); return; }

        app.gui.titleText.text = "Loading: " + file.name + "...";
        try {
            var ab = await file.arrayBuffer();
            if (sceneInstance.isDisposed || myGen !== window.__pptxGen) return;
            var ns = await parsePptx(ab);
            if (sceneInstance.isDisposed || myGen !== window.__pptxGen) return;
            if (ns.length > 0) {
                app.slides = ns; app.currentSlide = 0;
                buildThumbnails(app); renderSlide(app); updateThumbs(app); updateNotes(app); updateStatus(app);
                app.gui.titleText.text = file.name.replace(".pptx", "") + " - PowerPoint";
            }
        } catch (err) {
            console.error("PPTX parse error:", err);
            if (!sceneInstance.isDisposed) app.gui.titleText.text = "Error: " + err.message;
        }
    };

    document.addEventListener("dragenter", onDragEnter);
    document.addEventListener("dragleave", onDragLeave);
    document.addEventListener("dragover", onDragOver);
    document.addEventListener("drop", onDrop);

    sceneInstance.onDisposeObservable.add(function () {
        document.removeEventListener("dragenter", onDragEnter);
        document.removeEventListener("dragleave", onDragLeave);
        document.removeEventListener("dragover", onDragOver);
        document.removeEventListener("drop", onDrop);
        if (dropOverlay.parentNode) dropOverlay.parentNode.removeChild(dropOverlay);
    });

    console.log("[INIT/INPUT] done");
}

async function runAppInit() {
    console.log("[INIT] boot sequence start");
    await createEnginePhase();
    var sceneResult = await createScenePhase();
    var sceneInstance = sceneResult.sceneInstance;
    var sceneObjs = sceneResult.sceneObjs;
    var app = createUiPhase(sceneInstance, sceneObjs);
    registerInputPhase(sceneInstance, app);

    scene = Promise.resolve(sceneInstance);
    window.scene = scene;
    sceneToRender = sceneInstance;
    console.log("[INIT] boot sequence complete");
    return scene;
}

window.initFunction = runAppInit;

runAppInit().catch(function (err) {
    console.error("[INIT] failed", err);
    throw err;
});

// Resize
window.addEventListener("resize", function () {
    if (engine) engine.resize();
});

var createScene = runAppInit;
export default createScene;
