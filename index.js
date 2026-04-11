// ============================================================================
// index.js - Entry point: createScene, drag & drop, keyboard navigation
// ============================================================================

import { getDefaultSlides } from "./constants.js";
import { setupScene } from "./scene-setup.js";
import { buildGuiFrame } from "./gui-frame.js";
import { parsePptx } from "./pptx-parser.js";
import { renderSlide, buildThumbnails, updateThumbs, updateNotes, updateStatus } from "./slide-renderer.js";

// Suppress verbose parser/render debug logs in production.
// Enable full logs by setting window.__PPTX_DEBUG__ = true in DevTools.
(function configureDebugLogging() {
    if (typeof window === "undefined" || !window.console || typeof window.console.log !== "function") return;
    if (window.__PPTX_DEBUG__) return;

    var rawLog = window.console.log.bind(window.console);
    var debugPrefixes = /^(\[PPTX\]|\[RENDER\]|\[SP\]|\[TREE\]|\[PIC\]|\[BG\]|\[BLIP\]|\[LAYOUT\]|\[MASTER\]|\[GRPSP\]|\[GF\]|\[SmartArt\]|\[TBL\]|\[CXN\]|\[INIT\/ENGINE\]|\[INIT\/SCENE\]|\[INIT\/UI\]|\[INIT\/INPUT\]|\[INIT\/JSZIP\]|\[INIT\] boot sequence)/;

    window.console.log = function () {
        if (arguments.length > 0 && typeof arguments[0] === "string" && debugPrefixes.test(arguments[0])) {
            return;
        }
        return rawLog.apply(window.console, arguments);
    };
})();

var canvas = document.getElementById("renderCanvas");

function ensurePerfStore() {
    if (typeof window === "undefined") return null;
    if (!window.__PPTX_PERF__ || typeof window.__PPTX_PERF__ !== "object") {
        window.__PPTX_PERF__ = {};
    }
    if (!Array.isArray(window.__PPTX_PERF__.sessions)) window.__PPTX_PERF__.sessions = [];
    if (window.__PPTX_PERF__.lastLoad === undefined) window.__PPTX_PERF__.lastLoad = null;
    if (window.__PPTX_PERF__.lastRender === undefined) window.__PPTX_PERF__.lastRender = null;
    if (window.__PPTX_PERF__.lastThumbBuild === undefined) window.__PPTX_PERF__.lastThumbBuild = null;
    if (window.__PPTX_PERF__.lastParse === undefined) window.__PPTX_PERF__.lastParse = null;
    return window.__PPTX_PERF__;
}

function recordPerfSession(session) {
    var store = ensurePerfStore();
    if (!store) return;
    store.lastLoad = session;
    store.sessions.push(session);
    if (store.sessions.length > 20) store.sessions.shift();
}

var AppError = function(code, userMsg, devMsg) {
    this.code = code;  // e.g., 'CANVAS_NOT_FOUND', 'ENGINE_INIT_FAIL', 'JSZIP_LOAD_FAIL', 'SCENE_BUILD_FAIL'
    this.userMsg = userMsg;  // user-facing message for UI
    this.devMsg = devMsg;     // detailed message for console
};
AppError.prototype = Object.create(Error.prototype);
AppError.prototype.constructor = AppError;

var showErrorOnUI = function(titleTextElement, errObj) {
    if (!titleTextElement) return; // UI not ready
    if (errObj instanceof AppError) {
        titleTextElement.text = "❌ " + errObj.userMsg;
        titleTextElement.color = "#E81123";
    } else if (errObj && errObj.message) {
        titleTextElement.text = "❌ Error: " + errObj.message;
        titleTextElement.color = "#E81123";
    } else {
        titleTextElement.text = "❌ Unknown error occurred";
        titleTextElement.color = "#E81123";
    }
};

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

function resolveJsZipConstructor() {
    if (typeof globalThis === "undefined") return null;
    var candidate = globalThis.JSZip;
    if (typeof candidate === "function") return candidate;
    if (candidate && typeof candidate.default === "function") return candidate.default;
    if (candidate && typeof candidate.JSZip === "function") return candidate.JSZip;
    return null;
}

function ensureJsZipLoaded() {
    // JSZip is loaded statically via <script> in index.html
    // This function validates that it loaded correctly
    var JsZipCtor = resolveJsZipConstructor();
    if (!JsZipCtor) {
        throw new AppError(
            'JSZIP_LOAD_FAIL',
            'File library (JSZip) not available. Check that libs/jszip.min.js is present.',
            'JSZip constructor could not be resolved from globalThis.JSZip'
        );
    }

    // Normalize for downstream code paths that expect a global constructor.
    globalThis.JSZip = JsZipCtor;

    var hasLoadAsync = typeof JsZipCtor.loadAsync === "function";
    if (!hasLoadAsync) {
        var testZip = new JsZipCtor();
        hasLoadAsync = typeof testZip.loadAsync === "function";
    }
    if (!hasLoadAsync) {
        throw new AppError(
            'JSZIP_LOAD_FAIL',
            'File library loaded but is incompatible. Try clearing browser cache.',
            'Resolved JSZip constructor does not expose loadAsync (static or instance)'
        );
    }
    console.log("[INIT/JSZIP] JSZip validated OK");
}

async function createEnginePhase() {
    console.log("[INIT/ENGINE] starting");
    if (!canvas) {
        throw new AppError(
            'CANVAS_NOT_FOUND',
            'Render canvas element not found. Page may be incomplete.',
            'renderCanvas element not found in DOM'
        );
    }

    try {
        engine = createDefaultEngine();
    } catch (e) {
        console.warn("[INIT/ENGINE] createDefaultEngine failed once, retrying", e);
        try {
            engine = createDefaultEngine();
        } catch (e2) {
            throw new AppError(
                'ENGINE_INIT_FAIL',
                'Failed to initialize graphics engine. WebGL may not be supported.',
                'Babylon.js Engine creation failed: ' + (e2.message || e2)
            );
        }
    }

    if (!engine) {
        throw new AppError(
            'ENGINE_NULL',
            'Graphics engine initialization failed.',
            'engine is null after creation'
        );
    }
    window.engine = engine;
    startRenderLoop(engine, canvas);
    console.log("[INIT/ENGINE] done");
    return engine;
}

async function createScenePhase() {
    console.log("[INIT/SCENE] starting");
    try {
        ensureJsZipLoaded();
    } catch (e) {
        if (e instanceof AppError) throw e;
        throw new AppError(
            'JSZIP_LOAD_FAIL',
            'Failed to load file library. Check that libs/jszip.min.js is present.',
            (e && e.message) || String(e)
        );
    }

    try {
        var sceneInstance = new BABYLON.Scene(engine);
        sceneInstance.clearColor = new BABYLON.Color3(0.08, 0.08, 0.15);
        var sceneObjs = setupScene(sceneInstance, canvas);
        console.log("[INIT/SCENE] done");
        return { sceneInstance: sceneInstance, sceneObjs: sceneObjs };
    } catch (e) {
        throw new AppError(
            'SCENE_BUILD_FAIL',
            'Failed to build 3D scene. Your graphics drivers may need updating.',
            (e && e.message) || String(e)
        );
    }
}

function createUiPhase(sceneInstance, sceneObjs) {
    console.log("[INIT/UI] starting");
    try {
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
    } catch (e) {
        throw new AppError(
            'UI_BUILD_FAIL',
            'Failed to build user interface.',
            (e && e.message) || String(e)
        );
    }
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
            var perfSession = {
                fileName: file.name,
                startedAt: performance.now()
            };
            var ab = await file.arrayBuffer();
            perfSession.arrayBufferMs = performance.now() - perfSession.startedAt;
            if (sceneInstance.isDisposed || myGen !== window.__pptxGen) return;
            var parseStart = performance.now();
            // Phase 1 callback: called as soon as slide structure is parsed (no images yet).
            // This lets the user see text/shapes immediately while images load in background.
            var structureReadyCalled = false;
            var ns = await parsePptx(ab,
                function onStructureReady(partialSlides) {
                    if (sceneInstance.isDisposed || myGen !== window.__pptxGen) return;
                    if (partialSlides.length === 0) return;
                    structureReadyCalled = true;
                    app.slides = partialSlides; app.currentSlide = 0;
                    buildThumbnails(app);
                    renderSlide(app);
                    updateThumbs(app); updateNotes(app); updateStatus(app);
                    app.gui.titleText.text = file.name.replace(".pptx", "") + " - Loading images...";
                },
                function onSlideImagesReady(idx) {
                    if (sceneInstance.isDisposed || myGen !== window.__pptxGen) return;
                    if (idx === app.currentSlide) renderSlide(app);
                },
                function onAllImagesReady() {
                    if (sceneInstance.isDisposed || myGen !== window.__pptxGen) return;
                    if (!app.slides || app.slides.length === 0) return;
                    buildThumbnails(app);
                    renderSlide(app);
                    updateThumbs(app); updateNotes(app); updateStatus(app);
                    app.gui.titleText.text = file.name.replace(".pptx", "") + " - PowerPoint";
                }
            );
            perfSession.parseMs = performance.now() - parseStart;
            if (sceneInstance.isDisposed || myGen !== window.__pptxGen) return;
            if (ns.length > 0) {
                if (!structureReadyCalled) {
                    app.slides = ns; app.currentSlide = 0;
                    buildThumbnails(app);
                    renderSlide(app);
                    updateThumbs(app); updateNotes(app); updateStatus(app);
                }
                perfSession.totalMs = performance.now() - perfSession.startedAt;
                recordPerfSession(perfSession);
                console.info(
                    "[PERF] load " + file.name +
                    " total=" + perfSession.totalMs.toFixed(1) + "ms" +
                    " arrayBuffer=" + perfSession.arrayBufferMs.toFixed(1) + "ms" +
                    " parse(phase1)=" + perfSession.parseMs.toFixed(1) + "ms"
                );
                if (!structureReadyCalled) app.gui.titleText.text = file.name.replace(".pptx", "") + " - Loading images...";
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
    var app = null;
    try {
        await createEnginePhase();
        var sceneResult = await createScenePhase();
        var sceneInstance = sceneResult.sceneInstance;
        var sceneObjs = sceneResult.sceneObjs;
        app = createUiPhase(sceneInstance, sceneObjs);
        registerInputPhase(sceneInstance, app);

        scene = Promise.resolve(sceneInstance);
        window.scene = scene;
        sceneToRender = sceneInstance;
        console.log("[INIT] boot sequence complete");
        return scene;
    } catch (err) {
        console.error("[INIT] failed", err);
        if (err instanceof AppError) {
            console.error("[INIT] error code:", err.code);
            console.error("[INIT] dev message:", err.devMsg);
            // Try to display user-friendly message on UI if available
            if (app && app.gui && app.gui.titleText) {
                showErrorOnUI(app.gui.titleText, err);
            }
        }
        throw err;
    }
}

window.initFunction = runAppInit;

runAppInit().catch(function (err) {
    console.error("[INIT] unhandled error during boot", err);
});

// Resize
window.addEventListener("resize", function () {
    if (engine) engine.resize();
});

var createScene = runAppInit;
export default createScene;
