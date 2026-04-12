// ============================================================================
// robot-scene-setup.js - Robot model scene setup with slide display
// ============================================================================

/**
 * Sets up the scene with the Robot.babylon model from babylonjs.com.
 * The robot model contains a "Slide" mesh where presentation slides are displayed.
 * Animations:
 *   - Frames 0-265: Initial robot setup/idle
 *   - Frames 266-300: Page flip "out" animation  
 *   - Frames 301-354: Page flip "in" animation
 * 
 * @param {BABYLON.Scene} scene - The Babylon.js scene
 * @param {HTMLCanvasElement} canvas - The render canvas
 * @returns {Promise<Object>} - Scene objects including slideMesh and importedMeshes
 */
export async function setupRobotScene(scene, canvas) {
    // Camera - positioned to view the robot from slightly above
    var camera = new BABYLON.FreeCamera(
        "Camera",
        new BABYLON.Vector3(-190, 120, -243),
        scene
    );
    camera.setTarget(new BABYLON.Vector3(-189, 120, -243));
    camera.rotation = new BABYLON.Vector3(0.30, 1.31, 0);
    camera.minZ = 0.1;
    camera.speed = 2.5;
    camera.attachControl(canvas, true);

    // Main light
    var light = new BABYLON.HemisphericLight(
        "light",
        new BABYLON.Vector3(0, 1, 0),
        scene
    );
    light.intensity = 1.0;

    console.log("[ROBOT-DEBUG] Loading Robot.babylon model...");
    
    // Import Robot model from Babylon.js assets
    var result = await BABYLON.SceneLoader.ImportMeshAsync(
        "",
        "https://www.babylonjs.com/Scenes/Robot/Assets/",
        "Robot.babylon",
        scene
    );

    var importedMeshes = result.meshes.slice();
    console.log("[ROBOT-DEBUG] Model loaded, mesh count:", importedMeshes.length);
    console.log("[ROBOT-DEBUG] Mesh names:", importedMeshes.map(function(m) { return m.name; }).join(", "));

    // Keep our camera active after import (model may set its own camera)
    scene.activeCamera = camera;

    // Fix LumHalo opacity texture if present
    var lumHalo = scene.getMeshByName("LumHalo");
    if (lumHalo && lumHalo.material && lumHalo.material.opacityTexture) {
        lumHalo.material.opacityTexture.getAlphaFromRGB = true;
    }

    // Get the slide mesh where we display presentations
    var slideMesh = scene.getMeshByName("Slide");
    console.log("[ROBOT-DEBUG] slideMesh found:", slideMesh ? "YES" : "NO");
    if (slideMesh) {
        console.log("[ROBOT-DEBUG] slideMesh.material:", slideMesh.material ? slideMesh.material.name : "NO MATERIAL");
        console.log("[ROBOT-DEBUG] slideMesh.material type:", slideMesh.material ? slideMesh.material.getClassName() : "N/A");
    }

    // Play initial animation (robot setup, frames 0-265)
    for (var m = 0; m < importedMeshes.length; m++) {
        scene.stopAnimation(importedMeshes[m]);
        scene.beginAnimation(importedMeshes[m], 0, 265, false, 1.0);
    }

    return {
        camera: camera,
        slideMesh: slideMesh,
        importedMeshes: importedMeshes
    };
}

/**
 * Animates the robot's page flip:
 * - First plays frames 266-300 (flip out), then callback, then 301-354 (flip in)
 * 
 * @param {BABYLON.Scene} scene - The scene
 * @param {BABYLON.Mesh[]} importedMeshes - All meshes from the robot model
 * @param {Function} onMidFlip - Callback to execute mid-flip (e.g., change slide texture)
 */
export function animatePageFlip(scene, importedMeshes, onMidFlip) {
    if (!importedMeshes || importedMeshes.length === 0) {
        if (onMidFlip) onMidFlip();
        return;
    }

    // Primary mesh animation with callback at midpoint
    if (importedMeshes[0]) {
        scene.beginAnimation(importedMeshes[0], 266, 300, false, 1.0, function () {
            if (onMidFlip) onMidFlip();
            scene.beginAnimation(importedMeshes[0], 301, 354, false, 1.0);
        });
    } else {
        if (onMidFlip) onMidFlip();
    }

    // Secondary meshes animate full range
    for (var i = 1; i < importedMeshes.length; i++) {
        scene.beginAnimation(importedMeshes[i], 266, 354, false, 1.0);
    }
}
