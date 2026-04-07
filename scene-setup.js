// ============================================================================
// scene-setup.js - Camera, lights, and PC monitor model
// ============================================================================

export function setupScene(scene, canvas) {
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

    return { screenPlane: screenPlane };
}
