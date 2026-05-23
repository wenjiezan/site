import * as THREE from 'three';
import { OrbitControls } from 'three/addons/controls/OrbitControls.js';
import { presets } from './presets.js';
import { TransmittancePass } from './passes/transmittance-pass.js';
import { SkyViewPass } from './passes/sky-view-pass.js';
import { AerialPerspectivePass } from './passes/aerial-perspective-pass.js';
import { CompositionPass } from './passes/composition-pass.js';
import { setupGUI } from './gui.js';

export function createAtmosphere(renderer, scene, camera, options = {}) {
    const params = { ...presets.Earth, ...options };
    
    // Prepare uniforms for shaders (Three.js style)
    const uniformsParams = {
        'uParams.planetRadius': { value: params.planetRadius },
        'uParams.atmosphereThickness': { value: params.atmosphereThickness },
        'uParams.rayleighScattering': { value: new THREE.Vector3(...params.rayleighScattering) },
        'uParams.rayleighScaleHeight': { value: params.rayleighScaleHeight },
        'uParams.mieScattering': { value: new THREE.Vector3(...params.mieScattering) },
        'uParams.mieExtinction': { value: new THREE.Vector3(...params.mieExtinction) },
        'uParams.mieAnisotropy': { value: params.mieAnisotropy },
        'uParams.mieScaleHeight': { value: params.mieScaleHeight },
        'uParams.ozoneAbsorption': { value: new THREE.Vector3(...params.ozoneAbsorption) },
        'uParams.ozoneCenterAltitude': { value: params.ozoneCenterAltitude },
        'uParams.ozoneThickness': { value: params.ozoneThickness },
        'uParams.groundAlbedo': { value: new THREE.Vector3(...params.groundAlbedo) }
    };

    const transmittancePass = new TransmittancePass(renderer, uniformsParams);
    const skyViewPass = new SkyViewPass(renderer, uniformsParams, transmittancePass.texture);
    const aerialPerspectivePass = new AerialPerspectivePass(renderer, uniformsParams, transmittancePass.texture);
    const compositionPass = new CompositionPass(
        renderer, 
        uniformsParams, 
        transmittancePass.texture, 
        skyViewPass.texture, 
        aerialPerspectivePass.texture
    );

    let transmittanceDirty = true;

    return {
        update(dt, sunDirection) {
            if (transmittanceDirty) {
                transmittancePass.render();
                transmittanceDirty = false;
            }
            skyViewPass.render(sunDirection, camera.position);
            aerialPerspectivePass.render(sunDirection, camera);
        },
        
        setParameters(partial) {
            Object.keys(partial).forEach(key => {
                const uniformKey = `uParams.${key}`;
                if (uniformsParams[uniformKey]) {
                    if (uniformsParams[uniformKey].value instanceof THREE.Vector3) {
                        uniformsParams[uniformKey].value.set(...partial[key]);
                    } else {
                        uniformsParams[uniformKey].value = partial[key];
                    }
                }
            });
            transmittanceDirty = true;
        },
        
        render(sceneColor, sceneDepth, target = null) {
            compositionPass.render(this.sunDirection || new THREE.Vector3(0, 1, 0), camera, sceneColor, sceneDepth, target);
        },
        
        dispose() {
            // TODO: cleanup
        },

        get params() { 
            // Return a proxy or a helper to get/set params easily for GUI
            return new Proxy(uniformsParams, {
                get(target, prop) {
                    if (prop in target) return target[prop].value;
                    if (`uParams.${prop}` in target) return target[`uParams.${prop}`].value;
                    return undefined;
                },
                set(target, prop, value) {
                    const key = prop.startsWith('uParams.') ? prop : `uParams.${prop}`;
                    if (key in target) {
                        target[key].value = value;
                        return true;
                    }
                    return false;
                }
            });
        }
    };
}

export function bootstrap(container) {
    const width = container.clientWidth || 800;
    const height = container.clientHeight || 450;

    console.log("Atmosphere bootstrap started", { width, height });

    const renderer = new THREE.WebGLRenderer({ antialias: true });
    renderer.setSize(width, height);
    renderer.setPixelRatio(window.devicePixelRatio);
    renderer.setClearColor(0x000000, 1.0);
    container.appendChild(renderer.domElement);

    const scene = new THREE.Scene();
    scene.background = new THREE.Color(0x050505);
    const camera = new THREE.PerspectiveCamera(60, width / height, 0.1, 2000000);
    camera.position.set(0, 6360 + 1, 0);

    const controls = new OrbitControls(camera, renderer.domElement);
    controls.enableDamping = true;
    controls.minDistance = 0.1;
    controls.maxDistance = 500;
    controls.target.set(0, 6360 + 1, -50);
    controls.update();

    // Planet
    const planetGeometry = new THREE.SphereGeometry(6360, 128, 128);
    const planetMaterial = new THREE.MeshStandardMaterial({ 
        color: 0x444444,
        roughness: 0.8,
        metalness: 0.1
    });
    const planet = new THREE.Mesh(planetGeometry, planetMaterial);
    scene.add(planet);

    const ambientLight = new THREE.AmbientLight(0xffffff, 0.1);
    scene.add(ambientLight);

    const sunLight = new THREE.DirectionalLight(0xffffff, 2);
    scene.add(sunLight);
    
    const sunParams = {
        azimuth: 180,
        elevation: 30
    };
    const sunDirection = new THREE.Vector3();

    function updateSun() {
        const phi = (90 - sunParams.elevation) * Math.PI / 180;
        const theta = sunParams.azimuth * Math.PI / 180;
        sunDirection.set(
            Math.sin(phi) * Math.sin(theta),
            Math.cos(phi),
            Math.sin(phi) * Math.cos(theta)
        );
        sunLight.position.copy(sunDirection);
    }
    updateSun();

    // Scene render targets
    let sceneTarget;
    try {
        sceneTarget = new THREE.WebGLRenderTarget(width, height, {
            depthTexture: new THREE.DepthTexture(),
            type: THREE.HalfFloatType
        });
    } catch (e) {
        console.error("Failed to create HalfFloat render target, falling back to UnsignedByteType", e);
        sceneTarget = new THREE.WebGLRenderTarget(width, height, {
            depthTexture: new THREE.DepthTexture()
        });
    }

    const atmosphere = createAtmosphere(renderer, scene, camera);
    atmosphere.sunDirection = sunDirection;

    setupGUI(atmosphere, sunParams, updateSun);

    function animate() {
        requestAnimationFrame(animate);
        try {
            controls.update();
            
            // Update atmosphere (LUTs)
            atmosphere.update(0, sunDirection);
            
            // Render scene to target
            renderer.setRenderTarget(sceneTarget);
            renderer.render(scene, camera);
            
            // Composite
            atmosphere.render(sceneTarget.texture, sceneTarget.depthTexture, null);
        } catch (e) {
            console.error("Render loop error", e);
        }
    }

    animate();

    window.addEventListener('resize', () => {
        const w = container.clientWidth;
        const h = container.clientHeight;
        if (w && h) {
            renderer.setSize(w, h);
            camera.aspect = w / h;
            camera.updateProjectionMatrix();
            sceneTarget.setSize(w, h);
        }
    });
}
