import * as THREE from 'three';
import vertexShader from '../shaders/fullscreen.vert.js';
import fragmentShader from '../shaders/sky-view.frag.js';

export class SkyViewPass {
    constructor(renderer, params, transmittanceTexture) {
        this.renderer = renderer;
        this.renderTarget = new THREE.WebGLRenderTarget(192, 108, {
            type: THREE.HalfFloatType,
            format: THREE.RGBAFormat,
            minFilter: THREE.LinearFilter,
            magFilter: THREE.LinearFilter,
            generateMipmaps: false,
            depthBuffer: false
        });

        this.material = new THREE.ShaderMaterial({
            uniforms: {
                ...params,
                uTransmittanceLUT: { value: transmittanceTexture },
                uSunDirection: { value: new THREE.Vector3(0, 1, 0) },
                uCameraPosition: { value: new THREE.Vector3(0, 0, 0) }
            },
            vertexShader,
            fragmentShader
        });

        this.scene = new THREE.Scene();
        this.camera = new THREE.OrthographicCamera(-1, 1, 1, -1, 0, 1);
        this.quad = new THREE.Mesh(new THREE.PlaneGeometry(2, 2), this.material);
        this.scene.add(this.quad);
    }

    render(sunDirection, cameraPosition) {
        this.material.uniforms.uSunDirection.value.copy(sunDirection);
        this.material.uniforms.uCameraPosition.value.copy(cameraPosition);
        
        const oldTarget = this.renderer.getRenderTarget();
        this.renderer.setRenderTarget(this.renderTarget);
        this.renderer.render(this.scene, this.camera);
        this.renderer.setRenderTarget(oldTarget);
    }

    get texture() {
        return this.renderTarget.texture;
    }
}
