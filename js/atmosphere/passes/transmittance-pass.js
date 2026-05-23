import * as THREE from 'three';
import vertexShader from '../shaders/fullscreen.vert.js';
import fragmentShader from '../shaders/transmittance.frag.js';

export class TransmittancePass {
    constructor(renderer, params) {
        this.renderer = renderer;
        this.renderTarget = new THREE.WebGLRenderTarget(256, 64, {
            type: THREE.HalfFloatType,
            format: THREE.RGBAFormat,
            minFilter: THREE.LinearFilter,
            magFilter: THREE.LinearFilter,
            generateMipmaps: false,
            depthBuffer: false
        });

        this.material = new THREE.ShaderMaterial({
            uniforms: {
                ...params
            },
            vertexShader,
            fragmentShader
        });

        this.scene = new THREE.Scene();
        this.camera = new THREE.OrthographicCamera(-1, 1, 1, -1, 0, 1);
        this.quad = new THREE.Mesh(new THREE.PlaneGeometry(2, 2), this.material);
        this.scene.add(this.quad);
    }

    render() {
        const oldTarget = this.renderer.getRenderTarget();
        this.renderer.setRenderTarget(this.renderTarget);
        this.renderer.render(this.scene, this.camera);
        this.renderer.setRenderTarget(oldTarget);
    }

    get texture() {
        return this.renderTarget.texture;
    }
}
