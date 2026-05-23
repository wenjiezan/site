import * as THREE from 'three';
import vertexShader from '../shaders/fullscreen.vert.js';
import fragmentShader from '../shaders/composition.frag.js';

export class CompositionPass {
    constructor(renderer, params, transmittanceTexture, skyViewTexture, aerialPerspectiveTexture) {
        this.renderer = renderer;
        this.material = new THREE.ShaderMaterial({
            uniforms: {
                ...params,
                uTransmittanceLUT: { value: transmittanceTexture },
                uSkyViewLUT: { value: skyViewTexture },
                uAerialPerspectiveLUT: { value: aerialPerspectiveTexture },
                uSceneColor: { value: null },
                uSceneDepth: { value: null },
                uSunDirection: { value: new THREE.Vector3(0, 1, 0) },
                uCameraPosition: { value: new THREE.Vector3(0, 0, 0) },
                uInverseProjection: { value: new THREE.Matrix4() },
                uInverseView: { value: new THREE.Matrix4() }
            },
            vertexShader,
            fragmentShader
        });

        this.scene = new THREE.Scene();
        this.camera = new THREE.OrthographicCamera(-1, 1, 1, -1, 0, 1);
        this.quad = new THREE.Mesh(new THREE.PlaneGeometry(2, 2), this.material);
        this.scene.add(this.quad);
    }

    render(sunDirection, camera, sceneColor, sceneDepth, target = null) {
        this.material.uniforms.uSunDirection.value.copy(sunDirection);
        this.material.uniforms.uCameraPosition.value.copy(camera.position);
        this.material.uniforms.uInverseProjection.value.copy(camera.projectionMatrixInverse);
        this.material.uniforms.uInverseView.value.copy(camera.matrixWorld);
        this.material.uniforms.uSceneColor.value = sceneColor;
        this.material.uniforms.uSceneDepth.value = sceneDepth;
        
        const oldTarget = this.renderer.getRenderTarget();
        this.renderer.setRenderTarget(target);
        this.renderer.render(this.scene, this.camera);
        this.renderer.setRenderTarget(oldTarget);
    }
}
