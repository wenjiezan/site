import { GUI } from 'lil-gui';
import { presets } from './presets.js';

export function setupGUI(atmosphere, sunParams, onSunChange) {
    const gui = new GUI();
    
    const sunFolder = gui.addFolder('Sun');
    sunFolder.add(sunParams, 'azimuth', 0, 360).onChange(onSunChange);
    sunFolder.add(sunParams, 'elevation', -10, 90).onChange(onSunChange);

    const params = atmosphere.params;
    
    const atmosphereFolder = gui.addFolder('Atmosphere');
    
    const update = () => atmosphere.setParameters({}); // Trigger transmittance rebake

    atmosphereFolder.add(params, 'planetRadius', 1000, 10000).onChange(update);
    atmosphereFolder.add(params, 'atmosphereThickness', 10, 200).onChange(update);
    
    const rayleighFolder = atmosphereFolder.addFolder('Rayleigh');
    rayleighFolder.add(params.rayleighScattering, 'x', 0, 0.1).name('R').onChange(update);
    rayleighFolder.add(params.rayleighScattering, 'y', 0, 0.1).name('G').onChange(update);
    rayleighFolder.add(params.rayleighScattering, 'z', 0, 0.1).name('B').onChange(update);
    rayleighFolder.add(params, 'rayleighScaleHeight', 1, 20).onChange(update);

    const mieFolder = atmosphereFolder.addFolder('Mie');
    mieFolder.add(params.mieScattering, 'x', 0, 0.1).name('R').onChange(update);
    mieFolder.add(params.mieScattering, 'y', 0, 0.1).name('G').onChange(update);
    mieFolder.add(params.mieScattering, 'z', 0, 0.1).name('B').onChange(update);
    mieFolder.add(params, 'mieAnisotropy', 0, 0.99).onChange(update);
    mieFolder.add(params, 'mieScaleHeight', 0.1, 20).onChange(update);

    const ozoneFolder = atmosphereFolder.addFolder('Ozone');
    ozoneFolder.add(params.ozoneAbsorption, 'x', 0, 0.01).name('R').onChange(update);
    ozoneFolder.add(params.ozoneAbsorption, 'y', 0, 0.01).name('G').onChange(update);
    ozoneFolder.add(params.ozoneAbsorption, 'z', 0, 0.01).name('B').onChange(update);
    ozoneFolder.add(params, 'ozoneCenterAltitude', 0, 50).onChange(update);
    ozoneFolder.add(params, 'ozoneThickness', 1, 30).onChange(update);

    const presetParams = { preset: 'Earth' };
    gui.add(presetParams, 'preset', Object.keys(presets)).onChange(value => {
        atmosphere.setParameters(presets[value]);
        // Update GUI display
        gui.controllersRecursive().forEach(c => c.updateDisplay());
    });

    return gui;
}
