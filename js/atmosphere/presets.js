export const presets = {
    Earth: {
        planetRadius: 6360.0,
        atmosphereThickness: 100.0,
        rayleighScattering: [0.0058, 0.0135, 0.0331],
        rayleighScaleHeight: 8.0,
        mieScattering: [0.00399, 0.00399, 0.00399],
        mieExtinction: [0.00444, 0.00444, 0.00444],
        mieAnisotropy: 0.8,
        mieScaleHeight: 1.2,
        ozoneAbsorption: [0.00065, 0.00188, 0.000085],
        ozoneCenterAltitude: 25.0,
        ozoneThickness: 15.0,
        groundAlbedo: [0.1, 0.1, 0.1]
    },
    Mars: {
        planetRadius: 3389.5,
        atmosphereThickness: 60.0,
        rayleighScattering: [0.0199, 0.0129, 0.0075],
        rayleighScaleHeight: 11.1,
        mieScattering: [0.001, 0.001, 0.001],
        mieExtinction: [0.002, 0.002, 0.002],
        mieAnisotropy: 0.7,
        mieScaleHeight: 11.1,
        ozoneAbsorption: [0.0, 0.0, 0.0],
        ozoneCenterAltitude: 0.0,
        ozoneThickness: 1.0,
        groundAlbedo: [0.45, 0.25, 0.15]
    }
};
