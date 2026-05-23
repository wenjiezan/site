import common from './common.glsl.js';

export default /* glsl */`
${common}

uniform AtmosphereParameters uParams;
varying vec2 vUv;

void main() {
    float r, mu;
    transmittanceLUTParameters(vUv, uParams, r, mu);

    vec3 ro = vec3(0.0, r, 0.0);
    vec3 rd = vec3(0.0, mu, sqrt(max(0.0, 1.0 - mu * mu)));

    float Rt = uParams.planetRadius + uParams.atmosphereThickness;
    float Rg = uParams.planetRadius;

    vec2 atmosphereIntersections = raySphereIntersection(ro, rd, Rt);
    float tMax = atmosphereIntersections.y;

    vec2 planetIntersections = raySphereIntersection(ro, rd, Rg);
    if (planetIntersections.x > 0.0) {
        tMax = min(tMax, planetIntersections.x);
    }

    const int STEPS = 40;
    float dt = tMax / float(STEPS);
    vec3 opticalDepth = vec3(0.0);

    for (int i = 0; i < STEPS; i++) {
        float t = (float(i) + 0.5) * dt;
        vec3 p = ro + rd * t;
        float h = length(p) - Rg;
        
        vec3 scattering, extinction;
        getScatteringExtinction(h, uParams, scattering, extinction);
        opticalDepth += extinction * dt;
    }

    gl_FragColor = vec4(exp(-opticalDepth), 1.0);
}
`;
