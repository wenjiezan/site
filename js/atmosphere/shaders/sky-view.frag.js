import common from './common.glsl.js';

export default /* glsl */`
${common}

uniform AtmosphereParameters uParams;
uniform sampler2D uTransmittanceLUT;
uniform vec3 uSunDirection;
uniform vec3 uCameraPosition;
varying vec2 vUv;

vec3 getTransmittance(float r, float mu) {
    vec2 uv = transmittanceLUTCoords(r, mu, uParams);
    return texture2D(uTransmittanceLUT, uv).rgb;
}

void skyViewLUTParameters(vec2 uv, float r, out float mu, out float phi) {
    float Rg = uParams.planetRadius;
    float Rt = Rg + uParams.atmosphereThickness;
    
    float beta = asin(clamp(Rg / r, 0.0, 1.0));
    float alpha = uv.y;
    
    if (alpha > 0.5) {
        float coord = (alpha - 0.5) * 2.0;
        float theta = coord * coord * (PI / 2.0);
        mu = cos(PI / 2.0 - theta);
    } else {
        float coord = (0.5 - alpha) * 2.0;
        float theta = coord * coord * beta;
        mu = -sin(theta);
    }
    
    phi = (uv.x * 2.0 - 1.0) * PI;
}

void main() {
    float Rg = uParams.planetRadius;
    float Rt = Rg + uParams.atmosphereThickness;
    
    float r = length(uCameraPosition);
    // Orient the coordinate system so the camera is on the Y axis for LUT calculation
    // However, skyViewLUT is usually done in a local frame.
    // Let's assume uCameraPosition is just for altitude 'r'.
    r = clamp(r, Rg + 0.1, Rt - 0.1);
    
    float mu, phi;
    skyViewLUTParameters(vUv, r, mu, phi);
    
    vec3 ro = vec3(0.0, r, 0.0);
    float sinTheta = sqrt(max(0.0, 1.0 - mu * mu));
    vec3 rd = vec3(sinTheta * cos(phi), mu, sinTheta * sin(phi));
    
    // We need sun direction in the same local frame (where up is +Y)
    // For simplicity, let's assume uSunDirection is already transformed or we calculate it here.
    // In a real implementation, we'd pass the local sun direction.
    vec3 sunDir = normalize(uSunDirection); 
    // Wait, uSunDirection is world space. We need it relative to the 'up' vector at camera pos.
    vec3 up = normalize(uCameraPosition);
    vec3 forward = vec3(1.0, 0.0, 0.0); // Arbitrary
    if (abs(dot(up, forward)) > 0.99) forward = vec3(0.0, 0.0, 1.0);
    vec3 right = normalize(cross(up, forward));
    forward = cross(right, up);
    
    mat3 worldToLocal = mat3(right, up, forward); // up is Y
    vec3 localSunDir = sunDir * worldToLocal;

    vec2 atmosphereIntersections = raySphereIntersection(ro, rd, Rt);
    float tMax = atmosphereIntersections.y;
    
    vec2 planetIntersections = raySphereIntersection(ro, rd, Rg);
    if (planetIntersections.x > 0.0) {
        tMax = min(tMax, planetIntersections.x);
    }
    
    const int STEPS = 32;
    float dt = tMax / float(STEPS);
    
    vec3 L = vec3(0.0);
    vec3 throughput = vec3(1.0);
    
    float cosTheta = clamp(dot(rd, localSunDir), -1.0, 1.0);
    float phaseR = rayleighPhase(cosTheta);
    float phaseM = miePhase(cosTheta, uParams.mieAnisotropy);
    
    for (int i = 0; i < STEPS; i++) {
        float t = (float(i) + 0.5) * dt;
        vec3 p = ro + rd * t;
        float r_p = length(p);
        float h = r_p - Rg;
        
        float rR, rM, rO;
        getAtmosphereDensities(h, uParams, rR, rM, rO);
        vec3 scattering = rR * uParams.rayleighScattering + rM * uParams.mieScattering;
        vec3 extinction = scattering + rM * (uParams.mieExtinction - uParams.mieScattering) + rO * uParams.ozoneAbsorption;
        
        vec3 sampleTransmittance = exp(-extinction * dt);
        
        float mu_s = clamp(dot(p, localSunDir) / r_p, -1.0, 1.0);
        vec3 sunTransmittance = getTransmittance(r_p, mu_s);
        
        vec3 S = sunTransmittance * (rR * uParams.rayleighScattering * phaseR + rM * uParams.mieScattering * phaseM);
        
        // Integrated scattering over the step
        vec3 stepInScattering = (S - S * sampleTransmittance) / max(extinction, 0.0001);
        L += throughput * stepInScattering;
        throughput *= sampleTransmittance;
    }
    
    gl_FragColor = vec4(L, 1.0);
}
`;
