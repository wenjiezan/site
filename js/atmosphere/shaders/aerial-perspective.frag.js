import common from './common.glsl.js';

export default /* glsl */`
${common}

uniform AtmosphereParameters uParams;
uniform sampler2D uTransmittanceLUT;
uniform vec3 uSunDirection;
uniform vec3 uCameraPosition;
uniform mat4 uInverseProjection;
uniform mat4 uInverseView;

varying vec2 vUv;
// For 3D texture, we use z as well. If this is a 2D shader rendering slices, we need to handle that.
// Let's assume this shader is used with a 3D texture or we are rendering slices into a 2D atlas.
// The plan said "32 slices in a 2D atlas if 3D RT support is shaky". 
// Let's try to make it work for a 3D texture first.
// If I use a 3D texture, I'll probably need a specialized pass.
// For now, I'll write the logic for a single voxel.

vec3 getTransmittance(float r, float mu) {
    vec2 uv = transmittanceLUTCoords(r, mu, uParams);
    return texture2D(uTransmittanceLUT, uv).rgb;
}

void main() {
    // This shader would typically be run for each slice of the 3D LUT.
    // Or we use a compute shader if available. 
    // In vanilla Three.js without compute, we often use a 2D atlas.
    // Let's assume we are rendering a 3D texture via multiple passes or a specific extension.
    // Actually, let's stick to the 2D atlas approach for maximum compatibility as per plan.
    
    // UV is for the 2D atlas. We need to extract 3D coordinates.
    // 32x32x32 = 1024x32 or 256x128 etc.
    // Let's use 1024x32 (32 slices of 32x32 side by side).
    
    float sliceSize = 32.0;
    float sliceIndex = floor(vUv.x * sliceSize);
    vec2 sliceUv = vec2(fract(vUv.x * sliceSize), vUv.y);
    float z = (sliceIndex + 0.5) / sliceSize; // 0 to 1
    
    float Rg = uParams.planetRadius;
    float Rt = Rg + uParams.atmosphereThickness;
    
    // Convert sliceUv + z to world position in frustum
    // (Simplified: z is distance from camera)
    float maxDistance = 100.0; // km, tunable
    float dist = z * maxDistance;
    
    // Reconstruct ray direction from sliceUv (which represents screen UV)
    vec4 clipPos = vec4(sliceUv * 2.0 - 1.0, -1.0, 1.0);
    vec4 viewPos = uInverseProjection * clipPos;
    viewPos /= viewPos.w;
    vec3 rd = normalize((uInverseView * vec4(viewPos.xyz, 0.0)).xyz);
    
    vec3 ro = uCameraPosition;
    
    // Raymarch from camera to dist
    const int STEPS = 16;
    float dt = dist / float(STEPS);
    
    vec3 L = vec3(0.0);
    vec3 throughput = vec3(1.0);
    
    float cosTheta = clamp(dot(rd, uSunDirection), -1.0, 1.0);
    float phaseR = rayleighPhase(cosTheta);
    float phaseM = miePhase(cosTheta, uParams.mieAnisotropy);
    
    for (int i = 0; i < STEPS; i++) {
        float t = (float(i) + 0.5) * dt;
        vec3 p = ro + rd * t;
        float r_p = length(p);
        float h = r_p - Rg;
        
        if (h < 0.0) {
             // Inside planet, skip or stop
             break;
        }
        h = min(h, uParams.atmosphereThickness);
        
        float rR, rM, rO;
        getAtmosphereDensities(h, uParams, rR, rM, rO);
        vec3 scattering = rR * uParams.rayleighScattering + rM * uParams.mieScattering;
        vec3 extinction = scattering + rM * (uParams.mieExtinction - uParams.mieScattering) + rO * uParams.ozoneAbsorption;
        
        vec3 sampleTransmittance = exp(-extinction * dt);
        
        float mu_s = clamp(dot(p, uSunDirection) / r_p, -1.0, 1.0);
        vec3 sunTransmittance = getTransmittance(r_p, mu_s);
        
        vec3 S = sunTransmittance * (rR * uParams.rayleighScattering * phaseR + rM * uParams.mieScattering * phaseM);
        vec3 stepInScattering = (S - S * sampleTransmittance) / max(extinction, 0.0001);
        L += throughput * stepInScattering;
        throughput *= sampleTransmittance;
    }
    
    gl_FragColor = vec4(L, throughput.r); // Store average transmittance in alpha
}
`;
