import common from './common.glsl.js';

export default /* glsl */`
${common}

uniform AtmosphereParameters uParams;
uniform sampler2D uTransmittanceLUT;
uniform sampler2D uSkyViewLUT;
uniform sampler2D uAerialPerspectiveLUT;
uniform sampler2D uSceneColor;
uniform sampler2D uSceneDepth;

uniform vec3 uSunDirection;
uniform vec3 uCameraPosition;
uniform mat4 uInverseProjection;
uniform mat4 uInverseView;

varying vec2 vUv;

vec3 getTransmittance(float r, float mu) {
    vec2 uv = transmittanceLUTCoords(r, mu, uParams);
    return texture2D(uTransmittanceLUT, uv).rgb;
}

// Sample sky view LUT with inverse mapping
vec3 sampleSkyView(vec3 rd) {
    vec3 up = normalize(uCameraPosition);
    float Rg = uParams.planetRadius;
    float Rt = Rg + uParams.atmosphereThickness;
    float r = clamp(length(uCameraPosition), Rg + 0.1, Rt - 0.1);

    float mu = dot(rd, up);
    
    // We need phi in the local frame
    vec3 forward = vec3(1.0, 0.0, 0.0);
    if (abs(dot(up, forward)) > 0.99) forward = vec3(0.0, 0.0, 1.0);
    vec3 right = normalize(cross(up, forward));
    forward = cross(right, up);
    
    float phi = atan(dot(rd, right), dot(rd, forward));
    
    float beta = asin(Rg / r);
    
    float alpha;
    if (mu > 0.0) {
        float theta = acos(clamp(mu, 0.0, 1.0));
        float coord = sqrt(clamp((PI / 2.0 - theta) / (PI / 2.0), 0.0, 1.0));
        alpha = 0.5 + 0.5 * coord;
    } else {
        float theta = asin(clamp(-mu, 0.0, sin(beta)));
        float coord = sqrt(clamp(theta / beta, 0.0, 1.0));
        alpha = 0.5 - 0.5 * coord;
    }
    
    vec2 uv = vec2(phi / (2.0 * PI) + 0.5, alpha);
    return texture2D(uSkyViewLUT, uv).rgb;
}

void main() {
    float depth = texture2D(uSceneDepth, vUv).r;
    vec3 sceneColor = texture2D(uSceneColor, vUv).rgb;
    
    vec4 clipPos = vec4(vUv * 2.0 - 1.0, depth * 2.0 - 1.0, 1.0);
    vec4 viewPos = uInverseProjection * clipPos;
    viewPos /= viewPos.w;
    vec3 worldPos = (uInverseView * viewPos).xyz;
    vec3 rd = normalize(worldPos - uCameraPosition);
    
    float dist = length(worldPos - uCameraPosition);
    
    vec3 finalColor;
    
    if (depth > 0.999999) {
        // Sample sky
        finalColor = sampleSkyView(rd);
        
        // Add sun disk
        float cosTheta = clamp(dot(rd, uSunDirection), -1.0, 1.0);
        float sunAngularRadius = 0.00465; // ~0.5 degrees
        if (cosTheta > cos(sunAngularRadius)) {
            float r = length(uCameraPosition);
            float mu = clamp(dot(normalize(uCameraPosition), uSunDirection), -1.0, 1.0);
            vec3 transmittance = getTransmittance(r, mu);
            finalColor += transmittance * 100.0; // Bloom-like intensity
        }
    } else {
        // Sample aerial perspective
        // map dist to z in LUT
        float maxDistance = 100.0;
        float z = clamp(dist / maxDistance, 0.0, 1.0);
        
        float sliceSize = 32.0;
        float sliceIndex = floor(z * (sliceSize - 1.0));
        float nextSliceIndex = min(sliceIndex + 1.0, sliceSize - 1.0);
        float weight = fract(z * (sliceSize - 1.0));
        
        vec2 uv0 = vec2((sliceIndex + vUv.x) / sliceSize, vUv.y);
        vec2 uv1 = vec2((nextSliceIndex + vUv.x) / sliceSize, vUv.y);
        
        vec4 ap0 = texture2D(uAerialPerspectiveLUT, uv0);
        vec4 ap1 = texture2D(uAerialPerspectiveLUT, uv1);
        vec4 ap = mix(ap0, ap1, weight);
        
        finalColor = sceneColor * ap.a + ap.rgb;
    }
    
    // Tonemapping
    finalColor = max(vec3(0.0), finalColor);
    finalColor = 1.0 - exp(-finalColor);
    
    gl_FragColor = vec4(finalColor, 1.0);
}
`;
