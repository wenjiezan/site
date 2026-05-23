export default /* glsl */`
#define PI 3.14159265359

struct AtmosphereParameters {
    float planetRadius;
    float atmosphereThickness;
    vec3 rayleighScattering;
    float rayleighScaleHeight;
    vec3 mieScattering;
    vec3 mieExtinction;
    float mieAnisotropy;
    float mieScaleHeight;
    vec3 ozoneAbsorption;
    float ozoneCenterAltitude;
    float ozoneThickness;
    vec3 groundAlbedo;
};

// Returns density at height h (0 at planet surface, atmosphereThickness at top)
float getDensity(float h, float scaleHeight) {
    return exp(-h / scaleHeight);
}

// Ozone is a layer, modeled as a triangle or gaussian. Here using a simple tent function.
float getOzoneDensity(float h, AtmosphereParameters params) {
    return max(0.0, 1.0 - abs(h - params.ozoneCenterAltitude) / params.ozoneThickness);
}

void getAtmosphereDensities(float h, AtmosphereParameters params, out float rayleigh, out float mie, out float ozone) {
    rayleigh = getDensity(h, params.rayleighScaleHeight);
    mie = getDensity(h, params.mieScaleHeight);
    ozone = getOzoneDensity(h, params);
}

void getScatteringExtinction(float h, AtmosphereParameters params, out vec3 scattering, out vec3 extinction) {
    float r, m, o;
    getAtmosphereDensities(h, params, r, m, o);
    
    scattering = r * params.rayleighScattering + m * params.mieScattering;
    extinction = scattering + m * (params.mieExtinction - params.mieScattering) + o * params.ozoneAbsorption;
}

// Ray-sphere intersection. Returns distance to first and second intersection.
vec2 raySphereIntersection(vec3 ro, vec3 rd, float radius) {
    float b = dot(ro, rd);
    float c = dot(ro, ro) - radius * radius;
    float h = b * b - c;
    if (h < 0.0) return vec2(-1.0);
    h = sqrt(h);
    return vec2(-b - h, -b + h);
}

// Phase functions
float rayleighPhase(float cosTheta) {
    return 3.0 / (16.0 * PI) * (1.0 + cosTheta * cosTheta);
}

float miePhase(float cosTheta, float g) {
    float g2 = g * g;
    return 3.0 / (8.0 * PI) * ((1.0 - g2) * (1.0 + cosTheta * cosTheta)) / ((2.0 + g2) * pow(1.0 + g2 - 2.0 * g * cosTheta, 1.5));
}

// LUT Helpers (Transmittance)
// r: altitude (distance from planet center)
// mu: cos(view-zenith angle)
vec2 transmittanceLUTCoords(float r, float mu, AtmosphereParameters params) {
    float H = params.atmosphereThickness;
    float Rg = params.planetRadius;
    float Rt = Rg + H;
    
    float x_mu = (mu + 1.0) / 2.0;
    float x_r = (r - Rg) / H;
    return vec2(x_mu, x_r);
}

void transmittanceLUTParameters(vec2 uv, AtmosphereParameters params, out float r, out float mu) {
    float H = params.atmosphereThickness;
    float Rg = params.planetRadius;
    
    mu = uv.x * 2.0 - 1.0;
    r = Rg + uv.y * H;
}
`;
