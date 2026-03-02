// ── 4D Benchmark System: Channel × Vertical × GEO × Platform ───────────────────
// Formula: final = base[channel] × verticalMult[vertical] × geoMult[geoTier] × platformMult[platform]
// CTR values are in PERCENT (e.g. 1.2 = 1.2%), CPI in USD

// ── 1. Base values per channel ──────────────────────────────────────────────────

const CHANNEL_BENCHMARKS = {
    // OEM / preinstall sources
    xiaomi:     { ctr: 1.2,  crInstall: 2.5,  cpi: 1.00 },
    huawei:     { ctr: 1.0,  crInstall: 3.2,  cpi: 1.20 },
    samsung:    { ctr: 1.1,  crInstall: 0.5,  cpi: 1.10 },
    oppo:       { ctr: 1.0,  crInstall: 0.4,  cpi: 1.00 },
    vivo:       { ctr: 1.0,  crInstall: 0.4,  cpi: 1.00 },
    realme:     { ctr: 1.0,  crInstall: 0.4,  cpi: 0.90 },
    transsion:  { ctr: 0.9,  crInstall: 0.3,  cpi: 0.80 },
    // Performance channels
    facebook:   { ctr: 1.5,  crInstall: 3.0,  cpi: 2.50 },
    meta:       { ctr: 1.5,  crInstall: 3.0,  cpi: 2.50 },
    google:     { ctr: 2.0,  crInstall: 2.5,  cpi: 2.80 },
    tiktok:     { ctr: 0.8,  crInstall: 1.5,  cpi: 3.00 },
    unity:      { ctr: 1.8,  crInstall: 2.0,  cpi: 1.80 },
    applovin:   { ctr: 1.6,  crInstall: 1.8,  cpi: 2.20 },
    ironsource: { ctr: 1.5,  crInstall: 1.7,  cpi: 2.00 },
    mintegral:  { ctr: 1.2,  crInstall: 1.5,  cpi: 1.80 },
    appnext:    { ctr: 0.9,  crInstall: 5.0,  cpi: 1.50 },
    xapads:     { ctr: 0.8,  crInstall: 4.0,  cpi: 1.20 },
    yandex:     { ctr: 1.4,  crInstall: 2.0,  cpi: 1.50 },
    vk:         { ctr: 1.2,  crInstall: 1.5,  cpi: 1.60 },
    snapchat:   { ctr: 0.7,  crInstall: 1.3,  cpi: 3.50 },
    twitter:    { ctr: 0.5,  crInstall: 1.0,  cpi: 4.50 },
    x:          { ctr: 0.5,  crInstall: 1.0,  cpi: 4.50 },
    moloco:     { ctr: 1.3,  crInstall: 1.6,  cpi: 2.00 },
    liftoff:    { ctr: 1.2,  crInstall: 1.5,  cpi: 2.50 },
    digital_turbine: { ctr: 1.0, crInstall: 0.6, cpi: 0.90 },
    // Default fallback
    _default:   { ctr: 1.2,  crInstall: 1.0,  cpi: 2.00 }
};

// ── 2. Vertical multipliers ────────────────────────────────────────────────────
// Multiplier 1.0 = neutral (same as current behavior for "other")
// Based on industry benchmarks: AppsFlyer, Adjust, Singular 2024-2025

const VERTICAL_MULTIPLIERS = {
    gaming:        { ctr: 1.80, crInstall: 1.30, cpi: 0.50 },  // Cheap installs, high CTR
    entertainment: { ctr: 1.40, crInstall: 1.15, cpi: 0.70 },  // High engagement
    social:        { ctr: 1.30, crInstall: 1.20, cpi: 0.75 },  // Strong network effects
    health:        { ctr: 1.10, crInstall: 1.00, cpi: 1.10 },  // Moderate
    utilities:     { ctr: 1.10, crInstall: 0.95, cpi: 0.85 },  // Functional, moderate
    education:     { ctr: 1.00, crInstall: 0.90, cpi: 0.90 },  // Considered decision
    ecommerce:     { ctr: 0.90, crInstall: 0.85, cpi: 1.30 },  // Competitive market
    travel:        { ctr: 0.90, crInstall: 0.85, cpi: 1.25 },  // Seasonal, competitive
    delivery:      { ctr: 0.85, crInstall: 0.80, cpi: 1.20 },  // Market-dependent
    finance:       { ctr: 0.75, crInstall: 0.70, cpi: 2.50 },  // Most expensive installs
    pharma:        { ctr: 0.80, crInstall: 0.75, cpi: 2.00 },  // Regulated, expensive
    other:         { ctr: 1.00, crInstall: 1.00, cpi: 1.00 }   // Neutral default
};

// ── 3. Per-country GEO multipliers ─────────────────────────────────────────────
// Based on AppTweak 2025, Adjust H1 2024, Mapendo 2025 data.
// All values are relative to US (US = 1.00 baseline).
// Sources: AppTweak Apple Ads Benchmarks 2025 (38 countries),
//          Adjust eCPI Benchmarks H1 2024, Mapendo CPI reports 2024-2025.

const COUNTRY_GEO_MULTIPLIERS = {
    US: { ctr: 1.00, crInstall: 1.00, cpi: 1.00 },  // baseline
    GB: { ctr: 1.09, crInstall: 1.04, cpi: 0.64 },
    JP: { ctr: 0.97, crInstall: 0.84, cpi: 0.63 },
    CA: { ctr: 1.15, crInstall: 1.02, cpi: 0.55 },
    AU: { ctr: 1.18, crInstall: 1.02, cpi: 0.54 },
    DE: { ctr: 1.30, crInstall: 1.00, cpi: 0.53 },
    KR: { ctr: 0.95, crInstall: 0.93, cpi: 0.45 },
    FR: { ctr: 1.12, crInstall: 1.04, cpi: 0.44 },
    NL: { ctr: 1.15, crInstall: 1.04, cpi: 0.37 },
    BE: { ctr: 1.09, crInstall: 1.05, cpi: 0.36 },
    SE: { ctr: 1.17, crInstall: 1.02, cpi: 0.36 },
    ES: { ctr: 1.09, crInstall: 1.00, cpi: 0.35 },
    IE: { ctr: 1.26, crInstall: 1.11, cpi: 0.35 },
    AE: { ctr: 1.17, crInstall: 1.09, cpi: 0.34 },
    NO: { ctr: 1.15, crInstall: 1.05, cpi: 0.34 },
    DK: { ctr: 1.27, crInstall: 1.05, cpi: 0.33 },
    SG: { ctr: 1.05, crInstall: 0.96, cpi: 0.33 },
    AT: { ctr: 1.52, crInstall: 1.07, cpi: 0.33 },
    IT: { ctr: 1.23, crInstall: 1.02, cpi: 0.33 },
    NZ: { ctr: 1.27, crInstall: 1.09, cpi: 0.32 },
    SA: { ctr: 0.83, crInstall: 0.98, cpi: 0.32 },
    FI: { ctr: 1.27, crInstall: 1.04, cpi: 0.31 },
    PL: { ctr: 1.00, crInstall: 1.05, cpi: 0.28 },
    MX: { ctr: 1.20, crInstall: 0.98, cpi: 0.27 },
    TH: { ctr: 1.00, crInstall: 0.87, cpi: 0.27 },
    BR: { ctr: 1.39, crInstall: 0.98, cpi: 0.26 },
    RO: { ctr: 1.45, crInstall: 1.13, cpi: 0.26 },
    TR: { ctr: 1.39, crInstall: 0.87, cpi: 0.25 },
    ZA: { ctr: 1.48, crInstall: 1.07, cpi: 0.25 },
    PT: { ctr: 1.32, crInstall: 1.07, cpi: 0.25 },
    ID: { ctr: 0.94, crInstall: 0.84, cpi: 0.23 },
    AR: { ctr: 1.05, crInstall: 0.98, cpi: 0.23 },
    CL: { ctr: 1.24, crInstall: 1.07, cpi: 0.22 },
    IN: { ctr: 0.79, crInstall: 0.87, cpi: 0.22 },
    VN: { ctr: 0.85, crInstall: 0.89, cpi: 0.21 },
    CO: { ctr: 1.20, crInstall: 1.04, cpi: 0.20 },
    MY: { ctr: 1.11, crInstall: 0.98, cpi: 0.19 },
    PH: { ctr: 1.18, crInstall: 0.98, cpi: 0.18 },
};

// ── 3b. Regional fallback multipliers (for countries not in COUNTRY_GEO_MULTIPLIERS)
// Based on Adjust eCPI H1 2024 regional data, normalized to US baseline.

const REGION_FALLBACK_MULTIPLIERS = {
    north_america:  { ctr: 1.10, crInstall: 1.00, cpi: 0.52 },
    western_europe: { ctr: 1.20, crInstall: 1.03, cpi: 0.35 },
    eastern_europe: { ctr: 1.30, crInstall: 1.05, cpi: 0.27 },
    apac_developed: { ctr: 1.00, crInstall: 0.90, cpi: 0.35 },
    apac_emerging:  { ctr: 1.00, crInstall: 0.88, cpi: 0.22 },
    latam:          { ctr: 1.20, crInstall: 1.00, cpi: 0.24 },
    mena:           { ctr: 0.90, crInstall: 1.00, cpi: 0.30 },
    africa:         { ctr: 1.40, crInstall: 1.00, cpi: 0.20 },
    rest:           { ctr: 1.10, crInstall: 0.95, cpi: 0.20 },
};

// ── 4. Platform multipliers (Android vs iOS) ─────────────────────────────────
// Android = baseline (1.0). iOS adjustments based on AppsFlyer Performance Index.
// iOS: higher CTR (better ad creatives/engagement), lower CR (stricter users),
// significantly higher CPI (premium audience, higher LTV expectations).

const PLATFORM_MULTIPLIERS = {
    Android: { ctr: 1.00, crInstall: 1.00, cpi: 1.00 },  // baseline
    iOS:     { ctr: 1.25, crInstall: 0.85, cpi: 2.50 }   // iOS: +25% CTR, −15% CR, 2.5× CPI
};

// ── 5. Country → Region mapping (fallback for countries not in COUNTRY_GEO_MULTIPLIERS)
// Used only when a country has no direct entry in COUNTRY_GEO_MULTIPLIERS.

const COUNTRY_TO_REGION = {
    // North America
    US: 'north_america', CA: 'north_america',

    // Western Europe
    GB: 'western_europe', DE: 'western_europe', FR: 'western_europe',
    IT: 'western_europe', ES: 'western_europe', NL: 'western_europe',
    BE: 'western_europe', AT: 'western_europe', CH: 'western_europe',
    SE: 'western_europe', NO: 'western_europe', DK: 'western_europe',
    FI: 'western_europe', IE: 'western_europe', PT: 'western_europe',
    LU: 'western_europe', IS: 'western_europe', MT: 'western_europe',
    CY: 'western_europe',

    // Eastern Europe
    PL: 'eastern_europe', CZ: 'eastern_europe', RO: 'eastern_europe',
    HU: 'eastern_europe', GR: 'eastern_europe', HR: 'eastern_europe',
    BG: 'eastern_europe', RS: 'eastern_europe', SK: 'eastern_europe',
    SI: 'eastern_europe', LT: 'eastern_europe', LV: 'eastern_europe',
    EE: 'eastern_europe', BY: 'eastern_europe', UA: 'eastern_europe',
    MD: 'eastern_europe', BA: 'eastern_europe', MK: 'eastern_europe',
    ME: 'eastern_europe', AL: 'eastern_europe', GE: 'eastern_europe',
    AM: 'eastern_europe', AZ: 'eastern_europe', RU: 'eastern_europe',

    // APAC — Developed
    JP: 'apac_developed', KR: 'apac_developed', AU: 'apac_developed',
    NZ: 'apac_developed', SG: 'apac_developed', HK: 'apac_developed',
    TW: 'apac_developed',

    // APAC — Emerging
    IN: 'apac_emerging', ID: 'apac_emerging', TH: 'apac_emerging',
    VN: 'apac_emerging', PH: 'apac_emerging', MY: 'apac_emerging',
    CN: 'apac_emerging', BD: 'apac_emerging', PK: 'apac_emerging',
    MM: 'apac_emerging', KH: 'apac_emerging', LA: 'apac_emerging',
    NP: 'apac_emerging', LK: 'apac_emerging', MN: 'apac_emerging',

    // LATAM
    BR: 'latam', MX: 'latam', AR: 'latam', CL: 'latam', CO: 'latam',
    PE: 'latam', EC: 'latam', VE: 'latam', UY: 'latam', PY: 'latam',
    BO: 'latam', CR: 'latam', PA: 'latam', DO: 'latam', GT: 'latam',
    HN: 'latam', SV: 'latam', NI: 'latam', CU: 'latam', JM: 'latam',
    TT: 'latam',

    // MENA
    AE: 'mena', SA: 'mena', QA: 'mena', KW: 'mena', BH: 'mena',
    OM: 'mena', IL: 'mena', TR: 'mena', EG: 'mena', MA: 'mena',
    TN: 'mena', DZ: 'mena', IQ: 'mena', JO: 'mena', LB: 'mena',
    LY: 'mena', SY: 'mena', YE: 'mena', IR: 'mena',

    // Africa
    ZA: 'africa', NG: 'africa', KE: 'africa', GH: 'africa',
    TZ: 'africa', UG: 'africa', ET: 'africa', RW: 'africa',
    SN: 'africa', CM: 'africa', CI: 'africa', MZ: 'africa',
};

// ── 6. GEO text aliases ────────────────────────────────────────────────────────
// Common country name → ISO 2-letter code

const GEO_ALIASES = {
    'russia': 'RU', 'rossiya': 'RU', 'russian': 'RU',
    'usa': 'US', 'united states': 'US', 'america': 'US',
    'uk': 'GB', 'united kingdom': 'GB', 'england': 'GB', 'britain': 'GB',
    'germany': 'DE', 'deutschland': 'DE',
    'france': 'FR', 'japan': 'JP', 'korea': 'KR', 'south korea': 'KR',
    'brazil': 'BR', 'brasil': 'BR',
    'india': 'IN', 'indonesia': 'ID',
    'turkey': 'TR', 'turkiye': 'TR',
    'mexico': 'MX', 'thailand': 'TH',
    'kazakhstan': 'KZ', 'ukraine': 'UA',
    'china': 'CN', 'prc': 'CN',
    'philippines': 'PH', 'vietnam': 'VN', 'vn': 'VN',
    'malaysia': 'MY', 'singapore': 'SG',
    'egypt': 'EG', 'south africa': 'ZA',
    'poland': 'PL', 'czech': 'CZ', 'czechia': 'CZ',
    'argentina': 'AR', 'colombia': 'CO', 'peru': 'PE', 'chile': 'CL',
    'uae': 'AE', 'emirates': 'AE',
    'saudi': 'SA', 'saudi arabia': 'SA',
    'nigeria': 'NG', 'pakistan': 'PK', 'bangladesh': 'BD',
    'uzbekistan': 'UZ', 'kenya': 'KE',
    'italy': 'IT', 'spain': 'ES', 'portugal': 'PT',
    'netherlands': 'NL', 'holland': 'NL',
    'switzerland': 'CH', 'sweden': 'SE', 'norway': 'NO',
    'denmark': 'DK', 'finland': 'FI', 'austria': 'AT',
    'belgium': 'BE', 'ireland': 'IE',
    'australia': 'AU', 'new zealand': 'NZ', 'canada': 'CA',
    'israel': 'IL', 'taiwan': 'TW', 'hong kong': 'HK',
    'romania': 'RO', 'hungary': 'HU', 'greece': 'GR',
    'serbia': 'RS', 'croatia': 'HR', 'bulgaria': 'BG',
    'belarus': 'BY', 'georgia': 'GE', 'azerbaijan': 'AZ',
    'morocco': 'MA', 'qatar': 'QA', 'kuwait': 'KW',
};

// ── 7. Store category → our vertical mapping ──────────────────────────────────
// Maps App Store primaryGenreName and Google Play category strings to our verticals

const STORE_TO_VERTICAL = {
    // App Store categories
    'games':                'gaming',
    'entertainment':        'entertainment',
    'social networking':    'social',
    'health & fitness':     'health',
    'medical':              'pharma',
    'utilities':            'utilities',
    'education':            'education',
    'shopping':             'ecommerce',
    'travel':               'travel',
    'food & drink':         'delivery',
    'finance':              'finance',
    'lifestyle':            'other',
    'productivity':         'utilities',
    'business':             'finance',
    'news':                 'entertainment',
    'sports':               'entertainment',
    'music':                'entertainment',
    'photo & video':        'entertainment',
    'photography':          'entertainment',
    'navigation':           'utilities',
    'weather':              'utilities',
    'reference':            'education',
    'books':                'education',
    'developer tools':      'utilities',
    'graphics & design':    'utilities',

    // Google Play categories (lowercase)
    'game':                         'gaming',
    'game_action':                  'gaming',
    'game_adventure':               'gaming',
    'game_arcade':                  'gaming',
    'game_board':                   'gaming',
    'game_card':                    'gaming',
    'game_casino':                  'gaming',
    'game_casual':                  'gaming',
    'game_educational':             'gaming',
    'game_music':                   'gaming',
    'game_puzzle':                  'gaming',
    'game_racing':                  'gaming',
    'game_role_playing':            'gaming',
    'game_simulation':              'gaming',
    'game_sports':                  'gaming',
    'game_strategy':                'gaming',
    'game_trivia':                  'gaming',
    'game_word':                    'gaming',
    'social':                       'social',
    'communication':                'social',
    'dating':                       'social',
    'health_and_fitness':           'health',
    'medical':                      'pharma',
    'tools':                        'utilities',
    'productivity':                 'utilities',
    'education':                    'education',
    'shopping':                     'ecommerce',
    'travel_and_local':             'travel',
    'maps_and_navigation':          'travel',
    'food_and_drink':               'delivery',
    'finance':                      'finance',
    'business':                     'finance',
    'entertainment':                'entertainment',
    'music_and_audio':              'entertainment',
    'video_players':                'entertainment',
    'news_and_magazines':           'entertainment',
    'sports':                       'entertainment',
    'photography':                  'entertainment',
    'weather':                      'utilities',
    'books_and_reference':          'education',
    'lifestyle':                    'other',
    'house_and_home':               'other',
    'beauty':                       'health',
    'parenting':                    'education',
    'auto_and_vehicles':            'other',
    'events':                       'entertainment',
    'art_and_design':               'utilities',
    'comics':                       'entertainment',
    'libraries_and_demo':           'utilities',
    'personalization':              'utilities',
};

// ── 8. Full country list (ISO 3166-1) ────────────────────────────────────────
// Used for GEO datalist dropdown. Sorted alphabetically by name.

const COUNTRIES = [
    { code: 'AF', name: 'Afghanistan' },
    { code: 'AL', name: 'Albania' },
    { code: 'DZ', name: 'Algeria' },
    { code: 'AD', name: 'Andorra' },
    { code: 'AO', name: 'Angola' },
    { code: 'AG', name: 'Antigua and Barbuda' },
    { code: 'AR', name: 'Argentina' },
    { code: 'AM', name: 'Armenia' },
    { code: 'AU', name: 'Australia' },
    { code: 'AT', name: 'Austria' },
    { code: 'AZ', name: 'Azerbaijan' },
    { code: 'BS', name: 'Bahamas' },
    { code: 'BH', name: 'Bahrain' },
    { code: 'BD', name: 'Bangladesh' },
    { code: 'BB', name: 'Barbados' },
    { code: 'BY', name: 'Belarus' },
    { code: 'BE', name: 'Belgium' },
    { code: 'BZ', name: 'Belize' },
    { code: 'BJ', name: 'Benin' },
    { code: 'BT', name: 'Bhutan' },
    { code: 'BO', name: 'Bolivia' },
    { code: 'BA', name: 'Bosnia and Herzegovina' },
    { code: 'BW', name: 'Botswana' },
    { code: 'BR', name: 'Brazil' },
    { code: 'BN', name: 'Brunei' },
    { code: 'BG', name: 'Bulgaria' },
    { code: 'BF', name: 'Burkina Faso' },
    { code: 'BI', name: 'Burundi' },
    { code: 'CV', name: 'Cabo Verde' },
    { code: 'KH', name: 'Cambodia' },
    { code: 'CM', name: 'Cameroon' },
    { code: 'CA', name: 'Canada' },
    { code: 'CF', name: 'Central African Republic' },
    { code: 'TD', name: 'Chad' },
    { code: 'CL', name: 'Chile' },
    { code: 'CN', name: 'China' },
    { code: 'CO', name: 'Colombia' },
    { code: 'KM', name: 'Comoros' },
    { code: 'CG', name: 'Congo' },
    { code: 'CD', name: 'Congo (DRC)' },
    { code: 'CR', name: 'Costa Rica' },
    { code: 'CI', name: "Cote d'Ivoire" },
    { code: 'HR', name: 'Croatia' },
    { code: 'CU', name: 'Cuba' },
    { code: 'CY', name: 'Cyprus' },
    { code: 'CZ', name: 'Czech Republic' },
    { code: 'DK', name: 'Denmark' },
    { code: 'DJ', name: 'Djibouti' },
    { code: 'DM', name: 'Dominica' },
    { code: 'DO', name: 'Dominican Republic' },
    { code: 'EC', name: 'Ecuador' },
    { code: 'EG', name: 'Egypt' },
    { code: 'SV', name: 'El Salvador' },
    { code: 'GQ', name: 'Equatorial Guinea' },
    { code: 'ER', name: 'Eritrea' },
    { code: 'EE', name: 'Estonia' },
    { code: 'SZ', name: 'Eswatini' },
    { code: 'ET', name: 'Ethiopia' },
    { code: 'FJ', name: 'Fiji' },
    { code: 'FI', name: 'Finland' },
    { code: 'FR', name: 'France' },
    { code: 'GA', name: 'Gabon' },
    { code: 'GM', name: 'Gambia' },
    { code: 'GE', name: 'Georgia' },
    { code: 'DE', name: 'Germany' },
    { code: 'GH', name: 'Ghana' },
    { code: 'GR', name: 'Greece' },
    { code: 'GD', name: 'Grenada' },
    { code: 'GT', name: 'Guatemala' },
    { code: 'GN', name: 'Guinea' },
    { code: 'GW', name: 'Guinea-Bissau' },
    { code: 'GY', name: 'Guyana' },
    { code: 'HT', name: 'Haiti' },
    { code: 'HN', name: 'Honduras' },
    { code: 'HK', name: 'Hong Kong' },
    { code: 'HU', name: 'Hungary' },
    { code: 'IS', name: 'Iceland' },
    { code: 'IN', name: 'India' },
    { code: 'ID', name: 'Indonesia' },
    { code: 'IR', name: 'Iran' },
    { code: 'IQ', name: 'Iraq' },
    { code: 'IE', name: 'Ireland' },
    { code: 'IL', name: 'Israel' },
    { code: 'IT', name: 'Italy' },
    { code: 'JM', name: 'Jamaica' },
    { code: 'JP', name: 'Japan' },
    { code: 'JO', name: 'Jordan' },
    { code: 'KZ', name: 'Kazakhstan' },
    { code: 'KE', name: 'Kenya' },
    { code: 'KI', name: 'Kiribati' },
    { code: 'KP', name: 'Korea (North)' },
    { code: 'KR', name: 'Korea (South)' },
    { code: 'KW', name: 'Kuwait' },
    { code: 'KG', name: 'Kyrgyzstan' },
    { code: 'LA', name: 'Laos' },
    { code: 'LV', name: 'Latvia' },
    { code: 'LB', name: 'Lebanon' },
    { code: 'LS', name: 'Lesotho' },
    { code: 'LR', name: 'Liberia' },
    { code: 'LY', name: 'Libya' },
    { code: 'LI', name: 'Liechtenstein' },
    { code: 'LT', name: 'Lithuania' },
    { code: 'LU', name: 'Luxembourg' },
    { code: 'MO', name: 'Macao' },
    { code: 'MG', name: 'Madagascar' },
    { code: 'MW', name: 'Malawi' },
    { code: 'MY', name: 'Malaysia' },
    { code: 'MV', name: 'Maldives' },
    { code: 'ML', name: 'Mali' },
    { code: 'MT', name: 'Malta' },
    { code: 'MR', name: 'Mauritania' },
    { code: 'MU', name: 'Mauritius' },
    { code: 'MX', name: 'Mexico' },
    { code: 'MD', name: 'Moldova' },
    { code: 'MC', name: 'Monaco' },
    { code: 'MN', name: 'Mongolia' },
    { code: 'ME', name: 'Montenegro' },
    { code: 'MA', name: 'Morocco' },
    { code: 'MZ', name: 'Mozambique' },
    { code: 'MM', name: 'Myanmar' },
    { code: 'NA', name: 'Namibia' },
    { code: 'NP', name: 'Nepal' },
    { code: 'NL', name: 'Netherlands' },
    { code: 'NZ', name: 'New Zealand' },
    { code: 'NI', name: 'Nicaragua' },
    { code: 'NE', name: 'Niger' },
    { code: 'NG', name: 'Nigeria' },
    { code: 'MK', name: 'North Macedonia' },
    { code: 'NO', name: 'Norway' },
    { code: 'OM', name: 'Oman' },
    { code: 'PK', name: 'Pakistan' },
    { code: 'PA', name: 'Panama' },
    { code: 'PG', name: 'Papua New Guinea' },
    { code: 'PY', name: 'Paraguay' },
    { code: 'PE', name: 'Peru' },
    { code: 'PH', name: 'Philippines' },
    { code: 'PL', name: 'Poland' },
    { code: 'PT', name: 'Portugal' },
    { code: 'QA', name: 'Qatar' },
    { code: 'RO', name: 'Romania' },
    { code: 'RU', name: 'Russia' },
    { code: 'RW', name: 'Rwanda' },
    { code: 'SA', name: 'Saudi Arabia' },
    { code: 'SN', name: 'Senegal' },
    { code: 'RS', name: 'Serbia' },
    { code: 'SG', name: 'Singapore' },
    { code: 'SK', name: 'Slovakia' },
    { code: 'SI', name: 'Slovenia' },
    { code: 'SO', name: 'Somalia' },
    { code: 'ZA', name: 'South Africa' },
    { code: 'SS', name: 'South Sudan' },
    { code: 'ES', name: 'Spain' },
    { code: 'LK', name: 'Sri Lanka' },
    { code: 'SD', name: 'Sudan' },
    { code: 'SR', name: 'Suriname' },
    { code: 'SE', name: 'Sweden' },
    { code: 'CH', name: 'Switzerland' },
    { code: 'SY', name: 'Syria' },
    { code: 'TW', name: 'Taiwan' },
    { code: 'TJ', name: 'Tajikistan' },
    { code: 'TZ', name: 'Tanzania' },
    { code: 'TH', name: 'Thailand' },
    { code: 'TL', name: 'Timor-Leste' },
    { code: 'TG', name: 'Togo' },
    { code: 'TT', name: 'Trinidad and Tobago' },
    { code: 'TN', name: 'Tunisia' },
    { code: 'TR', name: 'Turkey' },
    { code: 'TM', name: 'Turkmenistan' },
    { code: 'UG', name: 'Uganda' },
    { code: 'UA', name: 'Ukraine' },
    { code: 'AE', name: 'United Arab Emirates' },
    { code: 'GB', name: 'United Kingdom' },
    { code: 'US', name: 'United States' },
    { code: 'UY', name: 'Uruguay' },
    { code: 'UZ', name: 'Uzbekistan' },
    { code: 'VE', name: 'Venezuela' },
    { code: 'VN', name: 'Vietnam' },
    { code: 'YE', name: 'Yemen' },
    { code: 'ZM', name: 'Zambia' },
    { code: 'ZW', name: 'Zimbabwe' },
];

// ── 9. Utility functions ───────────────────────────────────────────────────────

/**
 * Normalize a GEO input to a 2-letter ISO country code.
 * Accepts "RU", "Russia", "russia", etc.
 */
function normalizeGeoCode(raw) {
    if (!raw) return '';
    const cleaned = raw.trim().toLowerCase();
    if (!cleaned) return '';
    // If already a 2-letter code, return uppercased
    if (cleaned.length === 2) return cleaned.toUpperCase();
    // Check aliases
    if (GEO_ALIASES[cleaned]) return GEO_ALIASES[cleaned];
    // Fallback: uppercase first 2 chars (best guess)
    return cleaned.toUpperCase().slice(0, 2);
}

/**
 * Get GEO multipliers for a country code.
 * 1. Direct match in COUNTRY_GEO_MULTIPLIERS (38 countries with real data)
 * 2. Fallback to region via COUNTRY_TO_REGION → REGION_FALLBACK_MULTIPLIERS
 * 3. Ultimate fallback: 'rest' region
 * Empty GEO → US baseline (neutral, no penalty).
 */
function getGeoMultipliers(geoCode) {
    if (!geoCode) return { ctr: 1.00, crInstall: 1.00, cpi: 1.00 }; // empty = neutral (US baseline)
    const code = geoCode.trim().toUpperCase();
    // Direct country match
    if (COUNTRY_GEO_MULTIPLIERS[code]) return COUNTRY_GEO_MULTIPLIERS[code];
    // Regional fallback
    const region = COUNTRY_TO_REGION[code] || 'rest';
    return REGION_FALLBACK_MULTIPLIERS[region] || REGION_FALLBACK_MULTIPLIERS.rest;
}

/**
 * Map a store category string to our vertical key.
 */
function mapStoreCategory(category) {
    if (!category) return 'other';
    const key = category.toLowerCase().trim();
    return STORE_TO_VERTICAL[key] || 'other';
}

/**
 * Extract ISO 2-letter code from display value like "Russia (RU)" or plain "RU".
 */
function extractGeoCode(val) {
    if (!val) return '';
    const match = val.match(/\(([A-Z]{2})\)\s*$/);
    return match ? match[1] : normalizeGeoCode(val);
}

/**
 * Round to a sensible number of decimals for display.
 */
function roundBenchmark(val) {
    if (val >= 10) return Math.round(val * 10) / 10;   // e.g. 12.3
    if (val >= 1)  return Math.round(val * 100) / 100;  // e.g. 1.24
    return Math.round(val * 1000) / 1000;               // e.g. 0.168
}
