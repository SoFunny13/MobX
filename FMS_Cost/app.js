// ============================================================
// Config
// ============================================================
const STATS_HEADERS = { offer: 'Оффер', source: 'Источник', cost: 'Расход' };
const NO_CONVERT_SOURCES = new Set(['yandexdirect_int']);
const PLATFORM_SUFFIXES = [' Android', ' iOS', ' AOS', ' Web', ' Мобайл', ' android', ' ios'];
const SOURCE_ALIASES = {
    'xiaomiglobal_int': 'xiaomi',
    'xiaomiglobal_int_contract': 'xiaomi',
    'yandexdirect_int': 'Яндекс Директ',
};

const OFFERS_XLSX = 'offers.xlsx';
const OFFERS_SHEET = 'Лист3';
const OFFERS_DATA_START_ROW = 4;
const OFFERS_COL_ID = 0;
const OFFERS_COL_NAME = 1;

const SOURCES_XLSX = 'id_Source.xlsx';
const SOURCES_SHEET = 'Лист1';
const SOURCES_DATA_START_ROW = 2;
const SOURCES_COL_ID = 0;
const SOURCES_COL_NAME = 1;

// ============================================================
// Reference Data (localStorage)
// ============================================================
function getOffers() {
    const data = localStorage.getItem('fms_offers');
    return data ? JSON.parse(data) : null;
}

function getSources() {
    const data = localStorage.getItem('fms_sources');
    return data ? JSON.parse(data) : null;
}

function saveOffers(arr) {
    localStorage.setItem('fms_offers', JSON.stringify(arr));
}

function saveSources(arr) {
    localStorage.setItem('fms_sources', JSON.stringify(arr));
}

function addOffer(id, name) {
    let offers = getOffers() || [];
    const numId = parseInt(id);
    offers = offers.filter(o => o.id !== numId);
    offers.push({ id: numId, name: name.trim() });
    saveOffers(offers);
}

function deleteOffer(id) {
    let offers = getOffers() || [];
    offers = offers.filter(o => o.id !== parseInt(id));
    saveOffers(offers);
}

function addSource(id, name) {
    let sources = getSources() || [];
    const numId = parseInt(id);
    sources = sources.filter(s => s.id !== numId);
    sources.push({ id: numId, name: name.trim() });
    saveSources(sources);
}

function deleteSource(id) {
    let sources = getSources() || [];
    sources = sources.filter(s => s.id !== parseInt(id));
    saveSources(sources);
}

// Build lookup dict: normalized_name -> { id, name }
function buildOffersDict(offers) {
    const dict = {};
    for (const o of offers) {
        dict[normalize(o.name)] = { id: o.id, name: o.name };
    }
    return dict;
}

function buildSourcesDict(sources) {
    const dict = {};
    for (const s of sources) {
        dict[normalize(s.name)] = { id: s.id, name: s.name };
    }
    return dict;
}

// Load reference data from .xlsx files (first time only)
async function initReferenceData() {
    if (!getOffers()) {
        try {
            const resp = await fetch(OFFERS_XLSX);
            const buf = await resp.arrayBuffer();
            const wb = XLSX.read(buf, { type: 'array' });
            const ws = wb.Sheets[OFFERS_SHEET];
            if (ws) {
                const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const offers = [];
                for (let i = OFFERS_DATA_START_ROW - 1; i < rows.length; i++) {
                    const row = rows[i];
                    if (row && row[OFFERS_COL_ID] != null && row[OFFERS_COL_NAME]) {
                        offers.push({ id: parseInt(row[OFFERS_COL_ID]), name: String(row[OFFERS_COL_NAME]).trim() });
                    }
                }
                saveOffers(offers);
            }
        } catch (e) {
            console.error('Failed to load offers.xlsx:', e);
        }
    }

    if (!getSources()) {
        try {
            const resp = await fetch(SOURCES_XLSX);
            const buf = await resp.arrayBuffer();
            const wb = XLSX.read(buf, { type: 'array' });
            const ws = wb.Sheets[SOURCES_SHEET];
            if (ws) {
                const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const sources = [];
                for (let i = SOURCES_DATA_START_ROW - 1; i < rows.length; i++) {
                    const row = rows[i];
                    if (row && row[SOURCES_COL_ID] != null && row[SOURCES_COL_NAME]) {
                        sources.push({ id: parseInt(row[SOURCES_COL_ID]), name: String(row[SOURCES_COL_NAME]).trim() });
                    }
                }
                saveSources(sources);
            }
        } catch (e) {
            console.error('Failed to load id_Source.xlsx:', e);
        }
    }
}

// ============================================================
// Learned Mappings (localStorage)
// ============================================================
function getLearnedMappings() {
    const data = localStorage.getItem('fms_learned');
    return data ? JSON.parse(data) : { offers: {}, sources: {} };
}

function saveLearnedMappings(mappings) {
    localStorage.setItem('fms_learned', JSON.stringify(mappings));
}

function getLearnedOfferId(name) {
    const m = getLearnedMappings();
    return m.offers[name] !== undefined ? m.offers[name] : null;
}

function getLearnedSourceId(name) {
    const m = getLearnedMappings();
    return m.sources[name] !== undefined ? m.sources[name] : null;
}

function saveOfferMappings(newMappings) {
    const m = getLearnedMappings();
    Object.assign(m.offers, newMappings);
    saveLearnedMappings(m);
}

function saveSourceMappings(newMappings) {
    const m = getLearnedMappings();
    Object.assign(m.sources, newMappings);
    saveLearnedMappings(m);
}

// ============================================================
// Matching Logic
// ============================================================
function normalize(name) {
    return name.toLowerCase().trim().replace(/\s+/g, ' ');
}

function normalizeNoSpace(name) {
    return name.toLowerCase().trim().replace(/\s/g, '');
}

function stripPlatformSuffix(name) {
    for (const suffix of PLATFORM_SUFFIXES) {
        if (name.endsWith(suffix)) {
            return name.slice(0, -suffix.length);
        }
    }
    return name;
}

function matchOffer(statsName, offersDict) {
    const raw = statsName.trim();

    // 1. Learned mappings
    const learnedId = getLearnedOfferId(raw);
    if (learnedId !== null) return learnedId;

    // 2. Strip platform suffix and normalize
    const base = stripPlatformSuffix(raw);
    const key = normalize(base);

    // 3. Exact normalized match
    if (offersDict[key]) return offersDict[key].id;

    // 4. No-space match
    const keyNoSpace = normalizeNoSpace(base);
    for (const [refName, ref] of Object.entries(offersDict)) {
        if (normalizeNoSpace(refName) === keyNoSpace) return ref.id;
    }

    // 5. Substring containment
    for (const [refName, ref] of Object.entries(offersDict)) {
        if (key.includes(refName) || refName.includes(key)) return ref.id;
    }

    return null;
}

function matchSource(statsName, sourcesDict) {
    const raw = statsName.trim();

    // 1. Learned mappings
    const learnedId = getLearnedSourceId(raw);
    if (learnedId !== null) return learnedId;

    // 2. Aliases
    const alias = SOURCE_ALIASES[raw.toLowerCase()];
    if (alias) {
        const aliasKey = normalize(alias);
        if (sourcesDict[aliasKey]) return sourcesDict[aliasKey].id;
    }

    // 3. Exact normalized match
    const key = normalize(raw);
    if (sourcesDict[key]) return sourcesDict[key].id;

    // 4. Strip _int / _int_contract suffix
    const base = raw.toLowerCase().replace(/_int_contract$/, '').replace(/_int$/, '').trim();
    for (const [refName, ref] of Object.entries(sourcesDict)) {
        const refLower = refName.toLowerCase();
        if (base === refLower) return ref.id;
        if (refLower.replace(/\s/g, '').includes(base)) return ref.id;
    }

    return null;
}

// ============================================================
// Stats File Parsing
// ============================================================
function parseStatsFile(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    // Find header columns in first 5 rows
    let offerCol = null, sourceCol = null, costCol = null, headerRow = null;

    for (let r = 0; r < Math.min(5, rows.length); r++) {
        const row = rows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
            const val = row[c];
            if (val && typeof val === 'string') {
                const trimmed = val.trim();
                if (trimmed === STATS_HEADERS.offer) { offerCol = c; headerRow = r; }
                if (trimmed === STATS_HEADERS.source) { sourceCol = c; headerRow = Math.max(headerRow || 0, r); }
                if (trimmed === STATS_HEADERS.cost) { costCol = c; headerRow = Math.max(headerRow || 0, r); }
            }
        }
    }

    if (offerCol === null || sourceCol === null || costCol === null) {
        return { error: 'Не найдены обязательные столбцы: Оффер, Источник, Расход' };
    }

    const statsRows = [];
    for (let r = headerRow + 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;
        const offerName = row[offerCol];
        const sourceName = row[sourceCol];
        const costRub = row[costCol];

        if (offerName && sourceName) {
            statsRows.push({
                offer_name: String(offerName).trim(),
                source_name: String(sourceName).trim(),
                cost_rub: parseFloat(costRub) || 0,
            });
        }
    }

    if (statsRows.length === 0) {
        return { error: 'Файл не содержит данных' };
    }

    return { statsRows };
}

// ============================================================
// Report Generation
// ============================================================
function generateReport(dateStr, exchangeRate, statsRows, offerIdMap, sourceIdMap) {
    const data = statsRows.map(row => {
        const offerId = offerIdMap[row.offer_name] ?? '';
        const sourceId = sourceIdMap[row.source_name] ?? '';
        const costRub = parseFloat(row.cost_rub) || 0;
        let cost;
        if (NO_CONVERT_SOURCES.has(row.source_name.toLowerCase())) {
            cost = Math.round(costRub);
        } else {
            cost = exchangeRate ? Math.round(costRub / exchangeRate) : 0;
        }
        return [dateStr, offerId, sourceId, cost];
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 14 }, { wch: 12 }, { wch: 16 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, ws, 'Отчет');
    XLSX.writeFile(wb, `report_${dateStr}.xlsx`);
}

// ============================================================
// Utility
// ============================================================
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function escapeAttr(text) {
    return String(text || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function showStatus(elementId, message, type) {
    const el = document.getElementById(elementId);
    el.className = 'status status-' + type;
    el.textContent = message;
    el.classList.remove('hidden');
}

function clearStatus(elementId) {
    const el = document.getElementById(elementId);
    el.className = '';
    el.textContent = '';
}

function showStatusTimed(elementId, message, type, ms) {
    showStatus(elementId, message, type);
    setTimeout(() => clearStatus(elementId), ms || 3000);
}
