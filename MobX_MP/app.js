// ── Constants ──────────────────────────────────────────────────────────────────

const CURRENCY_SYMBOLS = {
    USD: '$', EUR: '\u20AC', RUB: '\u20BD', GBP: '\u00A3',
    KZT: '\u20B8', UAH: '\u20B4', BRL: 'R$', INR: '\u20B9'
};

// Exchange rates to USD (how many units of currency per 1 USD)
const CURRENCY_RATES = {
    USD: 1,
    EUR: 0.92,
    RUB: 96,
    GBP: 0.79,
    KZT: 470,
    UAH: 41.5,
    BRL: 5.1,
    INR: 83.5
};

// Benchmark data is now in benchmarks.js:
// CHANNEL_BENCHMARKS, VERTICAL_MULTIPLIERS, COUNTRY_GEO_MULTIPLIERS,
// REGION_FALLBACK_MULTIPLIERS, COUNTRY_TO_REGION, PLATFORM_MULTIPLIERS,
// GEO_ALIASES, STORE_TO_VERTICAL, getGeoMultipliers

// CORS proxy for Google Play (iTunes API works directly without proxy)
const GP_PROXY = 'https://api.codetabs.com/v1/proxy/?quest=';

// ── State ──────────────────────────────────────────────────────────────────────

let rowCounter = 0;
let currentCurrency = 'USD'; // tracks last selected currency for conversion
let selectedGeos = []; // selected GEO codes for multi-GEO row creation

// ── DOM refs ───────────────────────────────────────────────────────────────────

const thead = document.getElementById('mediaplan-head');
const tbody = document.getElementById('mediaplan-body');
const tfoot = document.getElementById('mediaplan-foot');
const addRowBtn = document.getElementById('add-row-btn');
const exportBtn = document.getElementById('export-btn');
const resetBtn = document.getElementById('reset-btn');
const cpaEventSelect = document.getElementById('cpa-event');
const cpaEventCustom = document.getElementById('cpa-event-custom');
const currencySelect = document.getElementById('currency');
const appLinkAndroid = document.getElementById('app-link-android');
const appLinkIos = document.getElementById('app-link-ios');
const appLinkAndroidStatus = document.getElementById('app-link-android-status');
const appLinkIosStatus = document.getElementById('app-link-ios-status');
const verticalSelect = document.getElementById('vertical');
const geoTagsContainer = document.getElementById('geo-tags');
const geoAddInput = document.getElementById('geo-add-input');
const geoAddDatalist = document.getElementById('geo-add-options');
const trafficTypeSelect = document.getElementById('traffic-type');

function getMode() {
    const v = cpaEventSelect.value;
    if (v === 'Installs') return 'installs';
    if (v === 'Commission') return 'commission';
    return 'purchases';
}

function isCommissionMode() {
    const m = getMode();
    return m === 'commission';
}

function getCommissionPct() {
    const el = document.getElementById('commission-pct');
    return parseFloat(el ? el.value : 0) || 0;
}

function getTrafficType() {
    return trafficTypeSelect.value; // 'app' | 'web'
}

// ── Initialization ─────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
    addRowBtn.addEventListener('click', addRow);
    exportBtn.addEventListener('click', exportToExcel);
    resetBtn.addEventListener('click', resetAll);
    cpaEventSelect.addEventListener('change', onCpaEventChange);
    trafficTypeSelect.addEventListener('change', onTrafficTypeChange);
    currencySelect.addEventListener('change', onCurrencyChange);
    verticalSelect.addEventListener('change', reapplyAllBenchmarks);
    appLinkAndroid.addEventListener('blur', () => onAppLinkBlur(appLinkAndroid, appLinkAndroidStatus));
    appLinkIos.addEventListener('blur', () => onAppLinkBlur(appLinkIos, appLinkIosStatus));
    const commPctInput = document.getElementById('commission-pct');
    if (commPctInput) commPctInput.addEventListener('input', updateVatVisibility);
    renderTableStructure();
    initGeoDatalist();
    initGeoMultiSelect();
    initSourceCheckboxes();
});

// ── GEO Datalist ──────────────────────────────────────────────────────────────

function initGeoDatalist() {
    const dl = document.createElement('datalist');
    dl.id = 'geo-options';
    COUNTRIES.forEach(c => {
        const opt = document.createElement('option');
        opt.value = `${c.name} (${c.code})`;
        dl.appendChild(opt);
    });
    document.body.appendChild(dl);
}

// ── GEO Multi-Select ─────────────────────────────────────────────────────────

function initGeoMultiSelect() {
    // Populate datalist for the add-input
    COUNTRIES.forEach(c => {
        const opt = document.createElement('option');
        opt.value = `${c.name} (${c.code})`;
        geoAddDatalist.appendChild(opt);
    });

    // When user selects a country from datalist
    geoAddInput.addEventListener('change', () => {
        const val = geoAddInput.value.trim();
        if (!val) return;
        const code = extractGeoCode(val);
        if (!code) return;
        geoAddInput.value = '';
        if (selectedGeos.includes(code)) return; // already selected
        addGeoTag(code);
    });

    // Also handle Enter key
    geoAddInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            geoAddInput.dispatchEvent(new Event('change'));
        }
    });

    // Click on container focuses input
    geoAddInput.closest('.geo-select-wrap').addEventListener('click', () => geoAddInput.focus());
}

function getCountryName(code) {
    const c = COUNTRIES.find(c => c.code === code);
    return c ? c.name : code;
}

function addGeoTag(code) {
    selectedGeos.push(code);
    renderGeoTags();
    // Add rows for this GEO for all checked sources
    addRowsForGeo(code);
}

function removeGeoTag(code) {
    selectedGeos = selectedGeos.filter(c => c !== code);
    renderGeoTags();
    // Remove rows with this GEO for all sources
    removeRowsForGeo(code);
}

function renderGeoTags() {
    geoTagsContainer.innerHTML = '';
    selectedGeos.forEach(code => {
        const tag = document.createElement('span');
        tag.className = 'geo-tag';
        tag.innerHTML = `${code} <button class="geo-tag-remove" title="Remove">&times;</button>`;
        tag.querySelector('.geo-tag-remove').addEventListener('click', (e) => {
            e.stopPropagation();
            removeGeoTag(code);
        });
        geoTagsContainer.appendChild(tag);
    });
}

/** Add rows for a newly added GEO — for all currently checked sources */
function addRowsForGeo(geoCode) {
    const geoDisplay = getCountryName(geoCode) + ' (' + geoCode + ')';
    const checkedSources = document.querySelectorAll('#source-select input[type="checkbox"]:checked');
    checkedSources.forEach(cb => {
        const sourceKey = cb.dataset.source;
        const label = cb.dataset.label;
        const platforms = getActivePlatforms(sourceKey);
        platforms.forEach(platform => {
            // Check if this exact combo already exists
            if (findRow(sourceKey, platform, geoCode)) return;
            createSourceRow(sourceKey, label, platform, geoDisplay);
        });
    });
}

/** Remove rows for a removed GEO — for all sources */
function removeRowsForGeo(geoCode) {
    tbody.querySelectorAll('tr').forEach(tr => {
        if (tr.dataset.geo === geoCode) {
            const sourceKey = tr.dataset.sourceKey;
            tr.remove();
            // If no rows left for this source, uncheck
            if (sourceKey && !tbody.querySelector(`tr[data-source-key="${sourceKey}"]`)) {
                const cb = document.querySelector(`#source-select input[data-source="${sourceKey}"]`);
                if (cb) cb.checked = false;
            }
        }
    });
    recalcTotals();
    updateVatVisibility();
}

/** Find a row by source + platform + geo combo */
function findRow(sourceKey, platform, geoCode) {
    return tbody.querySelector(`tr[data-source-key="${sourceKey}"][data-platform="${platform}"][data-geo="${geoCode}"]`);
}

// ── CPA Event Dropdown ────────────────────────────────────────────────────────

function getCpaEventValue() {
    return cpaEventSelect.value || '';
}

function onCpaEventChange() {
    cpaEventCustom.style.display = 'none';
    cpaEventCustom.value = '';
    // Clear rows and rebuild table structure for new mode
    tbody.innerHTML = '';
    // Uncheck all source checkboxes
    document.querySelectorAll('#source-select input[type="checkbox"]').forEach(cb => {
        cb.checked = false;
    });
    rowCounter = 0;
    renderTableStructure();
    recalcTotals();
    updateVatVisibility();
    const commGroup = document.getElementById('commission-pct-group');
    if (commGroup) commGroup.style.display = isCommissionMode() ? 'block' : 'none';
}

function onTrafficTypeChange() {
    tbody.innerHTML = '';
    rowCounter = 0;
    filterSourcesByTraffic();
    renderTableStructure();
    recalcTotals();
    updateVatVisibility();
}

// ── Dynamic Table Structure ──────────────────────────────────────────────────

function renderTableStructure() {
    if (getTrafficType() === 'web') {
        thead.innerHTML = `<tr class="header-row">
            <th class="col-channel">Channel</th>
            <th class="col-platform">Platform</th>
            <th class="col-geo">GEO</th>
            <th class="col-period">Period</th>
            <th class="col-purchases">Total Purchases</th>
            <th class="col-cpp">Cost per purchase</th>
            <th class="col-budget">Total Budget</th>
            <th class="col-actions"></th>
        </tr>`;
        tfoot.innerHTML = `<tr class="total-row" id="total-row">
            <td colspan="4" class="total-label">Total</td>
            <td id="total-purchases">0</td>
            <td>—</td>
            <td id="total-cost">0</td>
            <td></td>
        </tr>`;
        return;
    }
    const mode = getMode();
    if (mode === 'commission') {
        thead.innerHTML = `<tr class="header-row">
            <th class="col-channel">Channel</th>
            <th class="col-platform">Platform</th>
            <th class="col-geo">GEO</th>
            <th class="col-period">Period</th>
            <th class="col-views">Views</th>
            <th class="col-cpm">CPM</th>
            <th class="col-ctr">CTR, %</th>
            <th class="col-clicks">Total Clicks</th>
            <th class="col-cpc">CPC</th>
            <th class="col-cr-install">CR install to install, %</th>
            <th class="col-installs">Total Installs</th>
            <th class="col-cpi">CPI</th>
            <th class="col-cr-purchase">CR install to purchase, %</th>
            <th class="col-purchases">Total Purchases</th>
            <th class="col-cpp">Cost per purchase</th>
            <th class="col-budget">Total Cost</th>
            <th class="col-actions"></th>
        </tr>`;
        tfoot.innerHTML = `<tr class="total-row" id="total-row">
            <td colspan="4" class="total-label">Total</td>
            <td id="total-views">0</td>
            <td>—</td><td>—</td>
            <td id="total-clicks">0</td>
            <td>—</td><td>—</td>
            <td id="total-installs">0</td>
            <td>—</td><td>—</td>
            <td id="total-purchases">0</td>
            <td>—</td>
            <td id="total-cost">0</td>
            <td></td>
        </tr>`;
    } else if (mode === 'installs') {
        thead.innerHTML = `<tr class="header-row">
            <th class="col-channel">Channel</th>
            <th class="col-platform">Platform</th>
            <th class="col-geo">GEO</th>
            <th class="col-period">Period</th>
            <th class="col-installs">Total Installs</th>
            <th class="col-cpi">CPI</th>
            <th class="col-budget">Cost</th>
            <th class="col-views">Views</th>
            <th class="col-cpm">CPM</th>
            <th class="col-ctr">CTR, %</th>
            <th class="col-clicks">Total Clicks</th>
            <th class="col-cpc">CPC</th>
            <th class="col-cr-purchase">CR install to purchase, %</th>
            <th class="col-purchases">Total Purchases</th>
            <th class="col-cpp">Cost per purchase</th>
            <th class="col-actions"></th>
        </tr>`;
        tfoot.innerHTML = `<tr class="total-row" id="total-row">
            <td colspan="4" class="total-label">Total</td>
            <td id="total-installs">0</td>
            <td>—</td>
            <td id="total-cost">0</td>
            <td id="total-views">0</td>
            <td>—</td>
            <td>—</td>
            <td id="total-clicks">0</td>
            <td>—</td>
            <td>—</td>
            <td id="total-purchases">0</td>
            <td>—</td>
            <td></td>
        </tr>`;
    } else {
        thead.innerHTML = `<tr class="header-row">
            <th class="col-channel">Channel</th>
            <th class="col-platform">Platform</th>
            <th class="col-geo">GEO</th>
            <th class="col-period">Period</th>
            <th class="col-events">Total Purchases</th>
            <th class="col-cpa">CPA Rate</th>
            <th class="col-budget">Cost</th>
            <th class="col-views">Views</th>
            <th class="col-cpm">CPM</th>
            <th class="col-ctr">CTR, %</th>
            <th class="col-clicks">Total Clicks</th>
            <th class="col-cpc">CPC</th>
            <th class="col-cr-install">CR install/click, %</th>
            <th class="col-installs">Total Installs</th>
            <th class="col-cpi">CPI</th>
            <th class="col-actions"></th>
        </tr>`;
        tfoot.innerHTML = `<tr class="total-row" id="total-row">
            <td colspan="4" class="total-label">Total</td>
            <td id="total-events">0</td>
            <td>—</td>
            <td id="total-cost">0</td>
            <td id="total-views">0</td>
            <td>—</td>
            <td>—</td>
            <td id="total-clicks">0</td>
            <td>—</td>
            <td>—</td>
            <td id="total-installs">0</td>
            <td>—</td>
            <td></td>
        </tr>`;
    }
}

// ── Source Checkboxes ──────────────────────────────────────────────────────────

function initSourceCheckboxes() {
    const checkboxes = document.querySelectorAll('#source-select input[type="checkbox"]');
    checkboxes.forEach(cb => {
        cb.addEventListener('change', () => {
            if (cb.checked) {
                addSourceRow(cb.dataset.source, cb.dataset.label);
            } else {
                removeSourceRow(cb.dataset.source);
            }
        });
    });

    document.getElementById('select-all-btn').addEventListener('click', () => {
        document.querySelectorAll('#source-select input[type="checkbox"]').forEach(cb => {
            if (!cb.checked && cb.closest('.source-label').style.display !== 'none') {
                cb.checked = true;
                addSourceRow(cb.dataset.source, cb.dataset.label);
            }
        });
    });

    document.getElementById('clear-all-btn').addEventListener('click', () => {
        document.querySelectorAll('#source-select input[type="checkbox"]').forEach(cb => {
            if (cb.checked && cb.closest('.source-label').style.display !== 'none') {
                cb.checked = false;
                removeSourceRow(cb.dataset.source);
            }
        });
    });

    filterSourcesByTraffic();
}

function filterSourcesByTraffic() {
    const type = getTrafficType();
    document.querySelectorAll('#source-select .source-label').forEach(label => {
        const traffic = label.dataset.traffic || 'app';
        const visible = traffic === type;
        label.style.display = visible ? '' : 'none';
        if (!visible) {
            const cb = label.querySelector('input[type="checkbox"]');
            if (cb && cb.checked) {
                cb.checked = false;
                removeSourceRow(cb.dataset.source);
            }
        }
    });
}

/**
 * Sources that only support Android (even if iOS link is filled).
 */
const ANDROID_ONLY_SOURCES = ['xiaomi', 'xiaomi_codev', 'appnext', 'xapads', 'google', 'huawei'];

/**
 * Sources that use a custom platform name instead of Android/iOS.
 */
const CUSTOM_PLATFORM_SOURCES = {
    huawei: 'AppGallery'
};

/**
 * Sources that have no Views/CPM/CTR (display dash instead).
 */
const NO_VIEWS_SOURCES = ['xiaomi_codev'];

/**
 * Determine which platforms to create rows for based on filled App Link fields.
 * Returns ['Android'], ['iOS'], or ['Android', 'iOS'].
 * If neither link is filled, defaults to ['Android'].
 */
function getActivePlatforms(sourceKey) {
    const hasAndroid = appLinkAndroid.value.trim() !== '';
    const hasIos = appLinkIos.value.trim() !== '';
    // Custom platform sources (e.g. Huawei → AppGallery)
    if (CUSTOM_PLATFORM_SOURCES[sourceKey]) return [CUSTOM_PLATFORM_SOURCES[sourceKey]];
    // Android-only sources always get only Android rows
    if (ANDROID_ONLY_SOURCES.includes(sourceKey)) return ['Android'];
    if (hasAndroid && hasIos) return ['Android', 'iOS'];
    if (hasIos) return ['iOS'];
    return ['Android'];
}

function addSourceRow(sourceKey, label) {
    const platforms = getActivePlatforms(sourceKey);

    if (selectedGeos.length > 0) {
        // Multi-GEO: create rows for each GEO × platform
        selectedGeos.forEach(geoCode => {
            const geoDisplay = getCountryName(geoCode) + ' (' + geoCode + ')';
            platforms.forEach(platform => {
                // Skip only exact duplicates (same source + platform + geo)
                if (findRow(sourceKey, platform, geoCode)) return;
                createSourceRow(sourceKey, label, platform, geoDisplay);
            });
        });
    } else {
        // No GEOs selected: one row per platform, user fills GEO manually
        platforms.forEach(platform => {
            if (findRow(sourceKey, platform, '')) return;
            createSourceRow(sourceKey, label, platform);
        });
    }
}

function createSourceRow(sourceKey, label, platform, geo) {
    rowCounter++;
    const id = rowCounter;
    const tr = document.createElement('tr');
    tr.dataset.rowId = id;
    tr.dataset.sourceKey = sourceKey;
    tr.dataset.platform = platform;
    // Track GEO code in data attribute for multi-GEO management
    const geoCode = geo ? extractGeoCode(geo) : '';
    tr.dataset.geo = geoCode;
    const mode = getMode();

    const baseCols = `
        <td><input type="text" data-field="channel" value="${label}" class="input-channel" readonly style="font-weight:600;cursor:default"></td>
        <td><input type="text" data-field="platform" value="${platform}" class="input-channel" readonly style="cursor:default"></td>
        <td><input type="text" data-field="geo" list="geo-options" placeholder="Select country" autocomplete="off"></td>
        <td><input type="text" data-field="period" value="1 month" style="width:80px"></td>`;

    if (getTrafficType() === 'web') {
        tr.innerHTML = baseCols + `
        <td><input type="text" inputmode="decimal" data-field="purchases" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpp" placeholder="Cost per purchase"></td>
        <td><input type="text" inputmode="decimal" data-field="budget" placeholder="Budget"></td>
        <td><button class="btn-delete" title="Remove row">&times;</button></td>`;
        const channelInput = tr.querySelector('[data-field="channel"]');
        channelInput.removeAttribute('readonly');
        channelInput.style.cursor = '';
        channelInput.style.fontWeight = '';
        const platformInput = tr.querySelector('[data-field="platform"]');
        platformInput.value = 'Web';
        platformInput.removeAttribute('readonly');
        platformInput.style.cursor = '';
    } else if (mode === 'installs') {
        tr.innerHTML = baseCols + `
        <td><input type="text" inputmode="decimal" data-field="installs" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpi" placeholder="CPI" class="editable-calc"></td>
        <td><input type="text" inputmode="decimal" data-field="budget" placeholder="Budget"></td>
        <td><input type="text" inputmode="decimal" data-field="views" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpm" class="calculated" readonly tabindex="-1"></td>
        <td><span class="pct-wrap"><input type="text" inputmode="decimal" data-field="ctr" placeholder="e.g. 1.2" class="editable-calc"><span class="pct-sign">%</span></span></td>
        <td><input type="text" inputmode="decimal" data-field="clicks" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpc" class="calculated" readonly tabindex="-1"></td>
        <td><span class="pct-wrap"><input type="text" inputmode="decimal" data-field="crPurchase" placeholder="e.g. 5"><span class="pct-sign">%</span></span></td>
        <td><input type="text" inputmode="decimal" data-field="purchases" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpp" class="calculated" readonly tabindex="-1"></td>
        <td><button class="btn-delete" title="Remove row">&times;</button></td>`;
    } else if (isCommissionMode()) {
        tr.innerHTML = baseCols + `
        <td><input type="text" inputmode="decimal" data-field="views" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpm" class="calculated" readonly tabindex="-1"></td>
        <td><span class="pct-wrap"><input type="text" inputmode="decimal" data-field="ctr" placeholder="e.g. 1.2" class="editable-calc"><span class="pct-sign">%</span></span></td>
        <td><input type="text" inputmode="decimal" data-field="clicks" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpc" class="calculated" readonly tabindex="-1"></td>
        <td><span class="pct-wrap"><input type="text" inputmode="decimal" data-field="crInstall" placeholder="e.g. 3" class="editable-calc"><span class="pct-sign">%</span></span></td>
        <td><input type="text" inputmode="decimal" data-field="installs" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpi" placeholder="CPI" class="editable-calc"></td>
        <td><span class="pct-wrap"><input type="text" inputmode="decimal" data-field="crPurchase" placeholder="e.g. 5"><span class="pct-sign">%</span></span></td>
        <td><input type="text" inputmode="decimal" data-field="purchases" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpp" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="budget" placeholder="Total Cost"></td>
        <td><button class="btn-delete" title="Remove row">&times;</button></td>`;
    } else {
        tr.innerHTML = baseCols + `
        <td><input type="text" inputmode="decimal" data-field="events" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpa" placeholder="CPA"></td>
        <td><input type="text" inputmode="decimal" data-field="budget" placeholder="Budget"></td>
        <td><input type="text" inputmode="decimal" data-field="views" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpm" class="calculated" readonly tabindex="-1"></td>
        <td><span class="pct-wrap"><input type="text" inputmode="decimal" data-field="ctr" placeholder="e.g. 1.2" class="editable-calc"><span class="pct-sign">%</span></span></td>
        <td><input type="text" inputmode="decimal" data-field="clicks" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpc" class="calculated" readonly tabindex="-1"></td>
        <td><span class="pct-wrap"><input type="text" inputmode="decimal" data-field="crInstall" placeholder="e.g. 5" class="editable-calc"><span class="pct-sign">%</span></span></td>
        <td><input type="text" inputmode="decimal" data-field="installs" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpi" placeholder="CPI" class="editable-calc"></td>
        <td><button class="btn-delete" title="Remove row">&times;</button></td>`;
    }

    // Wire events
    tr.querySelectorAll('input[data-field]').forEach(inp => {
        if (inp.dataset.field === 'channel' || inp.dataset.field === 'platform' || inp.dataset.field === 'geo' || inp.dataset.field === 'period') return;
        inp.addEventListener('input', () => recalcRow(tr));
    });
    // Format editable numeric fields with spaces
    tr.querySelectorAll('input[data-field="budget"], input[data-field="cpa"], input[data-field="cpi"], input[data-field="purchases"], input[data-field="cpp"], input[data-field="crPurchase"], input[data-field="crInstall"], input[data-field="ctr"]').forEach(wireNumericFormat);

    const geoInput = tr.querySelector('[data-field="geo"]');
    if (geo) geoInput.value = geo;
    geoInput.addEventListener('change', () => {
        updateVatVisibility();
        ['ctr','cpi','crInstall'].forEach(f => {
            const el = tr.querySelector(`[data-field="${f}"]`);
            if (el) el.dataset.auto = '1';
        });
        applyBenchmarks(tr);
    });
    geoInput.addEventListener('blur', () => { updateVatVisibility(); applyBenchmarks(tr); });
    geoInput.addEventListener('input', () => {
        const code = extractGeoCode(geoInput.value.trim());
        if (code && code.length === 2) {
            ['ctr','cpi','crInstall'].forEach(f => {
                const el = tr.querySelector(`[data-field="${f}"]`);
                if (el) el.dataset.auto = '1';
            });
            applyBenchmarks(tr);
        }
    });

    // Track manual edits on benchmark fields
    const ctrInput = tr.querySelector('[data-field="ctr"]');
    const cpiInput = tr.querySelector('[data-field="cpi"]');
    if (ctrInput) ctrInput.addEventListener('input', () => { delete ctrInput.dataset.auto; });
    if (cpiInput) cpiInput.addEventListener('input', () => { delete cpiInput.dataset.auto; });

    if (mode === 'purchases' || isCommissionMode()) {
        const crInstallInput = tr.querySelector('[data-field="crInstall"]');
        if (crInstallInput) crInstallInput.addEventListener('input', () => { delete crInstallInput.dataset.auto; });
    }

    tr.querySelector('.btn-delete').addEventListener('click', () => {
        tr.remove();
        // Uncheck checkbox only if no rows left for this source
        if (!tbody.querySelector(`tr[data-source-key="${sourceKey}"]`)) {
            const cb = document.querySelector(`#source-select input[data-source="${sourceKey}"]`);
            if (cb) cb.checked = false;
        }
        recalcTotals();
        updateVatVisibility();
    });

    tbody.appendChild(tr);
    applyBenchmarks(tr);
}

function removeSourceRow(sourceKey) {
    const rows = tbody.querySelectorAll(`tr[data-source-key="${sourceKey}"]`);
    rows.forEach(tr => tr.remove());
    if (rows.length > 0) {
        recalcTotals();
        updateVatVisibility();
    }
}

// ── App Link → Auto-detect vertical ───────────────────────────────────────────

async function onAppLinkBlur(input, statusEl) {
    const url = input.value.trim();
    if (!url) {
        statusEl.textContent = '';
        statusEl.className = 'app-link-status';
        return;
    }

    try {
        statusEl.textContent = '⏳';
        statusEl.className = 'app-link-status loading';

        const result = await fetchAppCategory(url);
        if (result.vertical && result.vertical !== 'other') {
            verticalSelect.value = result.vertical;
            statusEl.textContent = '✓ ' + verticalSelect.options[verticalSelect.selectedIndex].text;
            statusEl.className = 'app-link-status success';
            reapplyAllBenchmarks();
        } else {
            statusEl.textContent = '— category not detected';
            statusEl.className = 'app-link-status error';
        }
        // Auto-fill client name if empty
        const clientInput = document.getElementById('client');
        if (result.appName && !clientInput.value.trim()) {
            clientInput.value = result.appName;
        }
    } catch (e) {
        console.warn('Failed to fetch app category:', e);
        statusEl.textContent = '— fetch failed';
        statusEl.className = 'app-link-status error';
    }
}

/**
 * Fetch app category and name from App Store or Google Play link.
 * Returns { vertical, appName }.
 */
async function fetchAppCategory(url) {
    // ── App Store ──
    const iosMatch = url.match(/apps\.apple\.com\/.*?\/app\/.*?\/id(\d+)/i)
                  || url.match(/itunes\.apple\.com\/.*?\/app\/.*?\/id(\d+)/i)
                  || url.match(/apps\.apple\.com\/.*?id(\d+)/i);
    if (iosMatch) {
        const countryMatch = url.match(/apps\.apple\.com\/([a-z]{2})\//i)
                          || url.match(/itunes\.apple\.com\/([a-z]{2})\//i);
        const country = countryMatch ? countryMatch[1] : '';
        return await fetchAppStoreCategory(iosMatch[1], country);
    }

    // ── Google Play ──
    const gpMatch = url.match(/play\.google\.com\/store\/apps\/details\?.*?id=([a-zA-Z0-9_.]+)/i);
    if (gpMatch) {
        return await fetchGooglePlayCategory(gpMatch[1]);
    }

    return { vertical: 'other', appName: '' };
}

async function fetchAppStoreCategory(appId, country) {
    // iTunes Search API — try direct first, then CORS proxy as fallback
    const attempts = [];
    if (country) attempts.push(country);
    attempts.push('');  // global (no country param)
    if (country !== 'us') attempts.push('us');
    for (const fb of ['mx', 'br', 'gb', 'de', 'ru', 'in', 'id', 'tr', 'sa', 'ng']) {
        if (!attempts.includes(fb)) attempts.push(fb);
    }

    // Try direct requests first (fast, no proxy)
    for (const cc of attempts) {
        let apiUrl = `https://itunes.apple.com/lookup?id=${appId}`;
        if (cc) apiUrl += `&country=${cc}`;
        try {
            const resp = await fetch(apiUrl);
            if (!resp.ok) continue;
            const data = await resp.json();
            if (data.results && data.results.length > 0) {
                const app = data.results[0];
                return { vertical: mapStoreCategory(app.primaryGenreName), appName: app.trackName || '' };
            }
        } catch (e) { continue; }
    }

    // Fallback: try via CORS proxy (for GitHub Pages etc.)
    for (const cc of attempts) {
        let baseUrl = `https://itunes.apple.com/lookup?id=${appId}`;
        if (cc) baseUrl += `&country=${cc}`;
        try {
            const resp = await fetch(GP_PROXY + encodeURIComponent(baseUrl));
            if (!resp.ok) continue;
            const data = await resp.json();
            if (data.results && data.results.length > 0) {
                const app = data.results[0];
                return { vertical: mapStoreCategory(app.primaryGenreName), appName: app.trackName || '' };
            }
        } catch (e) { continue; }
    }

    return { vertical: 'other', appName: '' };
}

async function fetchGooglePlayCategory(packageName) {
    // Google Play needs a CORS proxy — use codetabs
    const pageUrl = `https://play.google.com/store/apps/details%3Fid%3D${packageName}%26hl%3Den`;
    const resp = await fetch(GP_PROXY + pageUrl);
    if (!resp.ok) throw new Error('Google Play fetch error');
    const html = await resp.text();

    // Extract app name from <title> or JSON-LD
    const nameMatch = html.match(/"name"\s*:\s*"([^"]+)"/i)
                   || html.match(/<title>([^<]+?)(?:\s*-\s*Apps on Google Play)?<\/title>/i);
    const appName = nameMatch ? nameMatch[1].trim() : '';

    // Extract category from JSON-LD structured data or itemprop tags
    const categoryMatch = html.match(/"applicationCategory"\s*:\s*"([^"]+)"/i)
                       || html.match(/itemprop="genre"[^>]*content="([^"]+)"/i)
                       || html.match(/"genre"\s*:\s*"([^"]+)"/i)
                       || html.match(/itemprop="genre"[^>]*>([^<]+)</i);

    if (categoryMatch) {
        return { vertical: mapStoreCategory(categoryMatch[1]), appName };
    }
    return { vertical: 'other', appName };
}

// ── Row Management ─────────────────────────────────────────────────────────────

function addRow() {
    rowCounter++;
    const id = rowCounter;
    const tr = document.createElement('tr');
    tr.dataset.rowId = id;
    const mode = getMode();

    const baseCols = `
        <td><input type="text" data-field="channel" placeholder="e.g. Xiaomi" class="input-channel"></td>
        <td>
            <select data-field="platform">
                <option value="Android">Android</option>
                <option value="iOS">iOS</option>
            </select>
        </td>
        <td><input type="text" data-field="geo" list="geo-options" placeholder="Select country" autocomplete="off"></td>
        <td><input type="text" data-field="period" value="1 month" style="width:80px"></td>`;

    if (getTrafficType() === 'web') {
        tr.innerHTML = baseCols + `
        <td><input type="text" inputmode="decimal" data-field="purchases" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpp" placeholder="Cost per purchase"></td>
        <td><input type="text" inputmode="decimal" data-field="budget" placeholder="Budget"></td>
        <td><button class="btn-delete" title="Remove row">&times;</button></td>`;
    } else if (mode === 'installs') {
        tr.innerHTML = baseCols + `
        <td><input type="text" inputmode="decimal" data-field="installs" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpi" placeholder="CPI" class="editable-calc"></td>
        <td><input type="text" inputmode="decimal" data-field="budget" placeholder="Budget"></td>
        <td><input type="text" inputmode="decimal" data-field="views" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpm" class="calculated" readonly tabindex="-1"></td>
        <td><span class="pct-wrap"><input type="text" inputmode="decimal" data-field="ctr" placeholder="e.g. 1.2" class="editable-calc"><span class="pct-sign">%</span></span></td>
        <td><input type="text" inputmode="decimal" data-field="clicks" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpc" class="calculated" readonly tabindex="-1"></td>
        <td><span class="pct-wrap"><input type="text" inputmode="decimal" data-field="crPurchase" placeholder="e.g. 5"><span class="pct-sign">%</span></span></td>
        <td><input type="text" inputmode="decimal" data-field="purchases" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpp" class="calculated" readonly tabindex="-1"></td>
        <td><button class="btn-delete" title="Remove row">&times;</button></td>`;
    } else if (isCommissionMode()) {
        tr.innerHTML = baseCols + `
        <td><input type="text" inputmode="decimal" data-field="views" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpm" class="calculated" readonly tabindex="-1"></td>
        <td><span class="pct-wrap"><input type="text" inputmode="decimal" data-field="ctr" placeholder="e.g. 1.2" class="editable-calc"><span class="pct-sign">%</span></span></td>
        <td><input type="text" inputmode="decimal" data-field="clicks" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpc" class="calculated" readonly tabindex="-1"></td>
        <td><span class="pct-wrap"><input type="text" inputmode="decimal" data-field="crInstall" placeholder="e.g. 3" class="editable-calc"><span class="pct-sign">%</span></span></td>
        <td><input type="text" inputmode="decimal" data-field="installs" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpi" placeholder="CPI" class="editable-calc"></td>
        <td><span class="pct-wrap"><input type="text" inputmode="decimal" data-field="crPurchase" placeholder="e.g. 5"><span class="pct-sign">%</span></span></td>
        <td><input type="text" inputmode="decimal" data-field="purchases" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpp" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="budget" placeholder="Total Cost"></td>
        <td><button class="btn-delete" title="Remove row">&times;</button></td>`;
    } else {
        tr.innerHTML = baseCols + `
        <td><input type="text" inputmode="decimal" data-field="events" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpa" placeholder="CPA"></td>
        <td><input type="text" inputmode="decimal" data-field="budget" placeholder="Budget"></td>
        <td><input type="text" inputmode="decimal" data-field="views" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpm" class="calculated" readonly tabindex="-1"></td>
        <td><span class="pct-wrap"><input type="text" inputmode="decimal" data-field="ctr" placeholder="e.g. 1.2" class="editable-calc"><span class="pct-sign">%</span></span></td>
        <td><input type="text" inputmode="decimal" data-field="clicks" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpc" class="calculated" readonly tabindex="-1"></td>
        <td><span class="pct-wrap"><input type="text" inputmode="decimal" data-field="crInstall" placeholder="e.g. 5" class="editable-calc"><span class="pct-sign">%</span></span></td>
        <td><input type="text" inputmode="decimal" data-field="installs" class="calculated" readonly tabindex="-1"></td>
        <td><input type="text" inputmode="decimal" data-field="cpi" placeholder="CPI" class="editable-calc"></td>
        <td><button class="btn-delete" title="Remove row">&times;</button></td>`;
    }

    // Wire events
    const channelInput = tr.querySelector('[data-field="channel"]');
    channelInput.addEventListener('blur', () => applyBenchmarks(tr));

    tr.querySelectorAll('input[data-field]').forEach(inp => {
        if (inp.dataset.field === 'channel' || inp.dataset.field === 'platform' || inp.dataset.field === 'geo' || inp.dataset.field === 'period') return;
        inp.addEventListener('input', () => recalcRow(tr));
    });
    // Format editable numeric fields with spaces
    tr.querySelectorAll('input[data-field="budget"], input[data-field="cpa"], input[data-field="cpi"], input[data-field="purchases"], input[data-field="cpp"], input[data-field="crPurchase"], input[data-field="crInstall"], input[data-field="ctr"]').forEach(wireNumericFormat);

    const geoInput = tr.querySelector('[data-field="geo"]');
    geoInput.addEventListener('change', () => {
        updateVatVisibility();
        ['ctr','cpi','crInstall'].forEach(f => {
            const el = tr.querySelector(`[data-field="${f}"]`);
            if (el) el.dataset.auto = '1';
        });
        applyBenchmarks(tr);
    });
    geoInput.addEventListener('blur', () => { updateVatVisibility(); applyBenchmarks(tr); });
    geoInput.addEventListener('input', () => {
        const code = extractGeoCode(geoInput.value.trim());
        if (code && code.length === 2) {
            ['ctr','cpi','crInstall'].forEach(f => {
                const el = tr.querySelector(`[data-field="${f}"]`);
                if (el) el.dataset.auto = '1';
            });
            applyBenchmarks(tr);
        }
    });

    const ctrInput = tr.querySelector('[data-field="ctr"]');
    const cpiInput = tr.querySelector('[data-field="cpi"]');
    if (ctrInput) ctrInput.addEventListener('input', () => { delete ctrInput.dataset.auto; });
    if (cpiInput) cpiInput.addEventListener('input', () => { delete cpiInput.dataset.auto; });

    if (mode === 'purchases' || isCommissionMode()) {
        const crInstallInput = tr.querySelector('[data-field="crInstall"]');
        if (crInstallInput) crInstallInput.addEventListener('input', () => { delete crInstallInput.dataset.auto; });
    }

    tr.querySelector('.btn-delete').addEventListener('click', () => {
        tr.remove();
        recalcTotals();
        updateVatVisibility();
    });

    tbody.appendChild(tr);
    channelInput.focus();
}

// ── 4D Benchmarks (Channel × Vertical × GEO × Platform) ──────────────────────

function applyBenchmarks(tr) {
    const channel = getField(tr, 'channel').toLowerCase().trim();
    const geoRaw = getField(tr, 'geo').trim();
    const platform = getField(tr, 'platform') || 'Android';
    const vertical = verticalSelect.value;
    const mode = getMode();

    const ctrInput = tr.querySelector('[data-field="ctr"]');
    const cpiInput = tr.querySelector('[data-field="cpi"]');
    const crInstallInput = tr.querySelector('[data-field="crInstall"]'); // null in installs mode

    const ctrCanSet = !ctrInput.value || ctrInput.dataset.auto === '1';
    const cpiCanSet = !cpiInput.value || cpiInput.dataset.auto === '1';
    const crCanSet = crInstallInput ? (!crInstallInput.value || crInstallInput.dataset.auto === '1') : false;

    if (!ctrCanSet && !cpiCanSet && !crCanSet) return;

    // 1. Find channel base
    let base = CHANNEL_BENCHMARKS._default;
    for (const [key, val] of Object.entries(CHANNEL_BENCHMARKS)) {
        if (key !== '_default' && channel.includes(key)) {
            base = val;
            break;
        }
    }

    // 2. Get vertical multiplier
    const vertMult = VERTICAL_MULTIPLIERS[vertical] || VERTICAL_MULTIPLIERS.other;

    // 3. Get GEO multiplier (per-country or regional fallback)
    const geoCode = extractGeoCode(geoRaw);
    const geoMult = getGeoMultipliers(geoCode);

    // 4. Get platform multiplier
    const platMult = PLATFORM_MULTIPLIERS[platform] || PLATFORM_MULTIPLIERS.Android;

    // 5. Calculate final values (4D: base × vertical × geo × platform)
    const finalCtr = roundBenchmark(base.ctr * vertMult.ctr * geoMult.ctr * platMult.ctr);
    const cpiMult = (vertical === 'finance') ? 1.47 : 2.35;
    const curRate = CURRENCY_RATES[currentCurrency] || 1; // convert from USD to current currency
    const finalCpi = Math.round(base.cpi * vertMult.cpi * geoMult.cpi * cpiMult * platMult.cpi * curRate * 10) / 10;

    // 6. Apply
    if (ctrCanSet && !isNoViewsSource(tr)) {
        ctrInput.value = fmtSpaces(finalCtr);
        ctrInput.dataset.auto = '1';
    }
    if (cpiCanSet) {
        cpiInput.value = fmtSpaces(finalCpi);
        cpiInput.dataset.auto = '1';
    }
    if (crCanSet && crInstallInput) {
        const finalCr = roundBenchmark(base.crInstall * vertMult.crInstall * geoMult.crInstall * platMult.crInstall);
        crInstallInput.value = fmtSpaces(finalCr);
        crInstallInput.dataset.auto = '1';
    }

    recalcRow(tr);
}

/**
 * Re-apply benchmarks to all rows (called when vertical changes).
 */
function reapplyAllBenchmarks() {
    tbody.querySelectorAll('tr').forEach(tr => applyBenchmarks(tr));
}

// ── Calculations ───────────────────────────────────────────────────────────────
//
// PURCHASES mode:
//   installs = budget / cpi
//   clicks   = installs / crInstall
//   cpc      = budget / clicks
//   views    = clicks / ctr
//   cpm      = (budget / views) × 1000
//   events   = budget / cpa
//
// INSTALLS mode:
//   installs = budget / cpi
//   clicks   = installs / crInstall (crInstall from benchmarks)
//   cpc      = budget / clicks
//   views    = clicks / ctr
//   cpm      = (budget / views) × 1000
//   purchases = installs × crPurchase
//   cpp       = budget / purchases

function isNoViewsSource(tr) {
    return NO_VIEWS_SOURCES.includes(tr.dataset.sourceKey || '');
}

function setDash(tr, name) {
    const el = tr.querySelector(`[data-field="${name}"]`);
    if (el) el.value = '—';
}

function recalcRow(tr) {
    if (getTrafficType() === 'web') {
        const budget = numField(tr, 'budget');
        const cpp = numField(tr, 'cpp');
        const purchases = (cpp > 0) ? Math.round(budget / cpp) : 0;
        setCalc(tr, 'purchases', purchases);
        recalcTotals();
        return;
    }

    const mode = getMode();
    const budget = numField(tr, 'budget');
    const cpi = numField(tr, 'cpi');
    const noViews = isNoViewsSource(tr);

    const effectiveBudget = budget;

    // installs = effectiveBudget / cpi
    const installs = (cpi > 0) ? effectiveBudget / cpi : 0;
    setCalc(tr, 'installs', installs);

    if (isCommissionMode()) {
        const crInstallPct = numField(tr, 'crInstall'); // visible field
        const crInstall = crInstallPct / 100;

        // clicks and cpc calculated from visible crInstall field
        const clicks = (crInstall > 0) ? installs / crInstall : 0;
        setCalc(tr, 'clicks', clicks);
        const cpc = (clicks > 0) ? effectiveBudget / clicks : 0;
        setCalc(tr, 'cpc', cpc);

        if (noViews) {
            setDash(tr, 'views');
            setDash(tr, 'cpm');
            setDash(tr, 'ctr');
        } else {
            const ctrPct = numField(tr, 'ctr');
            const ctr = ctrPct / 100;
            const views = (ctr > 0) ? clicks / ctr : 0;
            setCalc(tr, 'views', views);
            const cpm = (views > 0) ? (effectiveBudget / views) * 1000 : 0;
            setCalc(tr, 'cpm', cpm);
        }

        // purchases = installs × crPurchase
        const crPurchasePct = numField(tr, 'crPurchase');
        const crPurchase = crPurchasePct / 100;
        const purchases = installs * crPurchase;
        setCalc(tr, 'purchases', Math.round(purchases));

        // cpp = effectiveBudget / purchases
        const cpp = (purchases > 0) ? effectiveBudget / purchases : 0;
        setCalc(tr, 'cpp', cpp);
    } else if (mode === 'installs') {
        const crInstallBenchmark = getInternalCrInstall(tr);

        // clicks and cpc always calculated from CR install benchmark
        const clicks = (crInstallBenchmark > 0) ? installs / crInstallBenchmark : 0;
        setCalc(tr, 'clicks', clicks);
        const cpc = (clicks > 0) ? effectiveBudget / clicks : 0;
        setCalc(tr, 'cpc', cpc);

        if (noViews) {
            setDash(tr, 'views');
            setDash(tr, 'cpm');
            setDash(tr, 'ctr');
        } else {
            const ctrPct = numField(tr, 'ctr');
            const ctr = ctrPct / 100;
            const views = (ctr > 0) ? clicks / ctr : 0;
            setCalc(tr, 'views', views);
            const cpm = (views > 0) ? (effectiveBudget / views) * 1000 : 0;
            setCalc(tr, 'cpm', cpm);
        }

        // purchases = installs × crPurchase
        const crPurchasePctI = numField(tr, 'crPurchase');
        const crPurchaseI = crPurchasePctI / 100;
        const purchasesI = installs * crPurchaseI;
        setCalc(tr, 'purchases', Math.round(purchasesI));

        // cpp = effectiveBudget / purchases
        const cppI = (purchasesI > 0) ? effectiveBudget / purchasesI : 0;
        setCalc(tr, 'cpp', cppI);
    } else {
        // Purchases mode
        const crInstallPct = numField(tr, 'crInstall');
        const crInstall = crInstallPct / 100;

        // clicks and cpc always calculated from CR install
        const clicks = (crInstall > 0) ? installs / crInstall : 0;
        setCalc(tr, 'clicks', clicks);
        const cpc = (clicks > 0) ? budget / clicks : 0;
        setCalc(tr, 'cpc', cpc);

        if (noViews) {
            setDash(tr, 'views');
            setDash(tr, 'cpm');
            setDash(tr, 'ctr');
        } else {
            const ctrPct = numField(tr, 'ctr');
            const ctr = ctrPct / 100;
            const views = (ctr > 0) ? clicks / ctr : 0;
            setCalc(tr, 'views', views);
            const cpm = (views > 0) ? (budget / views) * 1000 : 0;
            setCalc(tr, 'cpm', cpm);
        }

        const cpa = numField(tr, 'cpa');
        const events = (cpa > 0) ? budget / cpa : 0;
        setCalc(tr, 'events', Math.round(events));
    }

    recalcTotals();
}

/**
 * Get CR install/click value for installs mode (from benchmarks, computed internally).
 * In installs mode there is no crInstall field in the row, so we compute it from benchmarks.
 */
function getInternalCrInstall(tr) {
    const channel = getField(tr, 'channel').toLowerCase().trim();
    const geoRaw = getField(tr, 'geo').trim();
    const platform = getField(tr, 'platform') || 'Android';
    const vertical = verticalSelect.value;

    let base = CHANNEL_BENCHMARKS._default;
    for (const [key, val] of Object.entries(CHANNEL_BENCHMARKS)) {
        if (key !== '_default' && channel.includes(key)) {
            base = val;
            break;
        }
    }
    const vertMult = VERTICAL_MULTIPLIERS[vertical] || VERTICAL_MULTIPLIERS.other;
    const geoCode = extractGeoCode(geoRaw);
    const geoMult = getGeoMultipliers(geoCode);
    const platMult = PLATFORM_MULTIPLIERS[platform] || PLATFORM_MULTIPLIERS.Android;

    const crPct = roundBenchmark(base.crInstall * vertMult.crInstall * geoMult.crInstall * platMult.crInstall);
    return crPct / 100; // convert from percent to ratio
}

function recalcTotals() {
    const mode = getMode();
    let totalCost = 0, totalInstalls = 0, totalClicks = 0, totalViews = 0;

    tbody.querySelectorAll('tr').forEach(tr => {
        totalCost += numField(tr, 'budget');
        totalInstalls += numField(tr, 'installs');
        totalClicks += numField(tr, 'clicks');
        totalViews += numField(tr, 'views');
    });

    const sym = CURRENCY_SYMBOLS[currencySelect.value] || '';

    const costEl = document.getElementById('total-cost');
    const installsEl = document.getElementById('total-installs');
    const clicksEl = document.getElementById('total-clicks');
    const viewsEl = document.getElementById('total-views');

    if (costEl) costEl.textContent = sym + fmtNum(totalCost);
    if (installsEl) installsEl.textContent = fmtNum(totalInstalls);
    if (clicksEl) clicksEl.textContent = fmtNum(totalClicks);
    if (viewsEl) viewsEl.textContent = fmtNum(totalViews);

    if (getTrafficType() === 'web' || mode === 'installs' || isCommissionMode()) {
        let totalPurchases = 0;
        tbody.querySelectorAll('tr').forEach(tr => {
            totalPurchases += numField(tr, 'purchases');
        });
        const purchasesEl = document.getElementById('total-purchases');
        if (purchasesEl) purchasesEl.textContent = fmtNum(Math.round(totalPurchases));
    } else {
        let totalEvents = 0;
        tbody.querySelectorAll('tr').forEach(tr => {
            totalEvents += numField(tr, 'events');
        });
        const eventsEl = document.getElementById('total-events');
        if (eventsEl) eventsEl.textContent = fmtNum(Math.round(totalEvents));
    }

    // Cost summary
    const hasRu = Array.from(tbody.querySelectorAll('[data-field="geo"]'))
        .some(inp => extractGeoCode(inp.value) === 'RU');
    document.getElementById('vat-row-line').style.display = hasRu ? 'flex' : 'none';
    updateVat(totalCost, hasRu);
}

function recalcAll() {
    tbody.querySelectorAll('tr').forEach(tr => recalcRow(tr));
}

// ── Currency Conversion ─────────────────────────────────────────────────────

/** Fields that hold monetary values and need conversion */
const MONEY_FIELDS = ['budget', 'cpa', 'cpi', 'cpc', 'cpm', 'cpp'];

function onCurrencyChange() {
    const newCurrency = currencySelect.value;
    const oldRate = CURRENCY_RATES[currentCurrency] || 1;
    const newRate = CURRENCY_RATES[newCurrency] || 1;
    const factor = newRate / oldRate; // e.g. USD→RUB: 96/1 = 96

    tbody.querySelectorAll('tr').forEach(tr => {
        MONEY_FIELDS.forEach(field => {
            const el = tr.querySelector(`[data-field="${field}"]`);
            if (!el) return;
            const raw = el.value.replace(/\s/g, '').replace('—', '');
            const num = parseFloat(raw);
            if (!num || isNaN(num)) return;
            const converted = num * factor;
            const rounded = roundSmart(converted);
            el.value = fmtSpaces(rounded);
        });
    });

    currentCurrency = newCurrency;
    recalcTotals();
}

// ── VAT ────────────────────────────────────────────────────────────────────────

function updateVatVisibility() {
    const hasRu = Array.from(tbody.querySelectorAll('[data-field="geo"]'))
        .some(inp => extractGeoCode(inp.value) === 'RU');
    let totalCost = 0;
    tbody.querySelectorAll('tr').forEach(tr => {
        totalCost += numField(tr, 'budget');
    });
    updateVat(totalCost, hasRu);
}

function updateVat(totalBudget, hasRu) {
    const sym = CURRENCY_SYMBOLS[currencySelect.value] || '';
    const commissionRowLine = document.getElementById('commission-row-line');
    const vatRowLine = document.getElementById('vat-row-line');
    const netLabel = document.getElementById('vat-net-label');

    if (isCommissionMode()) {
        if (vatRowLine) vatRowLine.style.display = 'none';
        if (commissionRowLine) commissionRowLine.style.display = 'flex';
        const pct = getCommissionPct();
        const netSpend = totalBudget;
        const commissionAmt = netSpend * pct / 100;
        const grossTotal = netSpend + commissionAmt;
        if (netLabel) netLabel.textContent = 'Net Media Spend';
        document.getElementById('vat-net').textContent = sym + fmtNum(netSpend);
        document.getElementById('commission-amount').textContent = sym + fmtNum(commissionAmt);
        document.getElementById('vat-gross').textContent = sym + fmtNum(grossTotal);
    } else {
        if (commissionRowLine) commissionRowLine.style.display = 'none';
        if (vatRowLine) vatRowLine.style.display = hasRu ? 'flex' : 'none';
        if (netLabel) netLabel.textContent = 'Max Placement Cost Net';
        const vat = hasRu ? totalBudget * 0.2 : 0;
        const gross = totalBudget + vat;
        document.getElementById('vat-net').textContent = sym + fmtNum(totalBudget);
        document.getElementById('vat-amount').textContent = sym + fmtNum(vat);
        document.getElementById('vat-gross').textContent = sym + fmtNum(gross);
    }
}

// ── Helpers ────────────────────────────────────────────────────────────────────

function getField(tr, name) {
    const el = tr.querySelector(`[data-field="${name}"]`);
    return el ? el.value : '';
}

function numField(tr, name) {
    const raw = getField(tr, name).replace(/\s/g, '');
    return parseFloat(raw) || 0;
}

function setCalc(tr, name, value) {
    const el = tr.querySelector(`[data-field="${name}"]`);
    if (el) el.value = value > 0 ? fmtSpaces(roundSmart(value)) : '';
}

function roundSmart(n) {
    if (n >= 1000) return Math.round(n);
    if (n >= 1) return Math.round(n * 100) / 100;
    return Math.round(n * 10000) / 10000;
}

/** Format number with space as thousands separator */
function fmtSpaces(n) {
    const parts = String(n).split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    return parts.join('.');
}

/** Wire blur/focus formatting on an editable numeric input */
function wireNumericFormat(input) {
    input.addEventListener('focus', () => {
        input.value = input.value.replace(/\s/g, '');
    });
    input.addEventListener('blur', () => {
        const raw = input.value.replace(/\s/g, '');
        if (raw !== '' && !isNaN(raw)) {
            input.value = fmtSpaces(raw);
        }
    });
}

function fmtNum(n) {
    if (n === 0) return '0';
    return n.toLocaleString('en-US', {
        maximumFractionDigits: 2,
        minimumFractionDigits: 0
    });
}

// ── Reset ──────────────────────────────────────────────────────────────────────

function resetAll() {
    if (!confirm('Reset all data?')) return;
    document.getElementById('client').value = '';
    document.getElementById('campaign').value = 'MobX Agency';
    document.getElementById('period').value = '1 month';
    cpaEventSelect.value = '';
    cpaEventCustom.value = '';
    cpaEventCustom.style.display = 'none';
    currencySelect.value = 'USD';
    currentCurrency = 'USD';
    appLinkAndroid.value = '';
    appLinkIos.value = '';
    verticalSelect.value = 'other';
    appLinkAndroidStatus.textContent = '';
    appLinkAndroidStatus.className = 'app-link-status';
    appLinkIosStatus.textContent = '';
    appLinkIosStatus.className = 'app-link-status';
    // Clear GEO multi-select
    selectedGeos = [];
    renderGeoTags();
    geoAddInput.value = '';
    // Uncheck all source checkboxes
    document.querySelectorAll('#source-select input[type="checkbox"]').forEach(cb => {
        cb.checked = false;
    });
    tbody.innerHTML = '';
    rowCounter = 0;
    renderTableStructure();
    recalcTotals();
    updateVatVisibility();
}

// ── Excel Export (using xlsx-js-style for cell styling) ─────────────────────────

function exportToExcel() {
    const client = document.getElementById('client').value || '';
    const campaign = document.getElementById('campaign').value || '';
    const docTitle = 'Internet placement proposal';
    const period = document.getElementById('period').value || '';
    const cpaEvent = getCpaEventValue() || 'events';
    const currency = currencySelect.value;
    const verticalLabel = verticalSelect.options[verticalSelect.selectedIndex].text;

    const mode = getMode();

    // Collect rows
    const rows = [];
    tbody.querySelectorAll('tr').forEach(tr => {
        const noViews = isNoViewsSource(tr);
        const base = {
            channel: getField(tr, 'channel'),
            platform: getField(tr, 'platform'),
            geo: extractGeoCode(getField(tr, 'geo')) || getField(tr, 'geo'),
            period: getField(tr, 'period'),
            budget: numField(tr, 'budget'),
            views: noViews ? 0 : numField(tr, 'views'),
            ctr: noViews ? 0 : numField(tr, 'ctr') / 100,
            clicks: numField(tr, 'clicks'),
            cpc: numField(tr, 'cpc'),
            installs: numField(tr, 'installs'),
            cpm: 0,
            cpi: numField(tr, 'cpi'),
            noViews: noViews
        };
        if (!noViews && numField(tr, 'views') > 0) {
            base.cpm = (numField(tr, 'budget') / numField(tr, 'views')) * 1000;
        }
        if (getTrafficType() === 'web') {
            base.purchases = numField(tr, 'purchases');
            base.cpp = numField(tr, 'cpp');
        } else if (isCommissionMode()) {
            base.crInstall = numField(tr, 'crInstall') / 100; // from visible field
            base.crPurchase = numField(tr, 'crPurchase') / 100;
            base.purchases = numField(tr, 'purchases');
            base.cpp = numField(tr, 'cpp');
        } else if (mode === 'installs') {
            base.crPurchase = numField(tr, 'crPurchase') / 100;
            base.purchases = numField(tr, 'purchases');
            base.cpp = numField(tr, 'cpp');
            base.crInstallBenchmark = getInternalCrInstall(tr); // ratio for formula
            base.commFactor = 1;
        } else {
            base.events = numField(tr, 'events');
            base.cpa = numField(tr, 'cpa');
            base.crInstall = numField(tr, 'crInstall') / 100;
        }
        rows.push(base);
    });

    const hasRu = rows.some(r => r.geo === 'RU');

    // ── Currency number format for Excel ──
    const CURRENCY_NUMFMT = {
        USD: '#,##0\\ "$"',
        EUR: '#,##0\\ "\u20AC"',
        RUB: '#,##0\\ "\u20BD"',
        GBP: '#,##0\\ "\u00A3"',
        KZT: '#,##0\\ "\u20B8"',
        UAH: '#,##0\\ "\u20B4"',
        BRL: '#,##0\\ "R$"',
        INR: '#,##0\\ "\u20B9"',
    };
    const CURRENCY_NUMFMT_DEC = {
        USD: '#,##0.00\\ "$"',
        EUR: '#,##0.00\\ "\u20AC"',
        RUB: '#,##0.00\\ "\u20BD"',
        GBP: '#,##0.00\\ "\u00A3"',
        KZT: '#,##0.00\\ "\u20B8"',
        UAH: '#,##0.00\\ "\u20B4"',
        BRL: '#,##0.00\\ "R$"',
        INR: '#,##0.00\\ "\u20B9"',
    };
    const curFmt = CURRENCY_NUMFMT[currency] || '#,##0';
    const curFmtDec = CURRENCY_NUMFMT_DEC[currency] || '#,##0.00';

    // ── Styles ──
    const border = {
        top:    { style: 'thin', color: { rgb: 'B0B0B0' } },
        bottom: { style: 'thin', color: { rgb: 'B0B0B0' } },
        left:   { style: 'thin', color: { rgb: 'B0B0B0' } },
        right:  { style: 'thin', color: { rgb: 'B0B0B0' } },
    };

    const sLabel = { font: { bold: true, sz: 11, color: { rgb: '000000' } } };
    const sValue = { font: { sz: 11, color: { rgb: '000000' } } };

    const sHeader = {
        font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '4285F4' } },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border: border
    };

    const sData = {
        font: { sz: 11, color: { rgb: '000000' } },
        alignment: { horizontal: 'center' },
        border: border
    };

    const sDataBold = {
        font: { bold: true, sz: 10, color: { rgb: '000000' } },
        alignment: { horizontal: 'left' },
        border: border
    };

    const sPercent = {
        font: { sz: 11, color: { rgb: '000000' } },
        alignment: { horizontal: 'center' },
        border: border,
        numFmt: '0.00%'
    };

    const sNumber = {
        font: { sz: 11, color: { rgb: '000000' } },
        alignment: { horizontal: 'center' },
        border: border,
        numFmt: '#,##0'
    };

    // Currency: integer (for CPA, Budget, Total cost)
    const sCurrency = {
        font: { sz: 11, color: { rgb: '000000' } },
        alignment: { horizontal: 'center' },
        border: border,
        numFmt: curFmt
    };

    // Currency: decimal (for CPI, CPC)
    const sCurrencyDec = {
        font: { sz: 11, color: { rgb: '000000' } },
        alignment: { horizontal: 'center' },
        border: border,
        numFmt: curFmtDec
    };

    const sTotal = {
        font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '4285F4' } },
        alignment: { horizontal: 'center' },
        border: border
    };

    const sTotalNum = {
        font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '4285F4' } },
        alignment: { horizontal: 'center' },
        border: border,
        numFmt: '#,##0'
    };

    const sTotalCur = {
        font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '4285F4' } },
        alignment: { horizontal: 'center' },
        border: border,
        numFmt: curFmt
    };

    const sGray = {
        font: { sz: 11, color: { rgb: '000000' } },
        fill: { fgColor: { rgb: 'BFBFBF' } }
    };
    const sGrayCur = {
        font: { sz: 11, color: { rgb: '000000' } },
        fill: { fgColor: { rgb: 'BFBFBF' } },
        numFmt: curFmt
    };
    const sBlue = {
        font: { bold: true, sz: 11, color: { rgb: '000000' } },
        fill: { fgColor: { rgb: 'BDD6EE' } }
    };
    const sBlueCur = {
        font: { bold: true, sz: 11, color: { rgb: '000000' } },
        fill: { fgColor: { rgb: 'BDD6EE' } },
        numFmt: curFmt
    };

    // Empty separator style (no borders, no fill — just a blank gap)
    const sSep = {};

    // ── Helpers ──
    function sc(v, s) { // string cell
        return { v: v || '', t: 's', s: s || {} };
    }
    function nc(v, s) { // number cell
        return { v: v || 0, t: 'n', s: s || {} };
    }
    function fc(formula, s) { // formula cell
        return { f: formula, t: 'n', s: s || {} };
    }
    function empty(s) {
        return { v: '', t: 's', s: s || {} };
    }

    // ── Build worksheet cell-by-cell ──
    const ws = {};

    function put(r, c, cell) {
        ws[XLSX.utils.encode_cell({ r, c })] = cell;
    }

    // Rows 0-3: empty (logo area)
    // Row 4: Client
    put(4, 0, sc('Client:', sLabel));
    put(4, 2, sc(client, sValue));
    // Row 5: Campaign/Agency
    put(5, 0, sc('Campaign/Agency:', sLabel));
    put(5, 2, sc(campaign, sValue));
    // Row 6: Document
    put(6, 0, sc('Document:', sLabel));
    put(6, 2, sc(docTitle, sValue));
    // Row 7: Period
    put(7, 0, sc('Period:', sLabel));
    put(7, 2, sc(period, sValue));
    // Row 8: Vertical
    put(8, 0, sc('Vertical:', sLabel));
    put(8, 2, sc(verticalLabel, sValue));

    // Row 10: Headers
    let headerTexts, COLS, sepCols;
    if (getTrafficType() === 'web') {
        COLS = 9;
        sepCols = new Set([4]);
        headerTexts = [
            'Channel', 'Platform', 'Targeting', 'Period', '',
            'Total Purchases', 'Cost per Purchase', '', 'Total Budget'
        ];
    } else if (isCommissionMode()) {
        COLS = 18;
        sepCols = new Set([4, 16]);
        headerTexts = [
            'Channel', 'Platform', 'Targeting', 'Period', '',
            'Views', 'CPM', 'CTR', 'Total Clicks', 'CPC',
            'CR install to install', 'Total Installs', 'CPI',
            'CR install to purchase', 'Total Purchases', 'Cost per purchase',
            '', 'Total Cost'
        ];
    } else if (mode === 'installs') {
        COLS = 19;
        sepCols = new Set([4, 10]);
        headerTexts = [
            'Channel', 'Platform', 'Targeting', 'Period', '',
            'Total installs', 'CPI', '', '',
            'Total cost', '',
            'Views', 'CPM', 'CTR', 'Total clicks', 'CPC',
            'CR install to purchase', 'Total purchases', 'Cost per purchase'
        ];
    } else {
        COLS = 19;
        sepCols = new Set([4, 10]);
        headerTexts = [
            'Channel', 'Platform', 'Targeting', 'Period', '',
            'Total Purchases', 'Cost per Purchase', '', '',
            'Total cost', '',
            'Views', 'CPM', 'CTR', 'Total clicks', 'CPC',
            'CR install per click', 'Total installs', 'CPI'
        ];
    }
    for (let c = 0; c < COLS; c++) {
        put(10, c, sc(headerTexts[c] || '', sepCols.has(c) ? sSep : sHeader));
    }

    // Row 11+: Data rows (with Excel formulas)
    const dataStartRow = 11;
    const firstR = dataStartRow + 1; // Excel 1-indexed first data row
    rows.forEach((r, i) => {
        const row = dataStartRow + i;
        const R = row + 1; // Excel 1-indexed row number
        const excelChannel = r.channel.replace(/\s*CoDev$/i, '');
        put(row, 0,  sc(excelChannel, sDataBold));
        put(row, 1,  sc(r.platform, sData));
        put(row, 2,  sc(r.geo, sData));
        put(row, 3,  sc(r.period, sData));
        put(row, 4,  empty(sSep));

        if (getTrafficType() === 'web') {
            // F=Total Purchases (formula), G=CPP (input), I=Budget (input)
            put(row, 5, fc(`ROUND(I${R}/G${R},0)`, sNumber));   // F: purchases
            put(row, 6, nc(r.cpp, sCurrencyDec));               // G: CPP (input)
            put(row, 7, empty(sSep));
            put(row, 8, nc(r.budget, sCurrency));               // I: Budget (input)
        } else if (isCommissionMode()) {
            // Commission: new column order
            // Input: H=CTR, K=CRinst, M=CPI, N=CRpur, R=TotalCost
            // Formulas: F=views, G=cpm, I=clicks, J=cpc, L=installs, O=purchases, P=cpp
            if (r.noViews) {
                put(row, 5,  sc('—', sData));                                   // F: views — n/a
                put(row, 6,  sc('—', sData));                                   // G: cpm — n/a
                put(row, 7,  sc('—', sData));                                   // H: CTR — n/a
            } else {
                put(row, 5,  fc(`I${R}/H${R}`, sNumber));                       // F: views = clicks/ctr
                put(row, 6,  fc(`(R${R}/F${R})*1000`, sCurrencyDec));           // G: cpm = (budget/views)*1000
                put(row, 7,  nc(r.ctr, sPercent));                              // H: CTR (input)
            }
            put(row, 8,  fc(`L${R}/K${R}`, sNumber));                           // I: clicks = installs/crInstall
            put(row, 9,  fc(`R${R}/I${R}`, sCurrencyDec));                      // J: cpc = budget/clicks
            put(row, 10, nc(r.crInstall, sPercent));                            // K: CR install (input)
            put(row, 11, fc(`R${R}/M${R}`, sNumber));                           // L: installs = budget/cpi
            put(row, 12, nc(r.cpi, sCurrencyDec));                              // M: CPI (input)
            put(row, 13, nc(r.crPurchase, sPercent));                           // N: CR purchase (input)
            put(row, 14, fc(`L${R}*N${R}`, sNumber));                           // O: purchases = installs*crPurchase
            put(row, 15, fc(`R${R}/O${R}`, sCurrencyDec));                      // P: cpp = budget/purchases
            put(row, 16, empty(sSep));                                          // Q: sep
            put(row, 17, nc(r.budget, sCurrency));                              // R: Total Cost (input)
        } else if (mode === 'installs') {
            // Input: G=CPI, J=Budget, N=CTR, Q=CR purchase
            // Formulas: F=installs, L=views, M=cpm, O=clicks, P=cpc, R=purchases, S=cpp
            put(row, 5,  fc(`J${R}/G${R}`, sNumber));             // F: installs = budget/cpi
            put(row, 6,  nc(r.cpi, sCurrencyDec));                // G: CPI (input)
            put(row, 7,  sc('/', sData));
            put(row, 8,  nc(1, sData));
            put(row, 9,  nc(r.budget, sCurrency));                // J: Budget (input)
            put(row, 10, empty(sSep));
            // Clicks & CPC always calculated (from CR install benchmark)
            const crBench = r.crInstallBenchmark;                 // ratio e.g. 0.03
            put(row, 14, fc(`F${R}/${crBench}`, sNumber));        // O: clicks = installs/crInstall
            put(row, 15, fc(`J${R}/O${R}`, sCurrencyDec));        // P: cpc = budget/clicks
            if (r.noViews) {
                put(row, 11, sc('—', sData));                     // L: views — n/a
                put(row, 12, sc('—', sData));                     // M: cpm — n/a
                put(row, 13, sc('—', sData));                     // N: CTR — n/a
            } else {
                put(row, 11, fc(`O${R}/N${R}`, sNumber));                      // L: views = clicks/ctr
                put(row, 12, fc(`(J${R}/L${R})*1000`, sCurrencyDec));          // M: cpm = (budget/views)*1000
                put(row, 13, nc(r.ctr, sPercent));                             // N: CTR (input)
            }
            put(row, 16, nc(r.crPurchase, sPercent));             // Q: CR purchase (input)
            put(row, 17, fc(`F${R}*Q${R}`, sNumber));             // R: purchases = installs*crPurchase
            put(row, 18, fc(`J${R}/R${R}`, sCurrencyDec));        // S: cpp = budget/purchases
        } else {
            // Input: G=CPA, J=Budget, N=CTR, Q=CR install, S=CPI
            // Formulas: F=events, L=views, M=cpm, O=clicks, P=cpc, R=installs
            put(row, 5,  fc(`J${R}/G${R}`, sNumber));             // F: events = budget/cpa
            put(row, 6,  nc(r.cpa, sCurrency));                   // G: CPA (input)
            put(row, 7,  sc('/', sData));
            put(row, 8,  nc(1, sData));
            put(row, 9,  nc(r.budget, sCurrency));                // J: Budget (input)
            put(row, 10, empty(sSep));
            // Clicks & CPC always calculated (from CR install)
            put(row, 14, fc(`R${R}/Q${R}`, sNumber));             // O: clicks = installs/crInstall
            put(row, 15, fc(`J${R}/O${R}`, sCurrencyDec));        // P: cpc = budget/clicks
            if (r.noViews) {
                put(row, 11, sc('—', sData));                     // L: views — n/a
                put(row, 12, sc('—', sData));                     // M: cpm — n/a
                put(row, 13, sc('—', sData));                     // N: CTR — n/a
            } else {
                put(row, 11, fc(`O${R}/N${R}`, sNumber));             // L: views = clicks/ctr
                put(row, 12, fc(`(J${R}/L${R})*1000`, sCurrencyDec)); // M: cpm = (budget/views)*1000
                put(row, 13, nc(r.ctr, sPercent));                    // N: CTR (input)
            }
            put(row, 16, nc(r.crInstall, sPercent));              // Q: CR install (input)
            put(row, 17, fc(`J${R}/S${R}`, sNumber));             // R: installs = budget/cpi
            put(row, 18, nc(r.cpi, sCurrencyDec));                // S: CPI (input)
        }
    });

    // Totals row (SUM formulas)
    const totR = dataStartRow + rows.length;
    const lastDataR = dataStartRow + rows.length; // Excel 1-indexed last data row
    const totRx = totR + 1; // Excel 1-indexed totals row

    put(totR, 0, sc('Total', sTotal));
    for (let c = 1; c < COLS; c++) put(totR, c, empty(sepCols.has(c) ? sSep : sTotal));

    // SUM formulas for totals
    if (getTrafficType() === 'web') {
        put(totR, 5, fc(`SUM(F${firstR}:F${lastDataR})`, sTotalNum));  // Total purchases
        put(totR, 6, sc('—', sTotal));
        put(totR, 7, empty(sSep));
        put(totR, 8, fc(`SUM(I${firstR}:I${lastDataR})`, sTotalCur));  // Total budget
    } else if (isCommissionMode()) {
        put(totR, 5,  fc(`SUM(F${firstR}:F${lastDataR})`, sTotalNum));  // Total views
        put(totR, 8,  fc(`SUM(I${firstR}:I${lastDataR})`, sTotalNum));  // Total clicks
        put(totR, 11, fc(`SUM(L${firstR}:L${lastDataR})`, sTotalNum));  // Total installs
        put(totR, 14, fc(`SUM(O${firstR}:O${lastDataR})`, sTotalNum));  // Total purchases
        put(totR, 17, fc(`SUM(R${firstR}:R${lastDataR})`, sTotalCur));  // Total cost
    } else {
        put(totR, 5,  fc(`SUM(F${firstR}:F${lastDataR})`, sTotalNum));  // Total col F
        put(totR, 9,  fc(`SUM(J${firstR}:J${lastDataR})`, sTotalCur));  // Total budget
        put(totR, 11, fc(`SUM(L${firstR}:L${lastDataR})`, sTotalNum));  // Total views
        put(totR, 14, fc(`SUM(O${firstR}:O${lastDataR})`, sTotalNum));  // Total clicks
        put(totR, 17, fc(`SUM(R${firstR}:R${lastDataR})`, sTotalNum));  // Total col R
    }

    // Cost summary block (with formula references)
    let vr = totR + 2;
    const budgetRef = getTrafficType() === 'web' ? `I${totRx}` : isCommissionMode() ? `R${totRx}` : `J${totRx}`;
    let lastRow = vr;

    if (isCommissionMode()) {
        put(vr, 0, sc('Net Media Spend', sGray));
        put(vr, 1, empty(sGray));
        put(vr, 2, fc(budgetRef, sGrayCur));
        const netRowX = vr + 1;
        vr++;
        put(vr, 0, sc('Commission', sGray));
        put(vr, 1, empty(sGray));
        put(vr, 2, fc(`C${netRowX}*0`, sGrayCur));
        vr++;
        put(vr, 0, sc('Total Cost Gross', sBlue));
        put(vr, 1, empty(sBlue));
        put(vr, 2, fc(`C${netRowX}`, sBlueCur));
        lastRow = vr;
    } else {
        const netRow = vr + 1; // Excel 1-indexed
        put(vr, 0, sc('Max Placement Cost Net', sGray));
        put(vr, 1, empty(sGray));
        put(vr, 2, fc(budgetRef, sGrayCur)); // reference total budget
        lastRow = vr;

        if (hasRu) {
            vr++;
            put(vr, 0, sc('VAT', sGray));
            put(vr, 1, sc('20%', sGray));
            put(vr, 2, fc(`C${netRow}*0.2`, sGrayCur)); // VAT = net * 20%
            lastRow = vr;
        }

        vr++;
        put(vr, 0, sc('Total cost Gross', sBlue));
        put(vr, 1, empty(sBlue));
        if (hasRu) {
            put(vr, 2, fc(`C${netRow}*1.2`, sBlueCur)); // gross = net * 1.2
        } else {
            put(vr, 2, fc(`C${netRow}`, sBlueCur)); // gross = net (no VAT)
        }
        lastRow = vr;
    }

    // ── Worksheet metadata ──
    ws['!ref'] = XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: lastRow, c: COLS - 1 }
    });

    if (getTrafficType() === 'web') {
        ws['!cols'] = [
            { wch: 21 }, { wch: 12 }, { wch: 16 }, { wch: 8 }, { wch: 5 },
            { wch: 18 }, { wch: 18 }, { wch: 5 }, { wch: 20 },
        ];
        ws['!merges'] = [
            { s: { r: 6, c: 2 }, e: { r: 6, c: 4 } },
            { s: { r: totR, c: 0 }, e: { r: totR, c: 4 } },
        ];
    } else if (isCommissionMode()) {
        ws['!cols'] = [
            { wch: 21 }, { wch: 12 }, { wch: 16 }, { wch: 8 }, { wch: 5 },
            { wch: 14 }, { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 12 },
            { wch: 22 }, { wch: 14 }, { wch: 12 },
            { wch: 22 }, { wch: 16 }, { wch: 18 },
            { wch: 5 }, { wch: 20 },
        ];
        ws['!merges'] = [
            { s: { r: 6, c: 2 }, e: { r: 6, c: 4 } },       // Document merge
            { s: { r: totR, c: 0 }, e: { r: totR, c: 4 } },  // Total label merge
        ];
    } else {
        ws['!cols'] = [
            { wch: 21 }, { wch: 12 }, { wch: 16 }, { wch: 8 }, { wch: 5 },
            { wch: 15 }, { wch: 15 }, { wch: 3 }, { wch: 4 }, { wch: 23 },
            { wch: 4 },
            { wch: 15 }, { wch: 10 }, { wch: 10 }, { wch: 16 }, { wch: 10 },
            { wch: 20 }, { wch: 16 }, { wch: 16 },
        ];
        ws['!merges'] = [
            { s: { r: 6, c: 2 }, e: { r: 6, c: 4 } },       // Document merge
            { s: { r: 10, c: 3 }, e: { r: 10, c: 4 } },      // Period header (shifted)
            { s: { r: 10, c: 6 }, e: { r: 10, c: 8 } },      // Cost per event header (shifted)
            { s: { r: totR, c: 0 }, e: { r: totR, c: 4 } },   // Total label merge
        ];
    }

    // ── Create workbook & inject logo ──
    const wb = XLSX.utils.book_new();
    const sheetName = getTrafficType() === 'web'
        ? 'Web Traffic'
        : (cpaEvent ? `One ${cpaEvent.charAt(0).toUpperCase() + cpaEvent.slice(1)}` : 'Media Plan');
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    const fileName = `MediaPlan_${client || 'Draft'}_${new Date().toISOString().slice(0, 10)}.xlsx`;

    // Generate xlsx as array buffer, then inject logo via JSZip
    const xlsxData = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

    injectLogoAndDownload(xlsxData, fileName, sheetName);
}

// ── Logo injection into xlsx zip ────────────────────────────────────────────────

async function injectLogoAndDownload(xlsxArrayBuffer, fileName, sheetName) {
    const zip = await JSZip.loadAsync(xlsxArrayBuffer);

    // 1. Add the image file
    const logoBytes = base64ToUint8Array(LOGO_BASE64);
    zip.file('xl/media/image1.png', logoBytes);

    // 2. Add drawing1.xml (defines where the image sits)
    const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col>
      <xdr:colOff>85725</xdr:colOff>
      <xdr:row>0</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="2371725" cy="742950"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="Logo"/>
        <xdr:cNvPicPr preferRelativeResize="0"/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1" cstate="print"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData fLocksWithSheet="0"/>
  </xdr:oneCellAnchor>
</xdr:wsDr>`;
    zip.file('xl/drawings/drawing1.xml', drawingXml);

    // 3. Add drawing1.xml.rels (links drawing to image)
    const drawingRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`;
    zip.file('xl/drawings/_rels/drawing1.xml.rels', drawingRels);

    // 4. Add drawing relationship to sheet1.xml.rels
    const sheetRelsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
    let sheetRels;
    if (zip.file(sheetRelsPath)) {
        let existing = await zip.file(sheetRelsPath).async('string');
        sheetRels = existing.replace(
            '</Relationships>',
            '  <Relationship Id="rId_drawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>\n</Relationships>'
        );
    } else {
        sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId_drawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;
    }
    zip.file(sheetRelsPath, sheetRels);

    // 5. Add <drawing> reference to sheet1.xml
    const sheetPath = 'xl/worksheets/sheet1.xml';
    let sheetXml = await zip.file(sheetPath).async('string');
    if (!sheetXml.includes('<drawing')) {
        sheetXml = sheetXml.replace(
            '</worksheet>',
            '  <drawing r:id="rId_drawing1"/>\n</worksheet>'
        );
        if (!sheetXml.includes('xmlns:r=')) {
            sheetXml = sheetXml.replace(
                '<worksheet',
                '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
            );
        }
    }
    zip.file(sheetPath, sheetXml);

    // 6. Update [Content_Types].xml to include drawing and image types
    const ctPath = '[Content_Types].xml';
    let contentTypes = await zip.file(ctPath).async('string');
    if (!contentTypes.includes('image/jpeg')) {
        contentTypes = contentTypes.replace(
            '</Types>',
            '  <Default Extension="jpeg" ContentType="image/jpeg"/>\n</Types>'
        );
    }
    if (!contentTypes.includes('drawing+xml')) {
        contentTypes = contentTypes.replace(
            '</Types>',
            '  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>\n</Types>'
        );
    }
    zip.file(ctPath, contentTypes);

    // 7. Generate final xlsx and download
    const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function base64ToUint8Array(base64) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
}
