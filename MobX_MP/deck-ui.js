/* ═══════════════════════════════════════════════════════════════════════════
   Gravils presentation — integration with Media Plan Generator
   ───────────────────────────────────────────────────────────────────────────
   • Reads the media plan exactly as it is shown on the page (no recalculation).
   • Handles the "Presentation" form (cover title, subtitle, logo, e-mails,
     footnotes).
   • Shows a preview and prints it to PDF (Chrome: "Save as PDF").
   Nothing in app.js is modified; this file only reads the page.
   ═══════════════════════════════════════════════════════════════════════════ */
(function () {
'use strict';

const $ = id => document.getElementById(id);

const TITLE_PRESETS = {
    mua: ['Mobile user', 'acquisition proposal'],
    ipp: ['Internet placement', 'proposal']
};
const SUBTITLE_PRESETS = {
    perf: 'Performance marketing for {client}',
    perfgeo: 'Performance marketing for {client} — {geo}, {vertical}',
    media: '{period} of performance media for {client}',
    placement: 'Internet placement for {client}'
};
const NOTE_GOOGLE = 'Source can be launched only once the app has passed whitelisting in Google Ads.';
const NOTE_TIKTOK = 'Source can be launched only with a local {adj}license.';
const COUNTRY_ADJ = { Brazil: 'Brazilian', India: 'Indian', Mexico: 'Mexican', Indonesia: 'Indonesian', Philippines: 'Philippine', Vietnam: 'Vietnamese', Thailand: 'Thai', Turkey: 'Turkish', Nigeria: 'Nigerian', Kenya: 'Kenyan', Pakistan: 'Pakistani', Colombia: 'Colombian', Peru: 'Peruvian', Chile: 'Chilean', Argentina: 'Argentine', Kazakhstan: 'Kazakh', Russia: 'Russian', Egypt: 'Egyptian', 'South Africa': 'South African' };

// ── Column dictionary: page column key → presentation label + formatting ──
const COLS = {
    channel: { label: 'Channel', kind: 'text' },
    platform: { label: 'Platform', kind: 'text' },
    geo: { label: 'Targeting', kind: 'geo' },
    period: { label: 'Period', kind: 'text' },
    views: { label: 'Views', kind: 'int' },
    cpm: { label: 'CPM', kind: 'rate' },
    ctr: { label: 'CTR', kind: 'pct' },
    clicks: { label: 'Total clicks', kind: 'int' },
    cpc: { label: 'CPC', kind: 'rate' },
    'cr-install': { label: 'CR install per click', kind: 'pct' },
    installs: { label: 'Total installs', kind: 'int' },
    cpi: { label: 'CPI', kind: 'money' },
    'cr-purchase': { label: 'CR install to {event}', kind: 'pct' },
    purchases: { label: 'Total {events}', kind: 'int' },
    events: { label: 'Total {events}', kind: 'int' },
    cpp: { label: 'Cost per {event}', kind: 'money2' },
    cpa: { label: 'Cost per {event}', kind: 'money' },
    budget: { label: 'Total cost', kind: 'money' },
    'cr-reg': { label: 'CR install to registration', kind: 'pct' },
    registrations: { label: 'Total registrations', kind: 'int' },
    'cost-per-reg': { label: 'Cost per registration', kind: 'money2' }
};
const COUNT_KINDS = { int: 1, money: 1, money2: 1 };

// ── Number formatting (display only) ───────────────────────────────────────
const NBSP = '\u00A0';
function toNum(v) {
    if (v == null) return null;
    const s = String(v).replace(/[\s\u00A0,]/g, '').replace(/[^0-9.\-]/g, '');
    if (!s || s === '-' || s === '.') return null;
    const n = parseFloat(s);
    return isFinite(n) ? n : null;
}
function group3(i) { return i.replace(/\B(?=(\d{3})+(?!\d))/g, NBSP); }
function fmt(n, dec, fixed) {
    let s = Math.abs(n) >= 1000 && !fixed ? String(Math.round(n)) : n.toFixed(dec);
    if (!fixed && s.indexOf('.') >= 0) s = s.replace(/\.?0+$/, '');
    const neg = s[0] === '-';
    if (neg) s = s.slice(1);
    const parts = s.split('.');
    return (neg ? '-' : '') + group3(parts[0]) + (parts[1] ? '.' + parts[1] : '');
}
/** Currency sign: the dollar goes before the number ($8, $10 000), others after (10 000 ₽). */
function withSym(v, sym) { return !sym ? v : sym === '$' ? '$' + v : v + NBSP + sym; }
function formatCell(raw, kind, sym, unit) {
    const text = String(raw == null ? '' : raw).trim();
    if (kind === 'text') return text;
    if (kind === 'geo') return (typeof extractGeoCode === 'function' ? extractGeoCode(text) : '') || text;
    const n = toNum(text);
    if (n == null) return text === '' ? '' : '—';
    switch (kind) {
        case 'int': return fmt(Math.round(n), 0);
        case 'rate': return fmt(n, 3, true);
        case 'pct': return fmt(n, 2, true) + '%';
        case 'money2': return withSym(fmt(n, 2, n < 1000), sym);
        case 'money': return withSym(fmt(n, n < 1 ? 4 : 2, false), sym);   // no "/ 1" after the unit price
    }
    return text;
}

function currencySymbol() {
    const code = ($('currency') || {}).value || 'USD';
    if (typeof CURRENCY_SYMBOLS !== 'undefined' && CURRENCY_SYMBOLS[code]) return CURRENCY_SYMBOLS[code];
    return { USD: '$', EUR: '€', RUB: '₽', GBP: '£', KZT: '₸', UAH: '₴', BRL: 'R$', INR: '₹' }[code] || '';
}

function cellValue(td) {
    if (!td) return '';
    const inp = td.querySelector('input,select');
    return inp ? inp.value : td.textContent;
}

// ── Read the finished plan from the page ───────────────────────────────────
function readPlan(opts) {
    const sym = currencySymbol();
    const heads = Array.from(document.querySelectorAll('#mediaplan-head th'));
    const cols = heads.map((th, idx) => {
        const cls = Array.from(th.classList).find(c => c.indexOf('col-') === 0);
        return { idx, key: cls ? cls.slice(4) : '', raw: th.textContent.trim() };
    }).filter(c => c.key && c.key !== 'actions');

    // event name ("purchases") from the page header, e.g. "Total Purchases"
    const evCol = cols.find(c => c.key === 'purchases' || c.key === 'events');
    const events = evCol ? evCol.raw.replace(/^total\s+/i, '').toLowerCase() : 'purchases';
    const event = events.replace(/s$/, '');

    // rows (skip completely empty ones)
    const rows = Array.from(document.querySelectorAll('#mediaplan-body tr')).map(tr => {
        const tds = Array.from(tr.children);
        return cols.map(c => cellValue(tds[c.idx]));
    });
    const budgetPos = cols.findIndex(c => c.key === 'budget');
    const useful = rows.filter(r => budgetPos < 0 || toNum(r[budgetPos]) || r.some((v, i) => i > 3 && toNum(v)));

    // totals row (the first cell spans the four identity columns)
    const totals = [];
    Array.from(document.querySelectorAll('#mediaplan-foot td')).forEach(td => {
        const span = parseInt(td.getAttribute('colspan') || '1', 10);
        totals.push(td.textContent.trim());
        for (let i = 1; i < span; i++) totals.push('');
    });

    // which columns go to the main ("buy") table: identity + the unit price + volume + cost
    const bIdx = budgetPos;
    let mainIdx;
    if (bIdx >= 2) {
        const set = new Set([0, 1, 2, 3, bIdx - 2, bIdx - 1, bIdx]);
        mainIdx = cols.map((c, i) => i).filter(i => set.has(i));
    } else {
        mainIdx = cols.map((c, i) => i);
    }
    const funnelIdx = cols.map((c, i) => i).filter(i => mainIdx.indexOf(i) < 0);
    const unitIdx = bIdx >= 1 ? bIdx - 1 : -1;

    const google = opts.noteGoogle, tiktok = opts.noteTiktok;
    const mark = name => {
        let n = String(name).trim();
        if (google && /google/i.test(n) && !/\*$/.test(n)) n += '*';
        if (tiktok && /tiktok/i.test(n) && !/\*$/.test(n)) n += google ? '**' : '*';
        return n;
    };

    const label = c => (COLS[c.key] ? COLS[c.key].label : c.raw)
        .replace('{events}', events).replace('{event}', event);
    const kind = c => (COLS[c.key] ? COLS[c.key].kind : 'text');

    const makeGroup = (idxs, lead) => ({
        lead,
        cols: idxs.map(i => ({ label: label(cols[i]), key: cols[i].key })),
        rows: useful.map(r => idxs.map(i => {
            const v = formatCell(r[i], i === unitIdx ? 'money' : kind(cols[i]), sym, i === unitIdx);
            return cols[i].key === 'channel' ? mark(v) : v;
        })),
        total: idxs.map((i, pos) => {
            if (pos === 0 && lead) return 'Total';
            const k = kind(cols[i]);
            const t = totals[cols[i].idx];
            if (!COUNT_KINDS[k] || i === unitIdx || k === 'money2' || (cols[i].key === 'cpi')) return '';
            return formatCell(t, k, sym, false);
        })
    });

    const main = makeGroup(mainIdx, true);
    const funnel = funnelIdx.length ? makeGroup(funnelIdx, false) : null;

    // identifying columns repeated in front of the funnel when it stands alone
    const geos = new Set(main.rows.map(r => r[2]));
    const idCols = geos.size > 1 ? [0, 1, 2] : [0, 1];

    // which footnotes actually apply
    const channels = useful.map(r => String(r[0]));
    return {
        plan: { main, funnel, idCols },
        hasGoogle: channels.some(c => /google/i.test(c)),
        hasTiktok: channels.some(c => /tiktok/i.test(c)),
        rowCount: useful.length
    };
}

function readSummary() {
    const sym = currencySymbol();
    const money = t => { const n = toNum(t); return n == null ? t : withSym(fmt(n, 2, false), sym); };
    const out = [];
    const vis = el => el && el.style.display !== 'none';
    const netLabel = ($('vat-net-label') || {}).textContent || 'Max Placement Cost Net';
    out.push({ label: netLabel.trim(), value: money($('vat-net').textContent) });
    if (vis($('vat-row-line'))) out.push({ label: 'VAT (22%)', value: money($('vat-amount').textContent) });
    if (vis($('commission-row-line'))) {
        const pct = ($('commission-pct') || {}).value || '0';
        out.push({ label: `Commission (${pct}%)`, value: money($('commission-amount').textContent) });
    }
    out.push({ label: 'Total cost Gross', value: money($('vat-gross').textContent) });
    return out;
}

function geoNames() {
    const codes = (typeof selectedGeos !== 'undefined' && selectedGeos.length) ? selectedGeos
        : Array.from(new Set(Array.from(document.querySelectorAll('#mediaplan-body [data-field="geo"]'))
            .map(i => (typeof extractGeoCode === 'function' ? extractGeoCode(i.value) : i.value)).filter(Boolean)));
    return codes.map(c => (typeof getCountryName === 'function' ? getCountryName(c) : c));
}

function verticalText() {
    const v = $('vertical');
    if (!v || v.value === 'other' || !v.value) return '';
    return v.options[v.selectedIndex].text;
}

// ── Presentation form ──────────────────────────────────────────────────────
let clientLogo = null;
let subtitleDirty = false;

function resolveSubtitle(tpl) {
    const client = ($('client').value || '').trim() || 'Client';
    const period = ($('period').value || '').trim();
    const vert = verticalText();
    let s = tpl.replace('{client}', client)
        .replace('{geo}', geoNames().join(', '))
        .replace('{vertical}', vert ? vert.toLowerCase() : '')
        .replace('{period}', period ? period.charAt(0).toUpperCase() + period.slice(1) : '');
    return s.replace(/\s+—\s*,\s*$/, '').replace(/,\s*$/, '').replace(/—\s*,/, '—').replace(/\s{2,}/g, ' ').trim();
}

function refreshSubtitle() {
    const sel = $('deck-subtitle-preset');
    if (subtitleDirty || !SUBTITLE_PRESETS[sel.value]) return;
    $('deck-subtitle').value = resolveSubtitle(SUBTITLE_PRESETS[sel.value]);
}

function coverTitle() {
    const v = $('deck-title-preset').value;
    if (TITLE_PRESETS[v]) return TITLE_PRESETS[v].slice();
    return ($('deck-title-custom').value || '').split(/\n+/).map(s => s.trim()).filter(Boolean);
}

function emails() {
    return ($('deck-emails').value || '').split(/[\s,;]+/).map(s => s.trim()).filter(s => /@/.test(s));
}

function collectDeckData() {
    const noteGoogle = $('deck-note-google').checked;
    const noteTiktok = $('deck-note-tiktok').checked;
    const p = readPlan({ noteGoogle, noteTiktok });
    const g = noteGoogle && p.hasGoogle, t = noteTiktok && p.hasTiktok;
    const footnotes = [];
    if (g) footnotes.push('* ' + NOTE_GOOGLE);
    if (t) {
        const names = geoNames();
        const adj = names.length === 1 && COUNTRY_ADJ[names[0]] ? COUNTRY_ADJ[names[0]] + ' ' : '';
        footnotes.push((g ? '** ' : '* ') + NOTE_TIKTOK.replace('{adj}', adj));
    }
    const extra = ($('deck-note-custom').value || '').trim();
    if (extra) footnotes.push(extra);
    // re-read the plan if a mark was requested for a channel that is absent
    const plan = (noteGoogle && !p.hasGoogle) || (noteTiktok && !p.hasTiktok)
        ? readPlan({ noteGoogle: g, noteTiktok: t }).plan : p.plan;

    refreshSubtitle();
    const client = ($('client').value || '').trim();
    const info = [
        { label: 'Client', value: client || '—' },
        { label: 'Campaign/Agency', value: ($('campaign').value || '').trim() || 'Gravils Agency' },
        { label: 'Document', value: 'Internet placement proposal' },
        { label: 'Period', value: ($('period').value || '').trim() || '—' }
    ];
    const vert = verticalText();
    if (vert) info.push({ label: 'Vertical', value: vert });

    return {
        client,
        clientLogo,
        coverTitle: coverTitle(),
        subtitle: ($('deck-subtitle').value || '').trim(),
        emails: emails(),
        footnotes,
        info,
        summary: readSummary(),
        plan,
        _rowCount: p.rowCount
    };
}

// ── Preview & print ────────────────────────────────────────────────────────
function fileName() {
    const c = ($('client').value || 'Client').trim().replace(/[^\w\-]+/g, '_').replace(/^_+|_+$/g, '');
    return ($('brand').value === 'mobx' ? 'MobX_' : 'Gravils_') + (c || 'Client') + '_Media_Plan';
}

async function openPreview() {
    const brand = $('brand').value;
    if (brand === 'mobx' && window.MobXDeckUI) return window.MobXDeckUI.openPreview();
    if (brand !== 'gravils') { alert('Presentation is available for the Gravils and MobX brands.'); return; }
    const data = collectDeckData();
    if (!data._rowCount) { alert('Add at least one source with a budget to the media plan first.'); return; }

    const modal = $('deck-modal');
    const frame = $('deck-frame');
    const status = $('deck-status');
    modal.style.display = 'flex';
    status.textContent = 'Building slides…';
    $('deck-print').disabled = true;

    await new Promise(res => {
        frame.onload = res;
        frame.srcdoc = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${fileName()}</title></head><body></body></html>`;
    });
    const doc = frame.contentDocument;
    let report;
    try {
        report = await window.GravilsDeck.render(doc, data);
    } catch (e) {
        console.error(e);
        status.textContent = 'Could not build the presentation: ' + e.message;
        return;
    }
    const slides = doc.querySelectorAll('section.slide').length;
    frame.style.height = (slides * (window.GravilsDeck.SLIDE_H + 40)) + 'px';
    scalePreview();

    const layoutName = { side: 'tables side by side', stack: 'tables one under another', split: 'buy and funnel on separate slides', single: 'single table' }[report.layout] || report.layout;
    let msg = `${report.slides} slides · ${layoutName}` +
        (report.planSlides > (report.layout === 'split' ? 2 : 1) ? ` · up to ${report.rowsPerSlide} rows per slide` : '') +
        ` · table text ${Math.round(report.fontPx)}px`;
    if (!report.readable) msg += ' · ⚠ the table is very large, text is below the comfortable size';
    if (report.overflow.length) msg += ' · ⚠ ' + report.overflow.join('; ');
    status.textContent = msg;
    $('deck-print').disabled = false;
}

function scalePreview() {
    const frame = $('deck-frame');
    const wrap = $('deck-scroll');
    const k = Math.min(1, (wrap.clientWidth - 32) / window.GravilsDeck.SLIDE_W);
    frame.style.transform = `scale(${k})`;
    $('deck-sizer').style.height = (parseFloat(frame.style.height || 0) * k) + 'px';
    $('deck-sizer').style.width = (window.GravilsDeck.SLIDE_W * k) + 'px';
}

function printDeck() {
    const frame = $('deck-frame');
    const old = document.title;
    document.title = fileName();
    const restore = () => { document.title = old; window.removeEventListener('afterprint', restore); };
    window.addEventListener('afterprint', restore);
    frame.contentWindow.focus();
    frame.contentWindow.print();
    setTimeout(restore, 3000);
}

function closePreview() {
    $('deck-modal').style.display = 'none';
    $('deck-frame').srcdoc = '<!DOCTYPE html><html><body></body></html>';
}

// ── Logo upload ────────────────────────────────────────────────────────────
function onLogo(e) {
    const f = e.target.files && e.target.files[0];
    if (!f) return;
    if (f.size > 5 * 1024 * 1024) { alert('Logo file is too large (max 5 MB).'); e.target.value = ''; return; }
    const r = new FileReader();
    r.onload = () => {
        clientLogo = r.result;
        const img = $('deck-logo-preview');
        img.src = clientLogo;
        img.style.display = 'block';
        $('deck-logo-empty').style.display = 'none';
        $('deck-logo-clear').style.display = '';
    };
    r.readAsDataURL(f);
}
function clearLogo() {
    clientLogo = null;
    $('deck-logo-file').value = '';
    $('deck-logo-preview').style.display = 'none';
    $('deck-logo-empty').style.display = '';
    $('deck-logo-clear').style.display = 'none';
}

// ── Wiring ─────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    if (!$('deck-section')) return;
    $('deck-title-preset').addEventListener('change', () => {
        $('deck-title-custom').style.display = $('deck-title-preset').value === 'custom' ? '' : 'none';
    });
    $('deck-subtitle-preset').addEventListener('change', () => {
        subtitleDirty = false;
        if ($('deck-subtitle-preset').value === 'custom') { $('deck-subtitle').focus(); return; }
        refreshSubtitle();
    });
    $('deck-subtitle').addEventListener('input', () => {
        subtitleDirty = true;
        $('deck-subtitle-preset').value = 'custom';
    });
    ['client', 'period', 'vertical'].forEach(id => $(id) && $(id).addEventListener('input', refreshSubtitle));
    $('vertical') && $('vertical').addEventListener('change', refreshSubtitle);
    const geoTags = $('geo-tags');
    if (geoTags && window.MutationObserver) new MutationObserver(refreshSubtitle).observe(geoTags, { childList: true });
    refreshSubtitle();

    $('deck-logo-file').addEventListener('change', onLogo);
    $('deck-logo-clear').addEventListener('click', clearLogo);
    $('deck-btn').addEventListener('click', openPreview);
    $('deck-print').addEventListener('click', printDeck);
    $('deck-close').addEventListener('click', closePreview);
    $('deck-modal').addEventListener('click', e => { if (e.target.id === 'deck-modal') closePreview(); });
    document.addEventListener('keydown', e => { if (e.key === 'Escape' && $('deck-modal').style.display === 'flex') closePreview(); });
    window.addEventListener('resize', () => { if ($('deck-modal').style.display === 'flex') scalePreview(); });
});

// exposed for tests
window.GravilsDeckUI = { collectDeckData, readPlan, formatCell, openPreview, getLogo: () => clientLogo };
})();
