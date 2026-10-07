/* ═══════════════════════════════════════════════════════════════════════════
   MobX presentation — integration with Media Plan Generator
   ───────────────────────────────────────────────────────────────────────────
   • Reads the media plan exactly as it is shown on the page (no recalculation)
     and prepares it for deck-mobx.js in the format of the Metro reference deck:
     column names as in the Excel (Channel, Platform, …), numbers in Russian
     format with the currency sign (1 400, 41,27 ₽, 1,40%), the unit price as
     "220 ₽ / 1", rows grouped by channel.
   • Preview / print for the MobX brand reuse the modal of deck-ui.js.
   Nothing in app.js is modified; this file only reads the page.

   window.MobXDeckUI.collect(opts) → data for MobXDeck.render(doc, data)
     opts: { clientLogo | appIcon (data URL), noVat, currencySign (default true),
             coverTitle: ['line 1', 'line 2'], year }
   ═══════════════════════════════════════════════════════════════════════════ */
(function () {
'use strict';

const $ = id => document.getElementById(id);
const NBSP = ' ';

// ── Russian period for the cover ("Период: 3 месяца") ──────────────────────
function plural(n, one, few, many) {
    const a = Math.abs(Math.round(n)) % 100, b = a % 10;
    if (a > 10 && a < 20) return many;
    if (b > 1 && b < 5) return few;
    if (b === 1) return one;
    return many;
}
function periodRu(p) {
    const raw = String(p || '').trim();
    const m = raw.match(/^(\d+)\s*(days?|weeks?|months?|years?)$/i);
    if (!m) return raw;
    const n = parseInt(m[1], 10), u = m[2].toLowerCase();
    const forms = u.startsWith('day') ? ['день', 'дня', 'дней']
        : u.startsWith('week') ? ['неделя', 'недели', 'недель']
        : u.startsWith('month') ? ['месяц', 'месяца', 'месяцев'] : ['год', 'года', 'лет'];
    return n + ' ' + plural(n, forms[0], forms[1], forms[2]);
}

// ── Numbers (display only), as in the reference deck ───────────────────────
function toNum(v) {
    if (v == null) return null;
    const s = String(v).replace(/[\s ,]/g, '').replace(/[^0-9.\-]/g, '');
    if (!s || s === '-' || s === '.') return null;
    const n = parseFloat(s);
    return isFinite(n) ? n : null;
}
function ru(n, dec, trimZeros) {
    let s = n.toFixed(dec);
    if (trimZeros && /\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
    const neg = s[0] === '-';
    if (neg) s = s.slice(1);
    const parts = s.split('.');
    return (neg ? '-' : '') + parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, NBSP) + (parts[1] ? ',' + parts[1] : '');
}
function currencySymbol() {
    const code = ($('currency') || {}).value || 'USD';
    if (typeof CURRENCY_SYMBOLS !== 'undefined' && CURRENCY_SYMBOLS[code]) return CURRENCY_SYMBOLS[code];
    return { USD: '$', EUR: '€', RUB: '₽', GBP: '£', KZT: '₸', UAH: '₴', BRL: 'R$', INR: '₹' }[code] || '';
}
/**
 * kinds: text · geo · int (1 400) · pct (1,40%) · rate (CPM/CPC: 41,27 ₽) ·
 *        price (CPI/CPA/cost per …: 220 ₽, 14,29 $, 1 $) · money (budget: 1 000 000 ₽)
 */
function fmt(raw, kind, sym) {
    const text = String(raw == null ? '' : raw).trim();
    if (kind === 'text') return text;
    if (kind === 'geo') return (typeof extractGeoCode === 'function' ? extractGeoCode(text) : '') || text;
    const n = toNum(text);
    if (n == null) return text === '' ? '' : '—';
    const cur = v => v + (sym ? NBSP + sym : '');
    switch (kind) {
        case 'int': return ru(Math.round(n), 0);
        case 'pct': return ru(n, 2) + '%';
        case 'rate': return cur(ru(n, 2));
        case 'price': return cur(Math.abs(n) >= 100 ? ru(Math.round(n), 0) : ru(n, 2, true));
        case 'money': return cur(ru(Math.round(n), 0));
    }
    return text;
}

// ── Columns: names as in the Excel export ──────────────────────────────────
const COL = {
    channel: ['Channel', 'text'], platform: ['Platform', 'text'], geo: ['Targeting', 'geo'], period: ['Period', 'text'],
    installs: ['Total installs', 'int'], cpi: ['CPI', 'price'], budget: ['Total cost', 'money'],
    events: ['Total purchases', 'int'], cpa: ['Cost per purchase', 'price'],
    views: ['Views', 'int'], cpm: ['CPM', 'rate'], ctr: ['CTR', 'pct'], clicks: ['Total clicks', 'int'], cpc: ['CPC', 'rate'],
    'cr-install': ['CR install per click', 'pct'],
    'cr-reg': ['CR install to registration', 'pct'], registrations: ['Total registrations', 'int'], 'cost-per-reg': ['Cost per registration', 'price'],
    'cr-purchase': ['CR install to purchase', 'pct'], purchases: ['Total purchases', 'int'], cpp: ['Cost per purchase', 'price']
};
const TOTAL_ID = { installs: 'total-installs', budget: 'total-cost', views: 'total-views', clicks: 'total-clicks', purchases: 'total-purchases', events: 'total-events', registrations: 'total-registrations' };

function cellValue(td) {
    if (!td) return '';
    const inp = td.querySelector('input,select');
    return inp ? inp.value : td.textContent;
}

// ── Read the finished plan from the page ───────────────────────────────────
function collect(opts) {
    opts = opts || {};
    const sym = opts.currencySign === false ? '' : currencySymbol();
    const cols = Array.from(document.querySelectorAll('#mediaplan-head th')).map((th, idx) => {
        const cls = Array.from(th.classList).find(c => c.indexOf('col-') === 0);
        return { idx, key: cls ? cls.slice(4) : '' };
    }).filter(c => c.key && c.key !== 'actions');
    const pos = key => cols.findIndex(c => c.key === key);

    const rows = Array.from(document.querySelectorAll('#mediaplan-body tr')).map(tr => {
        const tds = Array.from(tr.children);
        return cols.map(c => cellValue(tds[c.idx]));
    });
    const bIdx = pos('budget');
    const useful = rows.filter(r => bIdx < 0 || toNum(r[bIdx]));
    const isCpa = pos('events') >= 0 || pos('cpa') >= 0;
    const label = i => (COL[cols[i].key] || [cols[i].key])[0];
    const kind = i => (COL[cols[i].key] || [0, 'text'])[1];
    const channelOf = r => String(r[pos('channel')] || '').trim().replace(/\s*CoDev$/i, '');   // as in the Excel
    const cell = (r, i, unit) => {
        let v = cols[i].key === 'channel' ? channelOf(r) : fmt(r[i], kind(i), sym);
        return unit && v && v !== '—' ? { v, unit: true } : v;
    };
    const total = key => {
        const el = $(TOTAL_ID[key] || '');
        return el ? fmt(el.textContent, (COL[key] || [0, 'int'])[1], sym) : '';
    };
    // channel groups: a dark line after the last row of each channel (rows of one channel are adjacent)
    const chans = useful.map(channelOf);
    const gb = chans.map((c, i) => i < chans.length - 1 && chans[i + 1] !== c);

    // (01) placement & budget: identity + volume + unit price ("/ 1") + budget
    const buyIdx = bIdx >= 2 ? cols.map((c, i) => i).filter(i => i <= 3 || (i >= bIdx - 2 && i <= bIdx)) : cols.map((c, i) => i);
    const unitIdx = bIdx >= 1 ? bIdx - 1 : -1;
    // (02) forecast: everything else; in a CPA plan the CPI column is left out (as in the reference)
    const funnelIdx = cols.map((c, i) => i).filter(i => buyIdx.indexOf(i) < 0 && !(isCpa && cols[i].key === 'cpi'));

    const group = (idxs, k, lead, unit) => ({
        kind: k, lead,
        cols: idxs.map(label),
        rows: useful.map(r => idxs.map(i => cell(r, i, unit && i === unitIdx))),
        gb: gb.slice(),
        total: idxs.map((i, p) => (p === 0 && lead) ? 'Total' : (TOTAL_ID[cols[i].key] ? total(cols[i].key) : ''))
    });
    const buy = group(buyIdx, 'b', true, true);
    let funnel = null, funnelLabelled = null;
    if (funnelIdx.length) {
        funnel = group(funnelIdx, 'k', false, false);
        // the forecast on a slide of its own: channel (and platform / geo when they vary) in front
        const ids = [pos('channel')];
        if (new Set(useful.map(r => r[pos('platform')])).size > 1) ids.push(pos('platform'));
        if (new Set(useful.map(r => r[pos('geo')])).size > 1) ids.push(pos('geo'));
        funnelLabelled = group(ids.concat(funnelIdx), 'k', true, false);
    }

    // budget cards
    const money = t => fmt(t, 'money', sym);
    const vis = el => el && el.style.display !== 'none';
    const vat = !opts.noVat && vis($('vat-row-line'));
    const comm = vis($('commission-row-line'));
    const summary = {
        net: money($('vat-net').textContent),
        vat: vat ? money($('vat-amount').textContent) : null,
        vatPct: '22%',
        commission: comm ? money($('commission-amount').textContent) : null,
        commissionPct: comm ? (($('commission-pct') || {}).value || '0') + '%' : null,
        gross: money((vat || comm ? $('vat-gross') : $('vat-net')).textContent)
    };

    const vsel = $('vertical');
    const vertical = vsel && vsel.value && vsel.value !== 'other' ? vsel.options[vsel.selectedIndex].text : '';
    const client = (($('client') || {}).value || '').trim();
    const period = (($('period') || {}).value || (useful[0] ? useful[0][pos('period')] : '') || '').trim();
    return {
        model: isCpa ? 'CPA' : 'CPI',
        client,
        clientLogo: opts.clientLogo || opts.appIcon || null,
        year: opts.year || new Date().getFullYear(),
        coverTitle: opts.coverTitle || null,
        coverPills: [{ text: 'Период: ' + periodRu(period) }].concat(vertical ? [{ text: vertical }] : []),
        sources: Array.from(new Set(chans)),
        info: [
            { label: 'Client', value: client || '—' },
            { label: 'Campaign / Agency', value: (($('campaign') || {}).value || '').trim() || 'MobX Agency' },
            { label: 'Document', value: 'Internet placement proposal' },
            { label: 'Period', value: period || '—' }
        ].concat(vertical ? [{ label: 'Vertical', value: vertical }] : []),
        plan: { buy, funnel, funnelLabelled },
        summary,
        _rowCount: useful.length
    };
}

// ── Preview & print in the tool (MobX brand), reusing the modal of deck-ui.js ──
function fileName() {
    const c = ($('client').value || 'Client').trim().replace(/[^\w\-]+/g, '_').replace(/^_+|_+$/g, '');
    return 'MobX_' + (c || 'Client') + '_Media_Plan';
}
async function openPreview() {
    const clientLogo = window.GravilsDeckUI && window.GravilsDeckUI.getLogo ? window.GravilsDeckUI.getLogo() : null;
    const data = collect({ clientLogo });
    if (!data._rowCount) { alert('Add at least one source with a budget to the media plan first.'); return; }
    const modal = $('deck-modal'), frame = $('deck-frame'), status = $('deck-status');
    modal.style.display = 'flex';
    status.textContent = 'Building slides…';
    $('deck-print').disabled = true;
    await new Promise(res => {
        frame.onload = res;
        frame.srcdoc = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${fileName()}</title></head><body></body></html>`;
    });
    let report;
    try {
        report = await window.MobXDeck.render(frame.contentDocument, data);
    } catch (e) {
        console.error(e);
        status.textContent = 'Could not build the presentation: ' + e.message;
        return;
    }
    frame.style.height = (report.slides * (window.MobXDeck.SLIDE_H + 40)) + 'px';
    const k = Math.min(1, ($('deck-scroll').clientWidth - 32) / window.MobXDeck.SLIDE_W);
    frame.style.transform = `scale(${k})`;
    $('deck-sizer').style.height = (parseFloat(frame.style.height) * k) + 'px';
    $('deck-sizer').style.width = (window.MobXDeck.SLIDE_W * k) + 'px';
    const layoutName = { side: 'tables side by side', split: 'tables on separate slides', single: 'single table' }[report.layout] || report.layout;
    let msg = `${report.slides} slides · ${layoutName} · table text ${Math.round(report.fontPx)}px`;
    if (!report.readable) msg += ' · ⚠ the table is very large, text is below the comfortable size';
    if (report.overflow.length) msg += ' · ⚠ ' + report.overflow.join('; ');
    status.textContent = msg;
    $('deck-print').disabled = false;
}

window.MobXDeckUI = { collect, openPreview, periodRu, fmt };
})();
