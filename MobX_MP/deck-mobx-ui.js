/* ═══════════════════════════════════════════════════════════════════════════
   MobX presentation — integration with Media Plan Generator
   ───────────────────────────────────────────────────────────────────────────
   • Reads the media plan exactly as it is shown on the page (no recalculation)
     and prepares it for deck-mobx.js: Russian labels, Russian number format
     (1 400, 0,027, 3,28%), currency sign, headline numbers and period phrase.
   • Preview / print for the MobX brand reuse the modal of deck-ui.js.
   Nothing in app.js is modified; this file only reads the page.

   window.MobXDeckUI.collect(opts) → data for MobXDeck.render(doc, data)
     opts: { appName, appIcon (data URL), emails:[...], noVat, currencySign (default true),
             event: { title, one, few, many, acc, gen }  // Russian words for the target action }
   ═══════════════════════════════════════════════════════════════════════════ */
(function () {
'use strict';

const $ = id => document.getElementById(id);
const NBSP = ' ';

// ── Russian words ──────────────────────────────────────────────────────────
function plural(n, one, few, many) {
    const a = Math.abs(Math.round(n)) % 100, b = a % 10;
    if (a > 10 && a < 20) return many;
    if (b > 1 && b < 5) return few;
    if (b === 1) return one;
    return many;
}
const WORDS = {
    installs: { title: 'Установки', one: 'установка', few: 'установки', many: 'установок', acc: 'установку', gen: 'установки' },
    purchases: { title: 'Покупки', one: 'покупка', few: 'покупки', many: 'покупок', acc: 'покупку', gen: 'покупки' },
    registrations: { title: 'Регистрации', one: 'регистрация', few: 'регистрации', many: 'регистраций', acc: 'регистрацию', gen: 'регистрации' }
};

/** "14 days" → { label: '14 дней', phrase: 'за 14 дней' }; "1 month" → { '1 месяц', 'за месяц' } */
function periodRu(p) {
    const raw = String(p || '').trim();
    const m = raw.match(/^(\d+)\s*(days?|weeks?|months?|years?)$/i);
    if (!m) return { label: raw, phrase: raw ? 'за ' + raw : '' };
    const n = parseInt(m[1], 10), u = m[2].toLowerCase();
    const forms = u.startsWith('day') ? ['день', 'дня', 'дней']
        : u.startsWith('week') ? ['неделя', 'недели', 'недель']
        : u.startsWith('month') ? ['месяц', 'месяца', 'месяцев'] : ['год', 'года', 'лет'];
    const acc = u.startsWith('week') ? ['неделю', 'недели', 'недель'] : forms;
    return {
        label: n + ' ' + plural(n, forms[0], forms[1], forms[2]),
        phrase: n === 1 ? 'за ' + acc[0] : 'за ' + n + ' ' + plural(n, acc[0], acc[1], acc[2])
    };
}

// ── Numbers (display only) ─────────────────────────────────────────────────
function toNum(v) {
    if (v == null) return null;
    const s = String(v).replace(/[\s ,]/g, '').replace(/[^0-9.\-]/g, '');
    if (!s || s === '-' || s === '.') return null;
    const n = parseFloat(s);
    return isFinite(n) ? n : null;
}
function ru(n, maxDec, fixed) {
    let s = n.toFixed(maxDec);
    if (!fixed && s.indexOf('.') >= 0) s = s.replace(/\.?0+$/, '');
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
function fmt(raw, kind, sym) {
    const text = String(raw == null ? '' : raw).trim();
    if (kind === 'text') return text;
    if (kind === 'geo') return (typeof extractGeoCode === 'function' ? extractGeoCode(text) : '') || text;
    if (kind === 'period') return periodRu(text).label;
    const n = toNum(text);
    if (n == null) return text === '' ? '' : '—';
    if (kind === 'int') return ru(Math.round(n), 0);
    if (kind === 'pct') return ru(n, 2) + '%';
    if (kind === 'money2') return ru(n, 2, true) + (sym ? NBSP + sym : '');   // CPM: 2 decimals, as in the Sheet
    if (kind === 'money') {
        // ≥100: whole units; 1–100: two decimals unless whole (1 $, 14,29 $, 2,10 ₽); <1: up to three (0,027 $)
        const a = Math.abs(n);
        const v = a >= 100 ? ru(n, 0) : a >= 1 ? (Math.abs(n - Math.round(n)) < 0.005 ? ru(n, 0) : ru(n, 2, true)) : ru(n, 3);
        return v + (sym ? NBSP + sym : '');
    }
    return text;
}

// ── Columns ────────────────────────────────────────────────────────────────
const KIND = {
    channel: 'text', platform: 'text', geo: 'geo', period: 'period',
    installs: 'int', views: 'int', clicks: 'int', purchases: 'int', events: 'int', registrations: 'int',
    cpi: 'money', budget: 'money', cpm: 'money2', cpc: 'money', cpp: 'money', cpa: 'money', 'cost-per-reg': 'money',
    ctr: 'pct', 'cr-install': 'pct', 'cr-purchase': 'pct', 'cr-reg': 'pct'
};
const TOTAL_ID = { installs: 'total-installs', budget: 'total-cost', views: 'total-views', clicks: 'total-clicks', purchases: 'total-purchases', events: 'total-events', registrations: 'total-registrations' };
function labelRu(key, ev) {
    return ({
        channel: 'Канал', platform: 'Площадка', geo: 'Таргетинг', period: 'Период',
        installs: 'Установки', cpi: 'Цена установки', budget: 'Бюджет',
        views: 'Показы', cpm: 'CPM', ctr: 'CTR', clicks: 'Клики', cpc: 'CPC',
        'cr-install': 'CR в установку', 'cr-purchase': 'CR в ' + ev.acc,
        purchases: ev.title, events: ev.title, cpp: 'Цена ' + ev.gen, cpa: 'Цена ' + ev.gen,
        'cr-reg': 'CR в регистрацию', registrations: 'Регистрации', 'cost-per-reg': 'Цена регистрации'
    })[key] || key;
}

function cellValue(td) {
    if (!td) return '';
    const inp = td.querySelector('input,select');
    return inp ? inp.value : td.textContent;
}

function eventWords(opts) {
    if (opts.event && opts.event.many) return Object.assign({}, WORDS.purchases, opts.event);
    const th = document.querySelector('#mediaplan-head th.col-purchases, #mediaplan-head th.col-events');
    const name = th ? th.textContent.replace(/^total\s+/i, '').trim().toLowerCase() : 'purchases';
    if (name === 'purchases') return WORDS.purchases;
    // custom CPA event typed in English: used as is
    return { title: name.charAt(0).toUpperCase() + name.slice(1), one: name, few: name, many: name, acc: name, gen: name };
}

// ── Read the finished plan from the page ───────────────────────────────────
function collect(opts) {
    opts = opts || {};
    const sym = opts.currencySign === false ? '' : currencySymbol();
    const ev = eventWords(opts);
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
    const cell = (r, i) => {
        let v = fmt(r[i], KIND[cols[i].key] || 'text', sym);
        if (cols[i].key === 'channel') v = v.replace(/\s*CoDev$/i, '');   // same as the Excel export
        return v;
    };
    const total = key => {
        const el = $(TOTAL_ID[key] || '');
        return el ? fmt(el.textContent, KIND[key], sym) : '';
    };

    // buy table: identity + volume + unit price + budget (same split as the Gravils deck)
    const buyIdx = bIdx >= 2 ? cols.map((c, i) => i).filter(i => i <= 3 || (i >= bIdx - 2 && i <= bIdx)) : cols.map((c, i) => i);
    // funnel: everything else; in CPA mode the CPI column is left out (as in the reference deck)
    const isCpa = pos('events') >= 0 || pos('cpa') >= 0;
    let funnelIdx = cols.map((c, i) => i).filter(i => buyIdx.indexOf(i) < 0 && !(isCpa && cols[i].key === 'cpi'));

    const group = (idxs, kind) => ({
        kind, lead: true,
        cols: idxs.map(i => labelRu(cols[i].key, ev)),
        rows: useful.map(r => idxs.map(i => cell(r, i))),
        total: idxs.map((i, p) => p === 0 ? 'Итого' : total(cols[i].key))
    });
    const buy = group(buyIdx, 'f');

    let funnel = null;
    if (funnelIdx.length) {
        // with several rows the funnel table gets identifying columns in front
        const ids = [];
        if (useful.length > 1) {
            ids.push(pos('channel'));
            const chans = useful.map(r => r[pos('channel')]);
            if (new Set(chans).size < chans.length) ids.push(pos('platform'));
            if (new Set(useful.map(r => r[pos('geo')])).size > 1) ids.push(pos('geo'));
        }
        funnel = group(ids.concat(funnelIdx), 'o');
        funnel.lead = ids.length > 0;
        funnel.total = ids.map((x, p) => p === 0 ? 'Итого' : '').concat(funnelIdx.map(i => total(cols[i].key)));
    }

    // headline: main volume first (installs for CPI plans, the event for CPA plans)
    const per = periodRu(($('period') || {}).value || (useful[0] ? useful[0][pos('period')] : ''));
    const totNum = key => { const el = $(TOTAL_ID[key] || ''); return el ? toNum(el.textContent) : null; };
    const evKey = pos('events') >= 0 ? 'events' : 'purchases';
    const nInst = totNum('installs'), nEv = totNum(evKey);
    const W = (n, w) => plural(n, w.one, w.few, w.many);
    let headline;
    if (isCpa) headline = { n1: ru(nEv || 0, 0), w1: W(nEv || 0, ev), n2: nInst != null ? ru(nInst, 0) : null, w2: nInst != null ? W(nInst, WORDS.installs) : '', period: per.phrase };
    else headline = { n1: ru(nInst || 0, 0), w1: W(nInst || 0, WORDS.installs), n2: nEv ? ru(nEv, 0) : null, w2: nEv ? W(nEv, ev) : '', period: per.phrase };

    // budget summary
    const money = t => fmt(t, 'money', sym);
    const vis = el => el && el.style.display !== 'none';
    const summary = [];
    const vat = !opts.noVat && vis($('vat-row-line'));
    const comm = vis($('commission-row-line'));
    if (vat || comm) summary.push({ label: 'Бюджет без НДС', value: money($('vat-net').textContent) });
    if (comm) summary.push({ label: `Комиссия ${($('commission-pct') || {}).value || '0'}%`, value: money($('commission-amount').textContent) });
    if (vat) summary.push({ label: 'НДС 22%', value: money($('vat-amount').textContent) });
    summary.push({ label: 'Общий бюджет (gross)', value: money((vat || comm ? $('vat-gross') : $('vat-net')).textContent) });

    const vsel = $('vertical');
    const vertical = vsel && vsel.value && vsel.value !== 'other' ? vsel.options[vsel.selectedIndex].text : '';
    const client = (($('client') || {}).value || '').trim();
    const appName = (opts.appName || client || '').trim();
    return {
        client,
        appName,
        appIcon: opts.appIcon || null,
        coverTitle: opts.coverTitle || ['Медиаплан', 'интернет-размещения'],
        coverPills: [{ text: 'Период: ' + per.label }].concat(vertical ? [{ text: vertical }] : []),
        planPills: [{ text: appName }].concat(vertical ? [{ text: vertical }] : []).filter(p => p.text),
        headline,
        plan: { buy, funnel },
        summary,
        emails: (opts.emails && opts.emails.length) ? opts.emails : ['go@mobx.agency'],
        _rowCount: useful.length
    };
}

// ── Preview & print in the tool (MobX brand), reusing the modal of deck-ui.js ──
function fileName() {
    const c = ($('client').value || 'Client').trim().replace(/[^\w\-]+/g, '_').replace(/^_+|_+$/g, '');
    return 'MobX_' + (c || 'Client') + '_Media_Plan';
}
async function openPreview() {
    const emails = (($('deck-emails') || {}).value || '').split(/[\s,;]+/).filter(s => /@/.test(s));
    const appIcon = window.GravilsDeckUI && window.GravilsDeckUI.getLogo ? window.GravilsDeckUI.getLogo() : null;
    const data = collect({ emails, appIcon });
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
    const layoutName = { both: 'tables on one slide', split: 'buy and funnel on separate slides', buyonly: 'single table' }[report.layout] || report.layout;
    let msg = `${report.slides} slides · ${layoutName} · table text ${Math.round(report.fontPx)}px`;
    if (!report.readable) msg += ' · ⚠ the table is very large, text is below the comfortable size';
    if (report.overflow.length) msg += ' · ⚠ ' + report.overflow.join('; ');
    status.textContent = msg;
    $('deck-print').disabled = false;
}

window.MobXDeckUI = { collect, openPreview, periodRu, plural, fmt };
})();
