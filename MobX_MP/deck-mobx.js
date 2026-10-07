/* ═══════════════════════════════════════════════════════════════════════════
   MobX presentation generator (template: Metro media plan, 07.10.2026)
   ───────────────────────────────────────────────────────────────────────────
   Takes a finished media plan (values exactly as shown in the plan, no
   recalculation) and lays it out on 1920×1080 slides in the MobX style:
     cover (blue) → media-plan slide(s) → closing slide (black)
   The plan slide repeats the Excel: info row (Client / Campaign / Document /
   Period / Vertical), "(01) Placement & budget" and "(02) Forecast metrics"
   side by side (data.lang 'en' by default, 'ru' = wording of the reference), and the budget cards "net + VAT = gross".
   Tables are auto-fitted: every table size derives from one scale (--k, 1 =
   reference size); the planner tries both tables on one slide, the tables on
   separate slides, and row pagination, and keeps the option with the fewest
   slides whose text stays readable, then the largest text.
   All colours are opaque (no transparency in the PDF).

   Public API (window.MobXDeck):
     render(doc, data) → Promise<report>
     deckCss()         → string
     SLIDE_W, SLIDE_H, K_MIN
   data: see deck-mobx-ui.js (collect()).
   ═══════════════════════════════════════════════════════════════════════════ */
(function () {
'use strict';

const SLIDE_W = 1920, SLIDE_H = 1080;
const BLUE = '#0000E1', BG = '#F8F8F8', INK = '#000000';
const GRAY40 = '#959595';       // black 40% over #F8F8F8 (tagline, labels, + / =)
const GRAY56 = '#707070';       // black 56% over white (budget card labels)
const LINE = '#E4E4E4';         // light separators
const ON_BLUE = '#C6C6F3';      // #F8F8F8 80% over blue (gross card label)
const ON_BLACK = '#AEAEAE';     // #F8F8F8 70% over black (closing footer)
const K_MIN = 0.74;             // smallest comfortable table scale (≈14px text)
const K_FLOOR = 0.45;
const MAX_SLIDES = 10;
const CARD_TOP = 410;           // top of the table cards
const SUM_H = 142, SUM_GAP = 22, SLIDE_BOTTOM = 1030;

const A = () => window.MOBX_DECK_ASSETS || {};

// Slide texts; data.lang = 'en' (default) | 'ru' (wording of the Metro reference deck)
const TEXT = {
    en: {
        cover: ['Internet placement', 'media plan'],
        agency: 'Mobile marketing agency',
        plan: m => `Media plan: ${m} model`,
        sec1: 'Placement & budget', sec2: 'Forecast metrics',
        closing: b => `<span style="color:${b}">Together</span> we build<br>projects that<br>change the market`
    },
    ru: {
        cover: ['Медиаплан', 'интернет\u2011размещения'],
        agency: 'Агентство мобильного маркетинга',
        plan: m => `Медиаплан: модель ${m}`,
        sec1: 'Размещение и бюджет', sec2: 'Прогнозные показатели',
        closing: b => `Создавать<br><span style="color:${b}">вместе</span> проекты,<br>меняющие рынок`
    }
};
const T = data => TEXT[data.lang] || TEXT.en;

// Golos Text: ascent 0.98, descent 0.22 em → in a line box of height L the
// baseline sits L/2 + 0.38·fs below its top.
const topFor = (baseline, fs, lh) => baseline - (lh || fs) / 2 - 0.38 * fs;

// ── Styles ─────────────────────────────────────────────────────────────────
function deckCss() {
    const a = A();
    return `${a.fontFaces || ''}
@page{size:${SLIDE_W}px ${SLIDE_H}px;margin:0}
*{box-sizing:border-box;margin:0;padding:0}
html,body{background:#d9d9d9}
body{font-family:"Golos Text",Arial,sans-serif;color:${INK};-webkit-print-color-adjust:exact;print-color-adjust:exact;-webkit-font-smoothing:antialiased}
.slide{position:relative;width:${SLIDE_W}px;height:${SLIDE_H}px;overflow:hidden;background:${BG};margin:0 auto 40px}
.slide.measure{position:absolute;left:-99999px;top:0;margin:0}
@media print{html,body{background:${BG}}.slide{margin:0;break-after:page;page-break-after:always}.slide:last-child{break-after:auto;page-break-after:auto}}
.abs{position:absolute}
.nw{white-space:nowrap}
.m{font-weight:500}
.svgbox svg{display:block;width:100%;height:100%}
.pill{display:inline-flex;align-items:center;white-space:nowrap;border-radius:999px;line-height:1;border:2px solid currentColor}

/* plan slide */
.info{display:flex;gap:64px}
.info .l{font-size:16px;line-height:16px;color:${GRAY40};white-space:nowrap;margin-bottom:13.6px}
.info .v{font-size:21px;line-height:21px;white-space:nowrap}
.sec{display:flex;align-items:baseline;gap:14px;white-space:nowrap;line-height:20px}
.sec .n{font-size:18px;font-weight:500;color:${BLUE}}
.sec .t{font-size:20px;font-weight:500;letter-spacing:-0.01em}
.card{position:absolute;background:#fff;border-radius:24px;padding:18px 20px}
.shade{position:absolute}

table.mt{--k:1;width:100%;border-collapse:separate;border-spacing:0;font-variant-numeric:tabular-nums}
.mt th{height:calc(54px*var(--k));background:${BLUE};color:${BG};font-weight:500;font-size:calc(var(--hf)*var(--k));line-height:1.06;text-align:center;padding:0 calc(var(--df)*0.75*var(--k));vertical-align:middle}
.mt.k th{background:${INK}}
.mt th:first-child{border-radius:calc(27px*var(--k)) 0 0 calc(27px*var(--k))}
.mt th:last-child{border-radius:0 calc(27px*var(--k)) calc(27px*var(--k)) 0}
.mt td{height:calc(52px*var(--k));font-size:calc(var(--df)*var(--k));line-height:1;text-align:center;white-space:nowrap;padding:0 calc(var(--df)*0.75*var(--k));border-bottom:2px solid ${LINE};vertical-align:middle}
.mt th:first-child,.mt td:first-child{padding-left:calc(24px*var(--k))}
.mt tr.gb td{border-bottom-color:${INK}}
.mt tr.last td{border-bottom-color:transparent}
.mt.lead th:first-child,.mt.lead td:first-child{text-align:left}
.mt.lead tbody tr:not(.tot) td:first-child{font-weight:500}
.mt tr.tot td{height:calc(54px*var(--k));font-weight:500;color:${BLUE};border-top:2px solid ${BLUE};border-bottom:2px solid ${BLUE}}
.mt.k tr.tot td{color:${INK};border-color:${INK}}
.mt tr.tot td:first-child{border-left:2px solid ${BLUE};border-radius:calc(27px*var(--k)) 0 0 calc(27px*var(--k))}
.mt tr.tot td:last-child{border-right:2px solid ${BLUE};border-radius:0 calc(27px*var(--k)) calc(27px*var(--k)) 0}
.mt.k tr.tot td:first-child,.mt.k tr.tot td:last-child{border-color:${INK}}
.mt.lead tr.tot td:first-child{padding-left:calc(22px*var(--k))}

/* budget cards */
.sum{display:flex;align-items:flex-start;gap:20px}
.sum .c{position:relative;height:${SUM_H}px;border-radius:24px;background:#fff;padding:0 36px;white-space:nowrap}
.sum .c.g{background:${BLUE};color:${BG};padding:0 43px}
.sum .vv{letter-spacing:-0.025em}
.sum .gv{letter-spacing:-0.03em}
.sum .pct{position:absolute;top:30px;height:30px;padding:0 12px 4px;font-size:16px;font-weight:500;color:${BLUE}}
.sum .op{font-size:40px;line-height:40px;font-weight:500;color:${GRAY40};width:22px;text-align:center;margin-top:${topFor(974, 40) - 888}px}
`;
}

// ── Small helpers ──────────────────────────────────────────────────────────
function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, c =>
        ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function h(doc, tag, attrs, html) {
    const el = doc.createElement(tag);
    if (attrs) for (const k in attrs) {
        if (k === 'style') el.style.cssText = attrs[k];
        else if (k === 'class') el.className = attrs[k];
        else el.setAttribute(k, attrs[k]);
    }
    if (html != null) el.innerHTML = html;
    return el;
}
function searchMax(lo, hi, fits, step) {
    step = step || 0.01;
    if (!fits(lo)) return { value: lo, fits: false };
    if (fits(hi)) return { value: hi, fits: true };
    let a = lo, b = hi;
    while (b - a > step) {
        const m = (a + b) / 2;
        if (fits(m)) a = m; else b = m;
    }
    fits(a);
    return { value: a, fits: true };
}
function waitImages(root) {
    const imgs = Array.from(root.querySelectorAll('img'));
    return Promise.all(imgs.map(img => img.complete ? null :
        new Promise(r => { img.onload = img.onerror = r; })));
}
/** Text placed by its first baseline (reference coordinates). */
function textAt(doc, html, o) {
    const lh = o.lh || o.fs;
    const st = [`font-size:${o.fs}px`, `line-height:${lh}px`, `font-weight:${o.w || 400}`, `top:${topFor(o.baseline, o.fs, lh).toFixed(2)}px`];
    if (o.left != null) st.push(`left:${o.left}px`);
    if (o.right != null) st.push(`right:${o.right}px;text-align:right`);
    if (o.color) st.push(`color:${o.color}`);
    return h(doc, 'div', { class: 'abs nw', style: st.join(';') }, html);
}
function svgBox(doc, svg, style) {
    return h(doc, 'div', { class: 'abs svgbox', style }, svg || '');
}
/** Opaque layered shadow under a card (pre-blended over #F8F8F8), as in the reference. */
function shadow(doc, x, y, w, hgt, r) {
    const frag = doc.createDocumentFragment();
    [[20, '#F7F7F7'], [16, '#F6F6F6'], [12, '#F5F5F5'], [8, '#F2F2F2'], [4, '#EFEFEF']].forEach(([o, c]) =>
        frag.appendChild(h(doc, 'div', { class: 'shade', style: `left:${x - o}px;top:${y + 6 - o}px;width:${w + 2 * o}px;height:${hgt + 2 * o}px;background:${c};border-radius:${r + o}px` })));
    return frag;
}

// ── Cover (blue) and closing (black) share the art ─────────────────────────
function brandFrame(doc, s, bandColor) {
    const a = A();
    s.appendChild(svgBox(doc, a.artSvg, 'left:0;top:0;width:1920px;height:1080px'));
    s.appendChild(h(doc, 'div', { class: 'abs', style: `left:0;top:226px;width:757px;height:80px;background:${bandColor}` }));
    const b = a.logoBigBox || [40, 40, 422, 116];
    s.appendChild(svgBox(doc, a.logoBigSvg, `left:${b[0]}px;top:${b[1]}px;width:${b[2]}px;height:${b[3]}px`));
}
function footer(doc, data, color) {
    return textAt(doc, esc(T(data).agency + ' · ' + (data.year || new Date().getFullYear())), { fs: 22, baseline: 1022, left: 40, color });
}

function buildCover(doc, data) {
    const s = h(doc, 'section', { class: 'slide', 'data-kind': 'cover', style: `background:${BLUE};color:${BG}` });
    brandFrame(doc, s, '#FFFFFF');

    let nameBox = null;
    if (data.clientLogo) {
        const logoBox = h(doc, 'div', { class: 'abs', style: 'left:40px;top:490px;height:96px;max-width:620px;display:flex' });
        logoBox.appendChild(h(doc, 'img', { src: data.clientLogo, alt: '', style: 'height:96px;width:auto;max-width:620px;object-fit:contain;border-radius:8px;display:block' }));
        s.appendChild(logoBox);
    } else if (data.client) {
        // no logo: the client name in its place
        nameBox = textAt(doc, esc(data.client), { fs: 44, baseline: 560, left: 40, w: 500, color: BG });
        s.appendChild(nameBox);
    }
    const lines = (data.coverTitle && data.coverTitle.length ? data.coverTitle : T(data).cover);
    const title = textAt(doc, lines.map(esc).join('<br>'), { fs: 64, lh: 68, w: 500, baseline: 698, left: 40, color: BG });
    title.style.letterSpacing = '-0.03em';
    s.appendChild(title);

    const pills = h(doc, 'div', { class: 'abs', style: 'left:40px;top:818px;display:flex;flex-direction:column;align-items:flex-start;gap:16px' });
    (data.coverPills || []).forEach(p => pills.appendChild(h(doc, 'div', { class: 'pill', style: `height:58px;padding:0 30px;font-size:26px;color:${BG}` }, esc(p.text))));
    s.appendChild(pills);
    s.appendChild(footer(doc, data, BG));

    s._fit = () => {
        if (nameBox) searchMax(24, 44, v => {
            nameBox.style.fontSize = v + 'px'; nameBox.style.lineHeight = v + 'px';
            nameBox.style.top = topFor(560, v) + 'px';
            return nameBox.offsetWidth <= 700;
        }, 0.5);
        searchMax(40, 64, v => {
            title.style.fontSize = v + 'px'; title.style.lineHeight = (v * 68 / 64) + 'px';
            title.style.top = topFor(698, v, v * 68 / 64) + 'px';
            return title.offsetWidth <= 717;
        }, 0.5);
    };
    return s;
}

function buildClosing(doc, data) {
    const s = h(doc, 'section', { class: 'slide', 'data-kind': 'closing', style: `background:#000;color:${BG}` });
    brandFrame(doc, s, BLUE);
    const t = textAt(doc, T(data).closing(BLUE),
        { fs: 70, lh: 77, w: 500, baseline: 790, left: 40, color: BG });
    t.style.letterSpacing = '-0.02em';
    s.appendChild(t);
    s.appendChild(footer(doc, data, ON_BLACK));
    s._fit = () => {};
    return s;
}

// ── Media-plan slide ───────────────────────────────────────────────────────
function cellHtml(c) { return esc(c); }
/** group = { kind:'b'|'k', lead, cols:[label], rows:[[cell]], gb:[bool], total:[cell]|null } */
function buildTable(doc, g, withTotal, hf, df) {
    const t = h(doc, 'table', { class: 'mt' + (g.kind === 'k' ? ' k' : '') + (g.lead ? ' lead' : ''), style: `--hf:${hf}px;--df:${df}px` });
    const thead = h(doc, 'thead'), trh = h(doc, 'tr');
    g.cols.forEach(c => trh.appendChild(h(doc, 'th', null, esc(c))));
    thead.appendChild(trh);
    t.appendChild(thead);
    const tb = h(doc, 'tbody');
    g.rows.forEach((r, i) => {
        const last = i === g.rows.length - 1;
        const tr = h(doc, 'tr', (last || g.gb[i]) ? { class: last ? 'last' : 'gb' } : null);
        r.forEach(v => tr.appendChild(h(doc, 'td', null, cellHtml(v))));
        tb.appendChild(tr);
    });
    if (withTotal && g.total) {
        const tr = h(doc, 'tr', { class: 'tot' });
        g.total.forEach(v => tr.appendChild(h(doc, 'td', null, cellHtml(v))));
        tb.appendChild(tr);
    }
    t.appendChild(tb);
    return t;
}

function buildSummary(doc, data) {
    const sum = data.summary || {};
    const box = h(doc, 'div', { class: 'sum abs', style: 'left:40px' });
    const plain = (label, value, pill) => {
        const c = h(doc, 'div', { class: 'c' });
        const l = textAt(doc, esc(label), { fs: 18, baseline: 52, left: 36, color: GRAY56 });
        c.appendChild(l);
        if (pill) { const pc = h(doc, 'div', { class: 'pill pct' }, esc(pill)); c.appendChild(pc); c._pill = [l, pc]; }
        const v = textAt(doc, esc(value), { fs: 44, w: 500, baseline: pill ? 106 : 104, left: 36 });
        v.classList.add('vv');
        c.appendChild(v);
        return c;
    };
    const op = ch => h(doc, 'div', { class: 'op' }, ch);
    const cards = [];
    const add = c => { cards.push(c); box.appendChild(c); };
    // net (+ commission) (+ VAT) = gross; without VAT simply net = gross, as in the Excel
    add(plain('Max placement cost net', sum.net));
    if (sum.commission) { box.appendChild(op('+')); add(plain('Commission', sum.commission, sum.commissionPct)); }
    if (sum.vat) { box.appendChild(op('+')); add(plain('VAT', sum.vat, sum.vatPct || '22%')); }
    box.appendChild(op('='));
    const g = h(doc, 'div', { class: 'c g' });
    g.appendChild(textAt(doc, esc('Total cost gross · ' + (data.model || '')), { fs: 18, baseline: 46, left: 43, color: ON_BLUE }));
    const gv = textAt(doc, esc(sum.gross), { fs: 56, w: 500, baseline: 108, left: 43, color: BG });
    gv.classList.add('gv');
    g.appendChild(gv);
    add(g);
    box._cards = cards;
    return box;
}
/** Budget cards: width follows the content (text is absolutely placed inside). */
function settleSummary(box) {
    box._cards.forEach(c => {
        if (c._pill) c._pill[1].style.left = (36 + c._pill[0].offsetWidth + 12) + 'px';   // "VAT  (22%)"
        const pad = c.classList.contains('g') ? 43 : 36;
        const w = Math.max(...Array.from(c.children).map(ch => ch.offsetLeft - pad + ch.offsetWidth));
        c.style.width = Math.ceil(w + 2 * pad) + 'px';
    });
}

function buildPlanSlide(doc, data, spec) {
    const a = A();
    const s = h(doc, 'section', { class: 'slide', 'data-kind': 'plan' });
    const b = a.logoBox || [40, 50, 151.87, 41.83];
    s.appendChild(svgBox(doc, a.logoSvg, `left:${b[0]}px;top:${b[1]}px;width:${b[2]}px;height:${b[3]}px`));
    s.appendChild(textAt(doc, esc(T(data).agency), { fs: 18, baseline: 70, right: 40, color: GRAY40 }));

    const title = textAt(doc, esc(T(data).plan(data.model || '')), { fs: 60, w: 500, baseline: 184, left: 40 });
    title.style.letterSpacing = '-0.03em';
    s.appendChild(title);
    const pills = h(doc, 'div', { class: 'abs', style: 'top:140px;display:flex;gap:12px' });
    (data.sources || []).forEach(t => pills.appendChild(h(doc, 'div', { class: 'pill m', style: `height:46px;padding:0 24px;font-size:22px;color:${BLUE}` }, esc(t))));
    s.appendChild(pills);
    let logo = null;
    if (data.clientLogo) {
        logo = h(doc, 'div', { class: 'abs', style: 'right:40px;top:138px;height:56px;display:flex;justify-content:flex-end' });
        logo.appendChild(h(doc, 'img', { src: data.clientLogo, alt: '', style: 'height:56px;width:auto;max-width:300px;object-fit:contain;border-radius:6px;display:block' }));
        s.appendChild(logo);
    }
    s.appendChild(h(doc, 'div', { class: 'abs', style: `left:40px;top:226px;width:1840px;height:2px;background:${LINE}` }));
    s.appendChild(h(doc, 'div', { class: 'abs', style: `left:40px;top:328px;width:1840px;height:2px;background:${LINE}` }));

    const info = h(doc, 'div', { class: 'info abs', style: `left:40px;top:${topFor(266, 16)}px` });
    (data.info || []).forEach((it, i) => {
        const col = h(doc, 'div');
        col.appendChild(h(doc, 'div', { class: 'l' }, esc(it.label)));
        col.appendChild(h(doc, 'div', { class: 'v' + (i === 0 ? ' m' : '') }, esc(it.value)));
        info.appendChild(col);
    });
    s.appendChild(info);

    // table cards
    const side = spec.groups.length === 2;
    const geo = side ? [[40, 850], [946, 934]] : [[40, 1840]];
    const cards = [];
    spec.groups.forEach((g, i) => {
        const [x, w] = geo[i];
        const sec = h(doc, 'div', { class: 'sec abs', style: `left:${x + 8}px;top:${topFor(388, 20)}px` });
        sec.appendChild(h(doc, 'span', { class: 'n' }, g.kind === 'k' ? '(02)' : '(01)'));
        sec.appendChild(h(doc, 'span', { class: 't' }, esc(g.kind === 'k' ? T(data).sec2 : T(data).sec1)));
        s.appendChild(sec);
        const card = h(doc, 'div', { class: 'card', style: `left:${x}px;top:${CARD_TOP}px;width:${w}px` });
        card.appendChild(buildTable(doc, g, spec.withTotal, g.kind === 'k' ? 16 : 17, g.kind === 'k' ? 18 : 19));
        cards.push({ card, x, w });
    });
    const shadeHost = h(doc, 'div', { class: 'abs', style: 'left:0;top:0' });
    s.appendChild(shadeHost);
    cards.forEach(c => s.appendChild(c.card));
    let divider = null;
    if (side) {
        divider = h(doc, 'div', { class: 'abs', style: `left:918px;top:372px;width:2px;background:${INK}` });
        s.appendChild(divider);
    }
    let sum = null;
    if (spec.isLast) {
        sum = buildSummary(doc, data);
        s.appendChild(sum);
    }
    s._title = title; s._pills = pills; s._logo = logo; s._info = info;
    s._cards = cards; s._shadeHost = shadeHost; s._divider = divider; s._sum = sum;
    return s;
}

/** Title row and info row: shrink when they would not fit. */
function settleChrome(s) {
    const titleRight = 40 + s._title.offsetWidth;
    const limit = (s._logo ? SLIDE_W - 40 - s._logo.offsetWidth - 40 : SLIDE_W - 40);
    s._pills.style.left = (titleRight + 30) + 'px';
    const pills = Array.from(s._pills.children);
    const fits = () => titleRight + 30 + s._pills.offsetWidth <= limit;
    // too many sources: smaller pills, then the tail goes into "+N"
    if (!fits()) searchMax(14, 22, v => {
        pills.forEach(p => { p.style.fontSize = v + 'px'; p.style.padding = `0 ${Math.round(v * 24 / 22)}px`; });
        return fits();
    }, 0.5);
    let hidden = 0;
    while (!fits() && pills.length > 1) {
        pills.pop().remove();
        hidden++;
        let more = s._pills.querySelector('.more');
        if (!more) { more = pills[0].cloneNode(); more.classList.add('more'); s._pills.appendChild(more); }
        more.textContent = '+' + hidden;
    }
    // info row wider than the slide: shrink its text
    if (s._info.offsetWidth > 1840) searchMax(0.6, 1, v => {
        s._info.style.gap = (64 * v) + 'px';
        s._info.querySelectorAll('.l').forEach(e => { e.style.fontSize = 16 * v + 'px'; });
        s._info.querySelectorAll('.v').forEach(e => { e.style.fontSize = 21 * v + 'px'; });
        return s._info.offsetWidth <= 1840;
    }, 0.02);
}

/** Apply table scale k; returns true when everything fits. */
function applyScale(s, k) {
    const bottomLimit = s._sum ? SLIDE_BOTTOM - SUM_H - SUM_GAP : SLIDE_BOTTOM;
    let ok = true, maxBottom = CARD_TOP;
    s._cards.forEach(({ card }) => {
        const t = card.querySelector('table');
        t.style.setProperty('--k', k);
        if (t.offsetWidth > card.clientWidth - 40 + 0.5) ok = false;          // wider than the card
        maxBottom = Math.max(maxBottom, CARD_TOP + card.offsetHeight);
    });
    if (maxBottom > bottomLimit + 0.5) ok = false;
    return ok;
}
/** After the scale is fixed: shadows, divider, budget cards under the tables. */
function finishPlan(s) {
    const doc = s.ownerDocument;
    let maxBottom = CARD_TOP;
    s._shadeHost.innerHTML = '';
    s._cards.forEach(({ card, x, w }) => {
        const hgt = card.offsetHeight;
        maxBottom = Math.max(maxBottom, CARD_TOP + hgt);
        s._shadeHost.appendChild(shadow(doc, x, CARD_TOP, w, hgt, 24));
    });
    if (s._divider) s._divider.style.height = (maxBottom - 6 - 372) + 'px';
    if (s._sum) {
        const top = SLIDE_BOTTOM - SUM_H;          // budget cards always at the bottom, as in the reference
        s._sum.style.top = top + 'px';
        settleSummary(s._sum);
        const host = h(doc, 'div', { class: 'abs', style: 'left:0;top:0' });
        s._sum._cards.forEach(c => {
            if (!c.classList.contains('g')) host.appendChild(shadow(doc, 40 + c.offsetLeft, top, c.offsetWidth, SUM_H, 24));
        });
        s.insertBefore(host, s._sum);
    }
}

// ── Planner ────────────────────────────────────────────────────────────────
function chunkBounds(n, k) {
    const out = [], size = Math.ceil(n / k);
    for (let i = 0; i < n; i += size) out.push([i, Math.min(n, i + size)]);
    return out;
}
function slice(g, a, b) {
    return { kind: g.kind, lead: g.lead, cols: g.cols, rows: g.rows.slice(a, b), gb: g.gb.slice(a, b), total: g.total };
}
function expandOption(plan, k, layout) {
    const n = plan.buy.rows.length;
    const bounds = chunkBounds(n, k);
    const specs = [];
    if (layout === 'side') {
        bounds.forEach(([a, b], i) => specs.push({ groups: [slice(plan.buy, a, b), slice(plan.funnel, a, b)], withTotal: i === bounds.length - 1 }));
    } else if (layout === 'split') {
        bounds.forEach(([a, b], i) => specs.push({ groups: [slice(plan.buy, a, b)], withTotal: i === bounds.length - 1 }));
        bounds.forEach(([a, b], i) => specs.push({ groups: [slice(plan.funnelLabelled || plan.funnel, a, b)], withTotal: i === bounds.length - 1 }));
    } else {
        bounds.forEach(([a, b], i) => specs.push({ groups: [slice(plan.buy, a, b)], withTotal: i === bounds.length - 1 }));
    }
    specs.forEach((sp, i) => { sp.isLast = i === specs.length - 1; });
    return specs;
}
function evaluate(doc, data, specs) {
    let min = Infinity;
    for (const sp of specs) {
        const s = buildPlanSlide(doc, data, sp);
        s.classList.add('measure');
        doc.body.appendChild(s);
        settleChrome(s);
        const r = searchMax(K_FLOOR, 1, k => applyScale(s, k), 0.01);
        s.remove();
        min = Math.min(min, r.value);
        if (min < K_FLOOR + 0.005) break;
    }
    return min;
}
function planTables(doc, data) {
    const plan = data.plan;
    const n = Math.max(1, plan.buy.rows.length);
    const layouts = plan.funnel ? ['side', 'split'] : ['single'];
    const options = [];
    for (let k = 1; k <= Math.min(n, MAX_SLIDES); k++) {
        layouts.forEach((l, pref) => {
            const slides = l === 'split' ? 2 * k : k;
            if (slides <= MAX_SLIDES) options.push({ k, layout: l, slides, pref });
        });
    }
    options.sort((p, q) => p.slides - q.slides || p.pref - q.pref);
    let best = null, fallback = null;
    for (let i = 0; i < options.length;) {
        const count = options[i].slides;
        const group = [];
        while (i < options.length && options[i].slides === count) group.push(options[i++]);
        for (const o of group) {
            o.specs = expandOption(plan, o.k, o.layout);
            o.scale = evaluate(doc, data, o.specs);
            if (!fallback || o.scale > fallback.scale + 0.001) fallback = o;
        }
        const ok = group.filter(o => o.scale >= K_MIN);
        if (ok.length) {
            ok.sort((p, q) => (q.scale - p.scale > 0.02 ? 1 : p.scale - q.scale > 0.02 ? -1 : p.pref - q.pref));
            best = ok[0];
            break;
        }
    }
    return { option: best || fallback, readable: !!best };
}

// ── Render ─────────────────────────────────────────────────────────────────
async function render(doc, data) {
    let style = doc.getElementById('mobx-deck-css');
    if (!style) {
        style = h(doc, 'style', { id: 'mobx-deck-css' });
        style.textContent = deckCss();
        doc.head.appendChild(style);
    }
    doc.body.innerHTML = '';
    const warm = h(doc, 'div', { style: 'position:absolute;left:-9999px;top:0' },
        '<span style="font-weight:400">Аa1$₽·</span><b style="font-weight:500">Аa1$₽·</b>');
    doc.body.appendChild(warm);
    try { await Promise.all([doc.fonts.load('400 20px "Golos Text"', 'Аa1$₽·'), doc.fonts.load('500 20px "Golos Text"', 'Аa1$₽·')]); } catch (e) { /* fall back silently */ }
    await doc.fonts.ready;
    warm.remove();
    if (data.clientLogo) {
        const probe = h(doc, 'img', { src: data.clientLogo, style: 'position:absolute;left:-9999px;height:56px' });
        doc.body.appendChild(probe);
        await waitImages(doc.body);
        probe.remove();
    }

    const { option, readable } = planTables(doc, data);
    const scale = Math.min(1, option.scale);

    const cover = buildCover(doc, data);
    doc.body.appendChild(cover);
    cover._fit();
    option.specs.forEach(sp => {
        const s = buildPlanSlide(doc, data, sp);
        doc.body.appendChild(s);
        settleChrome(s);
        applyScale(s, scale);          // same table size on every plan slide
        finishPlan(s);
    });
    const closing = buildClosing(doc, data);
    doc.body.appendChild(closing);
    closing._fit();

    await waitImages(doc.body);
    return {
        slides: option.specs.length + 2,
        planSlides: option.specs.length,
        layout: option.layout,
        rowsPerSlide: Math.ceil(Math.max(1, data.plan.buy.rows.length) / option.k),
        scale: Math.round(scale * 100) / 100,
        fontPx: Math.round(19 * scale * 10) / 10,
        readable,
        overflow: checkOverflow(doc)
    };
}

function checkOverflow(doc) {
    const issues = [];
    doc.querySelectorAll('section.slide[data-kind="plan"]').forEach((s, i) => {
        s.querySelectorAll('.card').forEach(card => {
            const t = card.querySelector('table');
            if (t && t.offsetWidth > card.clientWidth - 40 + 1) issues.push(`plan slide ${i + 1}: table wider than its card`);
            if (card.offsetTop + card.offsetHeight > SLIDE_BOTTOM + 1) issues.push(`plan slide ${i + 1}: table below the slide`);
        });
        const sum = s.querySelector('.sum');
        if (sum && sum.offsetTop + SUM_H > SLIDE_H) issues.push(`plan slide ${i + 1}: budget cards below the slide`);
    });
    return issues;
}

window.MobXDeck = { render, deckCss, SLIDE_W, SLIDE_H, K_MIN };
})();
