/* ═══════════════════════════════════════════════════════════════════════════
   Gravils presentation generator
   ───────────────────────────────────────────────────────────────────────────
   Takes a finished media plan (no recalculation — values are taken as they
   are shown in the plan) and lays it out on 1920×1080 slides:
     cover → media-plan slide(s) → contacts.
   Tables are auto-fitted: every size in a table derives from one CSS
   variable (--fs); the planner tries several layouts (side by side, stacked,
   split onto two slides, row pagination) and keeps the one with the fewest
   slides whose text stays readable, then the largest font.

   Public API (window.GravilsDeck):
     render(doc, data) → Promise<report>   build slides into a document
     deckCss()         → string            styles for that document
     SLIDE_W, SLIDE_H, FS_MIN, FS_MAX
   ═══════════════════════════════════════════════════════════════════════════ */
(function () {
'use strict';

const SLIDE_W = 1920, SLIDE_H = 1080;
const FS_MIN = 13;          // smallest acceptable table body text, px
const FS_MAX = 30;          // largest table body text, px
const FS_FLOOR = 7;         // search floor (used only to rank layouts that don't fit)
const MAX_SLIDES = 12;      // upper bound for table slides the planner will consider
const X = 70, INNER_W = SLIDE_W - 2 * X;
const TABLE_TOP = 300;
const SIDE_GAP = 44, STACK_GAP = 34;

const A = () => window.GRAVILS_DECK_ASSETS || {};

// ── Styles ─────────────────────────────────────────────────────────────────
/**
 * PDF-safe colours. Semi-transparent fills/gradients become soft-masked
 * transparency groups in the PDF, and several viewers (browser previews,
 * some Mac/Windows apps) render those wrongly — green turns magenta.
 * All colours are therefore made opaque, pre-blended over the black slide.
 */
function flat(css) {
    return css.replace(/rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([\d.]+)\s*\)/g, (m, r, g, b, al) => {
        const k = parseFloat(al);
        const c = v => Math.round(parseInt(v, 10) * k);
        return `rgb(${c(r)},${c(g)},${c(b)})`;
    });
}

function deckCss() {
    const a = A();
    return flat(`
@font-face{font-family:"Graphik LCG";src:url("${a.fontRegular}") format("truetype");font-weight:400;font-style:normal}
@font-face{font-family:"Graphik LCG";src:url("${a.fontBold}") format("truetype");font-weight:700;font-style:normal}
@font-face{font-family:"Montserrat";src:url("${a.fontMontserrat}") format("woff2");font-weight:500;font-style:normal}
@page{size:${SLIDE_W}px ${SLIDE_H}px;margin:0}
*{box-sizing:border-box;margin:0;padding:0}
html,body{background:#1b1b1b}
body{font-family:"Graphik LCG",Arial,sans-serif;color:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact;-webkit-font-smoothing:antialiased}
.slide{position:relative;width:${SLIDE_W}px;height:${SLIDE_H}px;overflow:hidden;background:#000;margin:0 auto 40px}
.slide.measure{position:absolute;left:-99999px;top:0;margin:0}
@media print{html,body{background:#000}.slide{margin:0;break-after:page;page-break-after:always}.slide:last-child{break-after:auto;page-break-after:auto}}
.bg{position:absolute;left:0;top:0;width:${SLIDE_W}px;height:${SLIDE_H}px}
.abs{position:absolute}
.dots{display:flex;gap:12px;align-items:center}
.dots i{display:block;width:20px;height:20px;border-radius:50%;background:#00F889}
.dots i:first-child{background:#fff}
.pageno{font-weight:700;font-size:30px;line-height:1;color:rgba(255,255,255,.45);text-align:right;width:200px}
.logo-box{display:flex;align-items:center}
.logo-box img{display:block;max-width:100%;max-height:100%;object-fit:contain}

/* info cards */
.info{display:flex;gap:14px;--cs:1}
.info .card{flex:1 1 auto;min-width:0;border-radius:22px;padding:calc(20px*var(--cs)) calc(28px*var(--cs));display:flex;flex-direction:column;gap:calc(8px*var(--cs));white-space:nowrap;background:linear-gradient(180deg,rgba(2,183,102,.26),rgba(2,183,102,.08) 55%,rgba(2,183,102,.03));border:1.5px solid rgba(2,183,102,.35)}
.info b{font-weight:700;font-size:calc(27px*var(--cs));line-height:1.1;overflow:hidden;text-overflow:ellipsis}
.info span{font-weight:400;font-size:calc(20px*var(--cs));line-height:1;color:rgba(255,255,255,.9)}

/* summary cards */
.sum{display:grid;gap:44px}
.sum .card{height:82px;border-radius:22px;padding:0 34px;display:flex;align-items:center;justify-content:space-between;gap:24px;white-space:nowrap;background:linear-gradient(180deg,rgba(2,183,102,.30),rgba(2,183,102,.10) 55%,rgba(2,183,102,.04));border:1.5px solid rgba(2,183,102,.4)}
.sum .l{font-weight:400;font-size:var(--sl,27px);line-height:1}
.sum .v{font-weight:700;font-size:var(--sv,31px);line-height:1}

/* table region */
.region{position:absolute;left:${X}px;width:${INNER_W}px;overflow:hidden;--fs:16px}
.region.side{display:grid;column-gap:${SIDE_GAP}px;align-items:start}
.region.stack{display:flex;flex-direction:column;gap:${STACK_GAP}px}
.tw{min-width:0;overflow:hidden;flex-shrink:0}
.foot{position:absolute;left:${X}px;width:${INNER_W}px;font-size:22px;line-height:1.35;font-weight:400;color:rgba(255,255,255,.85)}

table.mp{width:100%;border-collapse:separate;border-spacing:0;font-size:var(--fs);font-variant-numeric:tabular-nums;white-space:nowrap;--r:calc(var(--fs)*.875)}
.mp th{height:calc(var(--fs)*2.7);font-size:calc(var(--fs)*.875);font-weight:700;line-height:1.1;text-align:center;padding:calc(var(--fs)*.3) calc(var(--fs)*.55);background:linear-gradient(180deg,rgba(2,183,102,.32),rgba(2,183,102,.10));border-left:1px solid rgba(2,183,102,.35)}
.mp td{height:calc(var(--fs)*2.75);text-align:center;padding:0 calc(var(--fs)*.55);border-bottom:1px solid rgba(255,255,255,.14);border-left:1px solid rgba(2,183,102,.18)}
.mp th:first-child,.mp td:first-child{border-left:0}
.mp th:first-child{border-radius:var(--r) 0 0 var(--r)}
.mp th:last-child{border-radius:0 var(--r) var(--r) 0}
.mp.lead th:first-child,.mp.lead td:first-child{text-align:left;padding-left:calc(var(--fs)*1.125)}
.mp.lead td:first-child{font-weight:700}
.mp tr.last td{border-bottom:0}
.mp tr.total td{height:calc(var(--fs)*2.5);font-weight:700;font-size:calc(var(--fs)*1.0625);border:0;background:linear-gradient(180deg,rgba(2,183,102,.30),rgba(2,183,102,.10))}
.mp tr.total td:first-child{border-radius:var(--r) 0 0 var(--r)}
.mp tr.total td:last-child{border-radius:0 var(--r) var(--r) 0}

/* contacts */
.contacts{display:flex;flex-direction:column;gap:26px;--cf:35px}
.contacts .row{display:flex;gap:25px;align-items:center}
.contacts img{width:48px;height:48px;flex-shrink:0;object-fit:contain}
.contacts a{font-family:"Montserrat","Graphik LCG",Arial,sans-serif;font-weight:500;font-size:var(--cf);line-height:1.2;color:#fff;text-decoration:underline;white-space:nowrap}
`);
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
const pad2 = n => String(n).padStart(2, '0');

/** Break a long header label into two balanced lines. */
function splitLabel(label) {
    const s = String(label).trim();
    if (s.length <= 11 || s.indexOf(' ') < 0) return esc(s);
    const words = s.split(/\s+/);
    let best = null;
    for (let i = 1; i < words.length; i++) {
        const a = words.slice(0, i).join(' '), b = words.slice(i).join(' ');
        const score = Math.max(a.length, b.length);
        if (!best || score < best.score) best = { a, b, score };
    }
    return esc(best.a) + '<br>' + esc(best.b);
}

/** Wrap every occurrence of `needle` in bold. */
function boldify(text, needle) {
    const t = esc(text);
    const n = String(needle || '').trim();
    if (!n) return t;
    const e = esc(n).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return t.replace(new RegExp(e, 'g'), m => `<b style="font-weight:700">${m}</b>`);
}

/** Binary search for the largest value in [lo, hi] for which fits(v) is true. */
function searchMax(lo, hi, fits, step) {
    step = step || 0.25;
    if (!fits(lo)) return { value: lo, fits: false };
    if (fits(hi)) return { value: hi, fits: true };
    let a = lo, b = hi;
    while (b - a > step) {
        const m = (a + b) / 2;
        if (fits(m)) a = m; else b = m;
    }
    fits(a);            // leave the element at the found size
    return { value: a, fits: true };
}

function waitImages(root) {
    const imgs = Array.from(root.querySelectorAll('img'));
    return Promise.all(imgs.map(img => img.complete ? null :
        new Promise(r => { img.onload = img.onerror = r; })));
}

// ── Chrome shared by slides ────────────────────────────────────────────────
function logoImg(doc, src, alt, boxStyle, align) {
    const box = h(doc, 'div', { class: 'logo-box abs', style: boxStyle + ';justify-content:' + (align || 'flex-start') });
    if (src) box.appendChild(h(doc, 'img', { src, alt: alt || '' }));
    return box;
}
function dots(doc, style) {
    return h(doc, 'div', { class: 'dots abs', style }, '<i></i><i></i><i></i>');
}

// ── Gradient headlines ─────────────────────────────────────────────────────
// Drawn as SVG text with a gradient fill: in the PDF this is a plain vector
// gradient clipped by selectable text — no transparency group, so it looks
// the same in every viewer (CSS background-clip:text is exported as a
// transparency group, which some viewers render incorrectly).
const COVER_STOPS = [[0.04, '#FFFFFF'], [0.34, '#C6E8D6'], [0.62, '#3ECB8B'], [0.84, '#02B766']];
const THANKS_STOPS = [[0.42, '#FFFFFF'], [0.62, '#B7EFD1'], [0.88, '#00F889']];
let gradSeq = 0;
function gradientText(doc, lines, o) {
    const NS = 'http://www.w3.org/2000/svg';
    const svg = doc.createElementNS(NS, 'svg');
    svg.setAttribute('xmlns', NS);
    svg.style.display = 'block';
    svg.style.overflow = 'visible';
    const id = 'gvgrad' + (++gradSeq);
    const defs = doc.createElementNS(NS, 'defs');
    const lg = doc.createElementNS(NS, 'linearGradient');
    lg.setAttribute('id', id);
    lg.setAttribute('x1', '0'); lg.setAttribute('y1', String(o.y1));
    lg.setAttribute('x2', '1'); lg.setAttribute('y2', String(o.y2));
    o.stops.forEach(([off, col]) => {
        const st = doc.createElementNS(NS, 'stop');
        st.setAttribute('offset', String(off));
        st.setAttribute('stop-color', col);
        lg.appendChild(st);
    });
    defs.appendChild(lg);
    svg.appendChild(defs);
    const text = doc.createElementNS(NS, 'text');
    text.setAttribute('fill', 'url(#' + id + ')');
    text.style.fontFamily = '"Graphik LCG", Arial, sans-serif';
    text.style.fontWeight = '700';
    svg.appendChild(text);
    svg._setLines = ls => {
        while (text.firstChild) text.removeChild(text.firstChild);
        ls.forEach(line => {
            const t = doc.createElementNS(NS, 'tspan');
            t.setAttribute('x', '0');
            t.textContent = line;
            text.appendChild(t);
        });
    };
    svg._setLines(lines);
    svg._layout = fs => {
        text.style.fontSize = fs + 'px';
        text.style.letterSpacing = (o.ls * fs) + 'px';
        const spans = Array.from(text.childNodes);
        spans.forEach((t, i) => t.setAttribute('y', String(fs * 0.8 + i * fs * o.lh)));
        const w = Math.ceil(Math.max(1, ...spans.map(t => t.getComputedTextLength())));
        const hgt = Math.ceil(fs * o.lh * Math.max(0, spans.length - 1) + fs * 1.02);
        svg.setAttribute('width', String(w + 4));
        svg.setAttribute('height', String(hgt));
        return { w, h: hgt };
    };
    return svg;
}
/** Greedy word wrap of lines to maxW at a given font size (uses the SVG to measure). */
function wrapLines(svg, lines, fs, maxW) {
    const out = [];
    lines.forEach(line => {
        let cur = '';
        line.split(/\s+/).forEach(w => {
            const test = cur ? cur + ' ' + w : w;
            svg._setLines([test]);
            if (svg._layout(fs).w > maxW && cur) { out.push(cur); cur = w; } else cur = test;
        });
        if (cur) out.push(cur);
    });
    return out;
}

// ── Cover ──────────────────────────────────────────────────────────────────
function buildCover(doc, data) {
    const a = A();
    const s = h(doc, 'section', { class: 'slide', 'data-kind': 'cover' });
    s.appendChild(h(doc, 'img', { class: 'bg', src: a.bgCover, alt: '' }));

    const lock = h(doc, 'div', { class: 'abs', style: 'left:88px;top:96px;display:flex;align-items:center;gap:64px;height:110px' });
    lock.appendChild(h(doc, 'img', { src: a.logoGravils, alt: 'Gravils', style: 'height:104px;width:auto;display:block' }));
    if (data.clientLogo) {
        lock.appendChild(h(doc, 'div', { style: 'font-size:34px;font-weight:400;color:rgb(191,191,191)' }, '×'));
        const lb = h(doc, 'div', { class: 'logo-box', style: 'height:90px;max-width:620px' });
        lb.appendChild(h(doc, 'img', { src: data.clientLogo, alt: data.client || '', style: 'max-height:90px' }));
        lock.appendChild(lb);
    }
    s.appendChild(lock);

    const block = h(doc, 'div', { class: 'abs', style: 'left:88px;bottom:76px;width:1680px;display:flex;flex-direction:column;gap:36px' });
    let lines = (data.coverTitle || []).filter(Boolean);
    const title = gradientText(doc, lines, { stops: COVER_STOPS, y1: 0.39, y2: 0.61, lh: 0.98, ls: -0.02 });
    const sub = h(doc, 'div', { style: 'font-size:35px;line-height:1.25;font-weight:400;color:rgb(235,235,235);max-width:1680px' },
        boldify(data.subtitle || '', data.client));
    block.appendChild(title);
    block.appendChild(sub);
    s.appendChild(block);
    s._fit = () => {
        // Title: largest size (64–120px) at which every line fits the 1580px column and
        // the block stays under the logos; very long custom lines are wrapped by words.
        const ok = v => { const r = title._layout(v); return r.w <= 1580 && r.h <= 470; };
        if (!ok(64)) { lines = wrapLines(title, lines, 64, 1580); title._setLines(lines); }
        const best = searchMax(64, 120, ok, 1);
        title._layout(best.value);
        searchMax(24, 35, v => { sub.style.fontSize = v + 'px'; return sub.offsetHeight <= v * 1.25 * 2 + 2; }, 0.5);
    };
    return s;
}

// ── Contacts ───────────────────────────────────────────────────────────────
function buildContacts(doc, data) {
    const a = A();
    const s = h(doc, 'section', { class: 'slide', 'data-kind': 'contacts' });
    s.appendChild(h(doc, 'img', { class: 'bg', src: a.bgCover, alt: '' }));
    s.appendChild(dots(doc, 'left:70px;top:78px'));
    if (data.clientLogo) s.appendChild(logoImg(doc, data.clientLogo, data.client, 'right:70px;top:70px;width:520px;height:110px', 'flex-end'));
    const thanks = gradientText(doc, ["Let's grow together!"], { stops: THANKS_STOPS, y1: 0.45, y2: 0.55, lh: 1, ls: -0.01 });
    thanks.classList.add('abs');
    thanks.style.left = '70px';
    thanks.style.top = '268px';
    s.appendChild(thanks);
    s.appendChild(h(doc, 'div', { class: 'abs', style: 'left:70px;top:372px;font-weight:400;font-size:35px;line-height:1.1' }, 'Stay connected with Gravils'));

    const many = (data.emails || []).length > 3;
    const listTop = many ? 470 : 520;
    const list = h(doc, 'div', { class: 'contacts abs', style: 'left:70px;top:' + listTop + 'px' });
    const row = (icon, href, text) => {
        const r = h(doc, 'div', { class: 'row' });
        r.appendChild(h(doc, 'img', { src: icon, alt: '' }));
        r.appendChild(h(doc, 'a', { href }, esc(text)));
        list.appendChild(r);
    };
    row(a.iconWeb, 'https://gravils.com/', 'https://gravils.com/');
    (data.emails || []).forEach(e => row(a.iconMail, 'mailto:' + e, e));
    s.appendChild(list);

    const lock = h(doc, 'div', { class: 'abs', style: 'left:70px;bottom:78px;display:flex;flex-direction:column;gap:8px' });
    lock.appendChild(h(doc, 'img', { src: a.logoGravils, alt: 'Gravils', style: 'height:110px;width:auto;display:block' }));
    lock.appendChild(h(doc, 'span', { style: 'font-weight:400;font-size:26px;line-height:1.3' },
        'Mobile performance <em style="font-style:normal;color:#00F889">marketing agency</em>'));
    s.appendChild(lock);

    // Keep the list above the bottom lockup (which starts at ~y 850).
    s._fit = () => {
        thanks._layout(90);
        const maxH = 850 - listTop - 40;
        const fit = v => {
            list.style.setProperty('--cf', v + 'px');
            list.style.gap = Math.round(v * 0.74) + 'px';
            list.querySelectorAll('img').forEach(i => { i.style.width = i.style.height = Math.round(v * 1.37) + 'px'; });
            return list.offsetHeight <= maxH && list.scrollWidth <= INNER_W;
        };
        searchMax(18, 35, fit, 0.5);
    };
    return s;
}

// ── Media-plan tables ──────────────────────────────────────────────────────
/**
 * group = { cols:[{label, lead?}], rows:[[cell…]], total:[cell…]|null }
 * Returns a <table class="mp">.
 */
function buildTable(doc, group, rows, withTotal) {
    const t = h(doc, 'table', { class: 'mp' + (group.lead ? ' lead' : '') });
    const thead = h(doc, 'thead');
    const trh = h(doc, 'tr');
    group.cols.forEach(c => trh.appendChild(h(doc, 'th', null, splitLabel(c.label))));
    thead.appendChild(trh);
    t.appendChild(thead);
    const tb = h(doc, 'tbody');
    rows.forEach((r, i) => {
        const tr = h(doc, 'tr', (i === rows.length - 1 && withTotal) ? { class: 'last' } : null);
        r.forEach(v => tr.appendChild(h(doc, 'td', null, esc(v))));
        tb.appendChild(tr);
    });
    if (withTotal && group.total) {
        const tr = h(doc, 'tr', { class: 'total' });
        group.total.forEach(v => tr.appendChild(h(doc, 'td', null, esc(v))));
        tb.appendChild(tr);
    }
    t.appendChild(tb);
    return t;
}

/** Funnel table with identifying columns in front (used when it is not beside the main table). */
function labelledFunnel(plan) {
    const f = plan.funnel, m = plan.main;
    const ids = plan.idCols;                      // indexes into main cols
    return {
        lead: true,
        cols: ids.map(i => m.cols[i]).concat(f.cols),
        rows: m.rows.map((r, k) => ids.map(i => r[i]).concat(f.rows[k])),
        total: f.total ? ['Total'].concat(ids.slice(1).map(() => '')).concat(f.total) : null
    };
}

function sliceGroup(g, from, to) {
    return { lead: g.lead, cols: g.cols, rows: g.rows.slice(from, to), total: g.total };
}

/**
 * One media-plan slide.
 * spec = { layout:'side'|'stack'|'single', groups:[g…], withTotal:bool, isLast:bool }
 */
function buildPlanSlide(doc, data, spec) {
    const a = A();
    const s = h(doc, 'section', { class: 'slide', 'data-kind': 'plan' });
    s.appendChild(h(doc, 'img', { class: 'bg', src: a.bgTable, alt: '' }));

    // header: logos
    s.appendChild(logoImg(doc, a.logoGravils, 'Gravils', 'left:70px;top:60px;width:400px;height:52px'));
    if (data.clientLogo) s.appendChild(logoImg(doc, data.clientLogo, data.client, 'right:70px;top:60px;width:360px;height:52px', 'flex-end'));

    // info cards
    const info = h(doc, 'div', { class: 'info abs', style: `left:${X}px;top:150px;width:${INNER_W}px` });
    (data.info || []).forEach(c => info.appendChild(h(doc, 'div', { class: 'card' }, `<b>${esc(c.value)}</b><span>${esc(c.label)}</span>`)));
    s.appendChild(info);

    // bottom: summary cards (last plan slide only), footnotes, dots, page number
    let bottom = spec.isLast ? 900 : 985;
    if (spec.isLast && data.summary && data.summary.length) {
        const n = data.summary.length;
        const sum = h(doc, 'div', { class: 'sum abs', style: `left:${X}px;top:900px;width:${INNER_W}px;grid-template-columns:repeat(${n},1fr)` });
        data.summary.forEach(c => sum.appendChild(h(doc, 'div', { class: 'card' }, `<div class="l">${esc(c.label)}</div><div class="v">${esc(c.value)}</div>`)));
        s.appendChild(sum);
        s._sum = sum;
    }
    let foot = null;
    if (data.footnotes && data.footnotes.length) {
        foot = h(doc, 'div', { class: 'foot', style: `bottom:${SLIDE_H - bottom + 26}px` },
            data.footnotes.map(esc).join('<br>'));
        s.appendChild(foot);
    }
    s.appendChild(dots(doc, 'left:70px;top:1013px'));
    s._pageNo = h(doc, 'div', { class: 'pageno abs', style: 'right:70px;top:1010px' }, '');
    s.appendChild(s._pageNo);

    // table region
    const region = h(doc, 'div', { class: 'region ' + (spec.layout === 'side' ? 'side' : 'stack'), style: `top:${TABLE_TOP}px` });
    if (spec.layout === 'side') {
        region.style.gridTemplateColumns = spec.groups.map(g => g.cols.length + 'fr').join(' ');
    }
    spec.groups.forEach(g => {
        const w = h(doc, 'div', { class: 'tw' });
        w.appendChild(buildTable(doc, g, g.rows, spec.withTotal));
        region.appendChild(w);
    });
    s.appendChild(region);
    s._region = region;
    s._foot = foot;
    s._bottom = bottom;
    s._info = info;
    return s;
}

/** Lay out the fixed chrome of a plan slide (must be in the document). */
function settleChrome(s) {
    // info cards: shrink uniformly if they don't fit on one line
    const info = s._info;
    const cards = Array.from(info.children);
    searchMax(0.6, 1, v => {
        info.style.setProperty('--cs', v);
        return cards.every(c => Array.from(c.children).every(t => t.scrollWidth <= t.clientWidth + 1));
    }, 0.02);
    // still too long at the smallest size: the longest value is cut with an ellipsis (full text in title)
    const cut = cards.filter(c => Array.from(c.children).some(t => t.scrollWidth > t.clientWidth + 1));
    if (cut.length) {
        const longest = cards.reduce((m, c) => (c.scrollWidth > m.scrollWidth ? c : m), cards[0]);
        cards.forEach(c => { c.style.flex = c === longest ? '1 1 0' : '0 0 auto'; });
        longest.querySelector('b').title = longest.querySelector('b').textContent;
    }
    if (s._sum) {
        const sum = s._sum;
        searchMax(0.6, 1, v => {
            sum.style.setProperty('--sl', 27 * v + 'px');
            sum.style.setProperty('--sv', 31 * v + 'px');
            return Array.from(sum.children).every(c => c.scrollWidth <= c.clientWidth + 1);
        }, 0.02);
    }
    const infoBottom = info.offsetTop + info.offsetHeight;
    const top = Math.max(TABLE_TOP, infoBottom + 40);
    let limit = s._bottom - 30;
    if (s._foot) limit = s._foot.offsetTop - 22;
    s._region.style.top = top + 'px';
    s._region.style.height = Math.max(120, limit - top) + 'px';
}

/** Largest --fs at which every table fits the region. */
function fitRegion(s) {
    const region = s._region;
    const tables = Array.from(region.querySelectorAll('table'));
    if (region.classList.contains('side')) balanceSide(region, tables);
    const fits = v => {
        region.style.setProperty('--fs', v + 'px');
        if (region.scrollHeight > region.clientHeight + 0.5) return false;
        return tables.every(t => t.offsetWidth <= t.parentElement.clientWidth + 0.5);
    };
    return searchMax(FS_FLOOR, FS_MAX, fits, 0.25).value;
}

/** Side-by-side: share the width in proportion to what each table naturally needs. */
function balanceSide(region, tables) {
    region.style.setProperty('--fs', '16px');
    const w = tables.map(t => { t.style.width = 'max-content'; const v = t.offsetWidth; t.style.width = ''; return v; });
    region.style.gridTemplateColumns = w.map(v => 'minmax(0,' + Math.max(1, v) + 'fr)').join(' ');
}

/** Last plan slide: when the table is short, pull footnotes and totals up under it. */
function tightenBottom(s) {
    if (!s._sum) return;
    const region = s._region;
    const items = Array.from(region.children);
    const contentBottom = region.offsetTop + Math.max.apply(null, items.map(el => el.offsetTop + el.offsetHeight));
    let next = contentBottom + 56;
    if (s._foot) {
        const footTop = Math.min(contentBottom + 34, s._foot.offsetTop);
        s._foot.style.bottom = '';
        s._foot.style.top = footTop + 'px';
        next = footTop + s._foot.offsetHeight + 40;
    }
    // pull up only when the table is short; otherwise keep the brandbook position (y 900)
    if (900 - next > 220) s._sum.style.top = next + 'px';
    else if (s._foot) { s._foot.style.top = ''; s._foot.style.bottom = (SLIDE_H - 900 + 26) + 'px'; }
}

// ── Planner ────────────────────────────────────────────────────────────────
function chunkBounds(n, k) {
    const out = [], size = Math.ceil(n / k);
    for (let i = 0; i < n; i += size) out.push([i, Math.min(n, i + size)]);
    return out;
}

/** Expand a (k, layout) option into concrete slide specs. */
function expandOption(plan, k, layout) {
    const n = plan.main.rows.length;
    const bounds = chunkBounds(n, k);
    const specs = [];
    if (!plan.funnel) {
        bounds.forEach(([a, b], i) => specs.push({ layout: 'single', groups: [sliceGroup(plan.main, a, b)], withTotal: i === bounds.length - 1 }));
    } else if (layout === 'side' || layout === 'stack') {
        const lf = labelledFunnel(plan);
        bounds.forEach(([a, b], i) => specs.push({
            layout,
            groups: [sliceGroup(plan.main, a, b), layout === 'side' ? sliceGroup(plan.funnel, a, b) : sliceGroup(lf, a, b)],
            withTotal: i === bounds.length - 1
        }));
    } else { // split: all main slides, then all funnel slides
        const lf = labelledFunnel(plan);
        bounds.forEach(([a, b], i) => specs.push({ layout: 'single', groups: [sliceGroup(plan.main, a, b)], withTotal: i === bounds.length - 1 }));
        bounds.forEach(([a, b], i) => specs.push({ layout: 'single', groups: [sliceGroup(lf, a, b)], withTotal: i === bounds.length - 1 }));
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
        const fs = fitRegion(s);
        s.remove();
        min = Math.min(min, fs);
        if (min < FS_FLOOR + 0.01) break;
    }
    return min;
}

function planTables(doc, data) {
    const plan = data.plan;
    const n = Math.max(1, plan.main.rows.length);
    const layouts = plan.funnel ? ['side', 'stack', 'split'] : ['single'];
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
            o.fs = evaluate(doc, data, o.specs);
            if (!fallback || o.fs > fallback.fs + 0.01) fallback = o;
        }
        const ok = group.filter(o => o.fs >= FS_MIN);
        if (ok.length) {
            // biggest text wins; near-ties go to the preferred layout order
            ok.sort((p, q) => (q.fs - p.fs > 0.5 ? 1 : p.fs - q.fs > 0.5 ? -1 : p.pref - q.pref));
            best = ok[0];
            break;
        }
    }
    const chosen = best || fallback;
    return { option: chosen, readable: !!best };
}

// ── Render ─────────────────────────────────────────────────────────────────
async function render(doc, data) {
    // styles + fonts
    let style = doc.getElementById('gravils-deck-css');
    if (!style) {
        style = h(doc, 'style', { id: 'gravils-deck-css' });
        style.textContent = deckCss();
        doc.head.appendChild(style);
    }
    doc.body.innerHTML = '';
    // warm up fonts: put text with every face on the page, then wait
    const warm = h(doc, 'div', { style: 'position:absolute;left:-9999px;top:0' },
        '<span style="font-weight:400">Aa1</span><b style="font-weight:700">Aa1</b><span style="font-family:Montserrat;font-weight:500">Aa1</span>');
    doc.body.appendChild(warm);
    try { await Promise.all([doc.fonts.load('400 20px "Graphik LCG"'), doc.fonts.load('700 20px "Graphik LCG"'), doc.fonts.load('500 20px Montserrat')]); } catch (e) { /* fall back silently */ }
    await doc.fonts.ready;
    warm.remove();

    // wait for client logo so logo boxes have final size
    if (data.clientLogo) {
        const probe = h(doc, 'img', { src: data.clientLogo, style: 'position:absolute;left:-9999px' });
        doc.body.appendChild(probe);
        await waitImages(doc.body);
        probe.remove();
    }

    const { option, readable } = planTables(doc, data);
    const specs = option.specs;
    const fs = option.fs;
    const total = specs.length + 2;

    const cover = buildCover(doc, data);
    doc.body.appendChild(cover);
    cover._fit();

    specs.forEach((sp, i) => {
        const s = buildPlanSlide(doc, data, sp);
        doc.body.appendChild(s);
        settleChrome(s);
        // same text size on every plan slide for consistency
        if (s._region.classList.contains('side')) balanceSide(s._region, Array.from(s._region.querySelectorAll('table')));
        s._region.style.setProperty('--fs', fs + 'px');
        if (sp.isLast) tightenBottom(s);
        s._pageNo.textContent = pad2(i + 2) + ' | ' + pad2(total);
    });

    const contacts = buildContacts(doc, data);
    doc.body.appendChild(contacts);
    contacts._fit();

    await waitImages(doc.body);

    const report = {
        slides: total,
        planSlides: specs.length,
        layout: option.layout,
        rowsPerSlide: Math.ceil(Math.max(1, data.plan.main.rows.length) / option.k),
        fontPx: Math.round(fs * 100) / 100,
        readable,
        overflow: checkOverflow(doc)
    };
    return report;
}

/** Final safety net: report any table that spills out of its box. */
function checkOverflow(doc) {
    const issues = [];
    doc.querySelectorAll('section.slide').forEach((s, i) => {
        const r = s.querySelector('.region');
        if (!r) return;
        if (r.scrollHeight > r.clientHeight + 1) issues.push(`slide ${i + 1}: table taller than its area`);
        r.querySelectorAll('table').forEach(t => {
            if (t.offsetWidth > t.parentElement.clientWidth + 1) issues.push(`slide ${i + 1}: table wider than its area`);
        });
    });
    return issues;
}

window.GravilsDeck = { render, deckCss, SLIDE_W, SLIDE_H, FS_MIN, FS_MAX, _splitLabel: splitLabel };
})();
