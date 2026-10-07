/* ═══════════════════════════════════════════════════════════════════════════
   MobX presentation generator
   ───────────────────────────────────────────────────────────────────────────
   Same idea as deck.js (Gravils): takes a finished media plan (values exactly
   as shown in the plan, no recalculation) and lays it out on 1920×1080 slides
     cover → media-plan slide(s) → contacts
   in the MobX style (light background, blue #0000E1, Golos Text, Russian).
   Tables are auto-fitted: every table size derives from one variable (--fs,
   cell text size); the planner tries one slide with both tables, the buy and
   funnel tables on separate slides, and row pagination, and keeps the option
   with the fewest slides whose text stays readable, then the largest text.
   All colours are opaque (no transparency in the PDF), like deck.js v1.1.

   Public API (window.MobXDeck):
     render(doc, data) → Promise<report>
     deckCss()         → string
     SLIDE_W, SLIDE_H, FS_MIN, FS_MAX
   data: see deck-mobx-ui.js (collect()).
   ═══════════════════════════════════════════════════════════════════════════ */
(function () {
'use strict';

const SLIDE_W = 1920, SLIDE_H = 1080;
const BLUE = '#0000E1', BG = '#F8F8F8';
const GRAY_TAG = '#959595';     // black 40% over #F8F8F8 (tagline)
const GRAY_LABEL = '#474747';   // black 72% over white (summary labels)
const LINE_SOFT = '#DBDBDB';    // black 14% over white (dashed lines, separator)
const FS_MAX = 30;              // table text at reference size
const FS_MIN = 18;              // smallest comfortable table text, px
const FS_FLOOR = 9;             // search floor (only to rank options that don't fit)
const ROW_FACTORS = [2.67, 2.3, 2.0];   // row height = factor × text size (reference 80px / 30px)
const MAX_SLIDES = 10;
const TABLE_X = 80, TABLE_W = 1760, TABLE_TOP = 366, TABLE_BOTTOM = 872;

const A = () => window.MOBX_DECK_ASSETS || {};

// Golos Text: ascent 0.98, descent 0.22 (em). With line-height = font-size the
// baseline sits 0.88·fs below the top of the line box.
const topFor = (baseline, fs) => baseline - 0.88 * fs;

// ── Styles ─────────────────────────────────────────────────────────────────
function deckCss() {
    const a = A();
    return `${a.fontFaces || ''}
@page{size:${SLIDE_W}px ${SLIDE_H}px;margin:0}
*{box-sizing:border-box;margin:0;padding:0}
html,body{background:#d9d9d9}
body{font-family:"Golos Text",Arial,sans-serif;color:#000;-webkit-print-color-adjust:exact;print-color-adjust:exact;-webkit-font-smoothing:antialiased}
.slide{position:relative;width:${SLIDE_W}px;height:${SLIDE_H}px;overflow:hidden;background:${BG};margin:0 auto 40px}
.slide.measure{position:absolute;left:-99999px;top:0;margin:0}
@media print{html,body{background:${BG}}.slide{margin:0;break-after:page;page-break-after:always}.slide:last-child{break-after:auto;page-break-after:auto}}
.bg{position:absolute;left:0;top:0;width:${SLIDE_W}px;height:${SLIDE_H}px}
.abs{position:absolute}
.nw{white-space:nowrap}
.m{font-weight:500}
.b{color:${BLUE}}
.pills{display:flex}
.pill{display:flex;align-items:center;white-space:nowrap;border-radius:999px;line-height:1}
.pill.o{border:2px solid ${BLUE};color:${BLUE}}
.pill.f{background:${BLUE};color:${BG}}
a.pill{text-decoration:none}

/* plan card: opaque layered shadow (no transparency) */
.card,.shade{position:absolute;border-radius:24px}
.card{background:#fff}

/* tables */
.region{position:absolute;left:${TABLE_X}px;width:${TABLE_W}px;top:${TABLE_TOP}px;overflow:hidden;--fs:30px;--rf:2.67;display:flex;flex-direction:column;gap:calc(var(--fs)*1.2667)}
.t .grid{display:grid;column-gap:16px}
.t .hd{height:calc(var(--fs)*2)}
.t .hd > div{display:flex;align-items:center;justify-content:center;text-align:center;border-radius:999px;font-weight:500;font-size:calc(var(--fs)*.8);line-height:1.05;overflow:hidden;padding:0 10px}
.t .hd.f > div{background:${BLUE};color:${BG}}
.t .hd.o > div{border:2px solid ${BLUE};color:${BLUE}}
.t .dash{border-top:2px dashed ${LINE_SOFT}}
.t .hd + .dash{margin-top:calc(var(--fs)/3)}
.t .solid{border-top:2px solid #000}
.t .row{height:calc(var(--fs)*var(--rf))}
.t .row > div{display:flex;align-items:center;justify-content:center;font-size:var(--fs);line-height:1;white-space:nowrap;overflow:hidden;min-width:0}
.t .row > div.lead{font-weight:500}
.t .row.tot > div{font-weight:500;color:${BLUE}}
.t .row.tot > div.lead{color:#000}
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
    step = step || 0.25;
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
    const st = [`font-size:${o.fs}px`, `line-height:${o.lh || o.fs}px`, `font-weight:${o.w || 400}`];
    const lh = o.lh || o.fs;
    st.push(`top:${(o.baseline - lh / 2 - 0.38 * o.fs).toFixed(2)}px`);
    if (o.left != null) st.push(`left:${o.left}px`);
    if (o.right != null) st.push(`right:${o.right}px;text-align:right`);
    if (o.color) st.push(`color:${o.color}`);
    return h(doc, 'div', { class: 'abs nw', style: st.join(';') }, html);
}
/** Pill row; o = {fs, height, padX, gap, kind:'o'|'f', lift (px the text sits above centre)} */
function pills(doc, items, o) {
    const row = h(doc, 'div', { class: 'pills abs', style: `gap:${o.gap}px` });
    items.forEach(it => {
        const tag = it.href ? 'a' : 'div';
        const attrs = { class: 'pill ' + o.kind, style: `height:${o.height}px;padding:0 ${o.padX}px ${2 * (o.lift || 0)}px;font-size:${o.fs}px` };
        if (it.href) attrs.href = it.href;
        row.appendChild(h(doc, tag, attrs, esc(it.text)));
    });
    return row;
}
function logo(doc, style) {
    const a = A();
    const box = h(doc, 'div', { class: 'abs', style }, a.logoSvg || '');
    const svg = box.querySelector('svg');
    if (svg) { svg.style.width = '100%'; svg.style.height = '100%'; svg.style.display = 'block'; }
    return box;
}

// ── Cover ──────────────────────────────────────────────────────────────────
function buildCover(doc, data) {
    const a = A();
    const s = h(doc, 'section', { class: 'slide', 'data-kind': 'cover' });
    s.appendChild(h(doc, 'img', { class: 'bg', src: a.bgCover, alt: '' }));

    const lines = (data.coverTitle && data.coverTitle.length ? data.coverTitle : ['Медиаплан', 'интернет-размещения']);
    const title = textAt(doc, lines.map(esc).join('<br>'), { fs: 110, lh: 104, w: 500, baseline: 232, left: 50 });
    s.appendChild(title);

    const coverPills = pills(doc, data.coverPills || [], { fs: 32, height: 54, padX: 32, gap: 22, kind: 'o', lift: 1 });
    coverPills.style.left = '50px';
    coverPills.style.top = '390px';
    s.appendChild(coverPills);

    // app block, centred in the white field of the background (996..1845 × 744..1000)
    const app = h(doc, 'div', { class: 'abs', style: 'left:996px;top:744px;width:849px;height:256px;display:flex;align-items:center;justify-content:center;gap:32px;padding:0 36px' });
    if (data.appIcon) app.appendChild(h(doc, 'img', { src: data.appIcon, alt: '', style: 'width:140px;height:140px;border-radius:32px;object-fit:cover;flex-shrink:0;display:block' }));
    const name = h(doc, 'div', { class: 'm', style: 'font-size:52px;line-height:54px;min-width:0' }, '');
    app.appendChild(name);
    s.appendChild(app);

    s._fit = () => {
        // title: shrink only if a custom line is too long for the left column
        searchMax(64, 110, v => {
            title.style.fontSize = v + 'px';
            title.style.lineHeight = (v * 104 / 110) + 'px';
            title.style.top = (232 - (v * 104 / 110) / 2 - 0.38 * v) + 'px';
            return title.offsetWidth <= 1250;
        }, 1);
        // app name: one line if it fits, otherwise two balanced lines, then shrink
        const maxW = 849 - 72 - (data.appIcon ? 172 : 0);
        const text = String(data.appName || data.client || '').trim();
        const fitsAt = (html, v) => {
            name.innerHTML = html;
            name.style.fontSize = v + 'px';
            name.style.lineHeight = (v * 54 / 52) + 'px';
            name.style.whiteSpace = 'nowrap';
            return name.scrollWidth <= maxW;
        };
        let html = esc(text);
        if (!fitsAt(html, 52)) {
            const two = splitTwo(text);
            html = two ? esc(two[0]) + '<br>' + esc(two[1]) : esc(text);
        }
        searchMax(28, 52, v => fitsAt(html, v), 0.5);
    };
    return s;
}
/** Split a name into two balanced lines, preferring a break after ':' or '—'. */
function splitTwo(text) {
    const words = text.split(/\s+/);
    if (words.length < 2) return null;
    let best = null;
    for (let i = 1; i < words.length; i++) {
        const a = words.slice(0, i).join(' '), b = words.slice(i).join(' ');
        let score = Math.max(a.length, b.length);
        if (/[:—–-]$/.test(a)) score -= 4;
        if (!best || score < best.score) best = { a, b, score };
    }
    return [best.a, best.b];
}

// ── Contacts ───────────────────────────────────────────────────────────────
function buildContacts(doc, data) {
    const a = A();
    const s = h(doc, 'section', { class: 'slide', 'data-kind': 'contacts' });
    s.appendChild(h(doc, 'img', { class: 'bg', src: a.bgContacts, alt: '' }));
    const title = textAt(doc, 'Напишите нам', { fs: 255, w: 500, baseline: 342, left: 40 });
    s.appendChild(title);
    const items = [{ text: 'Web-site', href: 'https://mobx.agency/' }];
    (data.emails && data.emails.length ? data.emails : ['go@mobx.agency'])
        .forEach(e => items.push({ text: 'Почта: ' + e, href: 'mailto:' + e }));
    const row = pills(doc, items, { fs: 42, height: 82, padX: 42, gap: 40, kind: 'f', lift: 3 });
    row.style.left = '40px';
    row.style.top = '418px';
    row.style.flexWrap = 'wrap';
    row.style.maxWidth = '1840px';
    row.style.rowGap = '20px';
    s.appendChild(row);
    s.appendChild(textAt(doc, 'Мы в соцсетях', { fs: 44, baseline: 807, left: 42 }));
    s._fit = () => {
        // many e-mails: shrink the pills until they stay above the social block (y ≈ 740)
        searchMax(24, 42, v => {
            row.querySelectorAll('.pill').forEach(p => {
                p.style.fontSize = v + 'px';
                p.style.height = Math.round(v * 82 / 42) + 'px';
                p.style.padding = '0 ' + Math.round(v) + 'px ' + Math.round(v / 7) + 'px';
            });
            return row.offsetTop + row.offsetHeight <= 720;
        }, 0.5);
    };
    return s;
}

// ── Media-plan slide ───────────────────────────────────────────────────────
/** group = { kind:'f'|'o', cols:[label], rows:[[cell]], total:[cell]|null, lead:bool } */
function buildTable(doc, group, withTotal) {
    const n = group.cols.length;
    const t = h(doc, 'div', { class: 't' });
    const tpl = `repeat(${n},minmax(0,1fr));column-gap:${n >= 10 ? 10 : 16}px`;
    const hd = h(doc, 'div', { class: 'grid hd ' + group.kind, style: 'grid-template-columns:' + tpl });
    group.cols.forEach(c => hd.appendChild(h(doc, 'div', null, esc(c))));
    t.appendChild(hd);
    t.appendChild(h(doc, 'div', { class: 'dash' }));
    group.rows.forEach((r, i) => {
        if (i > 0) t.appendChild(h(doc, 'div', { class: 'dash' }));
        const row = h(doc, 'div', { class: 'grid row', style: 'grid-template-columns:' + tpl });
        r.forEach((v, j) => row.appendChild(h(doc, 'div', (j === 0 && group.lead) ? { class: 'lead' } : null, esc(v))));
        t.appendChild(row);
    });
    if (withTotal && group.total) {
        t.appendChild(h(doc, 'div', { class: 'solid' }));
        const row = h(doc, 'div', { class: 'grid row tot', style: 'grid-template-columns:' + tpl });
        group.total.forEach((v, j) => row.appendChild(h(doc, 'div', (j === 0 && group.lead) ? { class: 'lead' } : null, esc(v))));
        t.appendChild(row);
    }
    return t;
}

function headlineHtml(data) {
    const k = data.headline || {};
    const num = v => `<span class="b">${esc(v)}</span>`;
    const l1 = `${esc(k.prefix || 'Медиаплан:')} ${num(k.n1)} ${esc(k.w1)}`;
    const l2 = k.n2 != null ? `и ${num(k.n2)} ${esc(k.w2)}${k.period ? ' ' + esc(k.period) : ''}` : esc(k.period || '');
    return l2 ? l1 + '<br>' + l2 : l1;
}

function buildPlanSlide(doc, data, spec) {
    const s = h(doc, 'section', { class: 'slide', 'data-kind': 'plan' });
    s.appendChild(logo(doc, 'left:40px;top:50px;width:151.9px;height:41.8px'));
    s.appendChild(textAt(doc, 'Агентство мобильного маркетинга', { fs: 24, baseline: 76, right: 40, color: GRAY_TAG }));

    const head = textAt(doc, headlineHtml(data), { fs: 72, lh: 74, w: 500, baseline: 200, left: 40 });
    s.appendChild(head);
    const pr = pills(doc, data.planPills || [], { fs: 26, height: 58, padX: 31, gap: 16, kind: 'o' });
    pr.style.right = '40px';
    pr.style.top = '228px';
    s.appendChild(pr);

    // card with an opaque layered shadow (pre-blended over #F8F8F8)
    [[20, '#F7F7F7'], [16, '#F6F6F6'], [12, '#F5F5F5'], [8, '#F2F2F2'], [4, '#EFEFEF']].forEach(([o, c]) =>
        s.appendChild(h(doc, 'div', { class: 'shade', style: `left:${40 - o}px;top:${330 - o + 2}px;width:${1840 + 2 * o}px;height:${700 + 2 * o}px;background:${c};border-radius:${24 + o}px` })));
    s.appendChild(h(doc, 'div', { class: 'card', style: 'left:40px;top:330px;width:1840px;height:700px' }));

    const region = h(doc, 'div', { class: 'region', style: `height:${TABLE_BOTTOM - TABLE_TOP}px` });
    spec.groups.forEach(g => region.appendChild(buildTable(doc, g, spec.withTotal)));
    s.appendChild(region);

    // separator + budget summary (bottom of the card)
    s.appendChild(h(doc, 'div', { class: 'abs', style: `left:${TABLE_X}px;top:910px;width:${TABLE_W}px;border-top:2px solid ${LINE_SOFT}` }));
    if (spec.isLast) {
        const sum = data.summary || [];
        const box = h(doc, 'div', { class: 'abs nw', style: `right:${SLIDE_W - TABLE_X - TABLE_W}px;top:${topFor(992, 64)}px;display:flex;align-items:baseline;gap:32px` });
        sum.forEach((it, i) => {
            const last = i === sum.length - 1;
            box.appendChild(h(doc, 'div', { style: `font-size:30px;line-height:64px;color:${GRAY_LABEL};position:relative;top:-12px` }, esc(it.label)));
            box.appendChild(h(doc, 'div', { class: 'm', style: last ? `font-size:64px;line-height:64px;color:${BLUE}` : 'font-size:30px;line-height:64px;color:#000;position:relative;top:-12px;margin-right:24px' }, esc(it.value)));
        });
        s.appendChild(box);
    }
    s._region = region;
    s._head = head;
    s._pills = pr;
    return s;
}

/** Headline and pills share the top band: shrink the headline if they would touch. */
function settleChrome(s) {
    const pillsLeft = SLIDE_W - 40 - s._pills.offsetWidth;
    searchMax(48, 72, v => {
        s._head.style.fontSize = v + 'px';
        s._head.style.lineHeight = (v * 74 / 72) + 'px';
        // keep the second baseline at y 274
        s._head.style.top = (274 - (v * 74 / 72) * 1.5 - 0.38 * v) + 'px';
        return 40 + s._head.offsetWidth <= pillsLeft - 40;
    }, 0.5);
}

/** For every row factor: the largest --fs at which every table fits the region. */
function fitRegion(s) {
    const region = s._region;
    const cells = Array.from(region.querySelectorAll('.hd > div, .row > div'));
    const fitsAt = (v, rf) => {
        region.style.setProperty('--fs', v + 'px');
        region.style.setProperty('--rf', rf);
        if (region.scrollHeight > region.clientHeight + 0.5) return false;
        return cells.every(c => c.scrollWidth <= c.clientWidth + 0.5 && c.scrollHeight <= c.clientHeight + 0.5);
    };
    return ROW_FACTORS.map(rf => searchMax(FS_FLOOR, FS_MAX, v => fitsAt(v, rf), 0.25).value);
}
/** One row factor for the whole deck: prefer the airier one unless a denser one gains more than 1px. */
function pickRow(perRf) {
    let best = null;
    perRf.forEach((fs, i) => { if (!best || fs > best.fs + 1) best = { fs, rf: ROW_FACTORS[i] }; });
    return best;
}

// ── Planner ────────────────────────────────────────────────────────────────
function chunkBounds(n, k) {
    const out = [], size = Math.ceil(n / k);
    for (let i = 0; i < n; i += size) out.push([i, Math.min(n, i + size)]);
    return out;
}
function slice(g, a, b) { return { kind: g.kind, lead: g.lead, cols: g.cols, rows: g.rows.slice(a, b), total: g.total }; }

function expandOption(plan, k, layout) {
    const n = plan.buy.rows.length;
    const bounds = chunkBounds(n, k);
    const specs = [];
    if (layout === 'both') {
        bounds.forEach(([a, b], i) => specs.push({ groups: [slice(plan.buy, a, b), slice(plan.funnel, a, b)], withTotal: i === bounds.length - 1 }));
    } else { // split: buy slides, then funnel slides
        bounds.forEach(([a, b], i) => specs.push({ groups: [slice(plan.buy, a, b)], withTotal: i === bounds.length - 1 }));
        bounds.forEach(([a, b], i) => specs.push({ groups: [slice(plan.funnel, a, b)], withTotal: i === bounds.length - 1 }));
    }
    specs.forEach((sp, i) => { sp.isLast = i === specs.length - 1; });
    return specs;
}

function evaluate(doc, data, specs) {
    const min = ROW_FACTORS.map(() => Infinity);   // per row factor: smallest fit over all slides
    for (const sp of specs) {
        const s = buildPlanSlide(doc, data, sp);
        s.classList.add('measure');
        doc.body.appendChild(s);
        settleChrome(s);
        fitRegion(s).forEach((fs, i) => { min[i] = Math.min(min[i], fs); });
        s.remove();
        if (Math.max.apply(null, min) < FS_FLOOR + 0.01) break;
    }
    return pickRow(min);
}

function planTables(doc, data) {
    const plan = data.plan;
    const n = Math.max(1, plan.buy.rows.length);
    const layouts = plan.funnel ? ['both', 'split'] : ['buyonly'];
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
            o.specs = o.layout === 'buyonly'
                ? chunkBounds(n, o.k).map(([a, b], j, all) => ({ groups: [slice(plan.buy, a, b)], withTotal: j === all.length - 1, isLast: j === all.length - 1 }))
                : expandOption(plan, o.k, o.layout);
            const r = evaluate(doc, data, o.specs);
            o.fs = r.fs; o.rf = r.rf;
            if (!fallback || o.fs > fallback.fs + 0.01) fallback = o;
        }
        const ok = group.filter(o => o.fs >= FS_MIN);
        if (ok.length) {
            ok.sort((p, q) => (q.fs - p.fs > 0.5 ? 1 : p.fs - q.fs > 0.5 ? -1 : p.pref - q.pref));
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
        '<span style="font-weight:400">Аa1$₽</span><b style="font-weight:500">Аa1$₽</b>');
    doc.body.appendChild(warm);
    try { await Promise.all([doc.fonts.load('400 20px "Golos Text"', 'Аa1$₽'), doc.fonts.load('500 20px "Golos Text"', 'Аa1$₽')]); } catch (e) { /* fall back silently */ }
    await doc.fonts.ready;
    warm.remove();
    if (data.appIcon) {
        const probe = h(doc, 'img', { src: data.appIcon, style: 'position:absolute;left:-9999px' });
        doc.body.appendChild(probe);
        await waitImages(doc.body);
        probe.remove();
    }

    const { option, readable } = planTables(doc, data);
    const specs = option.specs;

    const cover = buildCover(doc, data);
    doc.body.appendChild(cover);
    cover._fit();

    specs.forEach(sp => {
        const s = buildPlanSlide(doc, data, sp);
        doc.body.appendChild(s);
        settleChrome(s);
        // same text size on every plan slide
        s._region.style.setProperty('--fs', option.fs + 'px');
        s._region.style.setProperty('--rf', option.rf);
    });

    const contacts = buildContacts(doc, data);
    doc.body.appendChild(contacts);
    contacts._fit();

    await waitImages(doc.body);
    return {
        slides: specs.length + 2,
        planSlides: specs.length,
        layout: option.layout,
        rowsPerSlide: Math.ceil(Math.max(1, data.plan.buy.rows.length) / option.k),
        fontPx: Math.round(option.fs * 100) / 100,
        rowFactor: option.rf,
        readable,
        overflow: checkOverflow(doc)
    };
}

function checkOverflow(doc) {
    const issues = [];
    doc.querySelectorAll('section.slide').forEach((s, i) => {
        const r = s.querySelector('.region');
        if (!r) return;
        if (r.scrollHeight > r.clientHeight + 1) issues.push(`slide ${i + 1}: table taller than its area`);
        r.querySelectorAll('.hd > div, .row > div').forEach(c => {
            if (c.scrollWidth > c.clientWidth + 1) issues.push(`slide ${i + 1}: "${c.textContent}" does not fit its cell`);
        });
    });
    return issues;
}

window.MobXDeck = { render, deckCss, SLIDE_W, SLIDE_H, FS_MIN, FS_MAX };
})();
