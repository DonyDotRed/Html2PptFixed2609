/* DECK·FORGE — HTML/Markdown → PPTX 작업대
   외부 통신 없음. 모든 처리는 브라우저 안에서 끝납니다. */
'use strict';

/* ═══ 0. 기본 도구 ═══════════════════════════════════════════════ */
const $ = (s, r) => (r || document).querySelector(s);
const $$ = (s, r) => [...(r || document).querySelectorAll(s)];
const clean = s => (s || '').replace(/\s+/g, ' ').trim();
const hex = c => String(c || '').replace('#', '').toUpperCase();
const clamp = (v, a, b) => Math.min(b, Math.max(a, v));
const esc = s => String(s).replace(/[&<>"]/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[m]));
const uid = () => 's' + Math.random().toString(36).slice(2, 9);

/* 글자 폭 어림: 한글·한자·가나는 2, 나머지는 1 */
function units(t) {
  let n = 0;
  for (const ch of String(t)) n += /[\u1100-\u11FF\u2E80-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFF00-\uFF60]/.test(ch) ? 2 : 1;
  return n;
}
function lum(h) {
  h = hex(h); const r = parseInt(h.slice(0, 2), 16) / 255, g = parseInt(h.slice(2, 4), 16) / 255, b = parseInt(h.slice(4, 6), 16) / 255;
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}
const onAccent = a => (lum(a) > 0.6 ? '1A1A1A' : 'FFFFFF');
function mix(a, b, t) {
  a = hex(a); b = hex(b); let o = '';
  for (let i = 0; i < 3; i++) {
    const v = Math.round(parseInt(a.substr(i * 2, 2), 16) * (1 - t) + parseInt(b.substr(i * 2, 2), 16) * t);
    o += v.toString(16).padStart(2, '0');
  }
  return o.toUpperCase();
}
function say(t, kind) { const m = $('#msg'); m.textContent = t; m.className = kind || ''; }

/* ═══ 1. 잠금 ════════════════════════════════════════════════════ */
const KEY_SHA = '09090c74379eedfd5cb1abf8fbcf40f0933937fd4e1a9d5f63982f43a18f304f';
const KEY_FNV = 'dca7b0f5';
function fnv(s) { let h = 2166136261 >>> 0; for (const c of s) { h ^= c.charCodeAt(0); h = Math.imul(h, 16777619) >>> 0; } return h.toString(16); }
async function keyOk(v) {
  try {
    if (window.crypto && crypto.subtle) {
      const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(v));
      return [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, '0')).join('') === KEY_SHA;
    }
  } catch (e) { /* 비보안 컨텍스트 → 아래 대체 경로 */ }
  return fnv(v) === KEY_FNV;
}
function openApp() {
  $('#gate').classList.add('hide');
  $('#app').classList.add('on');
  restore();
}
async function tryKey() {
  const v = $('#pw').value;
  if (await keyOk(v)) {
    if ($('#pwKeep').checked) { try { localStorage.setItem('df.key', '1'); } catch (e) { } }
    $('#slotA').setAttribute('fill', '#7fc08a');
    $('#slotB').setAttribute('stroke', '#7fc08a');
    setTimeout(openApp, 220);
  } else {
    $('#pwMsg').textContent = '열쇠말이 맞지 않습니다.';
    $('.gate-card').classList.remove('shake'); void $('.gate-card').offsetWidth; $('.gate-card').classList.add('shake');
    $('#pw').select();
  }
}

/* ═══ 2. 상태 ════════════════════════════════════════════════════ */
/* 판형 — 단위는 인치 (PowerPoint 한계 56인치 안) */
const DIM = {
  '16x9': { W: 10, H: 5.625, name: '와이드 16:9' },
  '16x10': { W: 10, H: 6.25, name: '16:10' },
  '4x3': { W: 10, H: 7.5, name: '표준 4:3' },
  'a4l': { W: 11.69, H: 8.27, name: 'A4 가로' },
  'a4p': { W: 8.27, H: 11.69, name: 'A4 세로' },
  'a3p': { W: 11.69, H: 16.54, name: 'A3 세로' },
  'a2p': { W: 16.54, H: 23.39, name: 'A2 세로', poster: true },
  'a1p': { W: 23.39, H: 33.11, name: 'A1 세로', poster: true },
  'a1l': { W: 33.11, H: 23.39, name: 'A1 가로', poster: true },
  'a0p': { W: 33.11, H: 46.81, name: 'A0 세로', poster: true },
  'a0l': { W: 46.81, H: 33.11, name: 'A0 가로', poster: true }
};
const isPoster = () => !!DIM[state.opts.aspect].poster;

const THEMES = [
  { id: 'reactor', name: '원자로실', bg: '0F1720', fg: 'DCE6F0', title: 'FFFFFF', accent: 'D9A441', muted: '8CA0B4', panel: '1B2836' },
  { id: 'paper', name: '백지', bg: 'FFFFFF', fg: '242424', title: '111111', accent: 'C0392B', muted: '8A8A8A', panel: 'F2F0ED' },
  { id: 'blue', name: '청사진', bg: '0E2A47', fg: 'DCE9F5', title: 'FFFFFF', accent: '6FB2C9', muted: '9DB6CC', panel: '143756' },
  { id: 'doc', name: '보고서', bg: 'FAF8F3', fg: '33302B', title: '1B2A41', accent: '1B6E7A', muted: '8B857C', panel: 'EFEAE0' },
  { id: 'lab', name: '실험실', bg: 'F5F7F6', fg: '26312C', title: '14201A', accent: '2E7D5B', muted: '7C8A83', panel: 'E4EAE6' },
  { id: 'night', name: '야간', bg: '12151A', fg: 'E4E7EB', title: 'FFFFFF', accent: '7FA8D9', muted: '96A0AC', panel: '1D2229' },
  { id: 'brief', name: '브리핑', bg: 'FFFDF7', fg: '342E2B', title: '5A1F1F', accent: 'A8452F', muted: '8C8078', panel: 'F4EBE2' },
  { id: 'hicon', name: '고대비', bg: '000000', fg: 'FFFFFF', title: 'FFE14D', accent: 'FFE14D', muted: 'BBBBBB', panel: '1A1A1A' }
];

const state = {
  raw: '',
  slides: [],
  sel: 0,
  page: 0,
  poster: { cols: 3, scale: 1, sec: 'bar', merge: true, banner: true },
  imgAR: {},                 // src → 가로/세로 비
  opts: {
    split: 'auto', customSel: '', aspect: '16x9', splitLong: true, firstTitle: true,
    images: true, tables: true, notes: true, live: false, shrink: true,
    titleSize: 30, bodySize: 17, headStyle: 'rule', font: '맑은 고딕'
  },
  theme: { ...THEMES[0] },
  deck: { title: '', sub: '', by: '' },
  master: { foot: '', page: true, showFoot: true, logo: null, logoAR: 2.5 },
  exp: { name: 'deck', notes: true }
};

/* ═══ 3. 마크다운 → HTML (가벼운 변환기) ═════════════════════════ */
function looksMD(t) {
  const s = t.trim();
  if (/^\s*<(!doctype|html|div|section|article|h[1-6]|p|ul|table)/i.test(s)) return false;
  return /^#{1,6}\s/m.test(s) || /^\s*[-*+]\s+/m.test(s) || /^\s*\d+\.\s/m.test(s) || /^\|.+\|/m.test(s) || /^---\s*$/m.test(s);
}
function mdInline(s) {
  return esc(s)
    .replace(/!\[([^\]]*)\]\(([^)\s]+)[^)]*\)/g, (m, a, u) => '<img alt="' + a + '" src="' + u + '">')
    .replace(/\[([^\]]+)\]\(([^)\s]+)[^)]*\)/g, '<a href="$2">$1</a>')
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/(^|[^*])\*([^*\n]+)\*/g, '$1<em>$2</em>')
    .replace(/~~([^~]+)~~/g, '<del>$1</del>');
}
function md2html(src) {
  const L = src.replace(/\r/g, '').split('\n');
  const out = []; let i = 0;
  const listStack = [];
  const closeLists = to => {
    while (listStack.length > to) { const s = listStack.pop(); out.push('</' + s.tag + '>' + (s.nested ? '</li>' : '')); }
  };
  while (i < L.length) {
    let ln = L[i];
    if (/^```/.test(ln)) {
      i++; const buf = [];
      while (i < L.length && !/^```/.test(L[i])) buf.push(L[i++]);
      i++; closeLists(0); out.push('<pre><code>' + esc(buf.join('\n')) + '</code></pre>'); continue;
    }
    if (/^\s*$/.test(ln)) { closeLists(0); i++; continue; }
    if (/^\s*(---|\*\*\*|___)\s*$/.test(ln)) { closeLists(0); out.push('<hr>'); i++; continue; }
    const h = ln.match(/^(#{1,6})\s+(.*)$/);
    if (h) { closeLists(0); out.push('<h' + h[1].length + '>' + mdInline(h[2]) + '</h' + h[1].length + '>'); i++; continue; }
    if (/^>\s?/.test(ln)) {
      const buf = [];
      while (i < L.length && /^>\s?/.test(L[i])) buf.push(L[i++].replace(/^>\s?/, ''));
      closeLists(0); out.push('<blockquote>' + mdInline(buf.join(' ')) + '</blockquote>'); continue;
    }
    if (/^\|.*\|/.test(ln) && i + 1 < L.length && /^\|[\s:|-]+\|/.test(L[i + 1])) {
      const cells = r => r.trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim());
      const head = cells(L[i]); i += 2; const body = [];
      while (i < L.length && /^\|.*\|/.test(L[i])) body.push(cells(L[i++]));
      closeLists(0);
      out.push('<table><thead><tr>' + head.map(c => '<th>' + mdInline(c) + '</th>').join('') + '</tr></thead><tbody>' +
        body.map(r => '<tr>' + r.map(c => '<td>' + mdInline(c) + '</td>').join('') + '</tr>').join('') + '</tbody></table>');
      continue;
    }
    const li = ln.match(/^(\s*)([-*+]|\d+\.)\s+(.*)$/);
    if (li) {
      const depth = Math.floor(li[1].replace(/\t/g, '  ').length / 2) + 1;
      const tag = /\d/.test(li[2]) ? 'ol' : 'ul';
      closeLists(depth);
      while (listStack.length < depth) {
        const nested = listStack.length > 0;
        if (nested) {                                  // 바로 앞 항목 안으로 넣습니다
          const last = out.length - 1;
          if (last >= 0 && /<\/li>$/.test(out[last])) out[last] = out[last].replace(/<\/li>$/, '');
        }
        out.push('<' + tag + '>'); listStack.push({ tag, nested });
      }
      out.push('<li>' + mdInline(li[3]) + '</li>'); i++; continue;
    }
    const buf = [];
    while (i < L.length && !/^\s*$/.test(L[i]) && !/^(#{1,6}\s|>|\||```|\s*([-*+]|\d+\.)\s)/.test(L[i])) buf.push(L[i++]);
    closeLists(0); out.push('<p>' + mdInline(buf.join(' ')) + '</p>');
  }
  closeLists(0);
  return out.join('\n');
}

/* ═══ 4. HTML → 슬라이드 ═════════════════════════════════════════ */
function runsOf(node, st) {
  st = st || {}; let out = [];
  node.childNodes.forEach(n => {
    if (n.nodeType === 3) { const t = n.textContent.replace(/\s+/g, ' '); if (t) out.push(Object.assign({ t }, st)); return; }
    if (n.nodeType !== 1) return;
    const g = n.tagName; const s2 = Object.assign({}, st);
    if (g === 'BR') { out.push(Object.assign({ t: '\n' }, st)); return; }
    if (g === 'STRONG' || g === 'B') s2.b = true;
    if (g === 'EM' || g === 'I') s2.i = true;
    if (g === 'CODE' || g === 'KBD' || g === 'SAMP' || g === 'TT') s2.c = true;
    if (g === 'A') s2.a = true;
    if (g === 'U' || g === 'INS') s2.u = true;
    if (g === 'S' || g === 'DEL' || g === 'STRIKE') s2.s = true;
    out = out.concat(runsOf(n, s2));
  });
  const m = [];
  out.forEach(r => {
    const p = m[m.length - 1];
    if (p && p.b === r.b && p.i === r.i && p.c === r.c && p.a === r.a && p.u === r.u && p.s === r.s) p.t += r.t;
    else m.push(Object.assign({}, r));
  });
  return m;
}
function trimRuns(rs) {
  const o = rs.slice();
  if (o.length) o[0] = Object.assign({}, o[0], { t: o[0].t.replace(/^\s+/, '') });
  if (o.length) o[o.length - 1] = Object.assign({}, o[o.length - 1], { t: o[o.length - 1].t.replace(/\s+$/, '') });
  return o.filter(r => r.t !== '');
}
const rtext = rs => (rs || []).map(r => r.t).join('');
const asRuns = t => String(t === undefined ? '' : t).split('\n').map(x => ({ t: x })).reduce((a, r, i) => (i ? a.concat([{ t: '\n' }, r]) : [r]), []);

function liItems(list, lvl, ord) {
  const out = []; let k = 0;
  [...list.children].forEach(ch => {
    if (ch.tagName === 'LI') {
      const cl = ch.cloneNode(true);
      $$('ul,ol', cl).forEach(n => n.remove());
      const rs = trimRuns(runsOf(cl));
      k++;
      if (rs.length) out.push({ runs: rs, lvl, num: ord ? k : 0 });
      [...ch.children].filter(c => c.tagName === 'UL' || c.tagName === 'OL')
        .forEach(sub => out.push(...liItems(sub, lvl + 1, sub.tagName === 'OL')));
    } else if (ch.tagName === 'UL' || ch.tagName === 'OL') {
      out.push(...liItems(ch, lvl + 1, ch.tagName === 'OL'));   // 앞 항목의 하위로 봅니다
    }
  });
  return out;
}

const CONTAINER = /^(DIV|MAIN|SECTION|ARTICLE|HEADER|FOOTER|NAV|DL|DETAILS|FIELDSET|CENTER|SPAN|LABEL)$/;

function absorb(el, S, o) {
  if (!el || el.nodeType !== 1) return;
  const tag = el.tagName;
  if (/^(SCRIPT|STYLE|NOSCRIPT|IFRAME|SVG|CANVAS|FORM|INPUT|BUTTON|SELECT|TEXTAREA|VIDEO|AUDIO)$/.test(tag)) return;
  if (o.notes && el.getAttribute && el.getAttribute('data-notes')) S.notes += (S.notes ? '\n' : '') + clean(el.getAttribute('data-notes'));
  if (el.matches && el.matches('aside.notes,.notes,.speaker-notes,.speakernotes,[data-role=notes]')) {
    if (o.notes) S.notes += (S.notes ? '\n' : '') + clean(el.textContent);
    return;
  }
  if (/^H[1-6]$/.test(tag)) {
    const t = clean(el.textContent); if (!t) return;
    if (!S.title) S.title = t;
    else if (!S.sub && !S.blocks.length && tag !== 'H1') S.sub = t;
    else S.blocks.push({ type: 'head', runs: trimRuns(runsOf(el)) });
    return;
  }
  if (tag === 'P' || tag === 'DD' || tag === 'DT' || tag === 'ADDRESS') {
    const imgs = $$('img', el);
    if (imgs.length && !clean(el.textContent)) { imgs.forEach(im => addImg(im, S, o)); return; }
    const rs = trimRuns(runsOf(el));
    if (rs.length) S.blocks.push({ type: 'text', runs: rs });
    imgs.forEach(im => addImg(im, S, o));
    return;
  }
  if (tag === 'UL' || tag === 'OL') {
    const items = liItems(el, 0, tag === 'OL');
    if (items.length) S.blocks.push({ type: 'list', ordered: tag === 'OL', items });
    return;
  }
  if (tag === 'PRE') { const t = el.textContent.replace(/\s+$/, ''); if (t) S.blocks.push({ type: 'code', text: t }); return; }
  if (tag === 'BLOCKQUOTE') {
    const rs = trimRuns(runsOf(el)); const cite = clean(($('cite', el) || {}).textContent || '');
    if (rs.length) S.blocks.push({ type: 'quote', runs: rs, cite });
    return;
  }
  if (tag === 'IMG') { addImg(el, S, o); return; }
  if (tag === 'FIGURE') {
    const im = $('img', el); if (im) addImg(im, S, o);
    const cap = clean(($('figcaption', el) || {}).textContent || '');
    if (cap) S.blocks.push({ type: 'cap', runs: [{ t: cap }] });
    const tb = $('table', el); if (tb) absorb(tb, S, o);
    return;
  }
  if (tag === 'TABLE') {
    if (!o.tables) { const t = clean(el.textContent); if (t) S.blocks.push({ type: 'text', runs: [{ t }] }); return; }
    const rows = [];
    $$('tr', el).forEach(tr => {
      const cs = [...tr.children].filter(c => /^(TD|TH)$/.test(c.tagName))
        .map(c => ({ t: clean(c.textContent), h: c.tagName === 'TH' }));
      if (cs.length) rows.push(cs);
    });
    if (rows.length) S.blocks.push({ type: 'table', rows, header: rows[0].some(c => c.h) || !!$('thead', el) });
    return;
  }
  if (tag === 'HR') return;
  if (CONTAINER.test(tag)) { [...el.childNodes].forEach(n => absorbNode(n, S, o)); return; }
  const t = clean(el.textContent);
  if (t) S.blocks.push({ type: 'text', runs: trimRuns(runsOf(el)) });
}
function absorbNode(n, S, o) {
  if (n.nodeType === 3) { const t = clean(n.textContent); if (t) S.blocks.push({ type: 'text', runs: [{ t }] }); return; }
  if (n.nodeType === 1) absorb(n, S, o);
}
function addImg(im, S, o) {
  if (!o.images) return;
  const src = im.getAttribute('src') || '';
  if (!src) return;
  const ar = (im.naturalWidth && im.naturalHeight) ? im.naturalWidth / im.naturalHeight : 0;
  if (ar) state.imgAR[src] = ar;
  S.blocks.push({ type: 'img', src, alt: im.getAttribute('alt') || '' });
}

function splitChunks(root, o) {
  const mode = o.split;
  if (mode === 'none') return [[...root.childNodes]];
  if (mode === 'section') {
    const secs = $$(':scope > section, :scope > article', root);
    if (secs.length) return secs.map(s => [...s.childNodes]);
  }
  let customs = [];
  if (mode === 'custom' && o.customSel.trim()) { try { customs = $$(o.customSel, root); } catch (e) { } }
  const isB = n => {
    const g = n.tagName;
    if (mode === 'custom') return customs.indexOf(n) >= 0;
    if (mode === 'marker') return false;
    if (mode === 'hr') return g === 'HR';
    if (mode === 'h1') return g === 'H1';
    if (mode === 'h2') return g === 'H1' || g === 'H2';
    if (mode === 'h3') return /^H[123]$/.test(g);
    if (mode === 'section') return g === 'SECTION' || g === 'ARTICLE' || g === 'H1' || g === 'H2' || g === 'HR';
    return g === 'H1' || g === 'H2' || g === 'HR' || g === 'SECTION' || g === 'ARTICLE';   // auto
  };
  const chunks = []; let cur = [];
  const flush = () => { if (cur.some(n => n.nodeType !== 3 || clean(n.textContent))) chunks.push(cur); cur = []; };
  [...root.childNodes].forEach(n => {
    if (n.nodeType === 8) { if (mode === 'marker' && /slide|슬라이드/i.test(n.textContent)) flush(); return; }
    if (n.nodeType === 3) { if (clean(n.textContent)) cur.push(n); return; }
    if (n.nodeType !== 1) return;
    if (isB(n)) { if (cur.length) flush(); if (n.tagName === 'HR') return; }
    cur.push(n);
  });
  flush();
  return chunks.length ? chunks : [[...root.childNodes]];
}

function parseSource(src, o) {
  let html = src;
  if (looksMD(src)) html = md2html(src);
  const doc = new DOMParser().parseFromString(html, 'text/html');
  $$('script,style,noscript,template', doc).forEach(n => n.remove());
  let root = doc.body;
  for (let i = 0; i < 4; i++) {
    const els = [...root.children];
    if (els.length === 1 && /^(DIV|MAIN|BODY)$/.test(els[0].tagName)) root = els[0]; else break;
  }
  const chunks = splitChunks(root, o);
  const slides = [];
  chunks.forEach(nodes => {
    const S = { id: uid(), title: '', sub: '', notes: '', blocks: [], on: true, layout: 'auto' };
    nodes.forEach(n => absorbNode(n, S, o));
    if (!S.title && !S.blocks.length && !S.notes) return;
    slides.push(S);
  });
  if (!slides.length) return [];
  // 표지
  if (o.firstTitle) {
    const f = slides[0];
    const light = f.blocks.filter(b => b.type !== 'cap').length <= 1;
    if (light) {
      f.layout = 'title';
      const p = f.blocks.find(b => b.type === 'text' || b.type === 'head');
      if (p && !f.sub) { f.sub = rtext(p.runs); f.blocks = f.blocks.filter(b => b !== p); }
    }
    else {
      slides.unshift({
        id: uid(), title: state.deck.title || f.title || '제목 없음',
        sub: state.deck.sub || '', notes: '', blocks: [], on: true, layout: 'title'
      });
    }
  }
  slides.forEach(s => { if (s.layout === 'auto') s.layout = autoLayout(s); });
  return slides;
}
function autoLayout(s) {
  const b = s.blocks.filter(x => x.type !== 'cap');
  if (!b.length && s.title) return 'section';
  if (b.length === 1 && b[0].type === 'img') return 'image';
  const imgs = b.filter(x => x.type === 'img').length;
  if (imgs === 1 && b.length >= 2 && b.length <= 5) return 'two';
  return 'content';
}

/* ═══ 5. 배치 엔진 (미리보기·PPTX 공용) ══════════════════════════ */
const GAP = 0.13;
function lineH(size) { return size * 1.32 / 72; }
function paraLines(runs, w, size, indent) {
  const txt = rtext(runs);
  const per = Math.max(6, (w - (indent || 0)) * 144 / size);
  let n = 0;
  txt.split('\n').forEach(seg => { n += Math.max(1, Math.ceil(units(seg) / per)); });
  return Math.max(1, n);
}
function blockHeight(b, w, S) {
  switch (b.type) {
    case 'head': return paraLines(b.runs, w, S.body + 3) * lineH(S.body + 3) + 0.06;
    case 'text': return paraLines(b.runs, w, S.body) * lineH(S.body);
    case 'cap': return paraLines(b.runs, w, S.body - 3) * lineH(S.body - 3);
    case 'list': return b.items.reduce((a, it) => a + paraLines(it.runs, w, S.body - (it.lvl ? 1 : 0), 0.26 + it.lvl * 0.26) * lineH(S.body) + 0.045, 0);
    case 'quote': return paraLines(b.runs, w - 0.5, S.body) * lineH(S.body) + 0.3 + (b.cite ? lineH(S.body - 3) : 0);
    case 'code': return b.text.split('\n').length * lineH(S.body - 3) + 0.22;
    case 'table': return b.rows.length * (S.body * 1.9 / 72 + 0.09) + 0.05;
    case 'img': {
      const ar = state.imgAR[b.src] || 1.6;
      return clamp(w * 0.62 / ar, 0.9, S.imgMax);
    }
    default: return 0.3;
  }
}
function blockShapes(b, x, y, w, S, T) {
  const out = [];
  const F = S.font;
  switch (b.type) {
    case 'head':
      out.push({ k: 'text', x, y, w, h: blockHeight(b, w, S), font: F, align: 'left', valign: 'top', size: S.body + 3, color: T.title, bold: true, paras: [{ runs: b.runs, size: S.body + 3, bold: true, color: T.title }] });
      break;
    case 'text':
      out.push({ k: 'text', x, y, w, h: blockHeight(b, w, S), font: F, align: 'left', valign: 'top', size: S.body, color: T.fg, paras: [{ runs: b.runs, size: S.body, color: T.fg }] });
      break;
    case 'cap':
      out.push({ k: 'text', x, y, w, h: blockHeight(b, w, S), font: F, align: 'center', valign: 'top', size: S.body - 3, color: T.muted, paras: [{ runs: b.runs, size: S.body - 3, color: T.muted, italic: true }] });
      break;
    case 'list':
      out.push({
        k: 'text', x, y, w, h: blockHeight(b, w, S), font: F, align: 'left', valign: 'top', size: S.body, color: T.fg,
        paras: b.items.map(it => ({
          runs: it.runs, size: S.body - (it.lvl ? 1 : 0), color: T.fg, lvl: it.lvl,
          bullet: true, num: it.num, marker: it.num ? it.num + '.' : (it.lvl ? '–' : '•'), markerColor: T.accent
        }))
      });
      break;
    case 'quote': {
      const h = blockHeight(b, w, S);
      out.push({ k: 'rect', x, y, w, h, fill: T.panel });
      out.push({ k: 'rect', x, y, w: 0.055, h, fill: T.accent });
      out.push({ k: 'text', x: x + 0.24, y: y + 0.13, w: w - 0.42, h: h - 0.26, font: F, align: 'left', valign: 'top', size: S.body, color: T.fg, paras: [{ runs: b.runs, size: S.body, color: T.fg, italic: true }].concat(b.cite ? [{ runs: [{ t: '— ' + b.cite }], size: S.body - 3, color: T.muted }] : []) });
      break;
    }
    case 'code': {
      const h = blockHeight(b, w, S);
      out.push({ k: 'rect', x, y, w, h, fill: T.panel });
      out.push({ k: 'text', x: x + 0.16, y: y + 0.1, w: w - 0.32, h: h - 0.2, font: 'Consolas', align: 'left', valign: 'top', size: S.body - 3, color: T.fg, mono: true, paras: b.text.split('\n').map(l => ({ runs: [{ t: l || ' ' }], size: S.body - 3, color: T.fg, mono: true })) });
      break;
    }
    case 'table': {
      const h = blockHeight(b, w, S);
      out.push({ k: 'table', x, y, w, h, rows: b.rows, header: b.header, size: Math.max(9, S.body - 3), font: F, color: T.fg, line: mix(T.fg, T.bg, 0.62), head: T.accent, headText: onAccent(T.accent), zebra: mix(T.bg, T.fg, 0.05) });
      break;
    }
    case 'img': {
      const h = blockHeight(b, w, S);
      const ar = state.imgAR[b.src] || 1.6;
      const iw = Math.min(w, h * ar);
      out.push({ k: 'img', src: b.src, x: x + (w - iw) / 2, y, w: iw, h, alt: b.alt });
      break;
    }
  }
  return out;
}

/* 한 장 구성 → {shapes, rest, over} */
function compose(sl, idx, total) {
  const o = state.opts, T = state.theme, M = state.master;
  const D = DIM[o.aspect], W = D.W, H = D.H;
  const sh = [{ k: 'rect', x: 0, y: 0, w: W, h: H, fill: T.bg }];
  const mx = 0.62;
  const isTitle = sl.layout === 'title', isSec = sl.layout === 'section';
  const footH = (M.showFoot && M.foot) || M.page ? 0.46 : 0.18;

  /* 표지 */
  if (isTitle) {
    sh.push({ k: 'rect', x: 0, y: H - 0.13, w: W, h: 0.13, fill: T.accent });
    sh.push({ k: 'rect', x: mx, y: H * 0.30, w: 1.5, h: 0.05, fill: T.accent });
    sh.push({ k: 'text', x: mx, y: H * 0.34, w: W - mx * 2, h: 1.5, font: o.font, size: o.titleSize + 8, color: T.title, bold: true, align: 'left', valign: 'top', paras: [{ runs: asRuns(state.deck.title || sl.title || ''), size: o.titleSize + 8, bold: true, color: T.title }] });
    const sub = state.deck.sub || sl.sub || '';
    if (sub) sh.push({ k: 'text', x: mx, y: H * 0.34 + 1.5, w: W - mx * 2, h: 0.6, font: o.font, size: o.bodySize + 2, color: T.fg, align: 'left', valign: 'top', paras: [{ runs: asRuns(sub), size: o.bodySize + 2, color: T.fg }] });
    if (state.deck.by) sh.push({ k: 'text', x: mx, y: H - 1.1, w: W - mx * 2, h: 0.4, font: o.font, size: o.bodySize - 1, color: T.muted, align: 'left', valign: 'top', paras: [{ runs: asRuns(state.deck.by), size: o.bodySize - 1, color: T.muted }] });
    if (M.logo) sh.push({ k: 'img', src: M.logo, x: W - mx - 1.0, y: 0.45, w: 1.0, h: 1.0 / (M.logoAR || 2.5), logo: true });
    return { shapes: sh, rest: [], over: false };
  }

  /* 제목 장식 */
  let top = 1.30, tx = mx, tw = W - mx * 2, tcolor = T.title;
  const hs = sl.layout === 'blank' ? (o.headStyle === 'side' ? 'side' : 'plain') : o.headStyle;
  if (hs === 'bar') { sh.push({ k: 'rect', x: 0, y: 0, w: W, h: 0.14, fill: T.accent }); top = 1.34; }
  if (hs === 'side') { sh.push({ k: 'rect', x: 0, y: 0, w: 0.17, h: H, fill: T.accent }); tx = mx + 0.14; tw = W - tx - mx; top = 1.24; }
  if (hs === 'block') { sh.push({ k: 'rect', x: 0, y: 0, w: W, h: 1.16, fill: T.accent }); tcolor = onAccent(T.accent); top = 1.46; }
  if (hs === 'plain') top = 1.18;

  if (isSec) {
    sh.push({ k: 'rect', x: tx, y: H / 2 - 0.62, w: 0.07, h: 1.24, fill: T.accent });
    sh.push({ k: 'text', x: tx + 0.28, y: H / 2 - 0.62, w: tw - 0.28, h: 1.24, font: o.font, size: o.titleSize + 2, color: T.title, bold: true, align: 'left', valign: 'middle', paras: [{ runs: asRuns(sl.title), size: o.titleSize + 2, bold: true, color: T.title }] });
    if (sl.sub) sh.push({ k: 'text', x: tx + 0.28, y: H / 2 + 0.66, w: tw - 0.28, h: 0.5, font: o.font, size: o.bodySize, color: T.muted, align: 'left', valign: 'top', paras: [{ runs: asRuns(sl.sub), size: o.bodySize, color: T.muted }] });
    addFoot(sh, W, H, mx, idx, total);
    return { shapes: sh, rest: [], over: false };
  }

  if (sl.layout === 'blank') {
    top = 0.5;
  } else if (sl.title) {
    const tS = o.titleSize - (units(sl.title) > 44 ? 6 : units(sl.title) > 30 ? 3 : 0);
    sh.push({ k: 'text', x: tx, y: hs === 'block' ? 0.30 : 0.40, w: tw - (M.logo ? 1.1 : 0), h: 0.72, font: o.font, size: tS, color: tcolor, bold: true, align: 'left', valign: 'middle', paras: [{ runs: asRuns(sl.title), size: tS, bold: true, color: tcolor }] });
    if (hs === 'rule') sh.push({ k: 'rect', x: tx, y: 1.12, w: tw, h: 0.022, fill: T.accent });
  } else top = 0.55;
  if (sl.sub) {
    sh.push({ k: 'text', x: tx, y: top - 0.06, w: tw, h: 0.34, font: o.font, size: o.bodySize, color: T.muted, align: 'left', valign: 'top', paras: [{ runs: asRuns(sl.sub), size: o.bodySize, color: T.muted }] });
    top += 0.36;
  }
  if (M.logo) sh.push({ k: 'img', src: M.logo, x: W - mx - 0.8, y: 0.3, w: 0.8, h: 0.8 / (M.logoAR || 2.5), logo: true });

  /* 본문 흐름 */
  const bottom = H - footH;
  const availH = bottom - top;
  const blocks = sl.blocks.slice();
  const scales = state.opts.shrink ? [1, .94, .88, .82, .76, .70, .64] : [1];
  let best = null;
  for (const sc of scales) {
    const S = { body: Math.max(9, Math.round(o.bodySize * sc)), font: o.font, imgMax: availH * 0.94 };
    const r = flow(blocks, tx, top, tw, availH, S, T, sl.layout);
    if (!r.rest.length) { best = r; break; }
    best = r;
  }
  sh.push(...best.shapes);
  addFoot(sh, W, H, mx, idx, total);
  return { shapes: sh, rest: best.rest, over: best.rest.length > 0 };
}

function flow(blocks, x0, y0, w0, availH, S, T, layout) {
  const out = []; const rest = []; const GAP = S.gap || 0.13;
  let x = x0, w = w0, y = y0;
  /* 두 단: 그림은 오른쪽 고정 */
  if (layout === 'two') {
    const im = blocks.find(b => b.type === 'img');
    if (im) {
      const rw = w0 * 0.42, ar = state.imgAR[im.src] || 1.4;
      const rh = Math.min(availH * 0.8, rw / ar);
      out.push({ k: 'img', src: im.src, x: x0 + w0 - rw, y: y0, w: rw, h: rh, alt: im.alt });
      w = w0 * 0.54;
      blocks = blocks.filter(b => b !== im);
    }
  }
  if (layout === 'image') {
    const im = blocks.find(b => b.type === 'img');
    if (im) {
      const ar = state.imgAR[im.src] || 1.6;
      let ih = availH - 0.1, iw = ih * ar;
      if (iw > w0) { iw = w0; ih = iw / ar; }
      out.push({ k: 'img', src: im.src, x: x0 + (w0 - iw) / 2, y: y0 + (availH - ih) / 2, w: iw, h: ih, alt: im.alt });
      blocks = blocks.filter(b => b !== im);
      y = y0 + (availH + ih) / 2 + 0.05;
    }
  }
  blocks.forEach(b => {
    if (rest.length) { rest.push(b); return; }
    const h = blockHeight(b, w, S);
    if (y + h > y0 + availH + 0.02) {
      if (out.filter(s => s.k !== 'img').length === 0 && !rest.length) {
        out.push(...blockShapes(b, x, y, w, S, T));   // 첫 덩어리는 잘리더라도 배치
        y += h + GAP; return;
      }
      rest.push(b); return;
    }
    out.push(...blockShapes(b, x, y, w, S, T));
    y += h + GAP;
  });
  return { shapes: out, rest };
}

function addFoot(sh, W, H, mx, idx, total) {
  const T = state.theme, M = state.master, o = state.opts;
  if (!(M.showFoot && M.foot) && !M.page) return;
  sh.push({ k: 'rect', x: mx, y: H - 0.42, w: W - mx * 2, h: 0.012, fill: mix(T.bg, T.fg, 0.22) });
  if (M.showFoot && M.foot)
    sh.push({ k: 'text', x: mx, y: H - 0.38, w: W - mx * 2 - 1, h: 0.3, font: o.font, size: 10, color: T.muted, align: 'left', valign: 'middle', paras: [{ runs: [{ t: M.foot }], size: 10, color: T.muted }] });
  if (M.page)
    sh.push({ k: 'text', x: W - mx - 1, y: H - 0.38, w: 1, h: 0.3, font: o.font, size: 10, color: T.muted, align: 'right', valign: 'middle', paras: [{ runs: [{ t: (idx + 1) + ' / ' + total }], size: 10, color: T.muted, align: 'right' }] });
}

/* 넘치는 슬라이드를 이어지는 장으로 나눔 */
function expandDeck() {
  const list = state.slides.filter(s => s.on);
  if (!state.opts.splitLong) return list;
  const out = [];
  list.forEach(s => {
    let cur = s, guard = 0;
    while (guard++ < 30) {
      const r = compose(cur, out.length, list.length);
      out.push(cur);
      if (!r.rest.length) break;
      const usedIds = r.rest;
      cur = { id: uid(), title: (s.title ? s.title + ' (계속)' : ''), sub: '', notes: '', blocks: usedIds, on: true, layout: 'content', cont: true };
    }
  });
  return out;
}

/* ═══ 5b. 포스터 조판 (A0·A1·A2) ════════════════════════════════
   슬라이드 한 장 = 포스터의 한 구역. 구역을 단 안에서 흘려 채웁니다. */
function composePoster(list) {
  const D = DIM[state.opts.aspect], W = D.W, H = D.H;
  const T = state.theme, o = state.opts, M = state.master, PO = state.poster;
  const u = Math.min(W, H) / 20;             // 판형 기본 단위 (A0 ≈ 1.66인치)
  const k = u * (PO.scale || 1);             // 글자 배율 (A0 본문 ≈ 28pt)
  const mx = 0.9 * u, gut = 0.55 * u;
  const cols = clamp(+PO.cols || 3, 1, 5);
  const colW = (W - 2 * mx - (cols - 1) * gut) / cols;
  const footH = ((M.showFoot && M.foot) || M.page) ? 0.62 * k : 0.3 * u;
  const bottom = H - footH;
  const secS = (o.bodySize + 5) * k, secPad = 0.16 * k;
  const S = { body: o.bodySize * k, font: o.font, gap: 0.16 * k, imgMax: (H - 2 * u) * 0.42 };
  const title = state.deck.title || (list[0] && list[0].title) || '';
  const sub = state.deck.sub, by = state.deck.by;
  const logoW = M.logo ? 1.5 * u : 0;
  const pages = [];
  let pg = null, ci = 0, y = 0, overflow = 0;

  const banner = () => {
    const sh = [{ k: 'rect', x: 0, y: 0, w: W, h: H, fill: T.bg }];
    if (!PO.banner) return { shapes: sh, top: 0.5 * u };
    const tS = clamp(o.titleSize * k * 2, 24, 130);
    const tw = W - 2 * mx - (logoW ? logoW + 0.3 * u : 0);
    const tl = paraLines(asRuns(title), tw, tS);
    const subS = (o.bodySize + 4) * k, byS = (o.bodySize + 1) * k;
    const hh = 0.42 * u + tl * lineH(tS) + (sub ? lineH(subS) + 0.1 * u : 0) + (by ? lineH(byS) + 0.06 * u : 0) + 0.42 * u;
    const fg = onAccent(T.accent);
    sh.push({ k: 'rect', x: 0, y: 0, w: W, h: hh, fill: T.accent });
    let ty = 0.42 * u;
    sh.push({ k: 'text', x: mx, y: ty, w: tw, h: tl * lineH(tS), font: o.font, size: tS, color: fg, bold: true, align: 'left', valign: 'top', paras: [{ runs: asRuns(title), size: tS, bold: true, color: fg }] });
    ty += tl * lineH(tS) + 0.1 * u;
    if (sub) { sh.push({ k: 'text', x: mx, y: ty, w: tw, h: lineH(subS), font: o.font, size: subS, color: fg, align: 'left', valign: 'top', paras: [{ runs: asRuns(sub), size: subS, color: fg }] }); ty += lineH(subS) + 0.06 * u; }
    if (by) sh.push({ k: 'text', x: mx, y: ty, w: tw, h: lineH(byS), font: o.font, size: byS, color: fg, align: 'left', valign: 'top', paras: [{ runs: asRuns(by), size: byS, color: fg }] });
    if (M.logo) sh.push({ k: 'img', src: M.logo, x: W - mx - logoW, y: 0.42 * u, w: logoW, h: logoW / (M.logoAR || 2.5), logo: true });
    return { shapes: sh, top: hh + 0.4 * u };
  };
  const contHead = () => {
    const sh = [{ k: 'rect', x: 0, y: 0, w: W, h: H, fill: T.bg }, { k: 'rect', x: 0, y: 0, w: W, h: 0.14 * u, fill: T.accent }];
    const s = (o.bodySize + 2) * k;
    sh.push({ k: 'text', x: mx, y: 0.3 * u, w: W - 2 * mx, h: lineH(s), font: o.font, size: s, color: T.muted, align: 'left', valign: 'top', paras: [{ runs: asRuns(title + ' (계속)'), size: s, color: T.muted }] });
    return { shapes: sh, top: 0.3 * u + lineH(s) + 0.3 * u };
  };
  const start = () => {
    pg = pages.length ? contHead() : banner();
    pg.secs = []; pg.notes = ''; pg.used = []; pages.push(pg); ci = 0; y = pg.top;
  };
  const mark = () => { pg.used[ci] = y; };
  const colX = () => mx + ci * (colW + gut);
  const nextCol = () => { ci++; if (ci >= cols) start(); else y = pg.top; };

  const secShapes = (t, x, yy) => {
    const lines = paraLines(asRuns(t), colW - secPad * 2, secS);
    const hh = lines * lineH(secS) + secPad * 1.6;
    const out = [];
    if (PO.sec === 'bar') {
      out.push({ k: 'rect', x, y: yy, w: colW, h: hh, fill: T.accent });
      out.push({ k: 'text', x: x + secPad, y: yy + secPad * .8, w: colW - secPad * 2, h: hh - secPad * 1.6, font: o.font, size: secS, color: onAccent(T.accent), bold: true, align: 'left', valign: 'top', paras: [{ runs: asRuns(t), size: secS, bold: true, color: onAccent(T.accent) }] });
    } else if (PO.sec === 'box') {
      out.push({ k: 'rect', x, y: yy, w: colW, h: hh, fill: T.panel });
      out.push({ k: 'rect', x, y: yy, w: 0.07 * u, h: hh, fill: T.accent });
      out.push({ k: 'text', x: x + secPad + 0.07 * u, y: yy + secPad * .8, w: colW - secPad * 2 - 0.07 * u, h: hh - secPad * 1.6, font: o.font, size: secS, color: T.title, bold: true, align: 'left', valign: 'top', paras: [{ runs: asRuns(t), size: secS, bold: true, color: T.title }] });
    } else {
      out.push({ k: 'text', x, y: yy, w: colW, h: hh - secPad * .6, font: o.font, size: secS, color: T.accent, bold: true, align: 'left', valign: 'top', paras: [{ runs: asRuns(t), size: secS, bold: true, color: T.accent }] });
      out.push({ k: 'rect', x, y: yy + hh - secPad * .5, w: colW, h: 0.025 * u, fill: T.accent });
    }
    return { out, hh };
  };

  start();
  list.forEach((sl, si) => {
    if (!PO.merge && si > 0) start();
    const items = [];
    if (sl.title) items.push({ sec: sl.title });
    if (sl.sub) items.push({ b: { type: 'text', runs: asRuns(sl.sub) } });
    sl.blocks.forEach(b => items.push({ b }));
    items.forEach((it, ii) => {
      if (it.sec) {
        const probe = secShapes(it.sec, 0, 0);
        const nb = items[ii + 1];
        const need = probe.hh + (nb && nb.b ? Math.min(blockHeight(nb.b, colW, S), 1.1 * u) : 0);
        if (y + need > bottom && !(ci === 0 && y === pg.top)) nextCol();
        const r = secShapes(it.sec, colX(), y);
        pg.shapes.push(...r.out);
        pg.secs.push(sl);
        y += r.hh + S.gap; mark();
      } else {
        const h = blockHeight(it.b, colW, S);
        if (y + h > bottom + 0.01 && y > pg.top + 0.01) nextCol();
        if (y + h > bottom + 0.01) overflow++;
        pg.shapes.push(...blockShapes(it.b, colX(), y, colW, S, T));
        y += h + S.gap; mark();
      }
    });
    if (sl.notes) pg.notes += (pg.notes ? '\n' : '') + sl.notes;
  });

  const fS = (o.bodySize - 3) * k;
  pages.forEach((p, i) => {
    p.title = title + (pages.length > 1 ? ' · ' + (i + 1) : '');
    p.over = overflow > 0;
    const span = bottom - p.top;
    let used = 0;
    for (let c = 0; c < cols; c++) used += Math.max(0, (p.used[c] || p.top) - p.top);
    p.fill = span > 0 ? used / (cols * span) : 1;
    if (footH <= 0.3 * u) return;
    p.shapes.push({ k: 'rect', x: mx, y: H - footH + 0.1 * u, w: W - 2 * mx, h: 0.02 * u, fill: mix(T.bg, T.fg, 0.25) });
    if (M.showFoot && M.foot)
      p.shapes.push({ k: 'text', x: mx, y: H - footH + 0.2 * u, w: W - 2 * mx - 2 * u, h: lineH(fS), font: o.font, size: fS, color: T.muted, align: 'left', valign: 'top', paras: [{ runs: [{ t: M.foot }], size: fS, color: T.muted }] });
    if (M.page && pages.length > 1)
      p.shapes.push({ k: 'text', x: W - mx - 2 * u, y: H - footH + 0.2 * u, w: 2 * u, h: lineH(fS), font: o.font, size: fS, color: T.muted, align: 'right', valign: 'top', paras: [{ runs: [{ t: (i + 1) + ' / ' + pages.length }], size: fS, color: T.muted, align: 'right' }] });
  });
  return pages;
}

/* 판형에 맞는 최종 페이지 목록 */
function buildPages() {
  const list = state.slides.filter(s => s.on);
  if (!list.length) return [];
  if (isPoster()) return composePoster(list);
  const deck = expandDeck();
  return deck.map((sl, i) => {
    const r = compose(sl, i, deck.length);
    return { shapes: r.shapes, notes: sl.notes, title: sl.title, over: r.over, secs: [sl] };
  });
}

/* ═══ 6. 미리보기 그리기 ═════════════════════════════════════════ */
function drawShapes(host, shapes, D, px) {
  host.innerHTML = '';
  const I = v => (v * px) + 'px';
  shapes.forEach(s => {
    if (s.k === 'rect') {
      const d = document.createElement('div');
      d.className = 'sh';
      d.style.cssText = 'left:' + I(s.x) + ';top:' + I(s.y) + ';width:' + I(s.w) + ';height:' + I(s.h) + ';background:#' + s.fill;
      host.appendChild(d);
    } else if (s.k === 'text') {
      const d = document.createElement('div');
      d.className = 'sh tx';
      d.style.cssText = 'left:' + I(s.x) + ';top:' + I(s.y) + ';width:' + I(s.w) + ';height:' + I(s.h) +
        ';display:flex;flex-direction:column;justify-content:' + (s.valign === 'middle' ? 'center' : 'flex-start') +
        ';overflow:hidden;font-family:' + (s.mono ? 'var(--mono)' : '"' + s.font + '",sans-serif') +
        ';text-align:' + (s.align || 'left');
      s.paras.forEach(p => {
        const el = document.createElement('div');
        const fs = p.size * px / 72;
        el.style.cssText = 'font-size:' + fs.toFixed(2) + 'px;line-height:1.32;color:#' + (p.color || s.color) +
          ';font-weight:' + (p.bold ? 700 : 400) + ';font-style:' + (p.italic ? 'italic' : 'normal') +
          ';padding-left:' + I((p.lvl || 0) * 0.26 + (p.bullet ? 0.26 : 0)) + ';text-indent:' + (p.bullet ? '-' + I(0.26) : '0') +
          ';margin-bottom:' + (p.bullet ? (0.045 * px).toFixed(1) + 'px' : '0') + ';white-space:pre-wrap';
        if (p.bullet) {
          const m = document.createElement('span');
          m.textContent = p.marker + ' ';
          m.style.color = '#' + (p.markerColor || p.color);
          el.appendChild(m);
        }
        p.runs.forEach(r => {
          const sp = document.createElement('span');
          sp.textContent = r.t;
          sp.style.cssText = (r.b ? 'font-weight:700;' : '') + (r.i ? 'font-style:italic;' : '') +
            (r.c ? 'font-family:var(--mono);font-size:.92em;' : '') + (r.u ? 'text-decoration:underline;' : '') +
            (r.s ? 'text-decoration:line-through;' : '');
          el.appendChild(sp);
        });
        d.appendChild(el);
      });
      host.appendChild(d);
    } else if (s.k === 'img') {
      const im = document.createElement('img');
      im.className = 'sh';
      im.src = s.src; im.alt = s.alt || '';
      im.style.cssText = 'left:' + I(s.x) + ';top:' + I(s.y) + ';width:' + I(s.w) + ';height:' + I(s.h) + ';object-fit:contain';
      im.onerror = () => { im.style.background = '#0002'; im.alt = '그림 없음'; };
      im.onload = () => {
        if (im.naturalWidth && !state.imgAR[s.src]) { state.imgAR[s.src] = im.naturalWidth / im.naturalHeight; }
      };
      host.appendChild(im);
    } else if (s.k === 'table') {
      const d = document.createElement('div');
      d.className = 'sh';
      d.style.cssText = 'left:' + I(s.x) + ';top:' + I(s.y) + ';width:' + I(s.w) + ';height:' + I(s.h) + ';overflow:hidden';
      const t = document.createElement('table');
      t.style.cssText = 'width:100%;border-collapse:collapse;font-size:' + (s.size * px / 72).toFixed(2) + 'px;font-family:"' + s.font + '",sans-serif';
      s.rows.forEach((row, ri) => {
        const tr = document.createElement('tr');
        row.forEach(c => {
          const td = document.createElement('td');
          td.textContent = c.t;
          const head = s.header && ri === 0;
          td.style.cssText = 'border:1px solid #' + s.line + ';padding:' + (0.05 * px).toFixed(1) + 'px ' + (0.08 * px).toFixed(1) + 'px;' +
            (head ? 'background:#' + s.head + ';color:#' + s.headText + ';font-weight:700;' : 'color:#' + s.color + ';' + (ri % 2 ? 'background:#' + s.zebra + ';' : ''));
          tr.appendChild(td);
        });
        t.appendChild(tr);
      });
      d.appendChild(t);
      host.appendChild(d);
    }
  });
}

function renderCanvas() {
  const pages = buildPages();
  const D = DIM[state.opts.aspect];
  const cv = $('#canvas'), wrap = $('#cwrap');
  $('#emptyMsg').hidden = pages.length > 0;
  cv.hidden = pages.length === 0;
  $('#statN').textContent = pages.length;
  $('#statAR').textContent = D.name;
  $('#statImg').textContent = state.slides.reduce((a, s) => a + s.blocks.filter(b => b.type === 'img').length, 0);
  setPrintSize(D);
  if (!pages.length) { $('#stageT').textContent = isPoster() ? '포스터 없음' : '슬라이드 없음'; return; }
  state.sel = clamp(state.sel, 0, state.slides.length - 1);
  let i;
  if (isPoster()) {
    const home = pages.findIndex(p => p.secs.indexOf(state.slides[state.sel]) >= 0);
    if (home >= 0) state.page = home;
    i = clamp(state.page, 0, pages.length - 1);
  } else {
    const home = pages.findIndex(p => p.secs[0] === state.slides[state.sel]);
    i = home >= 0 ? home : clamp(state.page, 0, pages.length - 1);
  }
  state.page = i;
  const r = pages[i];
  const availW = Math.max(320, wrap.clientWidth - 36);
  const availH = Math.max(200, wrap.clientHeight - 36);
  const px = Math.min(availW / D.W, availH / D.H);
  cv.style.width = (D.W * px) + 'px';
  cv.style.height = (D.H * px) + 'px';
  drawShapes(cv, r.shapes, D, px);
  $('#stageT').textContent = (i + 1) + ' / ' + pages.length + ' · ' + (r.title || '제목 없음') +
    (r.over ? ' · 내용이 넘칩니다' : '') +
    (isPoster() && r.fill !== undefined && r.fill < 0.55 ? ' · 여백이 많습니다 (배율을 올리거나 단 수를 줄여 보세요)' : '');
}
/* 내용이 한 장을 꽉 채우도록 배율을 찾습니다 */
function posterAutoFit() {
  if (!isPoster()) { say('포스터 판형에서만 쓸 수 있습니다.', 'warn'); return; }
  const keepC = state.poster.cols, keepS = state.poster.scale;
  const cand = [];
  for (let c = 2; c <= 4; c++) {          // 한 단짜리 포스터는 줄이 너무 길어 제외합니다
    for (let v = 0.6; v <= 1.601; v += 0.1) {
      state.poster.cols = c;
      state.poster.scale = Math.round(v * 100) / 100;
      const pg = buildPages();
      if (pg.length !== 1 || pg[0].over) continue;
      const f = pg[0].fill;
      cand.push({ c, v: state.poster.scale, f });
    }
  }
  let best = null;
  cand.forEach(x => { if (!best || x.f > best.f) best = x; });
  /* 채움률이 비슷하면 단이 많은 쪽이 포스터답습니다 */
  if (best) cand.forEach(x => { if (x.f >= best.f - 0.08 && x.c > best.c) best = x; });
  if (best) {
    state.poster.cols = best.c; state.poster.scale = best.v;
    syncControls(); refresh();
    say(best.c + '단 · 배율 ' + best.v.toFixed(2) + '배로 맞췄습니다 (채움 ' + Math.round(best.f * 100) + '%)', 'ok');
  } else {
    state.poster.cols = keepC; state.poster.scale = keepS;
    syncControls(); refresh();
    say('한 장에 담기지 않습니다. 내용을 줄이거나 더 큰 판형을 쓰세요.', 'warn');
  }
}

/* 포스터에서 쪽을 옮기면 그 쪽의 첫 구역을 선택합니다 */
function jumpToPage() {
  const pages = buildPages();
  const p = pages[clamp(state.page, 0, pages.length - 1)];
  if (p && p.secs && p.secs.length) {
    const i = state.slides.indexOf(p.secs[0]);
    if (i >= 0) state.sel = i;
  }
}
function setPrintSize(D) {
  const s = $('#pageStyle');
  if (s) s.textContent = '@page{size:' + D.W.toFixed(2) + 'in ' + D.H.toFixed(2) + 'in;margin:0}';
}

function renderStrip() {
  const strip = $('#strip');
  strip.innerHTML = '';
  const D = DIM[state.opts.aspect];
  const tw = 166, px = tw / D.W;
  state.slides.forEach((s, i) => {
    const c = document.createElement('div');
    c.className = 'card' + (s.on ? '' : ' off');
    c.setAttribute('aria-current', i === state.sel);
    c.draggable = true;
    c.dataset.i = i;
    const th = document.createElement('div');
    th.className = 'thumb';
    if (isPoster()) {
      th.style.cssText = 'height:62px;background:var(--ink3);padding:7px 8px;overflow:hidden';
      th.innerHTML = '<div style="font-size:11px;color:var(--brass);font-family:var(--mono)">구역 ' + (i + 1) + '</div>' +
        '<div style="font-size:12px;color:var(--fg);margin-top:2px;line-height:1.35;max-height:32px;overflow:hidden">' +
        esc(s.title || '제목 없음') + '</div>';
    } else {
      th.style.height = (D.H * px) + 'px';
      drawShapes(th, compose(s, i, state.slides.length).shapes, D, px);
    }
    const cap = document.createElement('div');
    cap.className = 'cap';
    cap.innerHTML = '<b>' + (i + 1) + '</b><span>' + esc(s.title || (rtext((s.blocks[0] || {}).runs) || '제목 없음').slice(0, 24)) + '</span>';
    const ops = document.createElement('div');
    ops.className = 'ops';
    ops.innerHTML = '<button title="켜기/끄기" data-a="on">' + (s.on ? '◉' : '○') + '</button>' +
      '<button title="복제" data-a="dup">⧉</button><button title="삭제" data-a="del">✕</button>';
    ops.onclick = e => {
      const a = e.target.dataset.a; if (!a) return;
      e.stopPropagation();
      if (a === 'on') s.on = !s.on;
      if (a === 'dup') state.slides.splice(i + 1, 0, JSON.parse(JSON.stringify(Object.assign({}, s, { id: uid() }))));
      if (a === 'del') state.slides.splice(i, 1);
      state.sel = clamp(state.sel, 0, state.slides.length - 1);
      refresh();
    };
    c.appendChild(th); c.appendChild(cap); c.appendChild(ops);
    c.onclick = () => { state.sel = i; refresh(); };
    c.ondragstart = e => { e.dataTransfer.setData('text/plain', i); c.classList.add('drag'); };
    c.ondragend = () => c.classList.remove('drag');
    c.ondragover = e => e.preventDefault();
    c.ondrop = e => {
      e.preventDefault();
      const from = +e.dataTransfer.getData('text/plain');
      if (isNaN(from) || from === i) return;
      const [m] = state.slides.splice(from, 1);
      state.slides.splice(i, 0, m);
      state.sel = i; refresh();
    };
    strip.appendChild(c);
  });
  const add = document.createElement('button');
  add.className = 'addcard'; add.textContent = '+'; add.title = '빈 슬라이드 추가';
  add.onclick = () => {
    state.slides.splice(state.sel + 1, 0, { id: uid(), title: '새 슬라이드', sub: '', notes: '', blocks: [], on: true, layout: 'content' });
    state.sel += 1; refresh();
  };
  strip.appendChild(add);
  const cur = strip.children[state.sel];
  if (cur && cur.scrollIntoView) cur.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

/* ═══ 7. 슬라이드 편집기 ═════════════════════════════════════════ */
const BLKNAME = { head: '소제목', text: '문단', list: '글머리 목록', quote: '인용', code: '코드', table: '표', img: '그림', cap: '그림 설명' };
function renderEditor() {
  const host = $('#editor');
  const s = state.slides[state.sel];
  if (!s) { host.innerHTML = '<p class="mini">슬라이드를 고르면 여기서 고칠 수 있습니다.</p>'; return; }
  host.innerHTML = '';
  const mk = (html) => { const d = document.createElement('div'); d.innerHTML = html; return d.firstElementChild; };

  const lay = mk('<div class="field"><label class="lab">배치</label><select>' +
    [['title', '표지'], ['section', '간지 (제목만)'], ['content', '제목 + 내용'], ['two', '내용 + 그림 (좌우)'], ['image', '그림 크게'], ['blank', '제목 없이 내용만']]
      .map(([v, n]) => '<option value="' + v + '"' + (s.layout === v ? ' selected' : '') + '>' + n + '</option>').join('') + '</select></div>');
  $('select', lay).onchange = e => { s.layout = e.target.value; refresh(); };
  host.appendChild(lay);

  const ti = mk('<div class="field"><label class="lab">제목</label><input type="text"></div>');
  $('input', ti).value = s.title;
  $('input', ti).oninput = e => { s.title = e.target.value; softRefresh(); };
  host.appendChild(ti);

  const su = mk('<div class="field"><label class="lab">소제목</label><input type="text"></div>');
  $('input', su).value = s.sub || '';
  $('input', su).oninput = e => { s.sub = e.target.value; softRefresh(); };
  host.appendChild(su);

  const bl = document.createElement('div');
  bl.innerHTML = '<label class="lab" style="margin-top:6px">내용</label>';
  s.blocks.forEach((b, bi) => bl.appendChild(blockEditor(s, b, bi)));
  host.appendChild(bl);

  const addRow = mk('<div class="row3" style="margin:6px 0 10px">' +
    '<button class="tbtn ghost" data-t="text">+문단</button>' +
    '<button class="tbtn ghost" data-t="list">+목록</button>' +
    '<button class="tbtn ghost" data-t="img">+그림</button></div>');
  addRow.onclick = e => {
    const t = e.target.dataset.t; if (!t) return;
    if (t === 'text') s.blocks.push({ type: 'text', runs: [{ t: '내용을 입력하세요' }] });
    if (t === 'list') s.blocks.push({ type: 'list', items: [{ runs: [{ t: '항목' }], lvl: 0, num: 0 }] });
    if (t === 'img') s.blocks.push({ type: 'img', src: '', alt: '' });
    refresh();
  };
  host.appendChild(addRow);

  const nt = mk('<div class="field"><label class="lab">발표자 노트</label><textarea style="height:70px"></textarea></div>');
  $('textarea', nt).value = s.notes || '';
  $('textarea', nt).oninput = e => { s.notes = e.target.value; };
  host.appendChild(nt);

  const tools = mk('<div class="row2"><button class="tbtn ghost" id="edSplit">여기서 나누기</button>' +
    '<button class="tbtn ghost" id="edDel">이 장 삭제</button></div>');
  host.appendChild(tools);
  $('#edSplit').onclick = () => {
    if (s.blocks.length < 2) { say('나눌 내용이 부족합니다.', 'warn'); return; }
    const half = Math.ceil(s.blocks.length / 2);
    const rest = s.blocks.splice(half);
    state.slides.splice(state.sel + 1, 0, { id: uid(), title: s.title + ' (계속)', sub: '', notes: '', blocks: rest, on: true, layout: 'content' });
    refresh();
  };
  $('#edDel').onclick = () => { state.slides.splice(state.sel, 1); state.sel = clamp(state.sel, 0, state.slides.length - 1); refresh(); };
}
function blockEditor(s, b, bi) {
  const d = document.createElement('div');
  d.className = 'blk';
  const h = document.createElement('div');
  h.className = 'bh';
  h.innerHTML = '<span class="tag">' + (BLKNAME[b.type] || b.type) + '</span><span class="sp"></span>' +
    '<button data-a="up" title="위로">↑</button><button data-a="down" title="아래로">↓</button><button data-a="del" title="삭제">✕</button>';
  h.onclick = e => {
    const a = e.target.dataset.a; if (!a) return;
    if (a === 'up' && bi > 0) { const [m] = s.blocks.splice(bi, 1); s.blocks.splice(bi - 1, 0, m); }
    if (a === 'down' && bi < s.blocks.length - 1) { const [m] = s.blocks.splice(bi, 1); s.blocks.splice(bi + 1, 0, m); }
    if (a === 'del') s.blocks.splice(bi, 1);
    refresh();
  };
  d.appendChild(h);

  if (b.type === 'img') {
    const v = document.createElement('div');
    v.className = 'imgv';
    v.innerHTML = '<img src="' + esc(b.src) + '" alt=""><input type="text" value="' + esc(b.src) + '" placeholder="그림 주소 또는 data:">';
    $('input', v).oninput = e => { b.src = e.target.value; $('img', v).src = b.src; softRefresh(); };
    d.appendChild(v);
    return d;
  }
  const ta = document.createElement('textarea');
  ta.spellcheck = false;
  if (b.type === 'list') ta.value = b.items.map(it => '  '.repeat(it.lvl) + '- ' + rtext(it.runs)).join('\n');
  else if (b.type === 'table') ta.value = b.rows.map(r => r.map(c => c.t).join('\t')).join('\n');
  else if (b.type === 'code') ta.value = b.text;
  else ta.value = rtext(b.runs);
  ta.oninput = () => {
    const v = ta.value;
    if (b.type === 'list') {
      b.items = v.split('\n').filter(l => l.trim()).map(l => {
        const m = l.match(/^(\s*)(?:[-*+]\s+|\d+\.\s+)?(.*)$/);
        return { runs: [{ t: m[2] }], lvl: Math.min(3, Math.floor(m[1].replace(/\t/g, '  ').length / 2)), num: 0 };
      });
    } else if (b.type === 'table') {
      b.rows = v.split('\n').filter(l => l.trim()).map((l, ri) => l.split('\t').map(c => ({ t: c.trim(), h: ri === 0 && b.header })));
    } else if (b.type === 'code') b.text = v;
    else b.runs = asRuns(v);
    softRefresh();
  };
  d.appendChild(ta);
  return d;
}

/* ═══ 8. PPTX 내보내기 ═══════════════════════════════════════════ */
async function toDataURL(src) {
  if (!src) return null;
  if (/^data:/.test(src)) return src;
  const res = await fetch(src, { mode: 'cors' });
  if (!res.ok) throw new Error('HTTP ' + res.status);
  const blob = await res.blob();
  return await new Promise((ok, no) => {
    const fr = new FileReader();
    fr.onload = () => ok(fr.result); fr.onerror = () => no(new Error('read'));
    fr.readAsDataURL(blob);
  });
}
async function resolveImages(list) {
  const cache = {}; let bad = 0;
  const srcs = new Set();
  list.forEach(s => s.blocks.forEach(b => { if (b.type === 'img' && b.src) srcs.add(b.src); }));
  if (state.master.logo) srcs.add(state.master.logo);
  for (const s of srcs) {
    try { cache[s] = await toDataURL(s); }
    catch (e) { cache[s] = null; bad++; }
  }
  return { cache, bad };
}
async function exportPptx() {
  if (!state.slides.length) { say('먼저 슬라이드를 만드세요.', 'warn'); return; }
  if (!libReady()) {
    checkLib();
    say('PPTX 엔진 파일이 없습니다. 위쪽 안내 막대에서 파일을 고르거나 vendor/pptxgen.bundle.js 를 함께 올려 주세요.', 'bad');
    return;
  }
  say('PPTX를 만드는 중…');
  const pages = buildPages();
  const { cache, bad } = await resolveImages(state.slides.filter(s => s.on));
  const D = DIM[state.opts.aspect];
  const P = new PptxGenJS();
  P.defineLayout({ name: 'DF', width: D.W, height: D.H });
  P.layout = 'DF';
  P.title = state.deck.title || state.exp.name;
  P.author = state.deck.by || '';
  P.company = 'DECK·FORGE';

  pages.forEach((r, i) => {
    const s = P.addSlide();
    s.background = { color: state.theme.bg };
    r.shapes.forEach(sh => {
      if (sh.k === 'rect') {
        s.addShape(P.ShapeType.rect, { x: sh.x, y: sh.y, w: sh.w, h: sh.h, fill: { color: sh.fill }, line: { type: 'none' } });
      } else if (sh.k === 'text') {
        const runs = [];
        sh.paras.forEach(p => {
          const rs = p.runs.length ? p.runs : [{ t: ' ' }];
          rs.forEach((r2, ri) => {
            String(r2.t).split('\n').forEach((piece, pi, arr) => {
              runs.push({
                text: piece,
                options: {
                  bold: !!(r2.b || p.bold), italic: !!(r2.i || p.italic),
                  underline: r2.u ? { style: 'sng' } : undefined,
                  strike: !!r2.s,
                  fontFace: (r2.c || p.mono) ? 'Consolas' : sh.font,
                  fontSize: p.size, color: p.color || sh.color,
                  bullet: p.bullet ? (p.num ? { type: 'number' } : { code: '2022' }) : false,
                  indentLevel: p.lvl || 0,
                  align: p.align || sh.align,
                  breakLine: (ri === rs.length - 1 && pi === arr.length - 1) || pi < arr.length - 1
                }
              });
            });
          });
        });
        s.addText(runs, {
          x: sh.x, y: sh.y, w: sh.w, h: sh.h, valign: sh.valign || 'top', align: sh.align || 'left',
          fontFace: sh.font, fontSize: sh.size, color: sh.color, margin: 0, lineSpacingMultiple: 1.15, wrap: true
        });
      } else if (sh.k === 'img') {
        const data = cache[sh.src];
        if (data) s.addImage({ data, x: sh.x, y: sh.y, w: sh.w, h: sh.h, sizing: { type: 'contain', w: sh.w, h: sh.h } });
        else s.addShape(P.ShapeType.rect, { x: sh.x, y: sh.y, w: sh.w, h: sh.h, fill: { color: state.theme.panel }, line: { color: state.theme.muted, width: 0.5 } });
      } else if (sh.k === 'table') {
        const rows = sh.rows.map((row, ri) => row.map(c => ({
          text: c.t,
          options: (sh.header && ri === 0)
            ? { bold: true, color: sh.headText, fill: { color: sh.head } }
            : { color: sh.color, fill: { color: ri % 2 ? sh.zebra : 'FFFFFF00' } }
        })));
        s.addTable(rows, {
          x: sh.x, y: sh.y, w: sh.w, colW: Array(sh.rows[0].length).fill(sh.w / sh.rows[0].length),
          fontFace: sh.font, fontSize: sh.size, border: { pt: 0.5, color: sh.line }, valign: 'middle', margin: 3
        });
      }
    });
    if (state.exp.notes && r.notes) s.addNotes(r.notes);
  });
  try {
    await P.writeFile({ fileName: (state.exp.name || 'deck') + '.pptx' });
    say('내려받았습니다. ' + (isPoster() ? '포스터 ' : '슬라이드 ') + pages.length + '장' + (bad ? ' · 그림 ' + bad + '개는 가져오지 못했습니다' : ''), bad ? 'warn' : 'ok');
  } catch (e) {
    say('PPTX를 만들지 못했습니다: ' + e.message, 'bad');
  }
}

/* ═══ 9. 저장·복원 ═══════════════════════════════════════════════ */
function snapshot() {
  return { v: 2, raw: state.raw, slides: state.slides, opts: state.opts, theme: state.theme, deck: state.deck, master: state.master, exp: state.exp, poster: state.poster };
}
let saveT = null;
function autosave() {
  clearTimeout(saveT);
  saveT = setTimeout(() => {
    try {
      localStorage.setItem('df.work', JSON.stringify(snapshot()));
      $('#autoInfo').textContent = '이 브라우저에 자동 저장됨 · ' + new Date().toLocaleTimeString('ko-KR');
    } catch (e) { $('#autoInfo').textContent = '자동 저장 실패 (용량 초과)'; }
  }, 700);
}
function applySnapshot(d) {
  if (!d) return;
  state.raw = d.raw || '';
  state.slides = d.slides || [];
  Object.assign(state.opts, d.opts || {});
  Object.assign(state.theme, d.theme || {});
  Object.assign(state.deck, d.deck || {});
  Object.assign(state.master, d.master || {});
  Object.assign(state.exp, d.exp || {});
  Object.assign(state.poster, d.poster || {});
  if (!DIM[state.opts.aspect]) state.opts.aspect = '16x9';
  syncControls();
  refresh();
}
function restore() {
  try {
    const raw = localStorage.getItem('df.work');
    if (raw) { applySnapshot(JSON.parse(raw)); say('지난 작업을 이어서 불러왔습니다.'); return; }
  } catch (e) { }
  syncControls();
  refresh();
}

/* ═══ 10. 화면 갱신 ══════════════════════════════════════════════ */
let softT = null;
function softRefresh() { clearTimeout(softT); softT = setTimeout(() => { renderCanvas(); renderStrip(); autosave(); }, 160); }
function refresh() { renderCanvas(); renderStrip(); renderEditor(); autosave(); }

function build() {
  const src = $('#src').value;
  state.raw = src;
  if (!src.trim()) { say('원본이 비어 있습니다.', 'warn'); return; }
  state.imgAR = {};
  const t0 = performance.now();
  try {
    state.slides = parseSource(src, state.opts);
  } catch (e) {
    say('원본을 읽지 못했습니다: ' + e.message, 'bad'); return;
  }
  state.sel = 0;
  if (!state.deck.title && state.slides.length) { state.deck.title = state.slides[0].title; $('#deckTitle').value = state.deck.title; }
  refresh();
  say(state.slides.length + '장을 만들었습니다 · ' + Math.round(performance.now() - t0) + 'ms', 'ok');
  /* 그림 크기를 알아낸 뒤 배치를 한 번 더 맞춥니다 */
  const srcs = [...new Set(state.slides.flatMap(s => s.blocks.filter(b => b.type === 'img').map(b => b.src)))];
  let left = srcs.length;
  srcs.forEach(u => {
    const im = new Image();
    im.onload = () => { state.imgAR[u] = im.naturalWidth / im.naturalHeight; if (!--left) softRefresh(); };
    im.onerror = () => { if (!--left) softRefresh(); };
    im.src = u;
  });
}

function syncControls() {
  const o = state.opts;
  $('#src').value = state.raw;
  $('#split').value = o.split;
  $('#customSel').value = o.customSel;
  $('#customWrap').hidden = o.split !== 'custom';
  $('#optSplitLong').checked = o.splitLong;
  $('#optFirstTitle').checked = o.firstTitle;
  $('#optImages').checked = o.images;
  $('#optTables').checked = o.tables;
  $('#optNotes').checked = o.notes;
  $('#optLive').checked = o.live;
  $('#optShrink').checked = o.shrink;
  $('#tSize').value = o.titleSize;
  $('#bSize').value = o.bodySize;
  $('#headStyle').value = o.headStyle;
  $('#fontSel').value = o.font;
  $('#pageSize').value = o.aspect;
  $('#posterPane').hidden = !isPoster();
  $('#pCols').value = state.poster.cols;
  $('#pScale').value = state.poster.scale;
  $('#pSec').value = state.poster.sec;
  $('#pMerge').checked = state.poster.merge;
  $('#pBanner').checked = state.poster.banner;
  if (isPoster()) {
    const D = DIM[o.aspect];
    const pk = Math.min(D.W, D.H) / 20 * state.poster.scale;
    $('#pInfo').textContent = D.name + ' · 본문 ' + Math.round(o.bodySize * pk) + 'pt · 구역제목 ' +
      Math.round((o.bodySize + 5) * pk) + 'pt · 큰제목 ' + Math.round(o.titleSize * pk * 2) + 'pt';
  }
  $('#deckTitle').value = state.deck.title;
  $('#deckSub').value = state.deck.sub;
  $('#deckBy').value = state.deck.by;
  $('#footText').value = state.master.foot;
  $('#optPage').checked = state.master.page;
  $('#optFoot').checked = state.master.showFoot;
  $('#fname').value = state.exp.name;
  $('#optExpNotes').checked = state.exp.notes;
  $('#cAccent').value = '#' + state.theme.accent;
  $('#cBg').value = '#' + state.theme.bg;
  $('#cText').value = '#' + state.theme.fg;
  $('#logoInfo').textContent = state.master.logo ? '넣음' : '없음';
  $('#srcInfo').textContent = state.raw ? (state.raw.length.toLocaleString() + '자') : '비어 있음';
  renderThemes();
}
function renderThemes() {
  const host = $('#themes');
  host.innerHTML = '';
  THEMES.forEach(t => {
    const b = document.createElement('button');
    b.className = 'sw';
    b.setAttribute('aria-pressed', t.id === state.theme.id);
    b.innerHTML = '<div class="bar"><i style="background:#' + t.bg + '"></i><i style="background:#' + t.accent + '"></i><i style="background:#' + t.title + '"></i></div><span>' + t.name + '</span>';
    b.onclick = () => {
      state.theme = Object.assign({}, t, { font: state.theme.font });
      syncControls(); refresh();
    };
    host.appendChild(b);
  });
}

/* ═══ 11. 예제 ═══════════════════════════════════════════════════ */
const SAMPLE = [
  '<h1>방사선 작업 안전 브리핑</h1>',
  '<p>2026년 상반기 정기 교육 · 방사선안전부</p>',
  '<h2>오늘 다룰 내용</h2>',
  '<ul><li>작업 전 <strong>선량 예측</strong>과 계획</li>',
  '<li>구역 구분과 출입 절차<ul><li>관리구역</li><li>제한구역</li></ul></li>',
  '<li>개인 선량계 착용 규칙</li><li>이상 상황 보고 체계</li></ul>',
  '<aside class="notes">첫 5분은 지난 분기 사례로 시작합니다.</aside>',
  '<h2>선량 한도 요약</h2>',
  '<table><thead><tr><th>대상</th><th>연간 유효선량</th><th>수정체</th></tr></thead>',
  '<tbody><tr><td>방사선작업종사자</td><td>50 mSv</td><td>20 mSv</td></tr>',
  '<tr><td>수시출입자</td><td>12 mSv</td><td>15 mSv</td></tr>',
  '<tr><td>일반인</td><td>1 mSv</td><td>15 mSv</td></tr></tbody></table>',
  '<p>5년 평균은 <em>연 20 mSv</em>를 넘지 않아야 합니다.</p>',
  '<h2>작업 전 점검</h2>',
  '<blockquote>측정하지 않은 선량은 관리할 수 없다.<cite>현장 수칙</cite></blockquote>',
  '<ol><li>작업 구역 선량률 측정</li><li>차폐·거리·시간 계획 수립</li><li>예상 집단선량 산정</li></ol>',
  '<h2>기록 양식 예시</h2>',
  '<pre><code>WORKER  ID     DOSE(mSv)  DATE\nKIM     A-1021  0.42      2026-03-11\nLEE     A-1044  0.18      2026-03-11</code></pre>',
  '<h2>정리</h2>',
  '<p>계획 · 측정 · 기록. 세 가지가 지켜지면 대부분의 사고는 예방됩니다.</p>'
].join('\n');


/* 포스터 판형에서 쓰는 예제 */
const SAMPLE_POSTER = [
  '<h1>작업자 실시간 선량 관리 체계의 현장 적용</h1>',
  '<h2>배경 및 목적</h2>',
  '<p>정기 정비 기간 중 고선량 구역 작업은 계획선량과 실제선량의 차이가 크다. 본 연구는 실시간 선량 전송과 구역별 선량률 지도를 결합해 작업 중 의사결정을 지원하는 체계를 시험 적용하고 그 효과를 평가하였다.</p>',
  '<ul><li>계획선량 대비 실제선량 편차 축소</li><li>고선량 구역 체류시간 단축</li><li>작업 중단 없이 관리자 개입 가능성 확인</li></ul>',
  '<h2>방법</h2>',
  '<p>2개 호기 정비 공정 12건, 연인원 340명을 대상으로 6개월간 적용하였다. 개인선량계 실시간 전송값을 30초 간격으로 수집하고, 구역별 기준선량률과 비교해 임계값 초과 시 현장 단말에 경보를 보냈다.</p>',
  '<ol><li>작업 전 구역 선량률 측정 및 지도 작성</li><li>공정별 계획선량 산정</li><li>작업 중 실시간 수집 및 경보</li><li>작업 후 편차 분석과 다음 공정 반영</li></ol>',
  '<h2>결과</h2>',
  '<table><thead><tr><th>구분</th><th>적용 전</th><th>적용 후</th><th>변화</th></tr></thead>',
  '<tbody><tr><td>계획 대비 편차</td><td>28 %</td><td>11 %</td><td>-17 %p</td></tr>',
  '<tr><td>공정당 집단선량</td><td>4.6 man·mSv</td><td>3.4 man·mSv</td><td>-26 %</td></tr>',
  '<tr><td>고선량 구역 체류</td><td>42 분</td><td>29 분</td><td>-31 %</td></tr>',
  '<tr><td>경보 후 조치 시간</td><td>-</td><td>1.8 분</td><td>-</td></tr></tbody></table>',
  '<p>편차 축소 효과는 선량률 구배가 큰 구역에서 두드러졌으며, 동일 공정을 반복할수록 계획 정확도가 개선되었다.</p>',
  '<h2>고찰</h2>',
  '<blockquote>측정값이 작업자에게 즉시 보이는 것만으로도 체류시간이 줄었다.<cite>현장 관찰</cite></blockquote>',
  '<p>경보 자체보다 선량률 지도의 사전 공유가 행동 변화에 더 크게 기여한 것으로 보인다. 다만 금속 구조물이 밀집한 구역에서는 전송 지연이 관측되어 보완이 필요하다.</p>',
  '<h2>결론</h2>',
  '<ul><li>계획 대비 편차 28 %에서 11 %로 감소</li><li>공정당 집단선량 26 % 저감</li><li>전송 지연 구간의 중계기 보강이 다음 과제</li></ul>',
  '<h2>참고문헌</h2>',
  '<p>1. ICRP Publication 103 (2007). 2. IAEA GSR Part 3 (2014). 3. 원자력안전위원회 고시 제2024-XX호.</p>'
].join('\n');

/* ═══ 12. 자체 검증 ══════════════════════════════════════════════ */
async function runTests() {
  const T = [];
  const t = (name, fn) => { try { const r = fn(); T.push({ name, ok: r === true, note: r === true ? '' : String(r) }); } catch (e) { T.push({ name, ok: false, note: e.message }); } };
  const O = Object.assign({}, state.opts);

  t('제목마다 장이 나뉜다', () => {
    const s = parseSource('<h1>A</h1><p>a</p><h1>B</h1><p>b</p>', Object.assign({}, O, { split: 'h1', firstTitle: false }));
    return s.length === 2 && s[0].title === 'A' && s[1].title === 'B' || 'len=' + s.length;
  });
  t('가로줄로도 나뉜다', () => {
    const s = parseSource('<p>a</p><hr><p>b</p>', Object.assign({}, O, { split: 'hr', firstTitle: false }));
    return s.length === 2 || 'len=' + s.length;
  });
  t('중첩 목록의 단계가 유지된다', () => {
    const s = parseSource('<h1>T</h1><ul><li>a<ul><li>b</li></ul></li></ul>', Object.assign({}, O, { split: 'h1', firstTitle: false }));
    const L = s[0].blocks.find(b => b.type === 'list');
    return (L && L.items.length === 2 && L.items[1].lvl === 1) || '단계 인식 실패';
  });
  t('굵게·기울임이 보존된다', () => {
    const s = parseSource('<h1>T</h1><p>보통 <strong>굵게</strong> <em>기울임</em></p>', Object.assign({}, O, { split: 'h1', firstTitle: false }));
    const r = s[0].blocks[0].runs;
    return (r.some(x => x.b) && r.some(x => x.i)) || '강조 유실';
  });
  t('표가 표로 들어온다', () => {
    const s = parseSource('<h1>T</h1><table><tr><th>a</th><th>b</th></tr><tr><td>1</td><td>2</td></tr></table>', Object.assign({}, O, { split: 'h1', firstTitle: false }));
    const tb = s[0].blocks.find(b => b.type === 'table');
    return (tb && tb.rows.length === 2 && tb.header) || '표 인식 실패';
  });
  t('발표자 노트를 읽는다', () => {
    const s = parseSource('<h1>T</h1><aside class="notes">메모</aside>', Object.assign({}, O, { split: 'h1', firstTitle: false }));
    return s[0].notes === '메모' || '노트=' + s[0].notes;
  });
  t('마크다운을 알아본다', () => {
    const h = md2html('# 제목\n\n- 하나\n- 둘\n\n| a | b |\n|---|---|\n| 1 | 2 |');
    return (/<h1>/.test(h) && /<li>/.test(h) && /<table>/.test(h)) || '변환 누락';
  });
  t('모든 도형이 슬라이드 안에 있다', () => {
    const s = parseSource(SAMPLE, Object.assign({}, O, { split: 'auto' }));
    const D = DIM[state.opts.aspect];
    let bad = 0;
    s.forEach((sl, i) => compose(sl, i, s.length).shapes.forEach(sh => {
      if (sh.x < -0.01 || sh.y < -0.01 || sh.x + sh.w > D.W + 0.01 || sh.y + sh.h > D.H + 0.2) bad++;
    }));
    return bad === 0 || bad + '개가 밖으로 나감';
  });
  t('내용이 넘치면 다음 장으로 넘어간다', () => {
    const long = '<h1>긴 장</h1>' + Array(40).fill('<p>이 문단은 넘침 처리를 확인하기 위한 충분히 긴 본문입니다.</p>').join('');
    const keep = state.slides, keepA = state.opts.aspect;
    state.opts.aspect = '16x9';                     // 슬라이드 판형에서 확인합니다
    state.slides = parseSource(long, Object.assign({}, O, { split: 'h1', firstTitle: false }));
    const n = expandDeck().length;
    state.slides = keep; state.opts.aspect = keepA;
    return n > 1 || '나뉘지 않음(' + n + ')';
  });
  t('마크다운 중첩 목록의 단계가 유지된다', () => {
    const s = parseSource('# T\n\n- 하나\n  - 안쪽\n- 둘\n', Object.assign({}, O, { split: 'h1', firstTitle: false }));
    const L = s[0].blocks.find(b => b.type === 'list');
    return (L && L.items.length === 3 && L.items[1].lvl === 1) || '단계 유실';
  });
  t('모든 제목 장식에서 도형이 넘치지 않는다', () => {
    const keep = O.headStyle; let bad = 0;
    const s = parseSource(SAMPLE, Object.assign({}, O, { split: 'auto' }));
    ['rule', 'bar', 'side', 'block', 'plain'].forEach(h => {
      state.opts.headStyle = h;
      const D = DIM[state.opts.aspect];
      s.forEach((sl, i) => compose(sl, i, s.length).shapes.forEach(sh => {
        if (sh.x < -0.01 || sh.y < -0.01 || sh.x + sh.w > D.W + 0.01 || sh.y + sh.h > D.H + 0.2) bad++;
      }));
    });
    state.opts.headStyle = keep;
    return bad === 0 || bad + '개 이탈';
  });
  t('A0 포스터가 한 장으로 조판된다', () => {
    const keep = state.opts.aspect, ks = state.slides;
    state.opts.aspect = 'a0p';
    state.slides = parseSource(SAMPLE, Object.assign({}, O, { split: 'auto', firstTitle: false }));
    const pg = buildPages();
    const D = DIM.a0p; let bad = 0;
    pg.forEach(p => p.shapes.forEach(sh => {
      if (sh.x < -0.01 || sh.y < -0.01 || sh.x + sh.w > D.W + 0.01 || sh.y + sh.h > D.H + 0.5) bad++;
    }));
    const n = pg.length;
    state.opts.aspect = keep; state.slides = ks;
    return (n === 1 && bad === 0) || ('쪽=' + n + ' 이탈=' + bad);
  });
  t('포스터 단 수를 바꾸면 배치가 달라진다', () => {
    const keep = state.opts.aspect, ks = state.slides, kc = state.poster.cols;
    state.opts.aspect = 'a0p';
    state.slides = parseSource(SAMPLE, Object.assign({}, O, { split: 'auto', firstTitle: false }));
    state.poster.cols = 2; const x2 = buildPages()[0].shapes.map(s => (s.x + s.w).toFixed(2)).join();
    state.poster.cols = 4; const x4 = buildPages()[0].shapes.map(s => (s.x + s.w).toFixed(2)).join();
    state.poster.cols = kc; state.opts.aspect = keep; state.slides = ks;
    return x2 !== x4 || '단 수가 반영되지 않음';
  });
  t('포스터 글자가 판형에 맞게 커진다', () => {
    const keep = state.opts.aspect, ks = state.slides;
    state.opts.aspect = 'a0p';
    state.slides = parseSource(SAMPLE, Object.assign({}, O, { split: 'auto', firstTitle: false }));
    const sizes = buildPages()[0].shapes.filter(s => s.k === 'text').map(s => s.size);
    const small = Math.min.apply(null, sizes), big = Math.max.apply(null, sizes);
    state.opts.aspect = keep; state.slides = ks;
    return (small > 18 && big > 70) || ('가장 작은 글자 ' + small.toFixed(0) + 'pt · 가장 큰 글자 ' + big.toFixed(0) + 'pt');
  });
  t('숨긴 요소가 실제로 숨겨진다', () => {
    const probe = document.createElement('div');
    probe.hidden = true; probe.style.display = 'flex';
    document.body.appendChild(probe);
    const d = getComputedStyle(probe).display;
    probe.remove();
    return d === 'none' || ('hidden 인데 display:' + d);
  });
  t('엔진이 준비되면 안내 막대가 보이지 않는다', () => {
    const w = $('#libWarn');
    if (!w) return '막대 요소 없음';
    if (!libReady()) return true;                       // 엔진이 없는 환경에서는 건너뜁니다
    const d = getComputedStyle(w).display;
    return (w.hidden && d === 'none') || ('hidden=' + w.hidden + ' display=' + d);
  });
  t('테마 색이 모두 정상 형식이다', () => {
    const bad = THEMES.filter(x => ['bg', 'fg', 'title', 'accent', 'muted', 'panel'].some(k => !/^[0-9A-F]{6}$/.test(x[k])));
    return bad.length === 0 || bad.map(b => b.id).join(',');
  });
  t('열쇠말 검사가 동작한다', () => (fnv('redpass') === KEY_FNV) || 'FNV 불일치');
  t('PPTX 엔진이 준비되어 있다', () => (typeof PptxGenJS !== 'undefined') || 'pptxgen 로드 실패');

  // 실제 파일 생성까지 확인
  let genNote = '';
  let genOk = false;
  try {
    const P = new PptxGenJS();
    P.defineLayout({ name: 'T', width: 10, height: 5.625 }); P.layout = 'T';
    const s = P.addSlide();
    s.addText([{ text: '검증', options: { fontSize: 20 } }], { x: 1, y: 1, w: 6, h: 1 });
    const blob = await P.write({ outputType: 'blob' });
    genOk = blob && blob.size > 2000;
    genNote = genOk ? (Math.round(blob.size / 1024) + 'KB 생성') : '크기가 너무 작습니다';
  } catch (e) { genNote = e.message; }
  T.push({ name: 'PPTX 파일이 실제로 만들어진다', ok: !!genOk, note: genNote });

  const pass = T.filter(x => x.ok).length;
  $('#testBody').innerHTML = '<p class="mini" style="font-size:13px;margin-bottom:10px">' + pass + ' / ' + T.length + ' 통과</p>' +
    '<table class="tests">' + T.map(x => '<tr><td class="' + (x.ok ? 'pass' : 'fail') + '">' + (x.ok ? '통과' : '실패') +
      '</td><td>' + esc(x.name) + (x.note ? '<br><span style="color:var(--faint)">' + esc(x.note) + '</span>' : '') + '</td></tr>').join('') + '</table>';
  $('#mTest').classList.add('on');
}

/* ═══ 13. 연결 ═══════════════════════════════════════════════════ */
function bind() {
  $('#pwGo').onclick = tryKey;
  $('#pw').onkeydown = e => { if (e.key === 'Enter') tryKey(); };

  $('#btnBuild').onclick = build;
  $('#btnExport').onclick = exportPptx;
  $('#btnExport2').onclick = exportPptx;
  $('#btnPrint').onclick = () => window.print();
  $('#btnTest').onclick = runTests;
  $('#btnHelp').onclick = () => $('#mHelp').classList.add('on');
  $$('[data-close]').forEach(b => b.onclick = () => $$('.modal').forEach(m => m.classList.remove('on')));
  $$('.modal').forEach(m => m.onclick = e => { if (e.target === m) m.classList.remove('on'); });

  $('#btnSample').onclick = () => {
    const src = isPoster() ? SAMPLE_POSTER : SAMPLE;
    $('#src').value = src; state.raw = src; state.deck.title = ''; state.deck.sub = ''; build();
    if (isPoster()) { state.deck.sub = '한국수력원자력 방사선안전부'; state.deck.by = '홍길동, 김철수, 이영희'; posterAutoFit(); }
  };
  $('#btnClear').onclick = () => { $('#src').value = ''; state.raw = ''; state.slides = []; refresh(); say('비웠습니다.'); };
  $('#btnPaste').onclick = async () => {
    try { const t = await navigator.clipboard.readText(); $('#src').value = t; state.raw = t; build(); }
    catch (e) { say('클립보드를 읽지 못했습니다. 직접 붙여넣어 주세요.', 'warn'); }
  };
  $('#btnFile').onclick = () => $('#file').click();
  $('#file').onchange = e => { const f = e.target.files[0]; if (f) readFile(f); };
  const dz = $('#drop');
  ['dragenter', 'dragover'].forEach(v => dz.addEventListener(v, e => { e.preventDefault(); dz.classList.add('hot'); }));
  ['dragleave', 'drop'].forEach(v => dz.addEventListener(v, e => { e.preventDefault(); dz.classList.remove('hot'); }));
  dz.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) readFile(f); });
  document.addEventListener('dragover', e => e.preventDefault());
  document.addEventListener('drop', e => e.preventDefault());

  $('#src').oninput = e => {
    state.raw = e.target.value;
    $('#srcInfo').textContent = state.raw ? state.raw.length.toLocaleString() + '자' : '비어 있음';
    if (state.opts.live) { clearTimeout(softT); softT = setTimeout(build, 500); }
    autosave();
  };

  const O = state.opts;
  $('#split').onchange = e => { O.split = e.target.value; $('#customWrap').hidden = O.split !== 'custom'; build(); };
  $('#customSel').oninput = e => { O.customSel = e.target.value; };
  $('#customSel').onchange = build;
  const rebuilds = { optSplitLong: 'splitLong', optFirstTitle: 'firstTitle', optImages: 'images', optTables: 'tables', optNotes: 'notes' };
  Object.keys(rebuilds).forEach(id => {
    $('#' + id).onchange = e => { O[rebuilds[id]] = e.target.checked; if (state.raw) build(); else refresh(); };
  });
  $('#optLive').onchange = e => { O.live = e.target.checked; autosave(); };
  $('#optShrink').onchange = e => { O.shrink = e.target.checked; refresh(); };
  $('#tSize').oninput = e => { O.titleSize = +e.target.value || 30; softRefresh(); };
  $('#bSize').oninput = e => { O.bodySize = +e.target.value || 17; softRefresh(); };
  $('#headStyle').onchange = e => { O.headStyle = e.target.value; refresh(); };
  $('#fontSel').onchange = e => { O.font = e.target.value; state.theme.font = e.target.value; refresh(); };
  $('#pageSize').onchange = e => {
    O.aspect = e.target.value; state.page = 0;
    if (isPoster() && state.opts.bodySize > 24) state.opts.bodySize = 17;
    syncControls(); refresh();
    say(DIM[O.aspect].name + ' 판형으로 바꿨습니다' + (isPoster() ? ' · 오른쪽에서 단 수와 배율을 조절하세요' : ''));
  };
  const PO = state.poster;
  $('#pCols').onchange = e => { PO.cols = +e.target.value; refresh(); };
  $('#pScale').oninput = e => { PO.scale = clamp(+e.target.value || 1, 0.5, 2); syncControls(); softRefresh(); };
  $('#pSec').onchange = e => { PO.sec = e.target.value; refresh(); };
  $('#pMerge').onchange = e => { PO.merge = e.target.checked; refresh(); };
  $('#pBanner').onchange = e => { PO.banner = e.target.checked; refresh(); };
  $('#pFit').onclick = posterAutoFit;

  $('#cAccent').oninput = e => { state.theme.accent = hex(e.target.value); state.theme.id = 'custom'; softRefresh(); };
  $('#cBg').oninput = e => { state.theme.bg = hex(e.target.value); state.theme.id = 'custom'; softRefresh(); };
  $('#cText').oninput = e => { state.theme.fg = hex(e.target.value); state.theme.id = 'custom'; softRefresh(); };

  $('#deckTitle').oninput = e => { state.deck.title = e.target.value; softRefresh(); };
  $('#deckSub').oninput = e => { state.deck.sub = e.target.value; softRefresh(); };
  $('#deckBy').oninput = e => { state.deck.by = e.target.value; softRefresh(); };
  $('#footText').oninput = e => { state.master.foot = e.target.value; softRefresh(); };
  $('#optPage').onchange = e => { state.master.page = e.target.checked; refresh(); };
  $('#optFoot').onchange = e => { state.master.showFoot = e.target.checked; refresh(); };
  $('#fname').oninput = e => { state.exp.name = e.target.value.replace(/[^\w.\-가-힣 ]/g, '') || 'deck'; };
  $('#optExpNotes').onchange = e => { state.exp.notes = e.target.checked; };

  $('#btnLogo').onclick = () => $('#logoFile').click();
  $('#logoFile').onchange = e => {
    const f = e.target.files[0]; if (!f) return;
    const fr = new FileReader();
    fr.onload = () => {
      state.master.logo = fr.result;
      const im = new Image();
      im.onload = () => { state.master.logoAR = im.naturalWidth / im.naturalHeight; $('#logoInfo').textContent = '넣음'; refresh(); };
      im.src = fr.result;
    };
    fr.readAsDataURL(f);
  };
  $('#btnLogoDel').onclick = () => { state.master.logo = null; $('#logoInfo').textContent = '없음'; refresh(); };

  $('#btnSave').onclick = () => {
    const b = new Blob([JSON.stringify(snapshot(), null, 1)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(b); a.download = (state.exp.name || 'deck') + '.json'; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 3000);
    say('작업을 저장했습니다.', 'ok');
  };
  $('#btnLoad').onclick = () => $('#projFile').click();
  $('#projFile').onchange = e => {
    const f = e.target.files[0]; if (!f) return;
    const fr = new FileReader();
    fr.onload = () => { try { applySnapshot(JSON.parse(fr.result)); say('작업을 불러왔습니다.', 'ok'); } catch (err) { say('파일을 읽지 못했습니다.', 'bad'); } };
    fr.readAsText(f);
  };

  $('#btnPrev').onclick = () => {
    if (isPoster()) { state.page = Math.max(0, state.page - 1); jumpToPage(); }
    else state.sel = Math.max(0, state.sel - 1);
    refresh();
  };
  $('#btnNext').onclick = () => {
    if (isPoster()) { state.page = Math.min(buildPages().length - 1, state.page + 1); jumpToPage(); }
    else state.sel = Math.min(state.slides.length - 1, state.sel + 1);
    refresh();
  };

  document.addEventListener('keydown', e => {
    if ($('#gate').classList.contains('hide') === false) return;
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); build(); }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); exportPptx(); }
    if (e.altKey && e.key === 'ArrowLeft') { e.preventDefault(); $('#btnPrev').click(); }
    if (e.altKey && e.key === 'ArrowRight') { e.preventDefault(); $('#btnNext').click(); }
    if (e.key === 'Escape') $$('.modal').forEach(m => m.classList.remove('on'));
  });
  let rz = null;
  window.addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(renderCanvas, 120); });
}

function readFile(f) {
  const fr = new FileReader();
  fr.onload = () => {
    $('#src').value = fr.result;
    state.raw = fr.result;
    if (!state.exp.name || state.exp.name === 'deck') {
      state.exp.name = f.name.replace(/\.[^.]+$/, '').replace(/[^\w.\-가-힣 ]/g, '') || 'deck';
      $('#fname').value = state.exp.name;
    }
    build();
  };
  fr.readAsText(f, 'utf-8');
}

/* ═══ 14. PPTX 엔진이 없을 때의 복구 ═════════════════════════════ */
function libReady() { return typeof PptxGenJS !== 'undefined'; }
function runLib(code, from) {
  try {
    const s = document.createElement('script');
    s.textContent = code;
    document.body.appendChild(s);
  } catch (e) { }
  if (libReady()) {
    $('#libWarn').hidden = true;
    say('PPTX 엔진을 ' + from + '에서 불러왔습니다. 이제 내려받기가 됩니다.', 'ok');
    return true;
  }
  say('그 파일에서는 엔진을 찾지 못했습니다. pptxgen.bundle.js 가 맞는지 확인하세요.', 'bad');
  return false;
}
function checkLib() {
  const warn = $('#libWarn');
  if (!warn) return;
  const st = $('#statLib');
  if (libReady()) {
    warn.hidden = true;
    if (st) { st.textContent = window.__pptxInline ? '내장 · 준비됨' : '외부 · 준비됨'; st.style.color = 'var(--ok)'; }
    return;
  }
  if (st && window.__pptxDone) { st.textContent = '없음'; st.style.color = 'var(--bad)'; }
  if (!window.__pptxDone) { warn.hidden = true; setTimeout(checkLib, 250); return; }  // 아직 찾는 중
  warn.hidden = false;
  const tried = (window.__pptxTried || []);
  console.log('DECK·FORGE — 엔진을 찾지 못했습니다. 시도한 위치:', tried.join(', '));
  const title = $('#libTitle'), msg = $('#libMsg'), where = $('#libWhere');
  const net = tried.some(u => /^https?:/.test(u));
  if (!window.__fileComplete) {
    title.textContent = 'HTML 파일이 잘렸습니다';
    msg.innerHTML = '이 페이지의 끝부분이 없습니다. 복사·붙여넣기로 옮기면 큰 파일이 잘립니다. ' +
      '파일을 <b>내려받아 그대로 업로드</b>하거나, 용량이 작은 판을 쓰세요.';
  } else if (window.__pptxInline) {
    title.textContent = '내장 엔진이 실행되지 않았습니다';
    msg.textContent = '파일이 손상되었을 수 있습니다. 다시 내려받아 그대로 올려 주세요.';
  } else if (net) {
    title.textContent = 'PPTX 엔진을 찾지 못했습니다';
    msg.innerHTML = '같은 폴더에 <code style="font-family:var(--mono)">deckforge-engine.js</code> 를 두거나, ' +
      '인터넷 연결을 확인하세요. 변환·미리보기·편집은 지금도 됩니다.';
  } else {
    title.textContent = 'PPTX 엔진을 찾지 못했습니다';
    msg.textContent = '변환·미리보기·편집은 지금도 됩니다.';
  }
  if (where) where.textContent = '파일: ' + location.pathname + ' · 시도 ' + tried.length + '곳';
  $('#btnLibFile').onclick = () => $('#libFile').click();
  $('#libFile').onchange = e => {
    const f = e.target.files[0]; if (!f) return;
    const fr = new FileReader();
    fr.onload = () => runLib(fr.result, f.name);
    fr.readAsText(f);
  };
  $('#btnLibCdn').onclick = () => {
    say('인터넷에서 엔진을 받는 중…');
    const s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
    s.onload = () => {
      $('#libWarn').hidden = libReady();
      say(libReady() ? 'PPTX 엔진을 받았습니다. 이 탭에서만 유효하니, 저장소에 vendor/pptxgen.bundle.js 를 넣어 두세요.' : '받았지만 엔진이 없습니다.', libReady() ? 'ok' : 'bad');
    };
    s.onerror = () => say('인터넷에서 받지 못했습니다. 파일을 직접 골라 주세요.', 'bad');
    document.body.appendChild(s);
  };
}

/* 시작 */
bind();
checkLib();
try { if (localStorage.getItem('df.key') === '1') openApp(); } catch (e) { }
$('#pw').focus();
