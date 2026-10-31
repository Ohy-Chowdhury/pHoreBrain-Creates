/* MCQ -> PPTX core. Works in browser (window) and Node (tests). */
(function (root) {
  'use strict';
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const EMU = 914400;
  const JSZipLib = () => root.JSZip || require('jszip');
  const parseXml = (s) => (root.DOMParser ? new root.DOMParser() : new (require('@xmldom/xmldom').DOMParser)()).parseFromString(s, 'application/xml');
  const toXml = (d) => {
    const s = (root.XMLSerializer ? new root.XMLSerializer() : new (require('@xmldom/xmldom').XMLSerializer)()).serializeToString(d);
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s.replace(/^<\?xml[^>]*\?>\s*/, '');
  };
  const kids = (el, name) => Array.from(el.childNodes).filter((n) => n.nodeType === 1 && n.localName === name);
  const all = (el, ns, name) => Array.from(el.getElementsByTagNameNS(ns, name));
  const shapeText = (sp) => all(sp, NS_A, 't').map((t) => t.textContent).join('');
  const shapes = (doc) => all(doc, NS_P, 'sp');
  const xfrmOf = (sp) => {
    const x = sp.getElementsByTagNameNS(NS_A, 'xfrm')[0];
    const off = x.getElementsByTagNameNS(NS_A, 'off')[0];
    const ext = x.getElementsByTagNameNS(NS_A, 'ext')[0];
    return { off, ext, x: +off.getAttribute('x'), y: +off.getAttribute('y'), cx: +ext.getAttribute('cx'), cy: +ext.getAttribute('cy') };
  };
  const isRoundRect = (sp) => { const g = sp.getElementsByTagNameNS(NS_A, 'prstGeom')[0]; return g && g.getAttribute('prst') === 'roundRect'; };
  const hasBengali = (s) => /[\u0980-\u09ff]/.test(s);

  /* ---------- input parsing ---------- */
  function parseItems(text) {
    const items = [], errors = [];
    text.split(/\r?\n/).forEach((raw, i) => {
      const line = raw.trim();
      if (!line || line[0] === '#') return;
      const p = line.split('|').map((s) => s.trim());
      if (p[0] === 'T') {
        if (!p[1]) errors.push(`Line ${i + 1}: Type heading khali`);
        else items.push({ kind: 'T', title: p[1] });
      } else if (p[0] === 'Q') {
        if (p.length < 7) { errors.push(`Line ${i + 1}: Q| e question, 4ta option ar answer lagbe`); return; }
        const ans = p[6].toUpperCase();
        if (!/^[A-D]$/.test(ans)) { errors.push(`Line ${i + 1}: answer A/B/C/D hote hobe`); return; }
        items.push({ kind: 'Q', q: p[1], opts: p.slice(2, 6), ans, extra: p[7] || '' });
      } else errors.push(`Line ${i + 1}: T| ba Q| diye shuru korte hobe`);
    });
    return { items, errors };
  }

  /* ---------- slide fillers ---------- */
  function fillQuestion(xml, item) {
    const doc = parseXml(xml);
    const sps = shapes(doc);
    const body = sps.find((sp) => /\(A\)/.test(shapeText(sp)));
    const ansSp = sps.find((sp) => /^\s*Ans/.test(shapeText(sp)));
    const rect = sps.find(isRoundRect);
    if (!body) throw new Error('Template-er question slide e "(A)" likha text box paoa jay nai');
    if (!ansSp) throw new Error('Template-er question slide e "Ans:" box paoa jay nai');
    const txBody = body.getElementsByTagNameNS(NS_P, 'txBody')[0];
    const paras = kids(txBody, 'p');
    const tplR = paras[0].getElementsByTagNameNS(NS_A, 'r')[0].cloneNode(true);
    const pPrEl = kids(paras[0], 'pPr')[0];
    const pPr = pPrEl ? pPrEl.cloneNode(true) : null;
    const rPr0 = tplR.getElementsByTagNameNS(NS_A, 'rPr')[0];
    const sz = +(rPr0 && rPr0.getAttribute('sz')) || 1800;
    const spc = pPr && pPr.getElementsByTagNameNS(NS_A, 'spcPct')[0];
    const pct = spc ? +spc.getAttribute('val') / 100000 : 1;
    const defTab = (pPr && +pPr.getAttribute('defTabSz')) || 914400;
    const tabIn = defTab / EMU;
    const xf = xfrmOf(body);
    const W = xf.cx / EMU - 0.4;
    const cw = 0.45 * sz / 7200;
    const meas = (s) => { let w = 0; for (const ch of s) { if (ch === '\t') continue; w += ch === ' ' ? cw * 0.5 : cw; } return w; };
    const nlines = (s) => Math.max(1, Math.ceil(meas(s) / W));
    const tabsTo = (s, k) => '\t'.repeat(Math.max(1, k - Math.floor(meas(s) / tabIn)));

    const ow = ['A', 'B', 'C', 'D'].map((c, i) => `(${c}) ${item.opts[i]}`);
    const m = item.q.match(/^(.*?)\s*(\[[^\]]*\])\s*$/);
    const qt = m ? m[1] : item.q, tag = m ? m[2] : '';
    const wq = meas(qt + ' ');
    const lines = []; let n;
    if (tag && wq < 0.46 * W) {
      const tgt = Math.round(0.72 * W / tabIn);
      lines.push(qt + ' ' + '\t'.repeat(Math.max(1, tgt - Math.floor(wq / tabIn))) + tag);
      n = (tgt * tabIn + meas(tag) > W - 0.5) ? 2 : 1;
    } else { const l = (qt + ' ' + tag).trim(); lines.push(l); n = nlines(l); }
    const k = Math.ceil(Math.max(meas(ow[0]), meas(ow[2])) / tabIn) + 2;
    if (k * tabIn + Math.max(meas(ow[1]), meas(ow[3])) <= W - 1) {
      lines.push(ow[0] + tabsTo(ow[0], k) + ow[1], ow[2] + tabsTo(ow[2], k) + ow[3]); n += 2;
    } else ow.forEach((x) => { lines.push(x); n += nlines(x); });
    if (item.extra) { lines.push(item.extra); n += nlines(item.extra); }

    paras.forEach((p) => txBody.removeChild(p));
    lines.forEach((text) => {
      const p = doc.createElementNS(NS_A, 'a:p');
      if (pPr) p.appendChild(pPr.cloneNode(true));
      const r = tplR.cloneNode(true);
      const rp = r.getElementsByTagNameNS(NS_A, 'rPr')[0];
      if (rp) { rp.setAttribute('lang', hasBengali(text) ? 'as-IN' : 'en-US'); rp.removeAttribute('err'); }
      r.getElementsByTagNameNS(NS_A, 't')[0].textContent = text;
      p.appendChild(r); txBody.appendChild(p);
    });
    const lineEmu = sz / 100 * pct * 12700;
    const h = Math.round(n * lineEmu + 183000);
    xf.ext.setAttribute('cy', String(h));
    if (rect) {
      const rx = xfrmOf(rect);
      rx.ext.setAttribute('cy', String(Math.round(Math.max(rx.cy, (xf.y - rx.y) * 2 + h + lineEmu))));
    }
    const runs = all(ansSp, NS_A, 'r');
    if (runs.length >= 2) runs[runs.length - 1].getElementsByTagNameNS(NS_A, 't')[0].textContent = item.ans;
    else { const t = runs[0].getElementsByTagNameNS(NS_A, 't')[0]; t.textContent = t.textContent.replace(/Ans:\s*[A-D]?/, 'Ans: ' + item.ans); }
    return toXml(doc);
  }

  function fillType(xml, title) {
    const doc = parseXml(xml);
    const sps = shapes(doc);
    const ts = sps.find((sp) => /^\s*Type/i.test(shapeText(sp)));
    if (!ts) throw new Error('Template-er type slide e "Type" likha text box paoa jay nai');
    const rect = sps.find(isRoundRect);
    const ci = title.indexOf(':');
    const label = (ci >= 0 ? title.slice(0, ci) : 'Type') + ': ';
    const name = (ci >= 0 ? title.slice(ci + 1) : title).trim();
    const runs = all(ts, NS_A, 'r');
    const setT = (r, s) => { r.getElementsByTagNameNS(NS_A, 't')[0].textContent = s; };
    if (runs.length >= 2) {
      setT(runs[0], label); setT(runs[1], name);
      for (let i = 2; i < runs.length; i++) runs[i].parentNode.removeChild(runs[i]);
    } else setT(runs[0], label + name);
    const rp = runs[0].getElementsByTagNameNS(NS_A, 'rPr')[0];
    const sz = +(rp && rp.getAttribute('sz')) || 1800;
    const spc = ts.getElementsByTagNameNS(NS_A, 'spcPct')[0];
    const pct = spc ? +spc.getAttribute('val') / 100000 : 1;
    const xf = xfrmOf(ts);
    const W = xf.cx / EMU - 0.4;
    const nl = Math.max(1, Math.ceil((label.length + name.length) * 0.27 * sz / 7200 / W));
    if (nl > 1) {
      const extra = Math.round((nl - 1) * sz / 100 * pct * 12700);
      if (rect) {
        const rx = xfrmOf(rect);
        const off = xf.y - rx.y;
        const ny = rx.y - Math.round(extra / 2);
        rx.off.setAttribute('y', String(ny)); rx.ext.setAttribute('cy', String(rx.cy + extra));
        xf.off.setAttribute('y', String(ny + off));
      }
      xf.ext.setAttribute('cy', String(xf.cy + extra));
    }
    return toXml(doc);
  }

  /* ---------- package handling ---------- */
  const resolve = (t) => (t.startsWith('/') ? t.slice(1) : 'ppt/' + t);
  async function readDeck(zip) {
    const pres = await zip.file('ppt/presentation.xml').async('string');
    const rels = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
    const map = {};
    (rels.match(/<Relationship\b[^>]*>/g) || []).forEach((r) => {
      const id = /\bId="([^"]+)"/.exec(r), tg = /\bTarget="([^"]+)"/.exec(r);
      if (id && tg) map[id[1]] = tg[1];
    });
    const order = (pres.match(/<p:sldId\b[^>]*>/g) || []).map((t) => {
      const id = /\bid="(\d+)"/.exec(t)[1], rid = /\br:id="([^"]+)"/.exec(t)[1];
      const path = resolve(map[rid]);
      return { id: +id, rid, path, relsPath: path.replace('slides/', 'slides/_rels/') + '.rels' };
    });
    return { pres, rels, order };
  }
  const slideText = async (zip, s) => all(parseXml(await zip.file(s.path).async('string')), NS_A, 't').map((t) => t.textContent).join('');

  async function inspect(buf) {
    const zip = await JSZipLib().loadAsync(buf);
    const { order } = await readDeck(zip);
    const texts = [];
    for (const s of order) texts.push(await slideText(zip, s));
    const q = texts.findIndex((t) => /Ans\s*:/.test(t));
    const ty = texts.findIndex((t, i) => i !== q && /(^|\s)Type/i.test(t) && !/Ans\s*:/.test(t));
    return { count: order.length, questionSlide: q + 1 || null, typeSlide: ty + 1 || null, texts };
  }

  async function generate(buf, items, opts = {}) {
    const zip = await JSZipLib().loadAsync(buf);
    const deck = await readDeck(zip);
    const info = await inspect(buf);
    const qNo = (opts.questionSlide || info.questionSlide), tNo = (opts.typeSlide || info.typeSlide);
    if (!qNo) throw new Error('Template e question sample slide paoa jay nai ("Ans:" thakte hobe)');
    const qIdx = qNo - 1, tIdx = tNo ? tNo - 1 : null;
    const qXml = await zip.file(deck.order[qIdx].path).async('string');
    const tXml = tIdx !== null ? await zip.file(deck.order[tIdx].path).async('string') : null;
    const qRels = (await zip.file(deck.order[qIdx].relsPath).async('string'));
    const tRels = tIdx !== null ? await zip.file(deck.order[tIdx].relsPath).async('string') : null;
    if (items.some((i) => i.kind === 'T') && tXml === null) throw new Error('Type heading ache kintu template e type sample slide nai');

    let maxN = 0;
    Object.keys(zip.files).forEach((f) => { const m = /^ppt\/slides\/slide(\d+)\.xml$/.exec(f); if (m) maxN = Math.max(maxN, +m[1]); });
    let ct = await zip.file('[Content_Types].xml').async('string');
    let prels = deck.rels, pres = deck.pres;
    const samples = [qIdx].concat(tIdx !== null ? [tIdx] : []);
    // remove sample slides (and their notes)
    for (const si of samples) {
      const s = deck.order[si];
      const r = await zip.file(s.relsPath).async('string');
      (r.match(/<Relationship\b[^>]*notesSlide[^>]*>/g) || []).forEach((rel) => {
        const tg = /Target="([^"]+)"/.exec(rel)[1].replace('../', 'ppt/');
        zip.remove(tg); zip.remove(tg.replace('notesSlides/', 'notesSlides/_rels/') + '.rels');
        ct = ct.replace(new RegExp('<Override[^>]*PartName="/' + tg.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*/>'), '');
      });
      zip.remove(s.path); zip.remove(s.relsPath);
      ct = ct.replace(new RegExp('<Override[^>]*PartName="/' + s.path.replace(/\./g, '\\.') + '"[^>]*/>'), '');
      prels = prels.replace(new RegExp('<Relationship\\b[^>]*Id="' + s.rid + '"[^>]*/>'), '');
    }
    const stripNotes = (r) => r.replace(/<Relationship\b[^>]*notesSlide[^>]*\/>/g, '');
    const gen = [];
    let nextId = Math.max(...deck.order.map((o) => o.id), 255) + 1, n = maxN;
    for (let i = 0; i < items.length; i++) {
      const it = items[i]; n++;
      const path = `ppt/slides/slide${n}.xml`, rid = `rIdGen${n}`;
      zip.file(path, it.kind === 'T' ? fillType(tXml, it.title) : fillQuestion(qXml, it));
      zip.file(`ppt/slides/_rels/slide${n}.xml.rels`, stripNotes(it.kind === 'T' ? tRels : qRels));
      ct = ct.replace('</Types>', `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
      prels = prels.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${n}.xml"/></Relationships>`);
      gen.push(`<p:sldId id="${nextId++}" r:id="${rid}"/>`);
    }
    const keep = deck.order.map((o, i) => ({ o, i })).filter((x) => !samples.includes(x.i));
    const at = Math.min(...samples);
    const before = keep.filter((x) => x.i < at).map((x) => `<p:sldId id="${x.o.id}" r:id="${x.o.rid}"/>`);
    const after = keep.filter((x) => x.i > at).map((x) => `<p:sldId id="${x.o.id}" r:id="${x.o.rid}"/>`);
    pres = pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, '<p:sldIdLst>' + before.join('') + gen.join('') + after.join('') + '</p:sldIdLst>');
    zip.file('ppt/presentation.xml', pres);
    zip.file('ppt/_rels/presentation.xml.rels', prels);
    zip.file('[Content_Types].xml', ct);
    if (zip.file('docProps/app.xml')) {
      const app = await zip.file('docProps/app.xml').async('string');
      zip.file('docProps/app.xml', app.replace(/<Slides>\d+<\/Slides>/, `<Slides>${before.length + gen.length + after.length}</Slides>`));
    }
    return zip.generateAsync({ type: opts.type || 'uint8array', compression: 'DEFLATE' });
  }

  const api = { parseItems, inspect, generate };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.MCQ = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
