/* Armodilo CI core: parse a FedEx commercial-invoice PDF (via pdf.js text items)
   and build the Armodilo-branded .docx (Google Docs compatible). Works in browser and Node. */
(function (root) {
  const L = 14.4;            // left margin (pt)
  const LH = 13.5;           // body line height (pt)
  const AMT_R = 501.9;       // right edge of the money column in the totals block (FedEx right-aligns these)
  const AWB_X = 444.6;       // "Air Waybill Nbr." column on the header rows (only when FedEx has one)
  // ACI number from the file name, e.g. "527617844816 - Douglas County Libraries.txt" -> 527617844816
  function aciFromName(name) { const m = String(name || '').replace(/^.*[\\/]/, '').match(/(?<!\d)\d{9,15}(?!\d)/); return m ? m[0] : ''; }
  // value printed under "Air Waybill Nbr." ('' = column not printed)
  // "Unit Qty: 10.00" is too wide for the EA column, so from 10 up the unit follows the number instead
  const qtySegs = it => String(it.qty).length > 4 ? [[185.5, `Unit Qty: ${it.qty} ${it.uom}`]] : [[185.5, `Unit Qty: ${it.qty}`], [252.2, it.uom]];
  // comment lines without their "1)" numbers; true when FedEx left them all blank
  const commentText = d => (d.comments || []).map(c => c.replace(/^\d\)\s*/, ''));
  const commentsBlank = d => !commentText(d).some(t => t.trim());
  // comments to print: FedEx's, or the standard ones from config.js when FedEx left them blank (unless edited on the page)
  function commentsFor(d, cfg) {
    if (d.comments_edited || !commentsBlank(d) || !(cfg.default_comments || []).length) return d.comments || [];
    return cfg.default_comments.map((t, i) => `${i + 1})${t}`);
  }
  const awbOut = (d, aci) => d.has_awb ? (d.awb_out != null ? d.awb_out : aci) : '';
  const SIG = { w: 54, h: 30 };

  // ---------- parsing ----------
  async function extractLines(pdfjs, data) {
    const doc = await pdfjs.getDocument({ data, isEvalSupported: false }).promise;
    const lines = [];
    let off = 0;
    for (let p = 1; p <= doc.numPages; p++) {
      const pg = await doc.getPage(p);
      const vp = pg.getViewport({ scale: 1 });
      const tc = await pg.getTextContent();
      const rows = new Map();
      for (const it of tc.items) {
        if (!it.str || !it.str.trim()) continue;
        const top = Math.round(vp.height - it.transform[5]) + off;
        const key = [...rows.keys()].find(k => Math.abs(k - top) <= 2) ?? top;
        if (!rows.has(key)) rows.set(key, []);
        rows.get(key).push({ x: it.transform[4], text: it.str.trim() });
      }
      [...rows.keys()].sort((a, b) => a - b).forEach(k => {
        const words = rows.get(k).sort((a, b) => a.x - b.x);
        lines.push({ top: k, words, text: words.map(w => w.text).join(' ') });
      });
      off += vp.height;
    }
    return lines;
  }


  // Fixed-width .txt export from FedEx: same structure as the PDF, columns found by character position.
  function extractTextLines(text) {
    const lines = [];
    text.replace(/\r/g, '').split('\n').forEach((raw, i) => {
      const words = [];
      const re = /\S+(?: \S+)*/g;          // runs separated by 2+ spaces
      let m;
      while ((m = re.exec(raw))) words.push({ x: m.index * 4, text: m[0] });  // ~4pt per character
      if (words.length) lines.push({ top: i * 12, words, text: words.map(w => w.text).join(' ') });
    });
    return lines;
  }

  const fmtZip = z => /^\d{9}$/.test(z) ? `${z.slice(0, 5)}-${z.slice(5)}` : z;
  const split = (line, x = 260) => [
    line.words.filter(w => w.x < x).map(w => w.text).join(' ').trim(),
    line.words.filter(w => w.x >= x).map(w => w.text).join(' ').trim()];

  function parseAddress(lines) {
    lines = lines.filter(Boolean);
    let id = '';
    if (lines.length && lines[lines.length - 1].startsWith('ID/EIN:')) id = lines.pop().slice(7).trim();
    let state = '', zip = '', country = '', phone = '';
    const m = lines.length ? lines[lines.length - 1].match(/^([A-Z]{2})\s+(\S+)\s+([A-Z]{2})\s*(.*)$/) : null;
    if (m) { state = m[1]; zip = fmtZip(m[2]); country = m[3]; phone = m[4].trim(); lines.pop(); }
    const city = lines.length ? lines.pop() : '';
    return { name: lines[0] || '', company: lines[1] || '', street: lines.slice(2), city, state, zip, country, phone, id };
  }

  function parse(ls) {
    const txt = ls.map(l => l.text).join('\n');
    const isRule = l => /^[-_]{5,}$/.test(l.text.replace(/\s/g, ''));
    const idx = re => { const i = ls.findIndex(l => re.test(l.text)); if (i < 0) throw new Error('This does not look like a FedEx commercial invoice (missing "' + re.source + '").'); return i; };
    const d = {};
    let i = idx(/^Invoice Number/);
    let w = ls[i + 1].text.split(/\s+/);
    [d.purpose, d.currency, d.ult_dest] = w;
    i = idx(/^Export Dt/);
    // read by column, so a multi-word C.I. Reference ("Tablet PC Display Enclosure") stays in one piece
    // Pkgs is the first whole number after the reference; everything between the date and it is the reference
    // (.txt columns are already separated by wide gaps, so those are used as they are)
    const ew = ls[i + 1].words.map(x => x.text);
    w = ls[i + 1].text.split(/\s+/);
    const pk = w.findIndex((x, k) => k > 1 && /^\d+$/.test(x));
    if ((ew.length === 3 || ew.length === 4) && /^\d+$/.test(ew[2])) [d.export_date, d.ci_ref, d.pkgs, d.bill_tc] = ew;
    else if (pk > 1) [d.export_date, d.ci_ref, d.pkgs, d.bill_tc] = [w[0], w.slice(1, pk).join(' '), w[pk], w.slice(pk + 1).join(' ')];
    else [d.export_date, d.ci_ref, d.pkgs, d.bill_tc] = w;
    const nxt = ls[i + 2].text.trim();
    d.pkg_type = /^[A-Z]+$/.test(nxt) ? nxt : '';
    const block = re => { const s = idx(re) + 1; let e = s; while (e < ls.length && !isRule(ls[e])) e++; return ls.slice(s, e); };
    d.consignee = parseAddress(block(/^Shipper:/).map(l => split(l)[1]));
    const imp = block(/^Broker:/);
    d.broker = imp.map(l => split(l)[0]).filter(Boolean);
    d.importer = parseAddress(imp.map(l => split(l)[1]));
    d.items = [];
    ls.forEach((l, s) => {
      if (!l.text.startsWith('Part Nbr:')) return;
      let e = s; while (e < ls.length && !isRule(ls[e])) e++;
      const b = ls.slice(s, e).map(x => x.text).join(' ');
      const g = (re, def = '') => { const m = b.match(re); return m ? m[1].trim() : def; };
      d.items.push({
        part: g(/Part Nbr:\s*(.*?)\s*Marks\/Nbrs:/) || g(/Part Nbr:\s*(\S+)/), marks: g(/Marks\/Nbrs:\s*(.*?)\s*C\/T/), mfr: g(/C\/T MFR:\s*(\S+)/),
        net_wgt: g(/Net\s+Wgt:\s*([\d.]+)/), wgt_unit: g(/Net\s+Wgt:\s*[\d.]+\s*(\w+)/, 'lbs'),
        hs: g(/HS Code:\s*(\S+)/), qty: g(/Unit Qty:\s*([\d.]+)/), uom: g(/Unit Qty:\s*[\d.]+\s*(\w+)/, 'EA'),
        commodity_value: g(/Commodity Value:\s*([\d.,]+)/), unit_value: g(/(?<![\d.])(\d+\.\d{6})(?![\d.])/),
        desc: g(/Desc:\s*(.*)$/)
      });
    });
    // lines FedEx puts between the products and the totals, e.g. "Following Repair/Processing"
    d.notes = [];
    const lastItem = ls.map(l => l.text).reduce((a, t, i) => t.startsWith('Part Nbr:') ? i : a, -1);
    const totalsAt = ls.findIndex(l => /^Total Shipment/.test(l.text));
    if (lastItem >= 0 && totalsAt > lastItem) {
      let k = lastItem; while (k < totalsAt && !isRule(ls[k])) k++;
      for (k++; k < totalsAt; k++) if (!isRule(ls[k]) && ls[k].text.trim()) d.notes.push(ls[k].text.trim());
    }
    const g = (re, def = '') => { const m = txt.match(re); return m ? m[1].trim() : def; };
    d.total_weight = g(/Total Shipment\s+([\d.]+)/);
    d.total_commodity = g(/Total Commodity\s*\n?([\d.,]+)/);
    d.terms = g(/Terms of Sale:\s*\n?(\w+)/);
    d.freight = g(/Freight:\s*\n?([\d.,]+)/);
    d.insurance = g(/Insurance:\s*\n?([\d.,]+)/);
    d.others = g(/Others:\s*\n?([\d.,]+)/);
    d.total_invoice = g(/Total Invoice Value:\s*\n?([\d.,]+)/);
    // comments run until the tracking line, the declaration or the signature (newer FedEx files have no tracking line)
    d.comments = (g(/Comments:\s*\n([\s\S]*?)(?:Tracking Numbers|I declare|Signature of shipper|$)/).match(/^\d\).*$/gm) || []);
    d.tracking = g(/Tracking Numbers:\s*([\d ,]+)/).split(/[\s,]+/).filter(Boolean);
    d.sig_date = g(/Signature of shipper:.*?(\d{1,2}\/\d{1,2}\/\d{4})/);
    // "Air Waybill Nbr" (when FedEx includes it) is the ACI number. Value on the same line or the line below.
    // d.has_awb: the file has an "Air Waybill Nbr" heading (even if FedEx left X's under it) -> it's printed on the invoice
    d.awb = '';
    const ai = ls.findIndex(l => /air\s*way\s*bill/i.test(l.text));
    d.has_awb = ai >= 0;
    if (ai >= 0) {
      const after = ls[ai].text.replace(/^.*?air\s*way\s*bill\s*(?:nbr|no|number|#)?\.?\s*:?/i, '');
      const pick = t => { const m = String(t).match(/\b\d{4}[ -]\d{4}[ -]\d{4}\b|\b\d{9,15}\b/); return m ? m[0].replace(/[\s-]/g, '') : ''; };
      d.awb = pick(after) || (ls[ai + 1] ? pick(ls[ai + 1].text) : '');
    }
    return d;
  }


  // ---------- matching a FedEx line to saved metal content ----------
  // norm: capitals, letters and digits only ("ALUR - 14" -> "ALUR14")
  const norm = s => String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  // product family = first word's leading letters: "ALUR14", "ALUR - 14", "ALUR TABLET KIOSK 14" -> "ALUR"
  const family = s => (String(s || '').toUpperCase().trim().match(/^[A-Z]+/) || [''])[0];
  const sizeOf = s => { const m = norm(s).match(/^[A-Z]+(\d+)$/); return m ? String(+m[1]) : null; };
  // sizes written in text: 14, 14", 14”, 14-inch, "Kiosk 14"
  const sizesIn = t => [...String(t || '').matchAll(/(?:^|[^\d.])(\d{2})(?=\s*(?:["\u201d\u2033]|in\b|inch|-?inch|\b))/gi)].map(m => String(+m[1]));

  /* Returns { key, how } for a confident match, or { key: null, options: [...], reason } when the
     user has to choose. Order: exact part number -> same product family, picked by size -> none. */
  // description match: FedEx "Desc:" vs a product's page-2 title (without its leading code and bracketed note)
  // or its saved description. Only used when it points to exactly one product.
  const words = t => String(t || '').toLowerCase().replace(/\(.*?\)/g, ' ').replace(/[^a-z0-9]+/g, ' ').trim();
  const titleCore = p => { const t = String(p.title || '').replace(/\s*\(.*$/, ''); return t.replace(/^(?:[A-Z0-9\/&-]+\s+)+(?=[A-Za-z])/, m => /[a-z]/.test(t) && m.trim() !== t.trim() ? '' : m); };
  function byDescription(cfg, keys, desc) {
    const d = words(desc);
    if (!d) return [];
    return keys.filter(k => { const p = cfg.parts[k]; return [words(titleCore(p)), words(p.title), words(p.item && p.item.desc)].includes(d); });
  }
  function matchPart(cfg, part, desc) {
    const m = matchPartCore(cfg, part, desc);
    if (m.key) return m;
    const keys = Object.keys(cfg.parts || {});
    const hits = byDescription(cfg, keys, desc);
    // prefer products from the same family as the part number ("ON/OFF-BA-01" -> ON...), else any single hit
    const fam = family(part), inFam = hits.filter(k => family(k) === fam);
    const pick = inFam.length === 1 ? inFam[0] : (hits.length === 1 ? hits[0] : null);
    if (pick) return { key: pick, how: 'the description on the invoice' };
    if (hits.length > 1) return { key: null, options: hits, reason: `the description matches more than one product (${hits.join(', ')})` };
    return m;
  }
  function matchPartCore(cfg, part, desc) {
    const keys = Object.keys(cfg.parts || {});
    const n = norm(part);
    const exact = keys.find(k => norm(k) === n);
    if (exact) return { key: exact, how: 'part number' };
    const fam = family(part);
    const sameFamily = fam ? keys.filter(k => family(k) === fam) : [];
    // one saved product with no sizes: only automatic when nothing else is added ("KBT02" -> "KBT-02"),
    // so a different product like "TILT BASE" is never mistaken for "TILT"
    if (sameFamily.length === 1 && !sizeOf(sameFamily[0])) {
      const np = norm(part), nk = norm(sameFamily[0]);
      const onlyDigitsDiffer = (np.startsWith(nk) && /^\d*$/.test(np.slice(nk.length))) || (nk.startsWith(np) && /^\d+$/.test(nk.slice(np.length)));
      if (onlyDigitsDiffer) return { key: sameFamily[0], how: 'part number' };
      return { key: null, options: sameFamily, reason: `not sure this is ${sameFamily[0]}` };
    }
    if (sameFamily.length) {
      const sized = sameFamily.filter(k => sizeOf(k));
      const wanted = [...new Set([sizeOf(part), ...sizesIn(part), ...sizesIn(desc)].filter(Boolean))];
      const hits = sized.filter(k => wanted.includes(sizeOf(k)));
      if (hits.length === 1) return { key: hits[0], how: `size ${sizeOf(hits[0])} from the invoice` };
      const list = sameFamily.join(', ');
      return { key: null, options: sameFamily, reason: wanted.length && sized.length
        ? `size ${wanted.join('/')} doesn’t match a saved ${fam} (saved: ${list})`
        : `the invoice doesn’t say which ${fam} it is (saved: ${list})` };
    }
    return { key: null, options: [], reason: 'no saved metal content for this part number' };
  }
  // the saved metal content for a line: a user's choice (it.metal_key) wins, '' means none
  // shipments of only Tablet Fit Kits have no metal: page 2 is "Product Information" and both headings are black
  const FIT_KIT = 'TABLETFITKIT';
  function fitKitOnly(d, cfg) {
    const keyOf = p => Object.keys(cfg.parts || {}).find(k => cfg.parts[k] === p) || '';
    return d.items.length > 0 && d.items.every(it => { const p = partFor(cfg, it); return p && keyOf(p).toUpperCase().replace(/[^A-Z0-9]/g, '') === FIT_KIT; });
  }
  function headings(d, cfg) {
    const fk = fitKitOnly(d, cfg);
    return { fitKit: fk, p1: `Please see Page ${metalPage(d)} for ${fk ? 'Product Information' : 'Steel/Aluminum Content Breakdown'}`,
             p2: fk ? 'Product Information' : 'Metal Content', black: fk };
  }
  function partFor(cfg, it) {
    if (it.metal_key !== undefined) return it.metal_key ? cfg.parts[it.metal_key] || null : null;
    const m = matchPart(cfg, it.part, it.desc);
    return m.key ? cfg.parts[m.key] : null;
  }
  const findPart = (cfg, part) => { const m = matchPart(cfg, part, ''); return m.key ? cfg.parts[m.key] : null; };

  // split "text (bracketed note)" so the bracketed part can be blue, like the original
  function splitParen(t) { const m = t.match(/^(.*?)( \(.*\))$/); return m ? [m[1], m[2]] : [t, '']; }
  // fields for one metal component, scaled by the line quantity: { label, value, note, bold }
  // (bold like the original: material, value, weight and countries)
  function componentFields(c, q) {
    const f = [];
    if (c.name) f.push({ label: 'Component', value: c.name, note: c.note ? `(${c.note})` : '' });
    if (c.hs) f.push({ label: 'H.S. Code', value: c.hs });
    if (c.qty != null) f.push({ label: 'Quantity', value: String(+(c.qty * q).toFixed(4)) });
    f.push({ label: 'Material', value: c.material, bold: true });
    f.push({ label: c.value_label || 'Total Value', value: `US$${(c.value * q).toFixed(2)}`, bold: true });
    const dec = (String(c.weight).split('.')[1] || '').length;
    const w = q === 1 ? c.weight : (parseFloat(c.weight) * q).toFixed(dec);
    // optional "6.51 lbs 12.5%" after the kg: lbs are recalculated for the quantity, the % share stays the same
    let sfx = c.weight_suffix || '';
    const lb = sfx.match(/^([\d.]+)\s*lbs\s*(.*)$/);
    if (lb && q !== 1) sfx = `${(parseFloat(w) * 2.20462).toFixed(2)} lbs${lb[2] ? ' ' + lb[2] : ''}`;
    f.push({ label: c.weight_label || 'Total Weight', value: `${w} kg${sfx ? ' ' + sfx : ''}`, bold: true });
    c.origin.forEach(([a, b]) => f.push({ label: a, value: b, bold: true }));
    f.push({ label: 'Reference', value: c.reference });
    return f;
  }
  // plain "Label: value (note)" lines (used for checks)
  const componentRows = (c, q) => componentFields(c, q).map(x => `${x.label}: ${x.value}${x.note ? ' ' + x.note : ''}`);
  const VALUE_X = 198;   // where the answers line up on page 2 (pt from page edge)


  // ---------- live material breakdown from the Google Doc ----------
  /* raw = { paragraphs: [{ runs: [{t, b, c}], img: [{mime, data, w, h}] }] } as sent by the Doc's Apps Script.
     Returns { parts, items, problems }: parts in the same shape as config.js, items = per-product defaults. */
  function parseBreakdown(raw) {
    const P = (raw && raw.paragraphs) || [];
    const txt = p => (p.runs || []).map(r => r.t).join('');
    const isBlue = c => /^#?0000ff$/i.test(c || '');
    const sectionRe = /^\s*(?:Part\s*)?Nbr\s*:\s*(.*?)\s*Marks\/Nbrs/i;
    const endRe = /^\s*(?:Page \d+ of \d+|Line \d+ of \d+|FILES TO ATTACH|Simplified Version|[-\u2014_]{10,})/i;
    const parts = {}, items = {}, problems = [];
    let sec = null, block = null, comp = null, blocksInSection = 0;
    const num = s => { const m = String(s).replace(/,/g, '').match(/-?\d*\.?\d+/); return m ? parseFloat(m[0]) : NaN; };
    const finishComp = () => { if (block && comp && Object.keys(comp).length) block.components.push(comp); comp = null; };
    const finishBlock = () => {
      finishComp();
      if (!block || block.pending) { block = null; return; }
      const b = block; block = null;
      const bad = b.components.filter(c => !c.material || isNaN(c.value) || !c.weight || isNaN(num(c.weight)));
      if (bad.length) { problems.push(`${b.key}: ${bad.length} component${bad.length > 1 ? 's are' : ' is'} missing a material, value or weight`); return; }
      if (parts[b.key]) return;                      // first one wins (the Doc repeats WALL)
      parts[b.key] = { title: b.title, title_color: b.color, components: b.components, images: b.images, source: 'doc' };
    };
    const capsKey = t => { const w = []; for (const x of t.split(/\s+/)) { if (/^[A-Z0-9\/&-]+$/.test(x) && /[A-Z]/.test(x)) w.push(x); else break; } return w.join(' '); };
    for (const p of P) {
      const t = txt(p).replace(/\u00a0/g, ' '), tt = t.trim();
      const sm = t.match(sectionRe);
      if (sm) {
        finishBlock();
        sec = { key: sm[1].replace(/\s+/g, ' ').trim() }; blocksInSection = 0;
        items[sec.key] = { part: sec.key };
        continue;
      }
      if (sec && !block) {                            // product header lines
        let m;
        if ((m = t.match(/HS Code:\s*([\d.]+)/i))) items[sec.key].hs = m[1];
        if ((m = tt.match(/^(\d+\.\d{4,})$/))) items[sec.key].unit_value = parseFloat(m[1]);
        if ((m = t.match(/^\s*Desc:\s*(.*)$/i))) items[sec.key].desc = m[1].replace(/\s*\(?\*.*$/, '').replace(/\s*\(\s*Diagram.*$/i, '').replace(/\s+/g, ' ').trim();
      }
      if (/^metal content$/i.test(tt)) { finishBlock(); block = { pending: true }; continue; }
      const lm = tt.match(/^Line \d+:\s*(.*)$/i);
      if (lm && sec && !/^Line \d+ of/i.test(tt) && (!block || block.pending)) {
        finishBlock();
        const runs = p.runs || [];
        const main = runs.filter(r => !isBlue(r.c)).map(r => r.t).join('').replace(/^\s*Line \d+:\s*/i, '').replace(/\s+/g, ' ').trim();
        const note = runs.filter(r => isBlue(r.c)).map(r => r.t).join('').replace(/\s+/g, ' ').trim();
        const first = runs.find(r => r.t.trim()) || {};
        const color = /^#?000000$/.test(first.c || '#ff0000') ? 'black' : 'red';
        const key = blocksInSection === 0 ? sec.key : (capsKey(main) || sec.key);
        blocksInSection++;
        block = { key, title: note ? `${main} ${note}` : main, color, components: [], images: [] };
        continue;
      }
      if (!block || block.pending) { if (p.img && p.img.length && block && !block.pending) block.images.push(...p.img); continue; }
      if (endRe.test(tt) || /^Line \d+ of \d+/i.test(tt)) { finishBlock(); continue; }
      if (p.img && p.img.length) block.images.push(...p.img);
      if (!tt) { finishComp(); continue; }
      const i = t.indexOf(':');
      if (i < 0) continue;
      const label = t.slice(0, i).replace(/\s+/g, ' ').trim();
      // value = text after the colon; blue runs inside it are the bracketed note
      let pos = 0, val = '', note = '';
      for (const r of p.runs || []) {
        const s0 = pos, s1 = pos + r.t.length; pos = s1;
        const part = s1 <= i + 1 ? '' : r.t.slice(Math.max(0, i + 1 - s0));
        if (isBlue(r.c)) note += part; else val += part;
      }
      val = val.replace(/\s+/g, ' ').trim(); note = note.replace(/\s+/g, ' ').trim();
      comp = comp || {};
      const L = label.toLowerCase();
      if (L === 'component') { if (comp.name || comp.material) { finishComp(); comp = {}; } comp.name = val; if (note) comp.note = note.replace(/^\(|\)$/g, ''); }
      else if (L === 'h.s. code' || L === 'hs code') comp.hs = val;
      else if (L === 'quantity') comp.qty = num(val);
      else if (L === 'material') { if (comp.material) { finishComp(); comp = {}; } comp.material = val; }
      else if (/total (component )?value/i.test(label)) { comp.value = num(val); if (/component/i.test(label)) comp.value_label = label; }
      else if (/total (component )?weight/i.test(label)) {
        const wm = val.match(/^(-?\d*\.?\d+)\s*kg\s*(.*)$/i);
        comp.weight = wm ? wm[1] : val; if (wm && wm[2].trim()) comp.weight_suffix = wm[2].replace(/\s+/g, ' ').trim();
        if (/component/i.test(label)) comp.weight_label = label;
      }
      else if (L === 'reference') comp.reference = val;
      else (comp.origin = comp.origin || []).push([label, val]);
    }
    finishBlock();
    Object.values(parts).forEach(p => p.components.forEach(c => { c.origin = c.origin || []; c.reference = c.reference || ''; }));
    return { parts, items, problems };
  }
  // put the Doc's breakdown on top of the built-in one (built-in stays for anything the Doc doesn't cover)
  function mergeBreakdown(cfg, live) {
    const parts = Object.assign({}, cfg.parts);
    for (const [k, p] of Object.entries(live.parts)) {
      const old = Object.keys(parts).find(x => norm(x) === norm(k));
      const base = old ? parts[old] : {};
      if (old && old !== k) delete parts[old];
      parts[old || k] = Object.assign({}, p, { item: Object.assign({}, base.item || {}, live.items[k] || {}) });
    }
    for (const [k, it] of Object.entries(live.items)) {
      const old = Object.keys(parts).find(x => norm(x) === norm(k));
      if (old && parts[old].source !== 'doc') parts[old] = Object.assign({}, parts[old], { item: Object.assign({}, parts[old].item || {}, it) });
    }
    return Object.assign({}, cfg, { parts });
  }

  // ---------- lines added by hand ----------
  const money = n => (Math.round(n * 100) / 100).toFixed(2);
  const num = s => parseFloat(String(s || '').replace(/,/g, '')) || 0;
  // a new invoice line, prefilled from a saved product's defaults
  function newLine(cfg, key) {
    const it = (cfg.parts[key] || {}).item || {};
    const line = { added: true, part: it.part || key || '', marks: '', mfr: 'CA', net_wgt: '0.00', wgt_unit: 'lbs',
      hs: it.hs || '', qty: '1.00', uom: 'EA', unit: it.unit_value != null ? String(it.unit_value) : '',
      desc: it.desc || '', metal_key: key || '', metal_how: key ? 'the product you picked' : '' };
    return fillLine(line);
  }
  // recompute the formatted fields of an added line from qty and unit value
  function fillLine(l) {
    l.qty = num(l.qty_in != null ? l.qty_in : l.qty).toFixed(2);
    l.unit_value = num(l.unit).toFixed(6);
    l.commodity_value = money(num(l.qty) * num(l.unit));
    return l;
  }
  // copy of the invoice with totals that include any added lines
  function withTotals(d) {
    if (!d.items.some(it => it.added)) return d;
    const commodity = d.items.reduce((t, it) => t + num(it.commodity_value), 0);
    const total = commodity + num(d.freight) + num(d.insurance) + num(d.others);
    return Object.assign({}, d, { total_commodity: money(commodity), total_invoice: money(total) });
  }

  // ---------- docx ----------
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const tw = pt => Math.round(pt * 20);
  const emu = pt => Math.round(pt * 12700);

  function run(text, o = {}) {
    let rpr = '';
    if (o.b) rpr += '<w:b/><w:bCs/>';
    if (o.color) rpr += `<w:color w:val="${o.color}"/>`;
    if (o.sz) rpr += `<w:sz w:val="${o.sz * 2}"/><w:szCs w:val="${o.sz * 2}"/>`;
    return `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
  }
  const TAB = '<w:r><w:tab/></w:r>';
  let picId = 1;
  function pic(rid, wPt, hPt, name) {
    const id = picId++;
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${emu(wPt)}" cy="${emu(hPt)}"/><wp:docPr id="${id}" name="${name}"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(wPt)}" cy="${emu(hPt)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  }
  // segs: [x, text, align?]; opts: center, b, sz, color, rule, line, pageBreak, auto
  function para(segs = [], o = {}) {
    let ppr = '';
    const tabs = [];
    let body = o.pageBreak ? '<w:r><w:br w:type="page"/></w:r>' : '';
    if (o.center !== undefined) body += run(o.center, o);
    for (const [x, text, align] of segs) {
      if (x > L + 0.5) { tabs.push(`<w:tab w:val="${align === 'right' ? 'right' : 'left'}" w:pos="${tw(x - L)}"/>`); body += TAB; }
      body += run(text, o);
    }
    if (o.raw) body += o.raw;
    if (o.rule) ppr += `<w:pBdr><w:bottom w:val="${o.rule}" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>`;
    if (tabs.length || o.tabs) ppr += `<w:tabs>${(o.tabs || []).concat(tabs).join('')}</w:tabs>`;
    if (o.auto) ppr += '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
    else if (o.line) ppr += `<w:spacing w:before="0" w:after="0" w:line="${tw(o.line)}" w:lineRule="exact"/>`;
    if (o.center !== undefined) ppr += '<w:jc w:val="center"/>';
    return `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${body}</w:p>`;
  }

  // page estimates for the Word file (the PDF counts exactly). Heights in points, measured on real output.
  const P1_CAP = 750, P2_CAP = 742;
  function p1Pages(d) {
    const n = (d.notes || []).length, k = d.items.length;
    const h = 722 + 78 * (k - 1) + 27 * n - (n ? 13.5 : 0);
    return Math.max(1, Math.ceil(h / P1_CAP));
  }
  function metalHeight(d, cfg) {
    let h = 16.8 + 13.5;
    d.items.forEach(it => {
      const p = partFor(cfg, it);
      h += 27;
      if (!p) return;
      p.components.forEach(c => { h += 13.5 * (componentFields(c, 1).length + 1); });
      (p.images || []).forEach(im => { h += fitImg(im).h + 13.5 + 4; });
    });
    return h + 13.5;
  }
  const metalPage = d => p1Pages(d) + 1;
  function totalPages(d, cfg) { return p1Pages(d) + Math.max(1, Math.ceil(metalHeight(d, cfg) / P2_CAP)); }
  // fit a picture inside the printable width (and a sensible height), keeping its shape
  function fitImg(im, maxW = 540, maxH = 420) {
    let w = im.w || 300, h = im.h || 200;
    const k = Math.min(1, maxW / w, maxH / h);
    return { w: w * k, h: h * k };
  }
  let MEDIA = [];
  function buildBody(d, cfg, aci, email) {
    picId = 1; MEDIA = [];
    const RED = 'FF0000', BLUE = '0000FF';
    const P = [];
    const P1 = p1Pages(d), NP = totalPages(d, cfg);   // page labels like the original: "Page 1 of 2"
    const blank = () => P.push(para());
    const hr = (k = 'dashed') => { P.push(para([], { rule: k, line: 7 })); P.push(para([], { line: 6.5 })); };
    const cn = d.consignee, im = d.importer;

    P.push(para([], { center: '', auto: true, raw: pic('rLogo', 161, 36.7, 'Armodilo logo') }));
    P.push(para([], { center: `ACI-${aci}`, b: true, sz: 12, line: 16.8 }));
    P.push(para([], { center: 'Commercial Invoice', b: true, sz: 12, line: 16.8 }));
    const HD = headings(d, cfg);
    P.push(para([], { center: HD.p1, b: true, sz: 11, color: HD.black ? '000000' : RED, line: 15.4 }));
    const awb = awbOut(d, aci);
    P.push(para([[L, 'Invoice Number'], [95.5, 'Purpose of Shipment'], [302.7, 'Curr'], [355.5, 'Ult Dest.'], ...(d.has_awb ? [[AWB_X, 'Air Waybill Nbr.']] : [])]));
    P.push(para([[122.4, d.purpose], [300.7, d.currency], [355.1, d.ult_dest], ...(awb ? [[AWB_X, awb]] : [])]));
    P.push(para([[L, 'Export Dt'], [81.1, 'C.I. References'], [275.5, 'Pkgs'], [333.8, 'Bill T/C']]));
    P.push(para([[L, d.export_date], [86.6, d.ci_ref], [222.8, d.pkg_type], [291.1, d.pkgs], [332.8, d.bill_tc]]));
    hr();
    P.push(para([[L, 'Shipper:'], [302.4, 'Consignee:']]));
    const left = cfg.shipper.lines;
    const right = [cn.name, cn.company, ...cn.street, `${cn.city},  ${cn.state} ${cn.zip} ${cn.country}`, cn.phone];
    for (let k = 0; k < Math.max(left.length, right.length); k++) {
      const s = []; if (k < left.length) s.push([L, left[k]]); if (k < right.length) s.push([302.4, right[k]]); P.push(para(s));
    }
    P.push(para([[L, `ID/EIN: ${cfg.shipper.id_ein}`], [268.1, 'ID/EIN:']]));       // consignee ID/EIN always blank
    hr();
    P.push(para([[L, 'Broker:'], [266.4, 'Importer:']]));
    const bl = [...d.broker, cfg.broker_extra_line];
    const ir = [im.name, im.company, ...im.street, `${im.city},  ${im.state} ${im.zip} ${im.country}`];
    for (let k = 0; k < Math.max(bl.length, ir.length); k++) {
      const s = []; if (k < bl.length) s.push([L, bl[k]]); if (k < ir.length) s.push([266.4, ir[k]]); P.push(para(s));
    }
    blank();
    const impId = im.id || cn.id || 'n/a';                                          // FedEx consignee ID moves here
    P.push(para([], { center: [`ID/EIN: ${impId}`, email, im.phone || cn.phone].filter(Boolean).join('  '), b: true, sz: 11, line: 15.4 }));
    hr();
    for (const it of d.items) {
      P.push(para([[L, `Part Nbr:${it.part}`], [230.5, 'Marks/Nbrs:' + (it.marks ? ' ' + it.marks : '')], [328.8, `C/T MFR:${it.mfr}`], [413.7, 'Net Wgt:'], [474.8, it.net_wgt], [511.0, it.wgt_unit]]));
      P.push(para([[L, `HS Code: ${it.hs}`], ...qtySegs(it), [307.2, 'Unit Value:'], [402.6, `Commodity Value: ${it.commodity_value}`]]));
      P.push(para([[302.4, it.unit_value]]));
      P.push(para([[L, `Desc: ${it.desc}`]]));
      hr();
    }
    (d.notes || []).forEach(n => { P.push(para([[L, n]])); hr(); });
    P.push(para([[194.9, 'Total Shipment'], [297.1, d.total_weight], [363.8, 'Total Commodity'], [AMT_R, d.total_commodity, 'right']]));
    P.push(para([[230.4, 'Weight:'], [308.5, 'lbs'], [410.2, 'Value:']]));
    for (const [lab, val] of [['Terms of Sale:', d.terms], ['Freight:', d.freight], ['Insurance:', d.insurance], ['Others:', d.others]])
      P.push(para([[410, lab, 'right'], [AMT_R, val, 'right']]));
    P.push(para([[410, 'Total Invoice Value:', 'right'], [AMT_R, d.currency === 'USD' ? `US$${d.total_invoice}` : d.total_invoice, 'right']]));
    if (!(d.notes || []).length) blank();
    P.push(para([[L, 'Comments:']]));
    commentsFor(d, cfg).forEach(c => P.push(para([[L, c]])));
    if (d.tracking.length) P.push(para([[L, 'Tracking Numbers:'], [108, d.tracking.join(', ')]]));
    blank();
    P.push(para([[L, 'I declare all information in this invoice to be true and correct.']]));
    // Signature line: "Signature of shipper:" [signature] "SHIPPING@ARMODILO.COM" [tab] date
    P.push(para([], {
      auto: true, tabs: [`<w:tab w:val="left" w:pos="${tw(348 - L)}"/>`],
      raw: run('Signature of shipper: ') + pic('rSig', SIG.w, SIG.h, 'Signature') + run(' SHIPPING@ARMODILO.COM') + TAB + run(d.sig_date)
    }));
    blank(); hr('single'); blank(); P.push(para([[L, `Page ${P1} of ${NP}`]]));

    // page 2
    P.push(para([], { center: HD.p2, b: true, sz: 12, color: HD.black ? '000000' : RED, line: 16.8, pageBreak: true }));
    blank();
    d.items.forEach((it, n) => {
      const p = partFor(cfg, it);
      if (!p) { P.push(para([[L, `Line ${n + 1}: ${it.part} ${it.desc} -- METAL CONTENT TO BE ADDED`]], { b: true, color: RED })); blank(); return; }
      const q = parseFloat(it.qty) || 1;
      const [t1, t2] = splitParen(`Line ${n + 1}: ${p.title}`);
      P.push(para([[L, t1]], { b: true, color: p.title_color === 'black' ? '000000' : RED, raw: t2 ? run(' ' + t2, { b: true, color: BLUE }) : '' }));
      blank();
      if (!p.components.length && p.images && p.images.length) { /* plastic parts: title + picture only */ }
      for (const c of p.components) {
        componentFields(c, q).forEach(x => P.push(para([], {
          tabs: [`<w:tab w:val="left" w:pos="${tw(50.4 - L)}"/>`, `<w:tab w:val="left" w:pos="${tw(VALUE_X - L)}"/>`],
          raw: TAB + run(x.label + ':', { color: RED }) + TAB + run(x.value, { color: RED, b: x.bold }) + (x.note ? run('  ' + x.note, { color: BLUE }) : '')
        })));
        blank();
      }
      (p.images || []).forEach(im => {             // pictures from the Doc, centred
        const rid = 'rImg' + (MEDIA.length + 1), sz = fitImg(im);
        MEDIA.push({ rid, ext: /png/i.test(im.mime) ? 'png' : 'jpeg', data: im.data });
        P.push(para([], { center: '', auto: true, raw: pic(rid, sz.w, sz.h, 'Picture') }));
        blank();
      });
    });
    P.push(para([[L, `Page ${NP} of ${NP}`]]));
    return P.join('');
  }

  async function buildDocx(JSZip, d, cfg, aci, email, logoB64, sigB64) {
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const bodyXml = buildBody(withTotals(d), cfg, aci, email);
    const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<w:body>${bodyXml}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="320" w:right="288" w:bottom="400" w:left="288" w:header="0" w:footer="200" w:gutter="0"/></w:sectPr></w:body></w:document>`;
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>` +
      `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>` +
      `<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="270" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>`;
    const z = new JSZip();
    z.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `</Types>`);
    z.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
    z.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>` +
      `<Relationship Id="rSig" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/signature.png"/>` +
      MEDIA.map(m => `<Relationship Id="${m.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${m.rid}.${m.ext}"/>`).join('') + `</Relationships>`);
    MEDIA.forEach(m => z.file(`word/media/${m.rid}.${m.ext}`, m.data, { base64: true }));
    z.file('word/document.xml', doc);
    z.file('word/styles.xml', styles);
    z.file('word/media/logo.png', logoB64, { base64: true });
    z.file('word/media/signature.png', sigB64, { base64: true });
    return z.generateAsync({ type: typeof window === 'undefined' ? 'nodebuffer' : 'blob', compression: 'DEFLATE',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }


  // ---------- PDF (jsPDF; Helvetica has the same letter widths as Arial) ----------
  function buildPdf(jsPDF, d0, cfg, aci, email, logoB64, sigB64) {
    const d = withTotals(d0);
    const doc = new jsPDF({ unit: 'pt', format: 'letter' });
    const W = 612, PH = 792, BOTTOM = 760, RED = [255, 0, 0], BLUE = [0, 0, 255], BLACK = [0, 0, 0];
    let y = 0;
    const t = (x, top, s, o = {}) => {
      doc.setFont('helvetica', o.b ? 'bold' : 'normal'); doc.setFontSize(o.sz || 10); doc.setTextColor(...(o.color || BLACK));
      doc.text(String(s), x, top + (o.sz || 10) * 0.905, o.right ? { align: 'right' } : undefined);
    };
    const tc = (top, s, o = {}) => { doc.setFont('helvetica', 'bold'); doc.setFontSize(o.sz || 12); doc.setTextColor(...(o.color || BLACK)); doc.text(String(s), W / 2, top + (o.sz || 12) * 0.905, { align: 'center' }); };
    const width = (s, b, sz = 10) => { doc.setFont('helvetica', b ? 'bold' : 'normal'); doc.setFontSize(sz); return doc.getTextWidth(s); };
    const rule = (top, solid) => { doc.setDrawColor(0); doc.setLineWidth(0.6); doc.setLineDashPattern(solid ? [] : [3, 2], 0); doc.line(L, top + 6, W - L, top + 6); doc.setLineDashPattern([], 0); };
    const room = h => { if (y + h > BOTTOM) { doc.addPage(); y = 30; } };   // start a new page if this block won't fit

    doc.addImage('data:image/png;base64,' + logoB64, 'PNG', 225, 16, 161, 36.6);
    tc(60, `ACI-${aci}`); tc(76.5, 'Commercial Invoice');
    y = 107.6;
    const awb = awbOut(d, aci);
    t(L, y, 'Invoice Number'); t(95.5, y, 'Purpose of Shipment'); t(302.7, y, 'Curr'); t(355.5, y, 'Ult Dest.'); if (d.has_awb) t(AWB_X, y, 'Air Waybill Nbr.');
    y += LH; t(122.4, y, d.purpose); t(300.7, y, d.currency); t(355.1, y, d.ult_dest); if (awb) t(AWB_X, y, awb);
    y += LH; t(L, y, 'Export Dt'); t(81.1, y, 'C.I. References'); t(275.5, y, 'Pkgs'); t(333.8, y, 'Bill T/C');
    y += LH; t(L, y, d.export_date); t(86.6, y, d.ci_ref); t(222.8, y, d.pkg_type); t(291.1, y, d.pkgs); t(332.8, y, d.bill_tc);
    y += LH; rule(y); y += LH;
    const cn = d.consignee, im = d.importer;
    t(L, y, 'Shipper:'); t(302.4, y, 'Consignee:');
    const left = cfg.shipper.lines, right = [cn.name, cn.company, ...cn.street, `${cn.city},  ${cn.state} ${cn.zip} ${cn.country}`, cn.phone];
    for (let k = 0; k < Math.max(left.length, right.length); k++) { y += LH; if (left[k]) t(L, y, left[k]); if (right[k]) t(302.4, y, right[k]); }
    y += LH; t(L, y, `ID/EIN: ${cfg.shipper.id_ein}`); t(268.1, y, 'ID/EIN:');
    y += LH; rule(y); y += LH;
    t(L, y, 'Broker:'); t(266.4, y, 'Importer:');
    const bl = [...d.broker, cfg.broker_extra_line], ir = [im.name, im.company, ...im.street, `${im.city},  ${im.state} ${im.zip} ${im.country}`];
    for (let k = 0; k < Math.max(bl.length, ir.length); k++) { y += LH; if (bl[k]) t(L, y, bl[k]); if (ir[k]) t(266.4, y, ir[k]); }
    y += LH * 2;
    tc(y, [`ID/EIN: ${im.id || cn.id || 'n/a'}`, email, im.phone || cn.phone].filter(Boolean).join('  '), { sz: 11 });
    y += 15; rule(y); y += LH;
    for (const it of d.items) {
      room(LH * 5);
      t(L, y, `Part Nbr:${it.part}`); t(230.5, y, 'Marks/Nbrs:' + (it.marks ? ' ' + it.marks : '')); t(328.8, y, `C/T MFR:${it.mfr}`); t(413.7, y, 'Net Wgt:'); t(474.8, y, it.net_wgt); t(511, y, it.wgt_unit);
      y += LH; t(L, y, `HS Code: ${it.hs}`); qtySegs(it).forEach(([x, v]) => t(x, y, v)); t(307.2, y, 'Unit Value:'); t(402.6, y, `Commodity Value: ${it.commodity_value}`);
      y += LH; t(302.4, y, it.unit_value);
      y += LH; t(L, y, `Desc: ${it.desc}`);
      y += LH; rule(y); y += LH;
    }
    (d.notes || []).forEach(n => { room(LH * 2); t(L, y, n); y += LH; rule(y); y += LH; });
    room(LH * 8);
    t(194.9, y, 'Total Shipment'); t(297.1, y, d.total_weight); t(363.8, y, 'Total Commodity'); t(AMT_R, y, d.total_commodity, { right: true });
    y += LH; t(230.4, y, 'Weight:'); t(308.5, y, 'lbs'); t(410.2, y, 'Value:');
    for (const [lab, val] of [['Terms of Sale:', d.terms], ['Freight:', d.freight], ['Insurance:', d.insurance], ['Others:', d.others]]) { y += LH; t(410, y, lab, { right: true }); t(AMT_R, y, val, { right: true }); }
    y += LH; t(410, y, 'Total Invoice Value:', { right: true }); t(AMT_R, y, d.currency === 'USD' ? `US$${d.total_invoice}` : d.total_invoice, { right: true });
    y += (d.notes || []).length ? LH : LH * 2; room(LH * (commentsFor(d, cfg).length + 7) + 17);
    t(L, y, 'Comments:');
    commentsFor(d, cfg).forEach(c => { y += LH; t(L, y, c); });
    if (d.tracking.length) { y += LH; t(L, y, 'Tracking Numbers:'); t(108, y, d.tracking.join(', ')); }
    y += LH * 2; t(L, y, 'I declare all information in this invoice to be true and correct.');
    y += LH + 17;
    const lab = 'Signature of shipper: ', lw = width(lab);
    t(L, y, lab);
    doc.addImage('data:image/png;base64,' + sigB64, 'PNG', L + lw, y + 9 - SIG.h + 6, SIG.w, SIG.h);
    t(L + lw + SIG.w + 3, y, 'SHIPPING@ARMODILO.COM'); t(348, y, d.sig_date);
    y += LH * 2; if (y + LH * 3 > 780) { doc.addPage(); y = 30; } rule(y, true);   // closing rule + page label may sit a little lower
    const p1LabelY = y + LH * 2, p1Page = doc.getNumberOfPages();

    // metal content always starts on a new page
    doc.addPage(); const metalPage = doc.getNumberOfPages();
    const HD = headings(d, cfg);
    tc(16.5, HD.p2, { color: HD.black ? BLACK : RED });
    y = 49.1;
    d.items.forEach((it, n) => {
      const p = partFor(cfg, it);
      room(LH * 3);
      if (!p) { t(L, y, `Line ${n + 1}: ${it.part} ${it.desc} -- METAL CONTENT TO BE ADDED`, { b: true, color: RED }); y += LH * 2; return; }
      const q = parseFloat(it.qty) || 1;
      const [t1, t2] = splitParen(`Line ${n + 1}: ${p.title}`);
      t(L, y, t1, { b: true, color: p.title_color === 'black' ? BLACK : RED });
      if (t2) t(L + width(t1 + '  ', true), y, t2.trim(), { b: true, color: BLUE });
      y += LH * 2;
      for (const c of p.components) {
        const rows = componentFields(c, q);
        room(LH * rows.length);
        for (const x of rows) {
          t(50.4, y, x.label + ':', { color: RED });
          t(VALUE_X, y, x.value, { b: x.bold, color: RED });
          if (x.note) t(VALUE_X + width(x.value + '  ', x.bold), y, x.note, { color: BLUE });
          y += LH;
        }
        y += LH;
      }
      (p.images || []).forEach(im => {
        const sz = fitImg(im);
        room(sz.h + LH);
        doc.addImage(`data:${im.mime || 'image/jpeg'};base64,${im.data}`, /png/i.test(im.mime) ? 'PNG' : 'JPEG', (W - sz.w) / 2, y, sz.w, sz.h);
        y += sz.h + LH;
      });
    });
    // page labels like the original: under the invoice and under the metal content
    room(LH); const total = doc.getNumberOfPages();
    t(L, y, `Page ${total} of ${total}`);
    doc.setPage(p1Page); t(L, p1LabelY, `Page ${p1Page} of ${total}`);
    doc.setPage(1);
    tc(92.8, `Please see Page ${metalPage} for ${HD.fitKit ? 'Product Information' : 'Steel/Aluminum Content Breakdown'}`, { sz: 11, color: HD.black ? BLACK : RED });
    return doc.output('blob');
  }

  function warnings(d, cfg, email) {
    const w = [];
    if (!email) w.push(`No email on file for ${d.importer.company || 'this importer'}. Add one above if it should appear on the invoice.`);
    d.items.forEach((it, i) => {
      if (partFor(cfg, it)) return;
      const m = matchPart(cfg, it.part, it.desc);
      w.push(`Line ${i + 1} (${it.part}): ${it.metal_manual ? 'set to no metal content' : m.reason}. Choose it under \u201cProducts and metal content\u201d, or page 2 will say “METAL CONTENT TO BE ADDED”.`);
    });
    if (d.items.some(it => it.added)) w.push(`Totals now include the added product${d.items.filter(it => it.added).length > 1 ? 's' : ''} (Total Commodity US$${withTotals(d).total_commodity}). Total Shipment Weight is still FedEx\u2019s ${d.total_weight} lbs, so check it covers everything.`);
    d.items.forEach(it => { if (it.added && (!it.part || !(num(it.unit) > 0))) w.push(`An added line is missing its part number or unit value.`); });
    d.items.forEach(it => { if (!it.added && parseFloat(it.net_wgt) === 0) w.push(`Part ${it.part} shows a net weight of 0.00 ${it.wgt_unit} on the FedEx invoice.`); });
    return w;
  }

  root.ArmodiloCI = { extractLines, extractTextLines, parse, buildDocx, warnings, findPart, matchPart, partFor, componentRows, componentFields, VALUE_X, newLine, fillLine, withTotals, metalPage, p1Pages, totalPages, buildPdf, parseBreakdown, aciFromName, AMT_R, AWB_X, qtySegs, commentText, commentsBlank, commentsFor, fitKitOnly, headings, mergeBreakdown, fitImg };
  if (typeof module !== 'undefined') module.exports = root.ArmodiloCI;
})(typeof window !== 'undefined' ? window : globalThis);
