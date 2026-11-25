// Convert update_deck.html (1920x1080 <section class="pg"> slides) into an editable .pptx.
// Native shapes for cards/chips/stickers/rows, editable text boxes (Poppins), images for photos and
// report screenshots, and one background image per slide (gradient + blob + dots + squiggle + dashed lines).
//
// Usage: PW=<path to @playwright/test> PPTXGEN=<path to pptxgenjs> node html_to_pptx.cjs <deck.html> <out.pptx>
const path = require('path');
const fs = require('fs');
const { chromium } = require(process.env.PW || '@playwright/test');
const PptxGenJS = require(process.env.PPTXGEN || 'pptxgenjs');

const PX = 1 / 144;          // 1920px = 13.333in
const PT = 0.5;              // 1px = 0.5pt on this canvas
const [src, out] = process.argv.slice(2);

const EXTRACT = () => {
  const toHex = (c) => {
    const m = c && c.match(/rgba?\(([^)]+)\)/);
    if (!m) return null;
    const [r, g, b, a = '1'] = m[1].split(',').map((s) => s.trim());
    if (parseFloat(a) === 0) return null;
    return { hex: [r, g, b].map((v) => (+v).toString(16).padStart(2, '0')).join('').toUpperCase(), a: parseFloat(a) };
  };
  const sec = document.querySelector('section.pg.cur');
  const S = sec.getBoundingClientRect();
  const rel = (r) => ({ x: r.left - S.left, y: r.top - S.top, w: r.width, h: r.height });
  const opacityOf = (el) => { let o = 1; for (let e = el; e && e !== sec.parentElement; e = e.parentElement) o *= parseFloat(getComputedStyle(e).opacity); return o; };
  const isDecor = (el) => el.matches('.blob, .dots, svg, svg *, .abs');
  const isImage = (el) => el.matches('.photo, .shot, img.logo');
  const isBlockish = (el) => !['inline'].includes(getComputedStyle(el).display);
  const items = [];

  const runsOf = (block) => {
    const runs = [];
    const walk = (node, inherited) => {
      for (const ch of node.childNodes) {
        if (ch.nodeType === 3) {
          let t = ch.nodeValue.replace(/\s+/g, ' ');
          if (!t.trim() && !t.length) continue;
          const el = ch.parentElement, cs = getComputedStyle(el);
          if (cs.textTransform === 'uppercase') t = t.toUpperCase();
          const col = toHex(cs.color);
          runs.push({ text: t, size: parseFloat(cs.fontSize) , weight: parseInt(cs.fontWeight, 10), color: col ? col.hex : '000000',
                      alpha: (col ? col.a : 1) * opacityOf(el), ls: parseFloat(cs.letterSpacing) || 0, italic: cs.fontStyle === 'italic' });
        } else if (ch.nodeType === 1) {
          if (ch.tagName === 'BR') { runs.push({ br: true }); continue; }
          if (ch.tagName === 'STYLE' || isBlockish(ch) || isImage(ch)) continue;
          walk(ch, inherited);
        }
      }
    };
    walk(block);
    // trim leading/trailing whitespace of each line
    const lines = [[]];
    runs.forEach((r) => (r.br ? lines.push([]) : lines[lines.length - 1].push(r)));
    lines.forEach((l) => { if (l.length) { l[0].text = l[0].text.replace(/^\s+/, ''); l[l.length - 1].text = l[l.length - 1].text.replace(/\s+$/, ''); } });
    return lines.map((l) => l.filter((r) => r.text.length));
  };

  const hasDirectInline = (el) => [...el.childNodes].some((ch) =>
    (ch.nodeType === 3 && ch.nodeValue.trim()) || (ch.nodeType === 1 && !isBlockish(ch) && !isImage(ch) && ch.tagName !== 'BR' && ch.textContent.trim()));

  const textRect = (el) => {
    const range = document.createRange();
    const rects = [];
    for (const ch of el.childNodes) {
      if (ch.nodeType === 1 && (isBlockish(ch) || isImage(ch))) continue;
      range.selectNode(ch);
      for (const r of range.getClientRects()) if (r.width > 0 && r.height > 0) rects.push(r);
    }
    if (!rects.length) return null;
    const x0 = Math.min(...rects.map((r) => r.left)), y0 = Math.min(...rects.map((r) => r.top));
    const x1 = Math.max(...rects.map((r) => r.right)), y1 = Math.max(...rects.map((r) => r.bottom));
    return { left: x0, top: y0, width: x1 - x0, height: y1 - y0 };
  };

  const rotationOf = (cs) => {
    if (!cs.transform || cs.transform === 'none') return 0;
    const m = cs.transform.match(/matrix\(([^)]+)\)/);
    if (!m) return 0;
    const [a, b] = m[1].split(',').map(parseFloat);
    return Math.round(Math.atan2(b, a) * 180 / Math.PI);
  };

  const visit = (el, insideShapeText) => {
    if (isDecor(el)) return;
    const cs = getComputedStyle(el);
    if (cs.display === 'none' || cs.visibility === 'hidden') return;
    if (isImage(el)) { items.push({ kind: 'image', id: el.dataset.pid, ...rel(el.getBoundingClientRect()) }); return; }
    if (el.tagName === 'STYLE') return;

    const bg = toHex(cs.backgroundColor);
    const bw = ['Top', 'Right', 'Bottom', 'Left'].map((s) => parseFloat(cs['border' + s + 'Width']) || 0);
    const op = opacityOf(el);
    const rot = rotationOf(cs);
    const shapeText = el.matches('.chip, .sticker, .dot');
    if (el !== document.querySelector('section.pg.cur') && (bg || bw.some((w) => w > 0))) {
      let r;
      if (rot) { const R = el.getBoundingClientRect(); r = { x: R.left - S.left + R.width / 2 - el.offsetWidth / 2, y: R.top - S.top + R.height / 2 - el.offsetHeight / 2, w: el.offsetWidth, h: el.offsetHeight }; }
      else r = rel(el.getBoundingClientRect());
      const full = bw.every((w) => w > 0) && cs.borderTopStyle !== 'dashed';
      const radius = parseFloat(cs.borderTopLeftRadius) || 0;
      if (bg || full) {
        items.push({ kind: 'shape', ...r, rot,
          shape: cs.clipPath && cs.clipPath.startsWith('polygon') ? 'homePlate' : (radius >= Math.min(r.w, r.h) / 2 - 1 ? (Math.abs(r.w - r.h) < 2 ? 'ellipse' : 'pill') : radius > 0 ? 'round' : 'rect'),
          radius, fill: bg && bg.hex, fillA: bg ? bg.a * op : 0,
          line: full ? toHex(cs.borderTopColor) : null, lineW: full ? bw[0] : 0,
          shadow: cs.boxShadow && cs.boxShadow !== 'none',
          text: shapeText ? runsOf(el) : null,
          lh: parseFloat(cs.lineHeight) || null });
      }
      // single-side borders become lines (rows, table cells)
      if (!full) {
        const lineAt = (y, w, c) => items.push({ kind: 'line', x: r.x, y, w: r.w, lw: w, color: c.hex, alpha: c.a * op, dash: cs.borderTopStyle === 'dashed' });
        if (bw[0] > 0 && cs.borderTopStyle !== 'dashed') lineAt(r.y, bw[0], toHex(cs.borderTopColor) || { hex: '000000', a: 0 });
        if (bw[2] > 0) lineAt(r.y + r.h, bw[2], toHex(cs.borderBottomColor) || { hex: '000000', a: 0 });
      }
      if (shapeText) return;
    }
    // custom bullet dots (li::before)
    if (el.tagName === 'LI') {
      const b = getComputedStyle(el, '::before'), c = toHex(b.backgroundColor), R = el.getBoundingClientRect();
      if (c) items.push({ kind: 'shape', shape: 'ellipse', x: R.left - S.left + parseFloat(b.left), y: R.top - S.top + parseFloat(b.top),
                          w: parseFloat(b.width), h: parseFloat(b.height), fill: c.hex, fillA: c.a * op, rot: 0 });
    }
    if (isBlockish(el) && hasDirectInline(el)) {
      const tr = textRect(el), R = el.getBoundingClientRect();
      if (tr) {
        const padL = parseFloat(cs.paddingLeft) || 0, padR = parseFloat(cs.paddingRight) || 0;
        const contentL = R.left + padL + (parseFloat(cs.borderLeftWidth) || 0), contentW = R.width - padL - padR;
        const lh = parseFloat(cs.lineHeight) || parseFloat(cs.fontSize) * 1.2;
        items.push({ kind: 'text', x: Math.min(contentL, tr.left) - S.left, y: tr.top - S.top, w: Math.max(contentW, tr.width), h: tr.height,
                     tx: tr.left - S.left, tw: tr.width, lines: Math.round(tr.height / lh), lh,
                     align: cs.textAlign === 'right' ? 'right' : cs.textAlign === 'center' ? 'center' : 'left',
                     runs: runsOf(el), rot });
      }
    }
    for (const ch of el.children) visit(ch);
  };
  visit(sec);
  return items;
};

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1920, height: 1080 }, deviceScaleFactor: 2 });
  await page.goto('file://' + path.resolve(src), { waitUntil: 'load' });
  await page.evaluate(() => document.fonts.ready);
  const n = await page.evaluate(() => document.querySelectorAll('section.pg').length);
  await page.addStyleTag({ content: 'html,body{margin:0;background:transparent}' });

  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'HD', width: 13.333, height: 7.5 });
  pres.layout = 'HD';
  pres.title = path.basename(out, '.pptx');
  pres.theme = { headFontFace: 'Poppins', bodyFontFace: 'Poppins' };

  for (let i = 0; i < n; i++) {
    await page.evaluate((k) => {
      document.querySelectorAll('section.pg').forEach((s, j) => { s.style.display = j === k ? '' : 'none'; s.classList.toggle('cur', j === k); });
      let pid = 0; document.querySelectorAll('.photo, .shot, img.logo').forEach((e) => { e.dataset.pid = String(pid++); });
      document.getElementById('mode') && document.getElementById('mode').remove();
    }, i);
    await page.evaluate(() => window.scrollTo(0, 0));
    const items = await page.evaluate(EXTRACT);

    // background pass: only gradient + decor visible
    await page.addStyleTag({ content: 'section.pg.cur *{visibility:hidden!important} section.pg .blob, section.pg .dots, section.pg svg, section.pg svg *, section.pg .abs{visibility:visible!important}' }).then((h) => page.evaluate((e) => e.id = 'mode', h));
    const bg = await page.screenshot({ clip: { x: 0, y: 0, width: 1920, height: 1080 }, type: 'jpeg', quality: 90 });
    await page.evaluate(() => document.getElementById('mode').remove());

    // image pass: transparent page, only photos/screenshots/logo visible
    await page.addStyleTag({ content: 'section.pg{background:transparent!important} section.pg.cur *{visibility:hidden!important} section.pg .photo, section.pg .shot, section.pg .shot *, section.pg img.logo{visibility:visible!important}' }).then((h) => page.evaluate((e) => e.id = 'mode', h));
    const imgs = {};
    for (const it of items.filter((x) => x.kind === 'image')) {
      imgs[it.id] = await page.screenshot({ clip: { x: it.x, y: it.y, width: it.w, height: it.h }, omitBackground: true, type: 'png' });
    }
    await page.evaluate(() => document.getElementById('mode').remove());

    const slide = pres.addSlide();
    slide.background = { data: 'image/jpeg;base64,' + bg.toString('base64') };
    for (const it of items) {
      if (it.kind === 'shape') {
        const type = { rect: pres.ShapeType.rect, round: pres.ShapeType.roundRect, pill: pres.ShapeType.roundRect, ellipse: pres.ShapeType.ellipse, homePlate: pres.ShapeType.homePlate }[it.shape];
        const o = { x: it.x * PX, y: it.y * PX, w: it.w * PX, h: it.h * PX, rotate: it.rot || 0,
                    fill: it.fill ? { color: it.fill, transparency: Math.round((1 - it.fillA) * 100) } : { type: 'none' },
                    line: it.line ? { color: it.line.hex, width: it.lineW * PT } : { type: 'none' } };
        if (it.shape === 'round' || it.shape === 'pill') o.rectRadius = Math.min(it.radius, Math.min(it.w, it.h) / 2) * PX;
        if (it.shape === 'homePlate') o.adjustments = [36 / it.h * 100000 / 100000 * 0.5 * 2];
        // no PPTX shadows: some viewers render them as dark bands behind cards
        if (it.text && it.text.flat().length) {
          const runs = toRuns(it.text);
          slide.addText(runs, { ...o, shape: type, align: 'center', valign: 'middle', margin: 0, fontFace: 'Poppins', isTextBox: false,
                                lineSpacing: it.lh ? it.lh * PT : undefined });
        } else slide.addShape(type, o);
      } else if (it.kind === 'line') {
        slide.addShape(pres.ShapeType.line, { x: it.x * PX, y: it.y * PX, w: it.w * PX, h: 0,
          line: { color: it.color, width: Math.max(it.lw * PT, 0.5), transparency: Math.round((1 - it.alpha) * 100), dashType: it.dash ? 'dash' : 'solid' } });
      } else if (it.kind === 'image') {
        slide.addImage({ data: 'image/png;base64,' + imgs[it.id].toString('base64'), x: it.x * PX, y: it.y * PX, w: it.w * PX, h: it.h * PX });
      }
    }
    for (const it of items.filter((x) => x.kind === 'text')) {
      const runs = toRuns(it.runs);
      if (!runs.length) continue;
      // single-line text gets 12% slack so a substitute font does not wrap it
      let x = it.x, w = it.w;
      if (it.lines <= 1) {
        const need = it.tw * 1.12;
        if (need > w) { if (it.align === 'right') x = it.tx + it.tw - need; w = need; }
        if (it.align === 'left') x = Math.min(x, it.tx);
      }
      slide.addText(runs, { x: x * PX, y: it.y * PX, w: w * PX, h: Math.max(it.h, it.lh) * PX, margin: 0, valign: 'top', align: it.align,
                            fontFace: 'Poppins', isTextBox: true, wrap: true, lineSpacing: it.lh * PT, rotate: it.rot || 0 });
    }
  }

  function toRuns(lines) {
    const runs = [];
    lines.forEach((line, li) => {
      line.forEach((r, ri) => {
        runs.push({ text: r.text, options: { fontSize: +(r.size * PT).toFixed(1), bold: r.weight >= 600, italic: r.italic, color: r.color,
          transparency: Math.round((1 - r.alpha) * 100) || undefined, charSpacing: (r.ls > 0 || (r.ls < 0 && r.size >= 30)) ? +(r.ls * PT).toFixed(2) : undefined,
          breakLine: ri === line.length - 1 && li < lines.length - 1 } });
      });
    });
    return runs;
  }

  await pres.writeFile({ fileName: out });
  await browser.close();
  console.log('wrote', out, n, 'slides');
})();
