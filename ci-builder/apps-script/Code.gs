/**
 * Armodilo CI Builder: sends the "Material Breakdown" Google Doc to the invoice page.
 *
 * Setup (once): open the Doc > Extensions > Apps Script > paste this file > Deploy > New deployment >
 * type "Web app", Execute as "Me", Who has access "Anyone" > Deploy > copy the Web app URL (ends in /exec)
 * into "breakdown_url" in config.js.
 *
 * Each time the page loads it asks this script for the Doc's current contents: every paragraph's text
 * with its bold/colour runs, plus any pictures. The page does the reading and checking itself.
 */
const DOC_ID = '1X4W7Iz6D3-hpE3uHp5QoEGu7Td_3EHj8cSt3WysZljw';

function doGet() {
  try {
    const body = DocumentApp.openById(DOC_ID).getBody();
    const paragraphs = [];
    for (let i = 0; i < body.getNumChildren(); i++) {
      const el = body.getChild(i);
      const type = el.getType();
      if (type !== DocumentApp.ElementType.PARAGRAPH && type !== DocumentApp.ElementType.LIST_ITEM) continue;
      const p = type === DocumentApp.ElementType.PARAGRAPH ? el.asParagraph() : el.asListItem();
      paragraphs.push({ runs: runsOf(p), img: imagesOf(p) });
    }
    return json({ fetchedAt: new Date().toISOString(), paragraphs });
  } catch (e) {
    return json({ error: String(e) });
  }
}

// text split into pieces that share the same bold + colour
function runsOf(p) {
  const text = p.editAsText();
  const s = text.getText();
  if (!s) return [];
  const starts = text.getTextAttributeIndices();
  const runs = [];
  for (let k = 0; k < starts.length; k++) {
    const a = starts[k], b = k + 1 < starts.length ? starts[k + 1] : s.length;
    if (b <= a) continue;
    runs.push({ t: s.substring(a, b), b: !!text.isBold(a), c: (text.getForegroundColor(a) || '#000000').toLowerCase() });
  }
  return runs;
}

// pictures in the paragraph (inline and "wrap text" ones); sizes sent in points
function imagesOf(p) {
  const out = [];
  const add = (blob, w, h) => out.push({ mime: blob.getContentType(), w: Math.round(w * 0.75 * 10) / 10, h: Math.round(h * 0.75 * 10) / 10, data: Utilities.base64Encode(blob.getBytes()) });
  for (let i = 0; i < p.getNumChildren(); i++) {
    const c = p.getChild(i);
    if (c.getType() === DocumentApp.ElementType.INLINE_IMAGE) { const im = c.asInlineImage(); add(im.getBlob(), im.getWidth(), im.getHeight()); }
  }
  (p.getPositionedImages() || []).forEach(im => add(im.getBlob(), im.getWidth(), im.getHeight()));
  return out;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
