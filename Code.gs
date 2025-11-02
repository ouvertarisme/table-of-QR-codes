/** 
 * Google Docs → QR Code Index from Footnote URLs
 * Specs: sequential hex refs (01..FF), title + URL, QR code image.
 * Menu: QR Tools → Build QR Table
 */

function onOpen() {
  DocumentApp.getUi()
    .createMenu('QR Tools')
    .addItem('Build QR Table', 'buildQrTable')
    .addToUi();
}

function buildQrTable() {
  const doc  = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const urls = extractUrlsFromAllFootnotes_(body);
  if (urls.length === 0) {
    DocumentApp.getUi().alert('No URLs found in footnotes.');
    return;
  }

  // Find placeholder "QRCodeTable"
  const search = body.findText('QRCodeTable');
  if (!search) {
    DocumentApp.getUi().alert('Placeholder "QRCodeTable" not found. Aborting.');
    return;
  }

  const elem = search.getElement();
  const startOffset = search.getStartOffset();
  const endOffset   = search.getEndOffsetInclusive();

  // Remove the placeholder text
  elem.asText().deleteText(startOffset, endOffset);

  // Collect entries
  const entries = urls.map((u, i) => {
    const ref   = toHex2_(i + 1);
    const title = fetchPageTitle_(u) || 'Title not available — please edit';
    return { ref, url: u, title };
  });

  // Insert the table in place of placeholder
  const parent = elem.getParent();
  const idx    = parent.getChildIndex(elem);

  const table = body.insertTable(body.getChildIndex(parent) + 1);
  const header = table.appendTableRow();
  header.appendTableCell('Référence');
  header.appendTableCell('Source et son URL');
  header.appendTableCell('QR Code');
  makeHeaderBold_(header);

  entries.forEach((e) => {
    const row = table.appendTableRow();

    row.appendTableCell(e.ref);

    const cell2 = row.appendTableCell('');
    cell2.appendParagraph(e.title);
    cell2.appendParagraph(e.url);

    const cell3 = row.appendTableCell('');
    try {
      const blob = buildQrBlob_(e.url, 600).setName(`qr_${e.ref}.png`);
      const img = cell3.insertImage(0, blob);
      img.setWidth(90).setHeight(90);
    } catch (err) {
      cell3.appendParagraph('QR failed');
    }
  });
}


/**
 * Returns the start index of the first table found AFTER the bookmark.
 */
function getTableStartIndexAfterBookmark_(docId, bookmarkId) {
  const d = Docs.Documents.get(docId);
  const bm = d.bookmarks && d.bookmarks[bookmarkId];
  if (!bm || !bm.position || bm.position.index == null) return null;

  const anchorIndex = bm.position.index;
  const content = d.body && d.body.content ? d.body.content : [];
  for (const el of content) {
    if (el.startIndex != null && el.table && el.startIndex > anchorIndex) {
      return el.startIndex; // table start
    }
  }
  return null;
}

/**
 * Merge cells and apply vertical-middle alignment via Docs API.
 * Each entry = two rows (top & bottom). Table has 3 columns.
 */
function applyMergesAndCellStyles_(docId, tableStartIndex, entryCount) {
  const requests = [];

  // Helpers to define table ranges
  const tableCellRange = (rowStart, rowEnd, colStart, colEnd) => ({
    tableRange: {
      tableCellLocation: { tableStartLocation: { index: tableStartIndex }, rowIndex: rowStart, columnIndex: colStart },
      rowSpan: rowEnd - rowStart + 1,
      columnSpan: colEnd - colStart + 1
    }
  });

  for (let i = 0; i < entryCount; i++) {
    const top = i * 2;       // top sub-row
    const bot = top + 1;     // bottom sub-row

    // a) Merge Ref vertically: (rows top..bot, col 0)
    requests.push({
      mergeTableCells: tableCellRange(top, bot, 0, 0)
    });

    // b) Merge URL horizontally on bottom row: (row bot, cols 1..2)
    requests.push({
      mergeTableCells: tableCellRange(bot, bot, 1, 2)
    });

    // c) Set vertical alignment MIDDLE for:
    //    - Ref merged cell (top..bot, col 0)
    //    - Title (top, col 1)
    //    - QR (top, col 2)
    const middleTargets = [
      tableCellRange(top, bot, 0, 0), // merged ref cell
      tableCellRange(top, top, 1, 1), // title cell
      tableCellRange(top, top, 2, 2)  // qr cell
    ];

    middleTargets.forEach(range => {
      requests.push({
        updateTableCellStyle: {
          tableRange: range.tableRange,
          tableCellStyle: { contentAlignment: 'MIDDLE' },
          fields: 'contentAlignment'
        }
      });
    });
  }

  Docs.Documents.batchUpdate({ requests }, docId);
}

/* ===================== Helpers ===================== */

/**
 * Traverse all footnotes and extract unique URLs in-order.
 */
function extractUrlsFromAllFootnotes_(body) {
  const found = [];
  const seen = new Set();

  let cursor = null;
  while (true) {
    const hit = body.findElement(DocumentApp.ElementType.FOOTNOTE, cursor);
    if (!hit) break;
    cursor = hit;
    const footnote = hit.getElement().asFootnote();
    const section = footnote.getFootnoteContents();
    const urlsInThis = extractUrlsFromContainer_(section);
    urlsInThis.forEach((u) => {
      if (!seen.has(u)) {
        seen.add(u);
        found.push(u);
      }
    });
  }
  return found;
}

/**
 * Recursively extract URLs from a container (FootnoteSection, Paragraph, etc.).
 * Looks for both linked text and plain-text URLs.
 */
function extractUrlsFromContainer_(container) {
  const urls = [];

  // Depth-first walk
  const stack = [container];
  while (stack.length) {
    const el = stack.pop();
    const type = el.getType();

    // If the element can have children, push them
    if (typeof el.getNumChildren === 'function') {
      for (let i = el.getNumChildren() - 1; i >= 0; i--) {
        stack.push(el.getChild(i));
      }
    }

    // Collect URLs from TEXT nodes
    if (type === DocumentApp.ElementType.TEXT) {
      const t = el.asText();
      // 1) Hyperlink attributes
      const idx = t.getTextAttributeIndices();
      idx.forEach((start) => {
        const link = t.getLinkUrl(start);
        if (link && isLikelyUrl_(link)) urls.push(normalizeUrl_(link));
      });

      // 2) Plain-text URLs via regex
      const s = t.getText();
      const regex = /https?:\/\/[^\s<>()\[\]{}"']+/g;
      let m;
      while ((m = regex.exec(s)) !== null) {
        const link = m[0];
        if (isLikelyUrl_(link)) urls.push(normalizeUrl_(link));
      }
    }
  }

  return urls;
}

function isLikelyUrl_(u) {
  return /^https?:\/\//i.test(u);
}

function normalizeUrl_(u) {
  try {
    // strip surrounding whitespace and normalize
    return new URL(u.trim()).toString();
  } catch (e) {
    return u.trim();
  }
}

function toHex2_(n) {
  // Clamp to 1..255, then format as 2-digit uppercase hex
  const x = Math.max(1, Math.min(255, n));
  return x.toString(16).toUpperCase().padStart(2, '0');
}

/**
 * Best-effort page title fetcher.
 */
function fetchPageTitle_(url) {
  try {
    const resp = UrlFetchApp.fetch(url, {
      followRedirects: true,
      muteHttpExceptions: true,
      validateHttpsCertificates: true
    });
    const code = resp.getResponseCode();
    if (code < 200 || code >= 400) return null;

    // Convert to string. If site returns bytes in another charset, Apps Script
    // still guesses reasonably well; otherwise we gracefully fail to null.
    const html = resp.getContentText();
    if (!html) return null;

    const m = /<title[^>]*>([\s\S]*?)<\/title>/i.exec(html);
    if (!m) return null;

    // Clean whitespace and entities
    let title = m[1]
      .replace(/\s+/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .trim();

    // Basic entity decode for common entities
    title = title
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&#39;|&apos;/g, "'")
      .replace(/&quot;/g, '"');

    return title || null;
  } catch (e) {
    return null;
  }
}

/**
 * Build a QR code image blob for the URL using Google Charts.
 * Note: the Charts Image API is unofficial/legacy but widely used and stable.
 * If you prefer another endpoint, swap the baseUrl below.
 */
function buildQrBlob_(url, sizePx) {
  const endpoints = [
    (u, s) => `https://quickchart.io/qr?text=${encodeURIComponent(u)}&size=${s}&margin=0`,
    (u, s) => `https://api.qrserver.com/v1/create-qr-code/?data=${encodeURIComponent(u)}&size=${s}x${s}&margin=0`,
    (u, s) => `https://zxing.org/w/chart?cht=qr&chs=${s}x${s}&chld=M&chl=${encodeURIComponent(u)}`
  ];

  const opts = {
    muteHttpExceptions: true,
    followRedirects: true,
    validateHttpsCertificates: true,
    headers: { 'Accept': 'image/png,*/*;q=0.8' }
  };

  for (let i = 0; i < endpoints.length; i++) {
    const qrUrl = endpoints[i](url, sizePx);
    try {
      const resp = UrlFetchApp.fetch(qrUrl, opts);
      const code = resp.getResponseCode();
      if (code >= 200 && code < 400) {
        const blob = resp.getBlob();
        const ct = (blob.getContentType() || '').toLowerCase();
        if ((ct.includes('image') || ct.includes('png')) && blob.getBytes().length > 0) {
          blob.setName('qr.png');
          return blob;
        }
      }
    } catch (e) {
      // try next endpoint
    }
    // Small pause helps with rate limits on large batches
    Utilities.sleep(60);
  }
  throw new Error('QR generation failed on all endpoints.');
}


function makeHeaderBold_(row) {
  for (let i = 0; i < row.getNumCells(); i++) {
    const p = row.getCell(i).getChild(0).asParagraph();
    const t = p.editAsText();
    t.setBold(true);
  }
}