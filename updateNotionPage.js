/***** CONFIG *****/
const NOTION_TOKEN   = '';      // secret_xxx from your integration
const DATABASE_ID    = '';       // 32-char (or with dashes) DB id
const UNIQUE_KEY_COL = 'Title';            // column used to upsert
const PAGE_ID_COL    = 'Notion Page ID';         // column where we store the page id
const NOTION_VER     = '2022-06-28';             // Notion API version
const RATE_DELAY_MS  = 350;          
const UPLOAD_THUMBS_AS_FILES = true;      // import thumbnail into Notion-managed file (best)

/***** SHEET → NOTION mapping (left = Sheet header, right = Notion property)
   "__TITLE__" is replaced by the DB's real Title prop name. *****/
const COLUMN_TO_PROPERTY = {
  'Title':        '__TITLE__',     // Title
  'STR':          'STR',           // Number (Percent)
  'APV':          'APV',           // Number (Percent)
  'AVD':          'AVD',           // Number (seconds)
  'Publish Date': 'Publish Date',  // Date
  'Video ID':     'Video ID',      // Rich text
  'Thumbnail':    'Thumbnail'      // Files & media
};

/***** MENU *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Notion Sync')
    .addItem('Sync Sheet → Notion', 'syncSheetToNotion')
    .addToUi();
}

/***** Ensure "Notion Page ID" column exists *****/
function ensurePageIdColumn(sheet, colName) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  if (!headers.includes(colName)) sheet.getRange(1, lastCol + 1).setValue(colName);
}

/***** URL & filename helpers *****/
function isValidImageUrl(u) {
  if (!u) return false;
  const s = String(u).trim();
  if (!/^https?:\/\//i.test(s)) return false;
  const base = s.split('?')[0];
  if (/\.(png|jpe?g|gif|webp)$/i.test(base)) return true;
  if (/\.ytimg\.com/i.test(s)) return true; // YouTube thumbs
  return false;
}
// Find an existing page by matching the Notion *title* property exactly
function findPageIdByUniqueKey(databaseId, titlePropName, value) {
  const payload = {
    filter: {
      property: titlePropName,
      title: { equals: String(value) }
    },
    page_size: 1
  };

  const res = UrlFetchApp.fetch(`https://api.notion.com/v1/databases/${databaseId}/query`, {
    method: 'post',
    headers: notionHeaders(),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code >= 300) {
    Logger.log(`findPageIdByUniqueKey failed (${code}): ${res.getContentText()}`);
    return '';
  }

  const data = JSON.parse(res.getContentText());
  return (data.results && data.results[0] && data.results[0].id) || '';
}

function parseImageFormula(formula) {
  // supports =IMAGE("url"...)
  const f = String(formula || '').trim();
  const m = f.match(/^=IMAGE\(\s*"([^"]+)"/i);
  return (m && m[1]) ? m[1] : '';
}
function ytThumbFromId(id) {
  const vid = String(id || '').trim();
  return vid ? `https://i.ytimg.com/vi/${vid}/hqdefault.jpg` : '';
}
function sanitizeFilename(name, fallbackExt) {
  const base = String(name || 'thumb').replace(/[^\w.-]+/g, '_').slice(0, 100);
  return base.endsWith(`.${fallbackExt}`) ? base : `${base}.${fallbackExt}`;
}
function fileExtFromUrl(u, dflt) {
  const m = String(u).split('?')[0].match(/\.(png|jpe?g|gif|webp)$/i);
  return m ? m[1].toLowerCase().replace('jpg', 'jpeg') : dflt;
}

/***** MAIN SYNC *****/
function syncSheetToNotion() {
  const sheet = SpreadsheetApp.getActiveSheet();
  ensurePageIdColumn(sheet, PAGE_ID_COL);

  const range    = sheet.getDataRange();
  const values   = range.getValues();
  const formulas = range.getFormulas(); // to read =IMAGE()
  if (values.length < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No data rows under the header');
    return;
  }

  const headers = values[0].map(h => (h || '').toString().trim());
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));
  if (!(UNIQUE_KEY_COL in idx)) throw new Error(`Missing required column: ${UNIQUE_KEY_COL}`);
  if (!(PAGE_ID_COL in idx))   throw new Error(`Missing required column: ${PAGE_ID_COL}`);

  const schema = getDatabaseSchema(DATABASE_ID);
  const titlePropName = Object.keys(schema).find(n => schema[n].type === 'title');

  let ok = 0, skipped = 0, failed = 0;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const fml = formulas[r];
    const key = String(row[idx[UNIQUE_KEY_COL]] || '').trim();
    if (!key) { skipped++; continue; }

    // Pick best thumbnail URL to use as cover candidate
    let coverUrl = '';
    let thumbUrls = [];
    if (idx['Thumbnail'] != null) {
      const raw = String(row[idx['Thumbnail']] || '').trim();
      if (raw) thumbUrls = raw.split(',').map(s => s.trim()).filter(Boolean);
      const fx  = fml[idx['Thumbnail']] || '';
      if (!thumbUrls.length && /^=IMAGE\(/i.test(fx)) {
        const u = parseImageFormula(fx);
        if (u) thumbUrls = [u];
      }
    }
    if (thumbUrls.length && isValidImageUrl(thumbUrls[0])) {
      coverUrl = thumbUrls[0];
    } else if (idx['Video ID'] != null) {
      const yt = ytThumbFromId(row[idx['Video ID']]);
      if (isValidImageUrl(yt)) coverUrl = yt;
    }

    let pageId = String(row[idx[PAGE_ID_COL]] || '').trim();
    const props = mapRowToNotionProperties(headers, row, fml, schema, titlePropName);

    try {
      if (!pageId) pageId = findPageIdByUniqueKey(DATABASE_ID, titlePropName, key);

      // CREATE (NO COVER) or UPDATE (NO COVER)
      if (pageId) {
        updatePageProps(pageId, props);
      } else {
        const res = createPageNoCover(DATABASE_ID, props);  // ← never sends cover on create
        pageId = res.id;
        sheet.getRange(r + 1, idx[PAGE_ID_COL] + 1).setValue(pageId);
      }

      // Set cover AFTER creation (safe) if we have a valid URL
      if (isValidImageUrl(coverUrl)) {
        try { patchCoverExternal(pageId, coverUrl); } catch (e) { Logger.log('patchCoverExternal: '+e.message); }
      }

      // Optional: import thumbnail into Notion as file + attach (cover + property)
      if (UPLOAD_THUMBS_AS_FILES && isValidImageUrl(coverUrl)) {
        const ext = fileExtFromUrl(coverUrl, 'jpg');
        const fname = sanitizeFilename('thumb-' + key, ext);
        const upId = startExternalUpload(coverUrl, fname);        // may be unsupported; safe to ignore
        if (upId && waitForUploadReady(upId, 6, 900)) {
          attachThumbAsFile(pageId, upId, fname, schema);
        }
      }

      ok++;
    } catch (e) {
      failed++;
      Logger.log(`Row ${r + 1} error: ${e.message}`);
      if (idx['Sync Error'] != null) sheet.getRange(r + 1, idx['Sync Error'] + 1).setValue(e.message);
    }

    Utilities.sleep(RATE_DELAY_MS);
  }

  const msg = `Notion sync → Success: ${ok}, Skipped: ${skipped}, Failed: ${failed}`;
  Logger.log(msg);
  SpreadsheetApp.getActiveSpreadsheet().toast(msg);
}

/***** NOTION HELPERS *****/
function notionHeaders() {
  return {
    'Authorization': `Bearer ${NOTION_TOKEN}`,
    'Notion-Version': NOTION_VER,
    'Content-Type': 'application/json'
  };
}
function getDatabaseSchema(databaseId) {
  const url = `https://api.notion.com/v1/databases/${databaseId}`;
  const res = UrlFetchApp.fetch(url, { method: 'get', headers: notionHeaders(), muteHttpExceptions: true });
  if (res.getResponseCode() >= 300) throw new Error(`Schema fetch failed: ${res.getContentText()}`);
  const data = JSON.parse(res.getContentText());
  const out = {};
  for (const [name, def] of Object.entries(data.properties)) out[name] = { type: def.type };
  return out;
}
function hhmmssToSeconds(s) {
  const txt = String(s).trim();
  if (!txt) return null;
  if (txt.includes(':')) {
    const parts = txt.split(':').map(n => Number(n));
    if (parts.length === 3) return parts[0]*3600 + parts[1]*60 + parts[2];
    if (parts.length === 2) return parts[0]*60 + parts[1];
  }
  const n = Number(txt);
  return isNaN(n) ? null : n;
}
function normalizeForProperty(propName, type, raw, formula) {
  if (raw === '' || raw == null) {
    if (propName === 'Thumbnail' && formula && /^=IMAGE\(/i.test(formula)) {
      const url = parseImageFormula(formula);
      return url ? [url] : null;
    }
    return null;
  }

  // Percent numbers (accept "44.5%", "44.5", or 0.445)
  if (propName === 'STR' || propName === 'APV') {
    const num = Number(String(raw).replace('%','').replace(',', '').trim());
    if (isNaN(num)) return null;
    return num > 1 ? num / 100 : num;  // Notion expects 0–1 for Percent
  }

  // AVD in seconds
  if (propName === 'AVD') return hhmmssToSeconds(raw);

  // Publish Date → YYYY-MM-DD
  if (propName.toLowerCase() === 'publish date') {
    const d = new Date(raw);
    if (isNaN(d)) return null;
    return d.toISOString().slice(0, 10);
  }

  // Files (Thumbnail): accept URLs in cell, comma-separated; or from =IMAGE()
  if (propName === 'Thumbnail') {
    const list = String(raw).split(',').map(s => s.trim()).filter(Boolean);
    if (formula && /^=IMAGE\(/i.test(formula)) {
      const u = parseImageFormula(formula);
      if (u) list.unshift(u);
    }
    const urls = list.filter(isValidImageUrl);
    return urls.length ? urls : null;
  }

  return raw;
}
function mapRowToNotionProperties(headers, row, rowFormulas, schema, titlePropName) {
  const props = {};
  headers.forEach((sheetCol, i) => {
    if (!(sheetCol in COLUMN_TO_PROPERTY)) return;  // only map listed columns
    let notionProp = COLUMN_TO_PROPERTY[sheetCol];
    if (notionProp === '__TITLE__') notionProp = titlePropName;
    if (!schema[notionProp]) return;

    const type = schema[notionProp].type;
    const val  = normalizeForProperty(notionProp, type, row[i], rowFormulas[i]);

    if (type === 'files') {
      const files = Array.isArray(val)
        ? val.map(u => ({ type: 'external', name: u.split('/').pop() || 'image', external: { url: u } }))
        : [];
      props[notionProp] = { files };
      return;
    }

    if (val == null) { props[notionProp] = { [type]: null }; return; }

    switch (type) {
      case 'title':
        props[notionProp] = { title: [{ text: { content: String(val) } }] };
        break;
      case 'number':
        props[notionProp] = { number: Number(val) };
        break;
      case 'url':
        props[notionProp] = { url: String(val) };
        break;
      case 'select':
        props[notionProp] = { select: { name: String(val) } };
        break;
      case 'multi_select':
        props[notionProp] = { multi_select: String(val).split(',').map(s => ({ name: s.trim() })).filter(x => x.name) };
        break;
      case 'checkbox':
        props[notionProp] = { checkbox: String(val).toLowerCase() === 'true' || String(val).toLowerCase() === 'yes' };
        break;
      case 'date':
        props[notionProp] = { date: { start: String(val) } };
        break;
      case 'rich_text':
        props[notionProp] = { rich_text: [{ text: { content: String(val) } }] };
        break;
      default:
        props[notionProp] = { rich_text: [{ text: { content: String(val) } }] };
    }
  });
  return props;
}

/***** Page create/update — never set cover on create *****/
function createPageNoCover(databaseId, properties) {
  const payload = { parent: { database_id: databaseId }, properties };
  const res = UrlFetchApp.fetch('https://api.notion.com/v1/pages', {
    method: 'post', headers: notionHeaders(),
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) throw new Error(`Create failed (${res.getResponseCode()}): ${res.getContentText()}`);
  return JSON.parse(res.getContentText());
}
function updatePageProps(pageId, properties) {
  const res = UrlFetchApp.fetch(`https://api.notion.com/v1/pages/${pageId}`, {
    method: 'patch', headers: notionHeaders(),
    payload: JSON.stringify({ properties }), muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) throw new Error(`Update failed (${res.getResponseCode()}): ${res.getContentText()}`);
  return JSON.parse(res.getContentText());
}
function patchCoverExternal(pageId, coverUrl) {
  if (!isValidImageUrl(coverUrl)) return;
  const payload = { cover: { type: 'external', external: { url: coverUrl } } };
  const res = UrlFetchApp.fetch(`https://api.notion.com/v1/pages/${pageId}`, {
    method: 'patch', headers: notionHeaders(),
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) Logger.log('patchCoverExternal failed: ' + res.getContentText());
}

/***** OPTIONAL: Import thumbnail into Notion as a FILE (with safe fallback) *****/
// 1) Start an external-url import; returns an uploadId (if supported)
function startExternalUpload(url, filename) {
  try {
    const res = UrlFetchApp.fetch('https://api.notion.com/v1/file_uploads', {
      method: 'post',
      headers: notionHeaders(),
      payload: JSON.stringify({ mode: 'external_url', external_url: url, filename }),
      muteHttpExceptions: true,
      contentType: 'application/json'
    });
    if (res.getResponseCode() !== 200) {
      Logger.log('startExternalUpload failed: ' + res.getResponseCode() + ' ' + res.getContentText());
      return '';
    }
    const data = JSON.parse(res.getContentText());
    return data.id || '';
  } catch (e) {
    Logger.log('startExternalUpload exception: ' + e.message);
    return '';
  }
}
// 2) Poll for the upload to be ready
function waitForUploadReady(uploadId, maxAttempts, waitMs) {
  for (let i = 0; i < (maxAttempts || 5); i++) {
    const res = UrlFetchApp.fetch(`https://api.notion.com/v1/file_uploads/${uploadId}`, {
      method: 'get', headers: notionHeaders(), muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      Logger.log('waitForUploadReady: ' + res.getResponseCode() + ' ' + res.getContentText());
      return false;
    }
    const data = JSON.parse(res.getContentText());
    if (data.status === 'uploaded') return true;
    if (data.status === 'failed') { Logger.log('Upload failed: ' + res.getContentText()); return false; }
    Utilities.sleep(waitMs || 800);
  }
  return false;
}
// 3) Attach uploaded file as cover + "Thumbnail" property
function attachThumbAsFile(pageId, uploadId, filename, schema) {
  try {
    const payload = {
      cover: { type: 'file_upload', file_upload: { id: uploadId } },
      properties: {}
    };
    if (schema['Thumbnail'] && schema['Thumbnail'].type === 'files') {
      payload.properties['Thumbnail'] = {
        type: 'files',
        files: [{ type: 'file_upload', name: filename, file_upload: { id: uploadId } }]
      };
    }
    const res = UrlFetchApp.fetch(`https://api.notion.com/v1/pages/${pageId}`, {
      method: 'patch', headers: notionHeaders(),
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      contentType: 'application/json'
    });
    if (res.getResponseCode() >= 300) {
      Logger.log('attachThumbAsFile failed: ' + res.getResponseCode() + ' ' + res.getContentText());
      return false;
    }
    return true;
  } catch (e) {
    Logger.log('attachThumbAsFile exception: ' + e.message);
    return false;
  }
}
