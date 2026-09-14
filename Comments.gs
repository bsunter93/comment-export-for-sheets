/**
 * Comment Export for Sheets: the actual work.
 *
 * Google's comments API will not tell you where a comment lives. Its anchor is
 * {"type":"workbook-range","uid":0,"range":"<opaque id>"} and that id resolves to nothing
 * reachable, so tab and cell are unrecoverable that way. The .xlsx export does carry them,
 * in xl/threadedComments, which is the route this takes: export, unzip, read.
 *
 * The part numbering is a trap worth knowing about. threadedComment2.xml is not the second
 * sheet. The mapping runs workbook.xml -> r:id -> workbook.xml.rels -> worksheets/sheetN.xml
 * -> that sheet's own rels -> the threadedComments part, and nothing else is reliable.
 */

var OUTPUT_TAB = 'Comment Export';
// Marker rather than prose, so the caller can re-ask for the grant instead of
// showing the user an error they cannot act on.
var NO_SCOPE = 'NO_FILE_SCOPE';
var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

var NS = {
  main:  XmlService.getNamespace('http://schemas.openxmlformats.org/spreadsheetml/2006/main'),
  r:     XmlService.getNamespace('http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
  pkg:   XmlService.getNamespace('http://schemas.openxmlformats.org/package/2006/relationships'),
  tc:    XmlService.getNamespace('http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments')
};

function exportCommentsToTab() {
  var ss = SpreadsheetApp.getActive();
  var parts = fetchArchive(ss.getId());
  var people = readPeople(parts);
  var threads = readThreads(parts, people);
  if (threads.length) writeTab(ss, threads);
  return { threads: threads.length };
}

/* ── fetch ──────────────────────────────────────────────────────────────── */

/** Returns the archive as { partName: text }. */
function fetchArchive(fileId) {
  var url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) +
            '/export?mimeType=' + encodeURIComponent(XLSX_MIME);
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  if (code !== 200) {
    var body = res.getContentText(), reason = '';
    try { reason = JSON.parse(body).error.errors[0].reason; } catch (err) {}

    // Drive answers 404, not 403, when drive.file has no grant on a file, deliberately,
    // so that it cannot be used to probe whether a file exists. Both mean "ask again".
    if (code === 404 || reason === 'notFound' || reason === 'appNotAuthorizedToFile') {
      throw new Error(NO_SCOPE);
    }
    if (reason === 'exportSizeLimitExceeded' || body.indexOf('fileSize') !== -1) {
      throw new Error('this spreadsheet is too large to export. Google caps it at 10MB.');
    }
    // Anything else is reported with its reason attached rather than reduced to a shrug.
    // An evening went into a silent failure that turned out to be one unenabled API.
    throw new Error('Drive ' + code + (reason ? ' (' + reason + ')' : ''));
  }

  var map = {};
  Utilities.unzip(res.getBlob().setContentType('application/zip')).forEach(function (p) {
    // Worksheet XML can be megabytes and is never read here, so it is skipped rather than
    // decoded; on a large workbook that is the difference between fast and timing out.
    var n = p.getName();
    if (n.indexOf('xl/worksheets/sheet') === 0 && n.indexOf('.rels') === -1) return;
    if (n.indexOf('xl/media/') === 0) return;
    map[n] = p.getDataAsString();
  });
  return map;
}

/* ── xml ────────────────────────────────────────────────────────────────── */

function parse(parts, name) {
  if (!parts[name]) return null;
  try {
    return XmlService.parse(parts[name]).getRootElement();
  } catch (e) {
    return null;
  }
}

/** personId -> display name. Absent for a file nobody has commented on. */
function readPeople(parts) {
  var out = {};
  var root = parse(parts, 'xl/persons/person.xml');
  if (!root) return out;
  root.getChildren('person', NS.tc).forEach(function (p) {
    out[p.getAttribute('id').getValue()] = p.getAttribute('displayName').getValue();
  });
  return out;
}

/** Relationship id -> target, resolved against the part's own directory. */
function readRels(parts, partName) {
  var root = parse(parts, relsPathFor(partName));
  var out = {};
  if (!root) return out;
  var base = dirOf(partName);
  root.getChildren('Relationship', NS.pkg).forEach(function (rel) {
    out[rel.getAttribute('Id').getValue()] = {
      type: rel.getAttribute('Type').getValue(),
      target: resolvePath(base, rel.getAttribute('Target').getValue())
    };
  });
  return out;
}

function relsPathFor(partName) {
  return dirOf(partName) + '_rels/' + partName.split('/').pop() + '.rels';
}

function dirOf(path) {
  var i = path.lastIndexOf('/');
  return i === -1 ? '' : path.slice(0, i + 1);
}

/** Targets are relative and routinely climb out of their directory with "../". */
function resolvePath(base, target) {
  if (target.charAt(0) === '/') return target.slice(1);
  var segs = (base + target).split('/');
  var out = [];
  segs.forEach(function (s) {
    if (s === '.' || s === '') return;
    if (s === '..') { out.pop(); return; }
    out.push(s);
  });
  return out.join('/');
}

/* ── threads ────────────────────────────────────────────────────────────── */

/**
 * Walks the workbook in tab order so the output reads the way the spreadsheet does,
 * then sorts within a tab by cell. A list that jumps between tabs is useless for
 * working through open comments, which is the whole job.
 */
function readThreads(parts, people) {
  var wb = parse(parts, 'xl/workbook.xml');
  if (!wb) throw new Error('the export did not contain a workbook');

  var wbRels = readRels(parts, 'xl/workbook.xml');
  var sheetsEl = wb.getChild('sheets', NS.main);
  if (!sheetsEl) return [];

  var rows = [];
  sheetsEl.getChildren('sheet', NS.main).forEach(function (sheet, order) {
    var name = sheet.getAttribute('name').getValue();
    var rId = sheet.getAttribute('id', NS.r);
    if (!rId) return;
    var rel = wbRels[rId.getValue()];
    if (!rel) return;

    // The sheet's own rels are the only honest link to its comments part.
    var sheetRels = readRels(parts, rel.target);
    var tcTarget = null;
    Object.keys(sheetRels).forEach(function (k) {
      if (sheetRels[k].target.indexOf('threadedComments/') !== -1) tcTarget = sheetRels[k].target;
    });
    if (!tcTarget) return;

    var root = parse(parts, tcTarget);
    if (!root) return;

    var comments = root.getChildren('threadedComment', NS.tc);
    var byId = {}, tops = [], replies = [];

    comments.forEach(function (c) {
      var item = {
        id:       attr(c, 'id'),
        parentId: attr(c, 'parentId'),
        ref:      attr(c, 'ref'),
        when:     attr(c, 'dT'),
        who:      people[attr(c, 'personId')] || '',
        done:     attr(c, 'done') === '1',
        text:     textOf(c)
      };
      byId[item.id] = item;
      (item.parentId ? replies : tops).push(item);
    });

    replies.forEach(function (r) {
      var parent = byId[r.parentId];
      if (!parent) { tops.push(r); return; }      // orphan: show it rather than drop it
      (parent.replies = parent.replies || []).push(r);
    });

    tops.sort(function (a, b) { return cellSort(a.ref, b.ref); });
    tops.forEach(function (t) {
      rows.push({ tab: name, order: order, thread: t });
    });
  });

  return rows;
}

function attr(el, name) {
  var a = el.getAttribute(name);
  return a ? a.getValue() : '';
}

function textOf(el) {
  var t = el.getChild('text', NS.tc);
  return t ? t.getText() : '';
}

/** Reading order: down the rows, then across the columns. */
function cellSort(a, b) {
  var pa = cellParts(a), pb = cellParts(b);
  return (pa.row - pb.row) || (pa.col - pb.col);
}

function cellParts(ref) {
  var m = /^([A-Z]+)(\d+)$/.exec(String(ref || '').toUpperCase());
  if (!m) return { col: 1e9, row: 1e9 };
  var col = 0;
  for (var i = 0; i < m[1].length; i++) col = col * 26 + (m[1].charCodeAt(i) - 64);
  return { col: col, row: parseInt(m[2], 10) };
}

/* ── output ─────────────────────────────────────────────────────────────── */

var HEADERS = ['Sheet', 'Cell', 'Status', 'Author', 'Date', 'Comment', 'Replies', 'Reply thread'];

function writeTab(ss, rows) {
  var sh = ss.getSheetByName(OUTPUT_TAB);
  if (sh) {
    sh.clear();
    if (sh.getFilter()) sh.getFilter().remove();
  } else {
    sh = ss.insertSheet(OUTPUT_TAB);
  }

  var data = [HEADERS];
  rows.forEach(function (r) {
    var t = r.thread;
    var reps = t.replies || [];
    data.push([
      r.tab,
      t.ref,
      t.done ? 'Resolved' : 'Open',
      t.who,
      toDate(t.when),
      t.text,
      reps.length,
      reps.map(function (x) { return (x.who || 'Unknown') + ': ' + x.text; }).join('\n')
    ]);
  });

  sh.getRange(1, 1, data.length, HEADERS.length).setValues(data);

  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, data.length, HEADERS.length).createFilter();
  sh.getRange(2, 5, Math.max(data.length - 1, 1), 1).setNumberFormat('yyyy-mm-dd hh:mm');
  // Comment text is long and the default row height would hide most of it.
  sh.getRange(1, 6, data.length, 1).setWrap(true);
  sh.getRange(1, 8, data.length, 1).setWrap(true);
  [110, 60, 80, 120, 130, 340, 70, 340].forEach(function (w, i) {
    sh.setColumnWidth(i + 1, w);
  });

  ss.setActiveSheet(sh);
}

/** dT is "2026-09-13T22:25:20.00", which Date parses as UTC. */
function toDate(s) {
  if (!s) return '';
  var d = new Date(/Z$/.test(s) ? s : s + 'Z');
  return isNaN(d.getTime()) ? s : d;
}
