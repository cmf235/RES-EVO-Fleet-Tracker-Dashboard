/**
 * EVO Fleet Tracker — file links (cert packets, rig up/pumpdown drawings) + pad info.
 * Restored from project history (Version 24). Feeds attachSheetExtras_() to the dashboard.
 */

function attachSheetExtras_(fleet, sheet) {
  try {
    var sh = sheet || fleetSheetFor_(fleet.num);
    if (!sh) return fleet;
    var vals = sh.getDataRange().getValues();

    fleet.files = filesFromValues_(vals);

    var pi = padInfoFromValues_(vals);
    fleet.rigUp   = fleet.rigUp   || {};
    fleet.rigDown = fleet.rigDown || {};
    ['address', 'mapsLink', 'gps'].forEach(function (k) {
      if (pi.rigUp[k])   fleet.rigUp[k]   = pi.rigUp[k];
      if (pi.rigDown[k]) fleet.rigDown[k] = pi.rigDown[k];
    });
  } catch (e) {
    // Never break the feed over these extras.
  }
  return fleet;
}

/** Find the tab whose name is "FLEET # <num> - ...". */
function fleetSheetFor_(num) {
  var want = String(num).replace(/[^0-9]/g, '');       // "#6" / "Fleet 6" -> "6"
  if (!want) return null;
  var shs = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < shs.length; i++) {
    var m = shs[i].getName().match(/FLEET\s*#\s*(\d+)/i);
    if (m && m[1] === want) return shs[i];
  }
  return null;
}

/* ── Upload URLs -> { drawing, dwg, pumpdrawing, cert } ─────────────────
 * Scans each row: captures a link only when the row has BOTH a known
 * label AND an http URL, so service rows ("PUMPDOWN IRON") are ignored. */
function filesFromValues_(vals) {
  var out = {};
  for (var r = 0; r < vals.length; r++) {
    var row = vals[r], kind = '', url = '';
    for (var c = 0; c < row.length; c++) {
      var s = row[c] == null ? '' : String(row[c]).trim();
      if (!s) continue;
      if (!kind) {
        if (/rig\s*up.*draw/i.test(s))   kind = 'drawing';
        else if (/pump\s*down/i.test(s)) kind = 'pumpdrawing';
        else if (/\bdwg\b/i.test(s))     kind = 'dwg';
        else if (/cert/i.test(s))        kind = (s.toLowerCase().indexOf('dwn') !== -1 || s.toLowerCase().indexOf('down') !== -1) ? 'certDown' : 'cert';
      }
      if (!url && /^https?:\/\//i.test(s)) url = s;
    }

      if (kind && url && !out[kind]) {
      var inline = (kind === 'drawing' || kind === 'pumpdrawing');  // images preview inline
      out[kind] = { url: directLink_(url, inline) };
    }
  }
  return out;
}

/* ── Rig Up/Down Pad Info -> { rigUp:{address,mapsLink,gps}, rigDown:{…} }
 * Tracks which "… Pad Info" header we're under, then reads the
 * Address / Maps Link / GPS rows beneath it. */
function padInfoFromValues_(vals) {
  var out = { rigUp: {}, rigDown: {} };
  var section = '';
  for (var r = 0; r < vals.length; r++) {
    var row = vals[r];
    var joined = row.map(function (c) { return c == null ? '' : String(c); }).join(' ');

    // Some tabs label the rig-up block "Rig Up Pad Info", others just "Pad Info".
    // Check "Rig Down Pad Info" FIRST, so the generic "Pad Info" match can safely
    // catch the rig-up block (both strings contain the words "Pad Info").
    if (joined.toLowerCase().indexOf('dwn pad info') !== -1 || joined.toLowerCase().indexOf('down pad info') !== -1) { section = 'rigDown'; continue; }
    else if (joined.toLowerCase().indexOf('pad info') !== -1) { section = 'rigUp'; continue; }
    if (!section) continue;

    var key = '', value = '';
    for (var c = 0; c < row.length; c++) {
      var s = row[c] == null ? '' : String(row[c]).trim();
      if (!s) continue;
      if (!key) {
        if (/^address$/i.test(s))       key = 'address';
        else if (/maps?\s*link/i.test(s)) key = 'mapsLink';
        else if (/^gps$/i.test(s))      key = 'gps';
        // when this cell IS the label, keep scanning for the value cell
      } else if (!isPadLabel_(s)) {
        value = s; break;                 // first non-label cell = the value
      }
    }
    if (key && value && !out[section][key]) {
      out[section][key] = (key === 'mapsLink') ? directLink_(value, false) : value;
    }
  }
  return out;
}

function isPadLabel_(s) {
  return /^address$/i.test(s) || /maps?\s*link/i.test(s) || /^gps$/i.test(s) || /pad\s*info/i.test(s);
}

/**
 * Turn a pasted Dropbox / Google Drive share link into a direct link.
 * inline=true  -> renders in an <img> preview (images: raw / thumbnail)
 * inline=false -> forces a download (DWG, .zip) / leaves other links usable
 */
function directLink_(raw, inline) {
  var u = String(raw || '').trim();
  if (!u) return '';

  // Google Drive: pull the file id out of any common share-link shape.
  var m = u.match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?[^]*?id=|thumbnail\?[^]*?id=)([-\w]{20,})/)
        || u.match(/[?&]id=([-\w]{20,})/);
  if (m) {
    var id = m[1];
    return inline
      ? 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1600'
      : 'https://drive.google.com/uc?export=download&id=' + id;
  }

  // Dropbox: keep ?rlkey=… (required), just flip the dl/raw flag.
  if (/dropbox\.com/i.test(u)) {
    u = u.replace(/([?&])(?:dl|raw)=\d+/gi, '$1');   // drop any dl= / raw=
    u = u.replace(/[?&]+$/, '');                     // trim trailing ? or &
    u = u.replace(/\?&/, '?').replace(/&&+/g, '&');  // tidy separators
    var sep = u.indexOf('?') >= 0 ? '&' : '?';
    return u + sep + (inline ? 'raw=1' : 'dl=1');
  }

  // Anything else (e.g. a maps.app.goo.gl link) is already a usable link.
  return u;
}

/* Optional: if your script loops over each fleet's TAB and you'd rather
 * pass the Sheet in directly, use these instead of attachSheetExtras_:
 *     fleet.files = filesFromValues_(sheet.getDataRange().getValues());
 * (kept for convenience; not required.) */
function filesForFleet_(num) {
  var sh = fleetSheetFor_(num);
  return sh ? filesFromValues_(sh.getDataRange().getValues()) : {};
}
