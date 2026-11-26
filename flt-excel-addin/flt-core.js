/*
 * flt-core.js - Logic tao file .flt (port tu generate_flt.py theo tai lieu ky thuat)
 * Khong phu thuoc Excel/DOM nen co the chay thu bang Node.
 */
(function (root) {
  "use strict";

  // ---------- Quy tac nhan dien cot (muc 2.2 cua tai lieu) ----------
  var IT_COL_NAMES = ["it code", "it. no", "it no", "it_no", "it_code", "itno", "itcode"];
  var IT_FALLBACK_COL_NAMES = ["no.", "no"];
  var GATE_COL_NAMES = [
    "cổng được phép quẹt vân tay",
    "máy được phép quẹt bằng vân tay",
    "máy được phép quẹt vân tay",
    "machine no.",
    "machine no",
    "machine_no"
  ];

  // ---------- Chuan hoa gia tri ----------
  function normHeader(v) {
    if (v === null || v === undefined) return "";
    return String(v).normalize("NFC").replace(/\s+/g, " ").trim().toLowerCase();
  }

  // Excel luu ma NV dang so thuc nguyen (59291.0) -> "59291"
  function cellToText(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "number") return Number.isInteger(v) ? String(v) : String(v);
    if (typeof v === "boolean") return "";
    return String(v).normalize("NFC").trim();
  }

  function isPlaceholder(s) {
    return /^column\d+$/i.test(s);
  }

  function findColumn(headers, names) {
    for (var i = 0; i < headers.length; i++) {
      if (names.indexOf(headers[i]) !== -1) return i;
    }
    return -1;
  }

  // ---------- Gom nhom theo tap cong/may (giong normalize_gate_key trong generate_flt.py) ----------
  // Khoa nhom gom ca tien to (cong/may) va tap hop cac phan tu da sap xep.
  function normalizeGateKey(text) {
    var s = String(text).normalize("NFC").replace(/\s+/g, " ").trim().toLowerCase();
    var m = s.match(/^(cổng|máy)\s+(.+)$/);
    var prefix = "", rest = s;
    if (m) { prefix = m[1]; rest = m[2]; }
    var parts = [];
    rest.split(",").forEach(function (p) {
      p = p.trim();
      if (prefix) p = p.replace(/^(cổng|máy)\s+/, "");
      if (p) parts.push(p);
    });
    parts.sort();
    return JSON.stringify([prefix, parts]);
  }

  // ---------- Dat ten file (giong sanitize_filename trong generate_flt.py) ----------
  function sanitizeFilename(name) {
    var s = String(name).trim().replace(/đ/g, "d").replace(/Đ/g, "D");
    s = s.normalize("NFD").replace(/\p{Mn}/gu, "");
    s = s.replace(/\s*,\s*/g, "_");
    s = s.replace(/[^A-Za-z0-9 _-]+/g, "_");
    s = s.replace(/_+/g, "_");
    s = s.replace(/\s+/g, " ");
    s = s.replace(/^[ _.]+|[ _.]+$/g, "");
    return s || "unknown_gate";
  }

  // ---------- Doc du lieu cac sheet ----------
  /*
   * sheetsData: [{ name: "New", rows: [[...], [...]] }]  (chi gom sheet dang hien thi)
   * Tra ve { groups, sheets, warnings }
   */
  function buildGroups(sheetsData) {
    var map = {};      // key -> group
    var order = [];    // giu thu tu xuat hien
    var sheets = [];
    var warnings = [];

    sheetsData.forEach(function (sh) {
      var rows = sh.rows || [];
      if (rows.length === 0) {
        sheets.push({ name: sh.name, used: false, note: "Sheet trống" });
        return;
      }
      var headers = rows[0].map(normHeader);
      var itIdx = findColumn(headers, IT_COL_NAMES);
      if (itIdx === -1) itIdx = findColumn(headers, IT_FALLBACK_COL_NAMES);
      var gateIdx = findColumn(headers, GATE_COL_NAMES);

      if (itIdx === -1 || gateIdx === -1) {
        var thieu = [];
        if (itIdx === -1) thieu.push("cột mã nhân viên");
        if (gateIdx === -1) thieu.push("cột cổng/máy");
        sheets.push({ name: sh.name, used: false, note: "Thiếu " + thieu.join(" và ") + " ở dòng 1" });
        return;
      }

      var count = 0;
      for (var r = 1; r < rows.length; r++) {
        var row = rows[r];
        var it = cellToText(row[itIdx]);
        var gate = cellToText(row[gateIdx]).replace(/\s+/g, " ");
        if (!it || !gate) continue;
        if (isPlaceholder(it) || isPlaceholder(gate)) continue;
        var key = normalizeGateKey(gate);
        var g = map[key];
        if (!g) {
          g = map[key] = { key: key, displayName: gate, codes: [], seen: {} };
          order.push(key);
        }
        if (!g.seen[it]) {
          g.seen[it] = true;
          g.codes.push(it);
        }
        count++;
      }
      sheets.push({ name: sh.name, used: true, note: count + " dòng hợp lệ" });
    });

    // Dat ten file, tranh trung ten sau khi chuan hoa
    var usedNames = {};
    var groups = order.map(function (k) {
      var g = map[k];
      var base = sanitizeFilename(g.displayName);
      var fname = base, n = 2;
      while (usedNames[fname.toLowerCase()]) {
        fname = base + "_" + n;
        n++;
      }
      usedNames[fname.toLowerCase()] = true;
      if (fname !== base) {
        warnings.push("Hai nhóm cổng/máy trùng tên file sau khi bỏ dấu, đã đổi thành " + fname + ".flt");
      }
      return { key: g.key, displayName: g.displayName, fileName: fname + ".flt", codes: g.codes };
    });

    return { groups: groups, sheets: sheets, warnings: warnings };
  }

  // ---------- Nhi phan .flt (muc 4) ----------
  function utf8(str) {
    return new TextEncoder().encode(str);
  }

  /* Tra ve { bytes, skipped } - ma NV dai hon 255 byte bi bo qua */
  function buildFltBytes(codes) {
    var enc = [];
    var skipped = [];
    codes.forEach(function (c) {
      var b = utf8(c);
      if (b.length > 255) skipped.push(c); else enc.push(b);
    });
    var total = 9;
    enc.forEach(function (b) { total += 2 + 4 + b.length + 5 + 1 + b.length; });
    var out = new Uint8Array(total);
    var dv = new DataView(out.buffer);
    var p = 0;
    out.set([0xFF, 0xFF, 0xFF, 0xFF, 0x01], p); p += 5;
    dv.setUint32(p, enc.length, true); p += 4;
    enc.forEach(function (b) {
      out[p++] = 0x00; out[p++] = 0x06;
      dv.setUint32(p, b.length, true); p += 4;
      out.set(b, p); p += b.length;
      out.set([0x01, 0x00, 0x00, 0x00, 0x06], p); p += 5;
      out[p++] = b.length;
      out.set(b, p); p += b.length;
    });
    return { bytes: out, skipped: skipped };
  }

  // ---------- ZIP (luu tru, khong nen) ----------
  var CRC_TABLE = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(data) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < data.length; i++) c = CRC_TABLE[(c ^ data[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function makeZip(files, now) {
    now = now || new Date();
    var dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    var dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    var locals = [], centrals = [], offset = 0;

    files.forEach(function (f) {
      var name = utf8(f.name);
      var crc = crc32(f.data);
      var size = f.data.length;

      var lh = new Uint8Array(30 + name.length);
      var l = new DataView(lh.buffer);
      l.setUint32(0, 0x04034b50, true);
      l.setUint16(4, 20, true);
      l.setUint16(6, 0x0800, true);
      l.setUint16(8, 0, true);
      l.setUint16(10, dosTime, true);
      l.setUint16(12, dosDate, true);
      l.setUint32(14, crc, true);
      l.setUint32(18, size, true);
      l.setUint32(22, size, true);
      l.setUint16(26, name.length, true);
      l.setUint16(28, 0, true);
      lh.set(name, 30);

      var ch = new Uint8Array(46 + name.length);
      var c = new DataView(ch.buffer);
      c.setUint32(0, 0x02014b50, true);
      c.setUint16(4, 20, true);
      c.setUint16(6, 20, true);
      c.setUint16(8, 0x0800, true);
      c.setUint16(10, 0, true);
      c.setUint16(12, dosTime, true);
      c.setUint16(14, dosDate, true);
      c.setUint32(16, crc, true);
      c.setUint32(20, size, true);
      c.setUint32(24, size, true);
      c.setUint16(28, name.length, true);
      c.setUint32(42, offset, true);
      ch.set(name, 46);

      locals.push(lh, f.data);
      centrals.push(ch);
      offset += lh.length + size;
    });

    var cdSize = 0;
    centrals.forEach(function (c) { cdSize += c.length; });
    var end = new Uint8Array(22);
    var e = new DataView(end.buffer);
    e.setUint32(0, 0x06054b50, true);
    e.setUint16(8, files.length, true);
    e.setUint16(10, files.length, true);
    e.setUint32(12, cdSize, true);
    e.setUint32(16, offset, true);

    var parts = locals.concat(centrals, [end]);
    var total = 0;
    parts.forEach(function (p) { total += p.length; });
    var out = new Uint8Array(total), pos = 0;
    parts.forEach(function (p) { out.set(p, pos); pos += p.length; });
    return out;
  }

  var api = {
    normalizeGateKey: normalizeGateKey,
    sanitizeFilename: sanitizeFilename,
    buildGroups: buildGroups,
    buildFltBytes: buildFltBytes,
    makeZip: makeZip,
    crc32: crc32
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.FltCore = api;
})(typeof self !== "undefined" ? self : this);
