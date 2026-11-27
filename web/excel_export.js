/*
 * Export Excel: fills the real "PROP_Calculator(Square) .xlsx" (embedded in
 * excel_template.js) with the values from the page and returns a new .xlsx.
 *
 *  - both sheets are filled, each with its own tab's values
 *  - formulas are kept; cached results are removed and Excel is told to
 *    recalculate on open, so the tables are Excel's own result
 *  - pictures, shapes and print settings are untouched (XML-level edit)
 *  - the empty balloon circle next to the angle size gets its number (51)
 *  - title-block "link" fields left blank keep their Excel formula
 *
 * No libraries: reads the zip with the browser's DecompressionStream and writes
 * changed parts uncompressed (valid .xlsx, Excel opens it normally).
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.ExcelExport = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const SHEET_NAMES = { "SST-BODY": "PROP TYPE SST-BODY", "SST-LEG": "PROP TYPE SST-LEG" };

  // cell -> [input key, kind]   kind: num | text | auto (number if it looks like one) | link
  const CELL_INPUTS = {
    C18: ["top_of_stub_height", "num"], C25: ["conc_to_gl", "num"], C40: ["stub_below_gl", "num"],
    B32: ["leg_width", "num"], B31: ["leg_size_text_2", "text"], B30: ["leg_size_text_3", "text"],
    Y30: ["hole_text", "text"], N33: ["slotted_hole_text", "text"],
    W31: ["plate_dim_1", "num"], W32: ["plate_dim_2", "num"], T35: ["hole_gap", "num"],
    S37: ["angle_size_text", "text"], AD32: ["slope_vertical", "num"], AE37: ["slope_horizontal", "num"],
    F91: ["base_width_conc_lvl", "num"], AD40: ["tower_slope_label", "text"],
    L60: ["rev_date", "text"], P60: ["rev_no", "auto"], R60: ["rev_desc", "text"], AG60: ["rev_checked_2", "text"],
    L65: ["client", "text"], N71: ["project", "text"],
    N75: ["dsgn_name", "text"], N76: ["drwn_name", "text"], N77: ["chkd_name", "text"], N78: ["appd_name", "text"],
    P75: ["dsgn_date", "text"], P76: ["drwn_date", "text"],
    // blank = keep the Excel formula
    AA60: ["rev_designed", "link"], AC60: ["rev_checked", "link"], AE60: ["rev_drawn", "link"],
    AI60: ["rev_approved", "link"], P77: ["chkd_date", "link"], P78: ["appd_date", "link"],
  };
  const DOC_NO_CELLS = "N O P Q R S T U V W X Y Z AA AB AC AD AE AF AG".split(" ").map(c => c + "80");

  // ---------------- zip ----------------
  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function readZip(buf) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    let eocd = -1;
    for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error("Excel template is not a valid .xlsx");
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const dec = new TextDecoder();
    const entries = [];
    for (let k = 0; k < count; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error("Bad zip directory");
      const method = dv.getUint16(p + 10, true), time = dv.getUint16(p + 12, true), date = dv.getUint16(p + 14, true);
      const crc = dv.getUint32(p + 16, true), csize = dv.getUint32(p + 20, true), usize = dv.getUint32(p + 24, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      const name = dec.decode(buf.subarray(p + 46, p + 46 + nlen));
      const start = off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true);
      entries.push({ name, method, time, date, crc, csize, usize, raw: buf.subarray(start, start + csize) });
      p += 46 + nlen + xlen + clen;
    }
    return entries;
  }

  async function inflate(entry) {
    if (entry.method === 0) return entry.raw;
    if (entry.method !== 8) throw new Error("Unsupported zip compression in template");
    const ds = new DecompressionStream("deflate-raw");
    const out = await new Response(new Blob([entry.raw]).stream().pipeThrough(ds)).arrayBuffer();
    return new Uint8Array(out);
  }

  function writeZip(entries) {
    const enc = new TextEncoder();
    const parts = [], central = [];
    let offset = 0;
    for (const e of entries) {
      const name = enc.encode(e.name);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0, true);
      h.setUint16(8, e.method, true); h.setUint16(10, e.time, true); h.setUint16(12, e.date, true);
      h.setUint32(14, e.crc, true); h.setUint32(18, e.raw.length, true); h.setUint32(22, e.usize, true);
      h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      parts.push(new Uint8Array(h.buffer), name, e.raw);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0, true);
      c.setUint16(10, e.method, true); c.setUint16(12, e.time, true); c.setUint16(14, e.date, true);
      c.setUint32(16, e.crc, true); c.setUint32(20, e.raw.length, true); c.setUint32(24, e.usize, true);
      c.setUint16(28, name.length, true); c.setUint32(42, offset, true);
      central.push(new Uint8Array(c.buffer), name);
      offset += 30 + name.length + e.raw.length;
    }
    const cdSize = central.reduce((n, a) => n + a.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
    return new Blob([...parts, ...central, new Uint8Array(end.buffer)],
      { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  // ---------------- xml edits ----------------
  const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const looksNumber = v => /^-?\d+(\.\d+)?$/.test(String(v).trim());

  function setCell(xml, ref, value, asNumber) {
    const re = new RegExp(`<c r="${ref}"(\\s+s="\\d+")?[^>]*?(?:/>|>[\\s\\S]*?</c>)`);
    const m = xml.match(re);
    if (!m) throw new Error("Cell " + ref + " not found in template");
    const style = m[1] || "";
    let cell;
    if (value === "" || value == null) cell = `<c r="${ref}"${style}/>`;
    else if (asNumber) cell = `<c r="${ref}"${style}><v>${Number(value)}</v></c>`;
    else cell = `<c r="${ref}"${style} t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
    return xml.replace(re, cell);
  }

  function fillSheet(xml, inputs) {
    for (const [ref, [key, kind]] of Object.entries(CELL_INPUTS)) {
      const v = inputs[key] == null ? "" : String(inputs[key]);
      if (kind === "link") { if (v.trim()) xml = setCell(xml, ref, v, false); continue; }  // blank: keep formula
      const num = kind === "num" || (kind === "auto" && looksNumber(v));
      if (kind === "num" && !isFinite(Number(v))) throw new Error(key + " must be a number");
      xml = setCell(xml, ref, v, num);
    }
    const doc = String(inputs.doc_no || "");
    DOC_NO_CELLS.forEach((ref, i) => { const ch = doc[i] || ""; xml = setCell(xml, ref, ch, /^\d$/.test(ch)); });
    // drop cached formula results so Excel shows its own fresh result
    xml = xml.replace(/<c ([^>]*?)>(<f[^>]*(?:\/>|>[^<]*<\/f>))<v>[^<]*<\/v><\/c>/g,
      (all, attrs, f) => `<c ${attrs.replace(/\s+t="[^"]*"/, "")}>${f}</c>`);
    return xml;
  }

  function fillBalloon(xml, mark) {
    const i = xml.indexOf('prst="flowChartConnector"');
    if (i < 0 || !String(mark || "").trim()) return xml;
    const end = xml.indexOf("</xdr:sp>", i);
    let sp = xml.slice(i, end);
    sp = sp.replace(/<a:bodyPr[^>]*\/>/, '<a:bodyPr vertOverflow="overflow" horzOverflow="overflow" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>')
           .replace(/<a:p>[\s\S]*<\/a:p>/, `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-IN" sz="2800"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial Narrow"/></a:rPr><a:t>${esc(mark)}</a:t></a:r></a:p>`);
    return xml.slice(0, i) + sp + xml.slice(end);
  }

  function partOf(entries, name) { return entries.find(e => e.name === name); }

  // workbook sheet name -> sheet xml path and its drawing xml path
  function sheetParts(wb, wbRels, rels) {
    const out = {};
    for (const [key, name] of Object.entries(SHEET_NAMES)) {
      const rid = (wb.match(new RegExp(`<sheet [^>]*name="${name}"[^>]*r:id="([^"]+)"`)) || [])[1];
      const target = (wbRels.match(new RegExp(`<Relationship [^>]*Id="${rid}"[^>]*Target="([^"]+)"`)) ||
                      wbRels.match(new RegExp(`<Relationship [^>]*Target="([^"]+)"[^>]*Id="${rid}"`)) || [])[1];
      if (!target) throw new Error("Sheet not found in template: " + name);
      const sheet = "xl/" + target.replace(/^\/?xl\//, "");
      const relXml = rels(sheet.replace("worksheets/", "worksheets/_rels/") + ".rels") || "";
      const d = (relXml.match(/Target="[^"]*drawings\/([^"]+)"/) || [])[1];
      out[key] = { sheet, drawing: d ? "xl/drawings/" + d : null };
    }
    return out;
  }

  /** states: { "SST-BODY": inputs, "SST-LEG": inputs }  ->  Promise<Blob> */
  async function build(templateBase64, states) {
    const bin = atob(templateBase64);
    const buf = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
    let entries = readZip(buf).filter(e => e.name !== "xl/calcChain.xml");
    const dec = new TextDecoder(), enc = new TextEncoder();
    const text = async name => { const e = partOf(entries, name); return e ? dec.decode(await inflate(e)) : null; };
    const relsCache = {};
    for (const e of entries) if (e.name.endsWith(".rels")) relsCache[e.name] = await text(e.name);

    const edits = {};
    const wb = await text("xl/workbook.xml");
    const parts = sheetParts(wb, relsCache["xl/_rels/workbook.xml.rels"], n => relsCache[n]);
    for (const [key, p] of Object.entries(parts)) {
      if (!states[key]) continue;
      edits[p.sheet] = fillSheet(await text(p.sheet), states[key]);
      if (p.drawing) edits[p.drawing] = fillBalloon(await text(p.drawing), states[key].balloon_mark);
    }
    // recalculate everything when Excel opens the file
    edits["xl/workbook.xml"] = wb.replace(/<calcPr[^>]*\/>/, m => m.replace(/\s*fullCalcOnLoad="[^"]*"/, "").replace("/>", ' fullCalcOnLoad="1"/>'));
    edits["[Content_Types].xml"] = (await text("[Content_Types].xml")).replace(/<Override PartName="\/xl\/calcChain.xml"[^>]*\/>/, "");
    edits["xl/_rels/workbook.xml.rels"] = relsCache["xl/_rels/workbook.xml.rels"].replace(/<Relationship [^>]*calcChain[^>]*\/>/, "");

    entries = entries.map(e => {
      if (!(e.name in edits)) return e;
      const data = enc.encode(edits[e.name]);
      return { name: e.name, method: 0, time: e.time, date: e.date, crc: crc32(data), usize: data.length, raw: data };
    });
    return writeZip(entries);
  }

  return { build, CELL_INPUTS, DOC_NO_CELLS };
});
