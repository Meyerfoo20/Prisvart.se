/* Läser en fil från butiken (Excel .xlsx eller CSV) och gör om raderna till varor.
   Allt sker i webbläsaren, inga extra bibliotek behövs. */

export const CATS = ["Mejeri", "Kött & fisk", "Bröd", "Frukt & grönt", "Färdigmat"];
export const MAX_ROWS = 1000;

const pad = n => String(n).padStart(2, "0");
export const dayKey = d => d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
const plusDays = (base, n) => { const d = new Date(base); d.setDate(d.getDate() + n); return dayKey(d); };
const tags = (el, n) => el.getElementsByTagNameNS("*", n);

export function validEan(code) {
  if (!/^\d{8}$|^\d{12,14}$/.test(code)) return false;
  const digits = code.split("").map(Number), check = digits.pop();
  let sum = 0;
  digits.reverse().forEach((d, i) => sum += d * (i % 2 === 0 ? 3 : 1));
  return (10 - (sum % 10)) % 10 === check;
}

/* ---------- Läs filen ---------- */

export async function readTable(file) {
  const buf = await file.arrayBuffer();
  const head = new Uint8Array(buf.slice(0, 4));
  if (head[0] === 0x50 && head[1] === 0x4b) return readXlsx(buf);
  if (head[0] === 0xd0 && head[1] === 0xcf) throw new Error("Gamla .xls-filer kan inte läsas. Spara filen som .xlsx eller CSV i Excel och försök igen.");
  let text;
  try { text = new TextDecoder("utf-8", { fatal: true }).decode(buf); }
  catch (e) { text = new TextDecoder("windows-1252").decode(buf); }
  return readCsv(text);
}

function readCsv(text) {
  text = text.replace(/^﻿/, "");
  const sample = text.split(/\r?\n/).slice(0, 10).join("\n");
  const count = c => sample.split(c).length - 1;
  const sep = [";", "\t", ","].sort((a, b) => count(b) - count(a))[0];
  const rows = []; let row = [], cell = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; }
      else if (ch === '"') q = false;
      else cell += ch;
    } else if (ch === '"' && cell === "") q = true;
    else if (ch === sep) { row.push(cell); cell = ""; }
    else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(cell); rows.push(row); row = []; cell = "";
    } else cell += ch;
  }
  if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
  return rows;
}

async function inflate(data) {
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function unzip(buf, wanted) {
  const u8 = new Uint8Array(buf), dv = new DataView(buf), dec = new TextDecoder();
  let end = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { end = i; break; }
  if (end < 0) throw new Error("Filen gick inte att öppna. Spara den som .xlsx eller CSV och försök igen.");
  let off = dv.getUint32(end + 16, true); const n = dv.getUint16(end + 10, true), out = {};
  for (let k = 0; k < n && dv.getUint32(off, true) === 0x02014b50; k++) {
    const method = dv.getUint16(off + 10, true), size = dv.getUint32(off + 20, true);
    const nl = dv.getUint16(off + 28, true), el = dv.getUint16(off + 30, true), cl = dv.getUint16(off + 32, true), local = dv.getUint32(off + 42, true);
    const name = dec.decode(u8.subarray(off + 46, off + 46 + nl));
    off += 46 + nl + el + cl;
    if (!wanted(name)) continue;
    const start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
    const raw = u8.subarray(start, start + size);
    out[name] = dec.decode(method === 0 ? raw : await inflate(raw));
  }
  return out;
}

function colIndex(letters) { let n = 0; for (const ch of letters.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; }
function serialDate(n) {
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);
  return d.getUTCFullYear() + "-" + pad(d.getUTCMonth() + 1) + "-" + pad(d.getUTCDate());
}

async function readXlsx(buf) {
  if (typeof DecompressionStream === "undefined") throw new Error("Webbläsaren kan inte läsa Excel-filer. Uppdatera den, eller spara filen som CSV.");
  const z = await unzip(buf, n => /^xl\/(sharedStrings\.xml|workbook\.xml|styles\.xml|_rels\/workbook\.xml\.rels|worksheets\/[^/]+\.xml)$/.test(n));
  const P = s => new DOMParser().parseFromString(s, "application/xml");
  if (!z["xl/workbook.xml"]) throw new Error("Det här ser inte ut som en Excel-fil. Spara den som .xlsx eller CSV.");
  const shared = z["xl/sharedStrings.xml"] ? [...tags(P(z["xl/sharedStrings.xml"]), "si")].map(si => [...tags(si, "t")].map(t => t.textContent).join("")) : [];

  let path = null;
  try {
    const sheet = tags(P(z["xl/workbook.xml"]), "sheet")[0];
    const rid = sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || sheet.getAttribute("r:id");
    for (const r of tags(P(z["xl/_rels/workbook.xml.rels"]), "Relationship")) {
      if (r.getAttribute("Id") === rid) { const t = r.getAttribute("Target").replace(/^\/?(xl\/)?/, ""); path = "xl/" + t; }
    }
  } catch (e) {}
  if (!path || !z[path]) path = Object.keys(z).filter(n => /^xl\/worksheets\/[^/]+\.xml$/.test(n)).sort()[0];
  if (!path) throw new Error("Hittade inget kalkylblad i filen.");

  const dateStyles = new Set();
  try {
    const st = P(z["xl/styles.xml"]), custom = {};
    for (const f of tags(st, "numFmt")) custom[f.getAttribute("numFmtId")] = f.getAttribute("formatCode");
    [...tags(tags(st, "cellXfs")[0], "xf")].forEach((xf, i) => {
      const id = +xf.getAttribute("numFmtId"), code = (custom[id] || "").replace(/\[[^\]]*\]|"[^"]*"/g, "");
      if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47) || /[dy]/i.test(code)) dateStyles.add(i);
    });
  } catch (e) {}

  const rows = [];
  for (const r of tags(P(z[path]), "row")) {
    const row = [];
    for (const c of tags(r, "c")) {
      const ref = (c.getAttribute("r") || "").replace(/\d+/g, ""), t = c.getAttribute("t"), v = tags(c, "v")[0];
      let val = "";
      if (t === "s") val = v ? shared[+v.textContent] ?? "" : "";
      else if (t === "inlineStr") val = [...tags(c, "t")].map(x => x.textContent).join("");
      else if (t === "str" || t === "e") val = v ? v.textContent : "";
      else if (t === "b") val = v && v.textContent === "1" ? "1" : "0";
      else if (v) { const num = +v.textContent; val = dateStyles.has(+c.getAttribute("s")) && num > 1000 ? serialDate(num) : num; }
      row[ref ? colIndex(ref) : row.length] = val;
    }
    rows.push(row);
  }
  return rows;
}

/* ---------- Tolka kolumnerna ---------- */

function norm(h) { return String(h ?? "").toLowerCase().replace(/[._:()\[\]*]/g, " ").replace(/\s+/g, " ").trim(); }

function colType(raw) {
  const h = norm(raw);
  if (!h || h.length > 40) return null;
  if (/jmf|jämför/.test(h)) return null;
  if (/ean|streck|gtin|barcode|upc|plu|art(ikel)? ?nr|artikelnummer/.test(h)) return "ean";
  if ((/nytt|^ny |nedsatt|nedsätt|sänkt|rea|kampanj|erbjudande|new|sale|offer/.test(h) && /pris|price|kr/.test(h)) || /^(nytt|nedsatt|rea)$/.test(h)) return "now";
  if (/rabatt|procent|%|nedsättning|discount|avdrag/.test(h)) return "disc";
  if (/bäst|före|datum|utgång|förbruk|expir|best|date|^bf$|^bfd$|^sbd$/.test(h)) return "expiry";
  if (/antal|^st$|stycken|kvantitet|qty|quantity|lager|kvar|mängd|count/.test(h)) return "qty";
  if (/kategori|avdelning|category|varugrupp|grupp|department/.test(h)) return "cat";
  if (/pris|price|ordinarie|^ord|belopp|kr$/.test(h)) return "price";
  if (/vara|varor|namn|produkt|benämning|artikel|beskrivning|name|product|item|text|titel/.test(h)) return "name";
  return null;
}

export function findHeader(rows) {
  let best = null;
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const cols = {};
    (rows[i] || []).forEach((cell, j) => { const t = typeof cell === "string" ? colType(cell) : null; if (t && cols[t] == null) cols[t] = j; });
    const n = Object.keys(cols).length;
    if (n >= 2 && (cols.price != null || cols.now != null) && (cols.name != null || cols.ean != null) && (!best || n > best.n)) best = { index: i, cols, n };
  }
  return best;
}

export const COL_LABELS = { name: "Vara", ean: "Streckkod", price: "Ordinarie pris", now: "Nytt pris", disc: "Rabatt", expiry: "Bäst före", qty: "Antal", cat: "Kategori" };

/* ---------- Tolka värdena ---------- */

function num(v) {
  if (typeof v === "number") return v;
  let s = String(v ?? "").replace(/\s| |kr|sek|:-/gi, "");
  if (!s) return null;
  if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, "");
  s = s.replace(",", ".").replace(/[^\d.\-]/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function discount(v) {
  if (v == null || v === "") return null;
  const n = num(String(v).replace("%", ""));
  if (n == null) return null;
  return Math.round(n > 0 && n <= 1 ? n * 100 : n);
}
const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, maj: 5, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, okt: 10, oct: 10, nov: 11, dec: 12 };
function mk(y, m, d) {
  const dt = new Date(y, m - 1, d);
  return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d ? dayKey(dt) : null;
}
function soonest(m, d, today) {
  const y = today.getFullYear(), k = mk(y, m, d);
  if (!k) return null;
  return k < plusDays(today, -60) ? mk(y + 1, m, d) : k;
}
export function parseDate(v, today = new Date()) {
  if (v == null || v === "") return null;
  if (typeof v === "number") {
    if (v > 30000 && v < 80000) return serialDate(v);
    v = String(Math.round(v));
  }
  const s = String(v).trim().toLowerCase();
  let m;
  if (/^i ?dag$/.test(s)) return dayKey(today);
  if (/^i ?morgon$/.test(s)) return plusDays(today, 1);
  if (/^i ?övermorgon$/.test(s)) return plusDays(today, 2);
  if ((m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/))) return mk(+m[1], +m[2], +m[3]);
  if ((m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})$/))) return mk(m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2], +m[1]);
  if ((m = s.match(/^(\d{1,2})[/.](\d{1,2})$/))) return soonest(+m[2], +m[1], today);
  if ((m = s.match(/^(\d{4})(\d{2})(\d{2})$/))) return mk(+m[1], +m[2], +m[3]);
  if ((m = s.match(/^(\d{2})(\d{2})(\d{2})$/))) return mk(2000 + +m[1], +m[2], +m[3]);
  if ((m = s.match(/^\d{5}$/)) && +s > 30000 && +s < 80000) return serialDate(+s);
  if ((m = s.match(/^(\d{1,2})\s*([a-zåäö]{3})[a-zåäö]*\.?\s*(\d{4})?$/)) && MONTHS[m[2]]) return m[3] ? mk(+m[3], MONTHS[m[2]], +m[1]) : soonest(MONTHS[m[2]], +m[1], today);
  return null;
}

const CAT_WORDS = [
  ["Färdigmat", /färdig|sushi|wrap|smörgås|sallad med|pizza|lasagne|soppa|(^|\s)paj|matlåda|lunch|pirog|ready meal/],
  ["Kött & fisk", /kött|fisk|kyckling|färs|fläsk|(^|\s)nöt|lamm|lax|torsk|sej|räk|skaldjur|korv|skinka|bacon|kalkon|chark|meat|fish|poultry|sausage/],
  ["Mejeri", /mejeri|mjölk|yoghurt|yogurt|(^|\s)fil(\s|$)|filmjölk|(^|\s)ost(\s|$|,)|ost\s\d|grädde|smör(\s|$)|kvarg|crème|creme|keso|kesella|ägg|dairy|milk|cheese|butter|cream/],
  ["Bröd", /bröd|bulle|bullar|baguette|limpa|fralla|frallor|(^|\s)kaka|kakor|bageri|croissant|bread|bakery|(^|\s)bun/],
  ["Frukt & grönt", /frukt|grönt|grönsak|banan|äpple|päron|apelsin|citron|sallad|tomat|gurka|spenat|avokado|potatis|(^|\s)lök|morot|morötter|paprika|(^|\s)bär|druv|melon|svamp|ruccola|fruit|vegetable/]
];
export function guessCat(text) {
  const t = String(text || "").toLowerCase();
  for (const [cat, re] of CAT_WORDS) if (re.test(t)) return cat;
  return "Färdigmat";
}
function catFrom(cell, name) {
  const c = String(cell ?? "").trim();
  const exact = CATS.find(k => k.toLowerCase() === c.toLowerCase());
  if (exact) return exact;
  if (c) { const t = c.toLowerCase(); for (const [cat, re] of CAT_WORDS) if (re.test(t)) return cat; }
  return guessCat(name);
}
const r2 = n => Math.round(n * 100) / 100;

/* Gör om alla rader till varor. Varje vara får err om den inte kan publiceras. */
export function buildItems(rows, header, opts) {
  const today = opts.today || new Date(), todayKey = dayKey(today), c = header.cols, out = [];
  for (let i = header.index + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    if (!row.some(v => String(v ?? "").trim() !== "")) continue;
    const get = t => c[t] != null ? row[c[t]] : undefined;
    const name = String(get("name") ?? "").trim().replace(/\s+/g, " ").slice(0, 120);
    if (/^(summa|totalt|total)\b/i.test(name)) continue;
    const eanDigits = String(get("ean") ?? "").replace(/\D/g, "");
    const it = { line: i + 1, name, ean: eanDigits.length <= 14 ? eanDigits : "", cat: catFrom(get("cat"), name), catFromCell: !!String(get("cat") ?? "").trim(), img: "", err: "" };

    const price = num(get("price")), newPrice = num(get("now")), disc = discount(get("disc"));
    it.price = price != null ? r2(price) : null;
    if (newPrice > 0 && price > 0) { it.now = r2(newPrice); it.disc = Math.round((1 - newPrice / price) * 100); }
    else if (disc != null && price > 0) { it.disc = disc; it.now = r2(price * (100 - disc) / 100); }
    else if (price > 0) { it.disc = opts.defDisc; it.now = r2(price * (100 - opts.defDisc) / 100); it.defDisc = true; }

    const rawDate = get("expiry");
    if (rawDate == null || String(rawDate).trim() === "") { it.expiry = opts.defExpiry; it.defDate = true; }
    else it.expiry = parseDate(rawDate, today);

    const q = num(get("qty"));
    it.qty = q == null || String(get("qty")).trim() === "" ? 1 : Math.min(999, Math.floor(q));

    if (!(price > 0)) it.err = "Ordinarie pris saknas";
    else if (price > 10000) it.err = "Priset är högre än 10 000 kr";
    else if (!(it.now > 0) || it.now >= it.price) it.err = "Nya priset är inte lägre än ordinarie pris";
    else if (it.disc < 10 || it.disc > 90) it.err = "Rabatten måste vara mellan 10 och 90 %";
    else if (!it.expiry) it.err = "Bäst före-datumet gick inte att läsa";
    else if (it.expiry < todayKey) it.err = "Bäst före-datumet har passerat";
    else if (!(it.qty >= 1)) it.err = "Antal är 0";
    if (!it.name) { if (validEan(it.ean)) it.needName = true; else if (!it.err) it.err = "Namn saknas"; }
    out.push(it);
  }
  return out;
}

export function templateCsv(today = new Date()) {
  const rows = [
    ["Vara", "Streckkod", "Ordinarie pris", "Rabatt %", "Bäst före", "Antal", "Kategori"],
    ["Mellanmjölk 1,5 l", "", "21,90", "50", plusDays(today, 1), "6", "Mejeri"],
    ["Kycklingfilé 900 g", "", "119,00", "40", plusDays(today, 2), "3", "Kött & fisk"]
  ];
  return "﻿" + rows.map(r => r.join(";")).join("\r\n") + "\r\n";
}
