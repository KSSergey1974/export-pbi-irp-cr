// xlsx.js — книга Excel (XLSX) из пакета выгрузки. Собирается в браузере без библиотек:
// части книги SpreadsheetML (XML) упаковываются в ZIP; сжатие — встроенный CompressionStream
// ("deflate-raw": Chrome/Edge 103+, Firefox 113+, Safari 16.4+), без него — ZIP без сжатия.
//
// Лист «Таблица»: все строки с настоящими типами (числа, проценты, даты), форматы и заливки как в PDF,
// закреплённая строка заголовков с фильтром, печать — альбомный лист по ширине с повтором заголовка.
// Лист «Сводка»: показатели, матрица статусов, виды работ, отбор (без диаграммы).

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const NS_R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD_FILL = "#E3DACF";
const RULE = "FFCFC7BB";
const DAY0 = 25569;                                  // серийный номер Excel для 01.01.1970
const UNIT_LABEL = { "": "", k: " тыс.", m: " млн", b: " млрд" };
const UNIT_SCALE = { "": "", k: ",", m: ",,", b: ",,," };   // запятые в конце формата делят на 1000
const H_ALIGN = { l: "left", c: "center", r: "right" };

const xesc = (s) => String(s)
  .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, "")
  .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** Буквенное имя столбца: 0 → A, 25 → Z, 26 → AA. */
export function colName(i) {
  let s = "";
  for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function argb(hex) {
  const h = String(hex || "").replace("#", "").trim();
  return /^[0-9a-f]{6}$/i.test(h) ? "FF" + h.toUpperCase() : null;
}

// ---------------------------------------------------------------- стили

class Styles {
  constructor() {
    this.fmtIds = new Map([["General", 0], ["0", 1], ["#,##0", 3], ["0%", 9], ["0.00%", 10]]);
    this.customFmts = [];
    this.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    this.fillIds = new Map();
    this.fonts = [
      '<font><sz val="10"/><name val="Arial"/><family val="2"/></font>',                         // 0 обычный
      '<font><b/><sz val="10"/><name val="Arial"/><family val="2"/></font>',                     // 1 жирный
      '<font><b/><sz val="14"/><name val="Arial"/><family val="2"/></font>',                     // 2 заголовок книги
      '<font><sz val="9"/><color rgb="FF666666"/><name val="Arial"/><family val="2"/></font>'    // 3 подпись
    ];
    this.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    this.xfIds = new Map();
  }

  fmt(code) {
    let id = this.fmtIds.get(code);
    if (id === undefined) {
      id = 164 + this.customFmts.length;
      this.customFmts.push(`<numFmt numFmtId="${id}" formatCode="${xesc(code)}"/>`);
      this.fmtIds.set(code, id);
    }
    return id;
  }

  fill(hex) {
    const c = argb(hex);
    if (!c) return 0;
    let id = this.fillIds.get(c);
    if (id === undefined) {
      id = this.fills.length;
      this.fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="${c}"/><bgColor indexed="64"/></patternFill></fill>`);
      this.fillIds.set(c, id);
    }
    return id;
  }

  /** Индекс формата ячейки для сочетания числового формата, шрифта, заливки, рамки и выравнивания. */
  xf({ fmt = "General", font = 0, fill = "", border = true, h = "", v = "center", wrap = false } = {}) {
    const key = `${fmt}|${font}|${fill}|${border}|${h}|${v}|${wrap}`;
    let id = this.xfIds.get(key);
    if (id === undefined) {
      const al = `<alignment${h ? ` horizontal="${h}"` : ""}${v ? ` vertical="${v}"` : ""}${wrap ? ' wrapText="1"' : ""}/>`;
      id = this.xfs.length;
      this.xfs.push(`<xf numFmtId="${this.fmt(fmt)}" fontId="${font}" fillId="${this.fill(fill)}" borderId="${border ? 1 : 0}" xfId="0"` +
        ` applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${al}</xf>`);
      this.xfIds.set(key, id);
    }
    return id;
  }

  xml() {
    return XML + `<styleSheet ${NS}>` +
      (this.customFmts.length ? `<numFmts count="${this.customFmts.length}">${this.customFmts.join("")}</numFmts>` : "") +
      `<fonts count="${this.fonts.length}">${this.fonts.join("")}</fonts>` +
      `<fills count="${this.fills.length}">${this.fills.join("")}</fills>` +
      `<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>` +
      `<border><left style="thin"><color rgb="${RULE}"/></left><right style="thin"><color rgb="${RULE}"/></right>` +
      `<top style="thin"><color rgb="${RULE}"/></top><bottom style="thin"><color rgb="${RULE}"/></bottom><diagonal/></border></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="${this.xfs.length}">${this.xfs.join("")}</cellXfs>` +
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
      `</styleSheet>`;
  }
}

/** Числовой формат Excel для столбца — те же знаки, единицы и суффикс, что в PDF. */
export function numFmtFor(c) {
  if (c.k === "p") return c.dec > 0 ? `0.${"0".repeat(c.dec)}%` : "0%";
  if (c.k === "d") return "dd.mm.yyyy";
  if (c.k === "n") {
    const base = c.dec > 0 ? `#,##0.${"0".repeat(c.dec)}` : "#,##0";
    const unit = UNIT_LABEL[c.u] || "";
    const label = (unit + (c.sx ? (unit ? c.sx.replace(/^\s*/, " ") : c.sx) : "")).replace(/"/g, "");
    return base + (UNIT_SCALE[c.u] || "") + (label ? `"${label}"` : "");
  }
  return "General";
}

// ---------------------------------------------------------------- ячейки и листы

function strCell(ref, s, text) {
  const t = xesc(text);
  const keep = /^\s|\s$/.test(t) ? ' xml:space="preserve"' : "";
  return `<c r="${ref}" s="${s}" t="inlineStr"><is><t${keep}>${t}</t></is></c>`;
}

function numCell(ref, s, v) {
  return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
}

function emptyCell(ref, s) {
  return `<c r="${ref}" s="${s}"/>`;
}

function worksheet({ dim, cols, rows, frozen, filter, merges, landscape, selected }) {
  const view = `<sheetViews><sheetView workbookViewId="0"${selected ? ' tabSelected="1"' : ""}>` +
    (frozen ? '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>' : "") +
    `</sheetView></sheetViews>`;
  return XML + `<worksheet ${NS} ${NS_R}>` +
    `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>` +
    `<dimension ref="${dim}"/>` + view +
    `<sheetFormatPr defaultRowHeight="12.75"/>` +
    `<cols>${cols.map((w, j) => `<col min="${j + 1}" max="${j + 1}" width="${w.toFixed(2)}" customWidth="1"/>`).join("")}</cols>` +
    `<sheetData>${rows.join("")}</sheetData>` +
    (filter ? `<autoFilter ref="${filter}"/>` : "") +
    (merges && merges.length ? `<mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>` : "") +
    `<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>` +
    `<pageSetup paperSize="9" orientation="${landscape ? "landscape" : "portrait"}" fitToWidth="1" fitToHeight="0"/>` +
    `</worksheet>`;
}

/** Ширина столбца Excel (в знаках) по длине значений и слов заголовка. */
function excelWidth(c, texts) {
  let max = 0;
  for (const t of texts) if (t.length > max) max = t.length;
  const head = Math.max(4, ...c.h.split(/\s+/).map((w) => w.length));
  const content = c.wr ? Math.min(max, 50) : max;
  return Math.min(Math.max(content, head) * 1.1 + 2, 62);
}

function tableSheet(m, st) {
  const cols = m.cols;
  const n = m.rows;
  const last = colName(cols.length - 1);
  const head = st.xf({ font: 1, fill: HEAD_FILL, h: "center", v: "center", wrap: true });
  const fmts = cols.map(numFmtFor);
  const rows = [
    `<row r="1" ht="30" customHeight="1">${cols.map((c, j) => strCell(colName(j) + "1", head, c.h)).join("")}</row>`
  ];
  const names = cols.map((c, j) => colName(j));
  for (let i = 0; i < n; i++) {
    const r = i + 2;
    const rowColor = m.rowColors[i] || "";
    let x = `<row r="${r}">`;
    for (let j = 0; j < cols.length; j++) {
      const c = cols[j];
      const v = m.values[j][i];
      const text = m.texts[j][i];
      const fill = c.fm === "row" ? rowColor : c.fm === "fixed" && text !== "" ? c.fc : "";
      const typed = typeof v === "number" && isFinite(v) && c.k !== "t" && c.k !== "b";
      const s = st.xf({ fmt: typed ? fmts[j] : "General", fill, h: H_ALIGN[m.aligns[j]] || "", wrap: c.wr });
      const ref = names[j] + r;
      if (v === null || v === undefined || text === "") x += emptyCell(ref, s);
      else if (typed) x += numCell(ref, s, c.k === "d" ? v + DAY0 : v);
      else x += strCell(ref, s, text);
    }
    rows.push(x + "</row>");
  }
  return {
    lastRef: `$A$1:$${last}$${n + 1}`,
    xml: worksheet({
      dim: `A1:${last}${n + 1}`,
      cols: cols.map((c, j) => excelWidth(c, m.texts[j])),
      rows, frozen: true, filter: `A1:${last}${n + 1}`, landscape: true, selected: true
    })
  };
}

const statusLabel = (s) => String(s || "").replace(/^\s*\d+\s*-\s*/, "");
const asNumber = (v) => (/^\s*-?\d+(?:[.,]\d+)?\s*$/.test(String(v)) ? Number(String(v).replace(",", ".")) : null);

function summarySheet(m, st) {
  const hdr = m.hdr;
  const periods = hdr ? hdr.p : [];
  const width = Math.max(5, 1 + 2 * (periods.length + 1));      // столбцов: статус + пары МКД/ВР (не меньше A:E)
  const rows = [];
  const merges = [];
  let r = 0;
  const row = (cells, ht) => { r++; rows.push(`<row r="${r}"${ht ? ` ht="${ht}" customHeight="1"` : ""}>${cells(r)}</row>`); };
  const blank = () => { r++; };
  /** Объединённая ячейка строки q со столбца c0 по c1: значение в первой, стиль (рамки) — во всех. */
  const span = (q, c0, c1, s, value) => {
    merges.push(`${colName(c0)}${q}:${colName(c1)}${q}`);
    let x = value === "" ? emptyCell(`${colName(c0)}${q}`, s) : strCell(`${colName(c0)}${q}`, s, value);
    for (let c = c0 + 1; c <= c1; c++) x += emptyCell(`${colName(c)}${q}`, s);
    return x;
  };

  const title = st.xf({ font: 2, border: false, v: "" });
  const muted = st.xf({ font: 3, border: false, v: "" });
  const plain = st.xf({ border: false, v: "" });
  const bold = st.xf({ font: 1, border: false, v: "" });
  const band = st.xf({ font: 1, fill: HEAD_FILL });
  const bandRight = st.xf({ font: 1, fill: HEAD_FILL, h: "right" });

  row((q) => strCell(`A${q}`, title, m.title || "Выгрузка"), 20);
  if (m.subtitle) row((q) => strCell(`A${q}`, plain, m.subtitle));
  row((q) => strCell(`A${q}`, muted, "Сформировано") + strCell(`B${q}`, bold, m.created));
  row((q) => strCell(`A${q}`, muted, "Строк в таблице") + numCell(`B${q}`, st.xf({ font: 1, border: false, v: "", fmt: "#,##0", h: "left" }), m.rows));

  if (hdr && m.showSummary && hdr.k.length) {
    blank();
    row((q) => strCell(`A${q}`, band, m.sumTitle || "Показатели") + span(q, 1, 4, band, ""));
    const label = st.xf({ font: 3 });
    const value = st.xf({ font: 1 });
    for (const [l, v] of hdr.k) row((q) => strCell(`A${q}`, label, l) + span(q, 1, 4, value, v));
  }

  if (hdr && m.showSummary && hdr.s.length) {
    blank();
    row((q) => strCell(`A${q}`, st.xf({ font: 1, border: false, v: "" }), "Статусы по периодам КП"));
    const th = st.xf({ font: 1, h: "center" });
    const heads = periods.concat(["Всего"]);
    let top = 0;
    row((q) => {
      top = q;
      let x = strCell(`A${q}`, th, "Статус");
      heads.forEach((p, i) => {
        const a = colName(1 + 2 * i), b = colName(2 + 2 * i);
        x += strCell(`${a}${q}`, th, p) + emptyCell(`${b}${q}`, th);
        merges.push(`${a}${q}:${b}${q}`);
      });
      return x;
    });
    row((q) => emptyCell(`A${q}`, th) + heads.map((_, i) => strCell(`${colName(1 + 2 * i)}${q}`, th, "МКД") + strCell(`${colName(2 + 2 * i)}${q}`, th, "ВР")).join(""));
    merges.push(`A${top}:A${top + 1}`);
    for (const rec of hdr.s) {
      const [color, lab, ...vals] = rec;
      const total = !color && lab === "Всего";
      row((q) => {
        const labStyle = total ? st.xf({ font: 1, fill: HEAD_FILL }) : st.xf({ fill: color });
        let x = strCell(`A${q}`, labStyle, statusLabel(lab));
        vals.forEach((v, i) => {
          const isVr = i % 2 === 1;
          const isAll = i >= vals.length - 2;
          const fill = total ? HEAD_FILL : !isVr && v !== "" ? color : "";
          const s = st.xf({ font: total || isAll ? 1 : 0, fill, h: "right", fmt: "0" });
          const ref = `${colName(1 + i)}${q}`;
          const num = asNumber(v);
          x += v === "" ? emptyCell(ref, s) : num !== null ? numCell(ref, s, num) : strCell(ref, s, v);
        });
        return x;
      });
    }
  }

  if (hdr && m.showSummary && hdr.w.length) {
    blank();
    row((q) => strCell(`A${q}`, band, "Виды работ") + span(q, 1, 2, bandRight, "МКД (ВР/лифтов)"));
    const name = st.xf({});
    const value = st.xf({ font: 1, h: "right", fmt: "0" });
    for (const [l, v] of hdr.w) {
      const num = asNumber(v);
      row((q) => {
        merges.push(`B${q}:C${q}`);
        return strCell(`A${q}`, name, l) + (num !== null ? numCell(`B${q}`, value, num) : strCell(`B${q}`, value, v)) + emptyCell(`C${q}`, value);
      });
    }
  }

  if (hdr && m.showFilters && hdr.f.length) {
    blank();
    row((q) => strCell(`A${q}`, band, "Отбор") + span(q, 1, 4, band, ""));
    const label = st.xf({ font: 1, border: false, v: "" });
    for (const [l, v] of hdr.f) row((q) => strCell(`A${q}`, label, l) + strCell(`B${q}`, plain, v));
  }

  const colsW = [Math.max(30, ...((hdr && hdr.k) || []).map(([l]) => l.length + 2))];
  for (let j = 1; j < width; j++) colsW.push(9);
  return worksheet({ dim: `A1:${colName(width - 1)}${Math.max(r, 1)}`, cols: colsW, rows, merges, landscape: true });
}

// ---------------------------------------------------------------- ZIP

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(u8) {
  let c = 0xffffffff;
  for (let i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

async function deflateRaw(u8) {
  if (typeof CompressionStream !== "function") return null;
  try {
    const stream = new Blob([u8]).stream().pipeThrough(new CompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  } catch (e) {
    return null;       // браузер без deflate-raw — запишем без сжатия
  }
}

function dosDateTime(d) {
  return {
    time: (d.getHours() << 11) | (d.getMinutes() << 5) | Math.floor(d.getSeconds() / 2),
    date: ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()
  };
}

export async function zip(files) {
  const enc = new TextEncoder();
  const { time, date } = dosDateTime(new Date());
  const local = [];
  const central = [];
  let offset = 0;
  for (const [name, text] of files) {
    const data = enc.encode(text);
    const packed = await deflateRaw(data);
    const body = packed || data;
    const method = packed ? 8 : 0;
    const nm = enc.encode(name);
    const crc = crc32(data);
    const h = new DataView(new ArrayBuffer(30));
    h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, method, true);
    h.setUint16(10, time, true); h.setUint16(12, date, true); h.setUint32(14, crc, true);
    h.setUint32(18, body.length, true); h.setUint32(22, data.length, true); h.setUint16(26, nm.length, true);
    local.push(new Uint8Array(h.buffer), nm, body);
    const c = new DataView(new ArrayBuffer(46));
    c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
    c.setUint16(10, method, true); c.setUint16(12, time, true); c.setUint16(14, date, true); c.setUint32(16, crc, true);
    c.setUint32(20, body.length, true); c.setUint32(24, data.length, true); c.setUint16(28, nm.length, true);
    c.setUint32(42, offset, true);
    central.push(new Uint8Array(c.buffer), nm);
    offset += 30 + nm.length + body.length;
  }
  const size = central.reduce((a, p) => a + p.length, 0);
  const e = new DataView(new ArrayBuffer(22));
  e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
  e.setUint32(12, size, true); e.setUint32(16, offset, true);
  return new Blob([...local, ...central, new Uint8Array(e.buffer)], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

// ---------------------------------------------------------------- книга

/**
 * m: { title, subtitle, created, rows, cols (ColumnSpec), values (по столбцам), texts (отформатированные),
 *      rowColors, aligns ("l" | "c" | "r"), hdr, sumTitle, showSummary, showFilters }
 */
export async function buildXlsx(m) {
  const st = new Styles();
  const table = tableSheet(m, st);
  const summary = summarySheet(m, st);
  const files = [
    ["[Content_Types].xml", XML +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      "</Types>"],
    ["_rels/.rels", XML +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ["xl/workbook.xml", XML + `<workbook ${NS} ${NS_R}>` +
      '<bookViews><workbookView activeTab="0"/></bookViews>' +
      '<sheets><sheet name="Таблица" sheetId="1" r:id="rId1"/><sheet name="Сводка" sheetId="2" r:id="rId2"/></sheets>' +
      "<definedNames>" +
      `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'Таблица'!${table.lastRef}</definedName>` +
      `<definedName name="_xlnm.Print_Titles" localSheetId="0">'Таблица'!$1:$1</definedName>` +
      "</definedNames></workbook>"],
    ["xl/_rels/workbook.xml.rels", XML +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/>` +
      `<Relationship Id="rId3" Type="${REL}/styles" Target="styles.xml"/></Relationships>`],
    ["xl/worksheets/sheet1.xml", table.xml],
    ["xl/worksheets/sheet2.xml", summary],
    ["xl/styles.xml", st.xml()]          // стили — после листов: форматы регистрируются по ходу
  ];
  return zip(files);
}
