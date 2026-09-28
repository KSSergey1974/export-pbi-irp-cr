// export.js — страница выгрузки: пакет из фрагмента адреса → документ для печати в PDF.
// Формат пакета совпадает с visual/src/payload.ts (ExportPayload, v: 2).

const PAPER = {
  A4L: { size: "A4 landscape", width: "281mm" },
  A4P: { size: "A4 portrait", width: "194mm" },
  A3L: { size: "A3 landscape", width: "404mm" }
};
const UNITS = { "": [1, ""], k: [1e3, " тыс."], m: [1e6, " млн"], b: [1e9, " млрд"] };

const bar = document.getElementById("bar");
const barTitle = document.getElementById("bar-title");
const barHint = document.getElementById("bar-hint");
const printBtn = document.getElementById("print");
const xlsxBtn = document.getElementById("xlsx");
let current = null;     // подготовленная выгрузка: общая для PDF (печать) и XLSX
const doc = document.getElementById("doc");

// ---------------------------------------------------------------- разбор пакета

function fnv1a(text) {
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return ("00000000" + (h >>> 0).toString(16)).slice(-8);
}

function fromBase64Url(s) {
  let b64 = s.replace(/-/g, "+").replace(/_/g, "/");
  const rem = b64.length % 4;
  if (rem === 1) throw new Error("некорректная длина пакета");
  if (rem) b64 += "=".repeat(4 - rem);
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

async function gunzipText(bytes) {
  if (typeof DecompressionStream !== "function") {
    throw new Error("браузер не поддерживает распаковку (нужен Chrome, Edge, Яндекс, Firefox 113+ или Safari 16.4+)");
  }
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"));
  return await new Response(stream).text();
}

export async function decodeHash(hash) {
  const body = (hash || "").replace(/^#/, "");
  if (!body) return { none: true };
  const parts = body.split(".");
  if (parts[0] !== "x" || parts.length !== 4) {
    throw new Error("адрес не похож на выгрузку из отчёта или был обрезан");
  }
  const [, gz, fnv, packed] = parts;
  if (fnv1a(packed) !== fnv) {
    throw new Error("данные в адресе обрезаны или искажены (не сошлась контрольная сумма)");
  }
  const bytes = fromBase64Url(packed);
  const text = gz === "1" ? await gunzipText(bytes) : new TextDecoder().decode(bytes);
  const payload = JSON.parse(text);
  if (!payload || payload.v !== 2) throw new Error("неизвестная версия пакета — обновите визуал или страницу");
  return { payload, urlLength: location.href.length };
}

function decodeColumn(enc) {
  if (Array.isArray(enc)) return enc;
  if (enc && Array.isArray(enc.d) && Array.isArray(enc.i)) return enc.i.map((k) => enc.d[k]);
  return [];
}

// ---------------------------------------------------------------- форматирование

const esc = (s) => String(s)
  .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function pad2(n) { return (n < 10 ? "0" : "") + n; }

function fmtDay(day) {
  const d = new Date(day * 86400000);
  return pad2(d.getUTCDate()) + "." + pad2(d.getUTCMonth() + 1) + "." + d.getUTCFullYear();
}

function fmtNum(v, dec) {
  return v.toLocaleString("ru-RU", { minimumFractionDigits: dec, maximumFractionDigits: dec });
}

export function formatCell(v, c) {
  if (v === null || v === undefined || v === "") return "";
  switch (c.k) {
    case "d": return typeof v === "number" ? fmtDay(v) : String(v);
    case "b": return v === 1 || v === true ? "Да" : "Нет";
    case "p": return typeof v === "number" ? fmtNum(v * 100, Math.max(c.dec, 0)) + "%" + (c.sx || "") : String(v);
    case "n": {
      if (typeof v !== "number") return String(v);
      const [div, label] = UNITS[c.u] || UNITS[""];
      return fmtNum(v / div, Math.max(c.dec, 0)) + label + (c.sx ? (label ? c.sx.replace(/^\s*/, " ") : c.sx) : "");
    }
    default: return String(v);
  }
}

function autoAlign(c) {
  if (c.al) return c.al;
  return c.k === "n" || c.k === "p" ? "r" : c.k === "d" || c.k === "b" ? "c" : "l";
}

const FONT_FAMILY = 'Arial, "Liberation Sans", "Helvetica Neue", Helvetica, sans-serif';
const MM = 96 / 25.4;                       // CSS-пикселей в миллиметре
const CELL_PAD = 1.8 * MM + 2;              // поля ячейки 2 × 0,8 мм, рамка и запас на округление
let measureCtx = null;

/** Измеритель ширины строки тем шрифтом, которым она будет напечатана (с кешем). */
function measurer(font) {
  if (!measureCtx) measureCtx = document.createElement("canvas").getContext("2d");
  const cache = new Map();
  return (text) => {
    let w = cache.get(text);
    if (w === undefined) {
      measureCtx.font = font;
      w = measureCtx.measureText(text).width;
      cache.set(text, w);
    }
    return w;
  };
}

const sum = (a) => a.reduce((x, y) => x + y, 0);

/**
 * Ширины столбцов в процентах ширины таблицы.
 * Минимум столбца — самое длинное неразрывное значение (без переноса — вся строка, с переносом —
 * самое длинное слово) и самое длинное слово заголовка. Желаемая ширина — 90-й процентиль строк
 * (с переносом) или самая длинная строка (без переноса). Остаток ширины получают столбцы,
 * где текст ещё переносится, затем — текстовые столбцы с переносом. Заданная в визуале
 * ширина (c.w, % ширины листа) соблюдается как есть.
 */
export function columnWidths(cols, texts, totalPx, fontPt) {
  const plain = measurer(`${fontPt}pt ${FONT_FAMILY}`);
  const bold = measurer(`bold ${fontPt}pt ${FONT_FAMILY}`);
  const cap = totalPx * 0.4;
  const min = [], pref = [], full = [], fixed = [];
  cols.forEach((c, j) => {
    const widths = [];
    let token = 0;
    let widest = 0;
    for (const s of texts[j]) {
      if (!s) continue;
      const w = plain(s);
      widths.push(w);
      if (w > widest) widest = w;
      if (c.wr) {
        for (const t of s.split(/\s+/)) if (t) { const tw = plain(t); if (tw > token) token = tw; }
      }
    }
    if (!c.wr) token = widest;
    const head = Math.max(0, ...c.h.split(/\s+/).filter(Boolean).map(bold));
    widths.sort((a, b) => a - b);
    const p90 = widths.length ? widths[Math.min(widths.length - 1, Math.floor(widths.length * 0.9))] : 0;
    min[j] = Math.min(Math.max(token, head) + CELL_PAD, cap);
    pref[j] = Math.min(Math.max(min[j], (c.wr ? p90 : widest) + CELL_PAD), Math.max(cap, min[j]));
    full[j] = Math.max(pref[j], Math.min(widest + CELL_PAD, cap));
    fixed[j] = c.w > 0 ? totalPx * Math.min(c.w, 90) / 100 : 0;
  });

  const width = cols.map((c, j) => fixed[j]);
  const auto = cols.map((c, j) => j).filter((j) => !fixed[j]);
  const rest = totalPx - sum(width);
  const sumMin = sum(auto.map((j) => min[j]));
  const sumPref = sum(auto.map((j) => pref[j]));

  if (auto.length === 0) {
    // все ширины заданы вручную — только нормируем
  } else if (rest >= sumPref) {
    auto.forEach((j) => { width[j] = pref[j]; });
    let extra = rest - sumPref;
    const needy = auto.filter((j) => full[j] > pref[j] + 0.5);           // текст там ещё переносится
    const deficit = sum(needy.map((j) => full[j] - pref[j]));
    if (deficit > 0) {
      const give = Math.min(extra, deficit);
      needy.forEach((j) => { width[j] += give * (full[j] - pref[j]) / deficit; });
      extra -= give;
    }
    if (extra > 0.5) {
      const wrapText = auto.filter((j) => cols[j].k === "t" && cols[j].wr);
      const pool = wrapText.length ? wrapText : auto;
      const base = sum(pool.map((j) => full[j])) || 1;
      pool.forEach((j) => { width[j] += extra * full[j] / base; });
    }
  } else if (rest >= sumMin) {
    const span = sumPref - sumMin;
    auto.forEach((j) => { width[j] = min[j] + (span > 0 ? (rest - sumMin) * (pref[j] - min[j]) / span : 0); });
  } else {
    auto.forEach((j) => { width[j] = min[j]; });                          // не помещается: сжимаем всё пропорционально
  }
  const total = sum(width) || 1;
  return width.map((w) => w / total * 100);
}

function stripStatusPrefix(s) {
  return String(s || "").replace(/^\s*\d+\s*-\s*/, "");
}

// ---------------------------------------------------------------- сборка документа

function renderHead(meta, now) {
  const sub = meta.subtitle ? `<p>${esc(meta.subtitle)}</p>` : "";
  return `<header class="head">
    <div><h1>${esc(meta.title || "Выгрузка")}</h1>${sub}</div>
    <dl class="head__meta">
      <dt>Сформировано</dt><dd>${esc(now)}</dd>
      <dt>Строк в таблице</dt><dd>${esc(meta.rows.toLocaleString("ru-RU"))}${meta.partial ? " (не все)" : ""}</dd>
    </dl>
  </header>`;
}

const PIE_UNIT_MM = 0.25;                               // 1 единица рисунка = 0,25 мм
const PIE_FONT = 7 / 72 * 25.4 / PIE_UNIT_MM;          // подписи 7 пт в единицах рисунка
const PIE_GAP = 1.2 / PIE_UNIT_MM;                      // зазор от рисунка до рамки, 1,2 мм
const PIE_MAX_W = 80 / PIE_UNIT_MM;                     // рисунок не шире 80 мм
const PX_TO_UNIT = 25.4 / 96 / PIE_UNIT_MM;             // CSS-пиксель → единица рисунка

const f1 = (x) => x.toFixed(1);

/** Помещается ли подпись (w × h) с центром (px, py) целиком в сектор [a0, a1] круга радиуса r (центр в 0,0). */
function labelInsideWedge(r, a0, a1, px, py, w, h) {
  const span = a1 - a0;
  for (const [dx, dy] of [[-w / 2, -h / 2], [w / 2, -h / 2], [-w / 2, h / 2], [w / 2, h / 2]]) {
    const x = px + dx, y = py + dy;
    if (Math.hypot(x, y) > r - 2) return false;
    let rel = Math.atan2(y, x) - a0;
    rel = ((rel % (2 * Math.PI)) + 2 * Math.PI) % (2 * Math.PI);
    if (rel < 0.03 || rel > span - 0.03) return false;
  }
  return true;
}

/** Доли диаграммы из матрицы статусов: МКД в столбце «Всего» (предпоследнее поле строки S). */
function pieItemsFromStatus(sRows) {
  return sRows
    .filter((r) => r[0] && r[1] !== "Всего")
    .map((r) => ({ color: r[0], value: Number(String(r[r.length - 2]).replace(/\s/g, "")) || 0 }));
}

/** Нормализация входа диаграммы: { s: строки матрицы } (как раньше) или готовый список { color, value }. */
function pieItems(src) {
  if (Array.isArray(src)) return src;
  return src && src.s ? pieItemsFromStatus(src.s) : [];
}

/** Доли диаграммы от 12 часов по часовой стрелке. */
function pieSlices(src) {
  const raw = pieItems(src).filter((x) => x.value > 0);
  const total = raw.reduce((a, x) => a + x.value, 0);
  if (!total) return null;
  const measure = measurer(`7pt ${FONT_FAMILY}`);
  let a0 = -Math.PI / 2;
  return raw.map((x) => {
    const share = x.value / total;
    const a1 = a0 + share * 2 * Math.PI;
    const text = (share * 100).toLocaleString("ru-RU", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + "%";
    const slice = { color: x.color, share, a0, a1, mid: (a0 + a1) / 2, text, tw: measure(text) * PX_TO_UNIT };
    a0 = a1;
    return slice;
  });
}

/** Раскладка при радиусе r (центр в 0,0): подписи внутри долей, где помещаются, иначе снаружи; рамка всего рисунка. */
function pieLayout(slices, r) {
  const th = PIE_FONT * 1.05;
  const inside = [];
  const outside = [];
  for (const x of slices) {
    let placed = false;
    for (const k of [0.62, 0.52, 0.72]) {
      const px = k * r * Math.cos(x.mid), py = k * r * Math.sin(x.mid);
      if (x.share > 0.9999 || labelInsideWedge(r, x.a0, x.a1, px, py, x.tw + 2, th)) {
        inside.push({ x: px, y: py, text: x.text });
        placed = true;
        break;
      }
    }
    if (!placed) outside.push({ mid: x.mid, text: x.text, tw: x.tw, right: Math.cos(x.mid) >= 0, y: (r + 9) * Math.sin(x.mid) });
  }
  // наружные подписи одной стороны раздвигаем по вертикали, если сошлись
  const gap = th * 1.1;
  for (const side of [true, false]) {
    const ls = outside.filter((l) => l.right === side).sort((a, b) => a.y - b.y);
    for (let i = 1; i < ls.length; i++) ls[i].y = Math.max(ls[i].y, ls[i - 1].y + gap);
  }
  const box = { x0: -r, y0: -r, x1: r, y1: r };
  const grow = (x, y) => { box.x0 = Math.min(box.x0, x); box.x1 = Math.max(box.x1, x); box.y0 = Math.min(box.y0, y); box.y1 = Math.max(box.y1, y); };
  for (const l of outside) {
    const c = Math.cos(l.mid), sn = Math.sin(l.mid);
    l.sx = (r + 1) * c; l.sy = (r + 1) * sn;                            // от края доли
    l.ex = (r + 6) * c; l.ey = (r + 6) * sn;                            // короткий отрезок по радиусу
    l.tx = l.ex + (l.right ? 3 : -3);                                   // и чуть в сторону
    const x0 = l.right ? l.tx + 1.5 : l.tx - 1.5 - l.tw;
    grow(l.ex, l.ey); grow(x0, l.y - th / 2); grow(x0 + l.tw, l.y + th / 2);
  }
  return { r, inside, outside, box };
}

/**
 * SVG диаграммы наибольшего размера, который с подписями и зазором PIE_GAP помещается
 * в высоту availMm (и не шире PIE_MAX_W). Подписи всегда 7 пт — растёт только круг.
 */
function pieSvg(src, availMm) {
  const slices = pieSlices(src);
  if (!slices) return "";
  const hMax = availMm / PIE_UNIT_MM;
  let L = null;
  for (let r = Math.floor(hMax / 2); r >= 20; r -= 0.5) {
    const t = pieLayout(slices, r);
    if (t.box.y1 - t.box.y0 + 2 * PIE_GAP <= hMax && t.box.x1 - t.box.x0 + 2 * PIE_GAP <= PIE_MAX_W) { L = t; break; }
  }
  if (!L) L = pieLayout(slices, 20);
  const { r, inside, outside, box } = L;
  const vx = box.x0 - PIE_GAP, vy = box.y0 - PIE_GAP;
  const vw = box.x1 - box.x0 + 2 * PIE_GAP, vh = box.y1 - box.y0 + 2 * PIE_GAP;

  const shapes = slices.map((x) => {
    if (x.share > 0.9999) return `<circle cx="0" cy="0" r="${f1(r)}" fill="${esc(x.color)}"/>`;
    const p0 = [r * Math.cos(x.a0), r * Math.sin(x.a0)], p1 = [r * Math.cos(x.a1), r * Math.sin(x.a1)];
    return `<path d="M0 0L${f1(p0[0])} ${f1(p0[1])}A${f1(r)} ${f1(r)} 0 ${x.share > 0.5 ? 1 : 0} 1 ${f1(p1[0])} ${f1(p1[1])}Z" fill="${esc(x.color)}"/>`;
  });
  const labels = inside.map((l) => `<text x="${f1(l.x)}" y="${f1(l.y)}" text-anchor="middle" dominant-baseline="central">${esc(l.text)}</text>`);
  for (const l of outside) {
    labels.push(`<polyline points="${f1(l.sx)},${f1(l.sy)} ${f1(l.ex)},${f1(l.ey)} ${f1(l.tx)},${f1(l.y)}" fill="none" stroke="#8c8c8c" stroke-width="0.6"/>` +
      `<text x="${f1(l.tx + (l.right ? 1.5 : -1.5))}" y="${f1(l.y)}" text-anchor="${l.right ? "start" : "end"}" dominant-baseline="central">${esc(l.text)}</text>`);
  }
  return `<svg class="pie" viewBox="${f1(vx)} ${f1(vy)} ${f1(vw)} ${f1(vh)}" style="width:${(vw * PIE_UNIT_MM).toFixed(2)}mm;height:${(vh * PIE_UNIT_MM).toFixed(2)}mm" role="img" aria-label="Доли МКД по статусам">` +
    `<g stroke="#ffffff" stroke-width="0.8">${shapes.join("")}</g>` +
    `<circle cx="0" cy="0" r="${f1(r)}" fill="none" stroke="#b3b3b3" stroke-width="0.6"/>` +     // светлые доли не сливаются с фоном
    `<g font-family='Arial, "Liberation Sans", sans-serif' font-size="${f1(PIE_FONT)}" fill="#262626">${labels.join("")}</g></svg>`;
}

const pendingPies = [];     // диаграммы, которые надо вписать в рамку после вёрстки

/**
 * Карточка диаграммы. Сначала рисунок небольшой (не растягивает ряд);
 * после вёрстки fitPies() подбирает его под высоту рамки, равную высоте соседнего блока.
 */
export function renderPie(src, title = "МКД по статусам") {
  const svg = pieSvg(src, 24);
  if (!svg) return "";
  const id = `pie-${pendingPies.length}`;
  pendingPies.push({ id, src });
  return `<section class="card card--pie"><h2>${esc(title)}</h2><div class="pie-box" id="${id}">${svg}</div></section>`;
}

/** Рисунок диаграммы под заданную высоту — тот же расчёт, что при вписывании в рамку (для проверок). */
export const pieSvgFor = (src, availMm) => pieSvg(src, availMm);

/** Вписывает диаграммы в рамки: в паре высоту задаёт соседний блок, отдельно стоящая — 38 мм. */
function fitPies() {
  for (const { id, src } of pendingPies.splice(0)) {
    const box = document.getElementById(id);
    if (!box) continue;
    const paired = !!box.closest(".summary__pair");
    const availMm = paired ? box.getBoundingClientRect().height * 25.4 / 96 : 38;
    box.innerHTML = pieSvg(src, Math.max(availMm, 20));
  }
}

// ---------------------------------------------------------------- шапка: блоки
//
// Формат PDF2: записи через ¶, поля через ¦. «B¦тип¦заголовок¦флаги» начинает блок, следующие
// записи относятся к нему. Типы: kpi (K¦подпись¦значение), list (W¦название¦значение; «- » —
// вложенный пункт), table (H — заголовки, X — стиль столбцов: «#цвет» заливка или «bar #цвет»
// полоса по проценту в ячейке, R¦цвет¦ячейки…, T — строка итога), status (P, S — матрица
// статусов), pie (V¦подпись¦значение¦цвет или доли из матрицы статусов), filters (F).
// Флаги через «;»: below — блок под предыдущим; fill — рамка до нижнего края ряда (нижние границы
// таких блоков на одном уровне, строки таблицы растягиваются); note=текст — подпись справа в заголовке.
// Пустой заголовок — блок без полосы заголовка.
// Старая шапка PDF1 («РРП_МО») переводится в те же блоки.

function block(type, title, extra = {}) {
  return Object.assign({ type, title, below: false, fill: false, note: "", k: [], p: [], s: [], w: [], f: [], h: [], x: [], rows: [], total: null, v: [] }, extra);
}

export function parseBlocks(raw) {
  const recs = String(raw || "").split("\u00b6");
  if (recs[0] !== "PDF2") return null;
  const blocks = [];
  let b = null;
  for (let i = 1; i < recs.length; i++) {
    if (!recs[i]) continue;
    const f = recs[i].split("\u00a6");
    const t = f.shift();
    if (t === "B") {
      const flags = f[2] || "";
      b = block(f[0] || "", f[1] || "", {
        below: /(^|;)\s*below\s*(;|$)/i.test(flags),
        fill: /(^|;)\s*fill\s*(;|$)/i.test(flags),
        note: (flags.match(/(?:^|;)\s*note=([^;]*)/i) || [])[1] || ""
      });
      blocks.push(b);
      continue;
    }
    if (!b) continue;
    if (t === "K") b.k.push(f);
    else if (t === "P") b.p = f;
    else if (t === "S") b.s.push(f);
    else if (t === "W") b.w.push(f);
    else if (t === "F") b.f.push(f);
    else if (t === "H") b.h = f;
    else if (t === "X") b.x = f;
    else if (t === "R") b.rows.push({ color: f[0] || "", cells: f.slice(1) });
    else if (t === "T") b.total = f;
    else if (t === "V") b.v.push(f);
  }
  return blocks;
}

/** Старая шапка «РРП_МО» (PDF1): показатели, диаграмма, матрица статусов, виды работ, отбор. */
export function legacyBlocks(meta, hdr) {
  const out = [];
  if (hdr.k.length) out.push(block("kpi", meta.sumTitle || "Показатели", { k: hdr.k }));
  if (hdr.s.length) {
    out.push(block("pie", "МКД по статусам"));
    out.push(block("status", "Статусы по периодам КП", { p: hdr.p, s: hdr.s }));
  }
  if (hdr.w.length) out.push(block("list", "Виды работ", { note: "МКД (ВР/лифт.)", w: hdr.w }));
  if (hdr.f.length) out.push(block("filters", "Отбор", { f: hdr.f }));
  return out;
}

/** Блоки шапки из пакета: новая строка PDF2 (hraw) или старая разобранная шапка (hdr). */
export function headerBlocks(p) {
  if (p.hraw) return parseBlocks(p.hraw) || [];
  return p.hdr ? legacyBlocks(p.meta, p.hdr) : [];
}

function kpiCard(b) {
  if (!b.k.length) return "";
  return `<section class="card card--kpi"><h2>${esc(b.title || "Показатели")}</h2><table class="kt"><tbody>${
    b.k.map(([l, v]) => `<tr><td><span class="kt__label">${esc(l)}</span><span class="kt__value">${esc(v)}</span></td></tr>`).join("")
  }</tbody></table></section>`;
}

function listCard(b) {
  if (!b.w.length) return "";
  const head = b.note
    ? `<h2 class="split"><span>${esc(b.title)}</span><span class="note">${esc(b.note)}</span></h2>`
    : `<h2>${esc(b.title)}</h2>`;
  const rows = b.w.map(([l, v]) => {
    const sub = /^\s*-\s+/.test(l || "");                           // «- » — вложенный пункт
    return `<tr><td${sub ? ' class="kt__sub"' : ""}>${esc(sub ? l.replace(/^\s*-\s+/, "") : l)}</td><td class="kt__num">${esc(v === undefined ? "" : v)}</td></tr>`;
  }).join("");
  return `<section class="card card--list">${head}<table class="kt kt--list"><tbody>${rows}</tbody></table></section>`;
}

function statusCard(b) {
  if (!b.s.length) return "";
  const periods = b.p;
  const heads = periods.concat(["Всего"]);
  // Столбец статусов — по самому длинному названию, остальные столбцы делят ширину поровну
  const plain = measurer(`7.5pt ${FONT_FAMILY}`);
  const bold = measurer(`bold 7.5pt ${FONT_FAMILY}`);
  const pad = 2.4 * MM + 2;                                   // поля 2 × 1,2 мм, рамка, запас
  const labelPx = Math.ceil(Math.max(bold("Статус"), ...b.s.map((r) => (r[0] ? plain : bold)(stripStatusPrefix(r[1])))) + pad);
  const valuePx = Math.max(bold("МКД"), bold("ВР"), ...b.s.flatMap((r) => r.slice(2).map((v) => bold(String(v))))) + pad;
  const pairPx = Math.max(...heads.map((h) => bold(h) + pad));
  const colMin = Math.max(valuePx, pairPx / 2);
  const minPx = Math.ceil(labelPx + colMin * 2 * heads.length + 3);

  const head1 = `<tr><th rowspan="2">Статус</th>${periods.map((h) => `<th colspan="2">${esc(h)}</th>`).join("")}<th colspan="2">Всего</th></tr>`;
  const head2 = `<tr>${heads.map(() => "<th>МКД</th><th>ВР</th>").join("")}</tr>`;
  const body = b.s.map((row) => {
    const [color, label, ...vals] = row;
    const total = !color && label === "Всего";
    const bg = color ? ` style="background:${esc(color)}"` : "";
    // Цветом статуса — подпись строки и непустые ячейки МКД; ячейки ВР и пустые — белые (как в отчёте)
    const cells = vals.map((v, i) => {
      const fill = color && i % 2 === 0 && v !== "" ? bg : "";
      return `<td class="${i >= vals.length - 2 ? "st__all" : ""}"${fill}>${esc(v)}</td>`;
    }).join("");
    return `<tr class="${total ? "st__total" : ""}"><td class="st__label"${bg}>${esc(stripStatusPrefix(label))}</td>${cells}</tr>`;
  }).join("");
  return `<section class="card card--st" style="min-width:${minPx}px"><h2>${esc(b.title || "Статусы по периодам КП")}</h2>` +
    `<table class="st"><colgroup><col style="width:${labelPx}px"><col span="${2 * heads.length}"></colgroup>` +
    `<thead>${head1}${head2}</thead><tbody>${body}</tbody></table></section>`;
}

/** Стиль столбца таблицы из записи X: «#цвет» — заливка, «bar #цвет» — полоса по проценту из текста ячейки. */
function colStyle(spec) {
  const m = /^\s*(bar\s+)?(#[0-9a-f]{6})\s*$/i.exec(spec || "");
  return m ? { bar: !!m[1], color: m[2] } : null;
}

function percentOf(text) {
  const m = /(-?\d+(?:[.,]\d+)?)\s*%/.exec(String(text || "").replace(/\s/g, ""));
  return m ? Math.min(Math.max(parseFloat(m[1].replace(",", ".")), 0), 100) : null;
}

function tableCard(b) {
  if (!b.rows.length && !b.total) return "";
  const n = Math.max(b.h.length, ...b.rows.map((r) => r.cells.length), b.total ? b.total.length : 0);
  const styles = Array.from({ length: n }, (_, j) => colStyle(b.x[j]));
  const cell = (v, j, rowColor) => {
    const cs = styles[j];
    let style = "";
    if (j === 0 && rowColor) style = `background:${rowColor};`;
    else if (cs && cs.bar) {
      const pct = percentOf(v);
      if (pct !== null) style = `background:linear-gradient(to right, ${cs.color} ${pct.toFixed(1)}%, transparent ${pct.toFixed(1)}%);`;
    } else if (cs) style = `background:${cs.color};`;
    return `<td class="${j === 0 ? "st__label" : ""}"${style ? ` style="${esc(style)}"` : ""}>${esc(v === undefined ? "" : v)}</td>`;
  };
  const head = b.h.length ? `<thead><tr>${Array.from({ length: n }, (_, j) => `<th>${esc(b.h[j] || "")}</th>`).join("")}</tr></thead>` : "";
  const body = b.rows.map((r) => `<tr>${Array.from({ length: n }, (_, j) => cell(r.cells[j], j, r.color)).join("")}</tr>`).join("");
  const total = b.total ? `<tr class="st__total">${Array.from({ length: n }, (_, j) => `<td class="${j === 0 ? "st__label" : ""}">${esc(b.total[j] || "")}</td>`).join("")}</tr>` : "";
  const h2 = b.title ? `<h2>${esc(b.title)}</h2>` : "";
  return `<section class="card card--tbl">${h2}<table class="st st--auto">${head}<tbody>${body}${total}</tbody></table></section>`;
}

function pieCard(b, blocks) {
  const items = b.v.length
    ? b.v.map((r) => ({ color: r[2] || "#cccccc", value: Number(String(r[1]).replace(/\s/g, "").replace(",", ".")) || 0 }))
    : pieItemsFromStatus((blocks.find((x) => x.type === "status") || { s: [] }).s);
  return renderPie(items, b.title || "МКД по статусам");
}

function blockCard(b, blocks) {
  switch (b.type) {
    case "kpi": return kpiCard(b);
    case "list": return listCard(b);
    case "status": return statusCard(b);
    case "table": return tableCard(b);
    case "pie": return pieCard(b, blocks);
    default: return "";
  }
}

/**
 * Сводка на первом листе: блоки по порядку, слева направо, колонками. «below» ставит блок
 * под предыдущим, диаграмма встаёт в пару с предыдущим блоком (та же высота рамки).
 * Остаток ширины получает колонка с матрицей статусов или самой широкой таблицей; что не
 * помещается в ряд, переносится на следующий.
 */
function renderSummary(meta, blocks) {
  if (!meta.showSummary) return "";
  const cols = [];
  for (const b of blocks) {
    if (b.type === "filters") continue;
    let html = blockCard(b, blocks);
    if (!html) continue;
    if (b.fill) html = html.replace('<section class="card ', '<section class="card fill ');
    const last = cols[cols.length - 1];
    if (b.type === "pie" && last && last.cards.length === 1 && !last.pair) { last.pair = true; last.cards.push(html); last.blocks.push(b); continue; }
    if (b.below && last) { last.cards.push(html); last.blocks.push(b); continue; }
    cols.push({ cards: [html], blocks: [b], pair: false });
  }
  if (!cols.length) return "";
  const weight = (c) => Math.max(...c.blocks.map((b) => (b.type === "status" ? 1000 : b.type === "table" ? Math.max(b.h.length, ...b.rows.map((r) => r.cells.length)) : 0)));
  let grow = -1, best = 3;                                     // таблица из 4 и более столбцов или матрица
  cols.forEach((c, i) => { const w = weight(c); if (w > best) { best = w; grow = i; } });
  return `<div class="summary">${cols.map((c, i) => {
    const inner = c.pair ? `<div class="summary__pair">${c.cards.join("")}</div>` : c.cards.join("");
    const fill = c.blocks.some((b) => b.fill) ? " fill" : "";
    return `<div class="summary__col${i === grow ? " grow" : ""}${fill}">${inner}</div>`;
  }).join("")}</div>`;
}

function renderFilters(meta, blocks) {
  const f = blocks.filter((b) => b.type === "filters").flatMap((b) => b.f);
  if (!meta.showFilters || !f.length) return "";
  const items = f.map(([l, v]) => `<b>${esc(l)}:</b> ${esc(v)}`).join(";&emsp;");
  return `<section class="filters"><h2>Отбор</h2><p>${items}</p></section>`;
}

/** Значения, тексты и цвета строк — один раз для документа и для XLSX. */
function prepareTable(p) {
  const cols = p.cols;
  const data = p.data.map(decodeColumn);
  const colors = p.rc ? decodeColumn(p.rc) : [];
  const texts = cols.map((c, j) => data[j].map((v) => formatCell(v, c)));
  return { cols, data, colors, texts, aligns: cols.map(autoAlign), n: p.meta.rows };
}

function renderTable(p, t) {
  const { cols, data, colors, texts, aligns, n } = t;
  const paper = PAPER[p.meta.paper] || PAPER.A4L;
  const tablePx = (parseFloat(paper.width) - 1.2) * MM;                 // ширина области печати минус поля .doc
  const shares = columnWidths(cols, texts, tablePx, p.meta.font || 7);
  const colgroup = shares.map((w) => `<col style="width:${w.toFixed(3)}%">`).join("");
  const thead = cols.map((c) => `<th>${esc(c.h)}</th>`).join("");

  const out = [];
  for (let i = 0; i < n; i++) {
    const rowColor = colors[i] || "";
    let tr = "<tr>";
    for (let j = 0; j < cols.length; j++) {
      const c = cols[j];
      const text = texts[j][i];
      let fill = "";
      if (c.fm === "row") fill = rowColor;
      else if (c.fm === "fixed" && text !== "") fill = c.fc;
      let style = fill ? `background:${fill};` : "";
      const v = data[j][i];
      if (c.bar && typeof v === "number") {
        const share = Math.min(Math.max(c.k === "p" ? v : v / 100, 0), 1) * 100;
        const base = fill || "transparent";
        style = `background:linear-gradient(to ${c.br ? "left" : "right"}, ${c.bc} ${share.toFixed(1)}%, ${base} ${share.toFixed(1)}%);`;
      }
      const cls = aligns[j] + (c.wr ? "" : " nw");
      tr += `<td class="${cls}"${style ? ` style="${esc(style)}"` : ""}>${esc(text)}</td>`;
    }
    out.push(tr + "</tr>");
  }
  return `<table class="grid"><colgroup>${colgroup}</colgroup><thead><tr>${thead}</tr></thead><tbody>${out.join("")}</tbody></table>`;
}

function cssString(s) {
  return '"' + String(s || "").replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/[\r\n]+/g, " ") + '"';
}

function applyPageRules(meta) {
  const paper = PAPER[meta.paper] || PAPER.A4L;
  document.documentElement.style.setProperty("--page-w", paper.width);
  document.documentElement.style.setProperty("--font", (meta.font || 7) + "pt");
  document.getElementById("page-rules").textContent =
    `@page { size: ${paper.size}; margin: 9mm 8mm 12mm 8mm;` +
    ` @bottom-left { content: ${cssString(meta.footer)}; font: 7pt Arial, sans-serif; color: #666666; }` +
    ` @bottom-right { content: "Стр. " counter(page) " из " counter(pages); font: 7pt Arial, sans-serif; color: #666666; } }`;
}

function nowText() {
  const d = new Date();
  return `${pad2(d.getDate())}.${pad2(d.getMonth() + 1)}.${d.getFullYear()} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
}

export function renderDocument(p) {
  const now = nowText();
  const table = prepareTable(p);
  applyPageRules(p.meta);
  doc.dataset.paper = PAPER[p.meta.paper] ? p.meta.paper : "A4L";   // раскладка сводки зависит от формата листа
  const blocks = headerBlocks(p);
  doc.innerHTML =
    renderHead(p.meta, now) +
    renderSummary(p.meta, blocks) +
    renderFilters(p.meta, blocks) +
    renderTable(p, table);
  fitPies();
  document.title = `${p.meta.title || "Выгрузка"} — ${now.slice(0, 10)}`;
  current = { p, table, now, blocks };
}

// ---------------------------------------------------------------- XLSX

/** Имя файла: заголовок и дата, без символов, запрещённых в Windows. */
function fileName(title, now) {
  const base = `${title || "Выгрузка"} — ${now.slice(0, 10)}`.replace(/[\\/:*?"<>|\u0000-\u001f]+/g, "_").trim();
  return (base || "Выгрузка") + ".xlsx";
}

function saveBlob(blob, name) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 60000);
}

async function downloadXlsx() {
  if (!current) return;
  const label = xlsxBtn.textContent;
  xlsxBtn.disabled = true;
  xlsxBtn.textContent = "Готовлю XLSX…";
  try {
    const { buildXlsx } = await import("./xlsx.js");    // модуль грузится только при выборе XLSX
    const { p, table, now, blocks } = current;
    const blob = await buildXlsx({
      title: p.meta.title, subtitle: p.meta.subtitle, created: now, rows: table.n,
      cols: table.cols, values: table.data, texts: table.texts, rowColors: table.colors, aligns: table.aligns,
      blocks, showSummary: p.meta.showSummary, showFilters: p.meta.showFilters
    });
    const name = fileName(p.meta.title, now);
    saveBlob(blob, name);
    barHint.textContent = `Файл «${name}» (${Math.max(1, Math.round(blob.size / 1024)).toLocaleString("ru-RU")} КБ) сохранён в папку загрузок.`;
    document.body.dataset.xlsx = "1";
  } catch (e) {
    barHint.textContent = `Не удалось собрать XLSX: ${e.message}`;
  } finally {
    xlsxBtn.disabled = false;
    xlsxBtn.textContent = label;
  }
}

// ---------------------------------------------------------------- запуск

function showError(title, hint) {
  bar.classList.add("bar--error");
  barTitle.textContent = title;
  barHint.textContent = hint;
  printBtn.disabled = true;
  xlsxBtn.disabled = true;
  current = null;
}

async function main() {
  bar.classList.remove("bar--error");
  const hash = location.hash;
  if (hash) history.replaceState(null, "", location.pathname + location.search);   // данные не остаются в адресе
  let decoded;
  try {
    decoded = await decodeHash(hash);
  } catch (e) {
    showError("Не удалось открыть выгрузку", `${e.message}. Нажмите кнопку в отчёте ещё раз; если повторится — сузьте отбор.`);
    return;
  }
  if (decoded.none) {
    showError("Страница открыта без данных", "Откройте её кнопкой выгрузки в отчёте Power BI.");
    doc.innerHTML = "";
    return;
  }
  const p = decoded.payload;
  const t0 = performance.now();
  renderDocument(p);
  const ms = Math.round(performance.now() - t0);
  barTitle.textContent = "Документ готов — выберите формат";
  // Колонтитул с номерами страниц рисует сама страница (Chrome, Edge, Яндекс); Firefox его не поддерживает.
  // Принтер «Microsoft Print to PDF» печатает на книжный лист и поворачивает альбомную страницу — нужен встроенный PDF браузера.
  const pdfTip = /Firefox\//.test(navigator.userAgent)
    ? "PDF: в окне печати назначение «Сохранить в PDF» (не «Microsoft Print to PDF» — он поворачивает лист); номера страниц — в «Колонтитулах» Firefox."
    : "PDF: в окне печати назначение «Сохранить как PDF» (не «Microsoft Print to PDF» — он поворачивает лист), галочку «Колонтитулы» снимите.";
  barHint.textContent = `Строк: ${p.meta.rows.toLocaleString("ru-RU")}. ${pdfTip} XLSX: файл сохранится в папку загрузок.`;
  barHint.title = `Подготовка документа заняла ${ms} мс`;
  printBtn.disabled = false;
  xlsxBtn.disabled = false;
  document.body.dataset.ready = "1";
}

printBtn.addEventListener("click", () => window.print());
xlsxBtn.addEventListener("click", () => { void downloadXlsx(); });
// Новая ссылка, вставленная в уже открытую вкладку, меняет только фрагмент — перерисовываем
window.addEventListener("hashchange", () => { if (location.hash) main(); });
main();
