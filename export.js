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

/** Доли диаграммы: МКД в столбце «Всего» матрицы статусов, от 12 часов по часовой стрелке. */
function pieSlices(hdr) {
  const raw = hdr.s
    .filter((r) => r[0] && r[1] !== "Всего")
    .map((r) => ({ color: r[0], value: Number(String(r[r.length - 2]).replace(/\s/g, "")) || 0 }))
    .filter((x) => x.value > 0);
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
function pieSvg(hdr, availMm) {
  const slices = pieSlices(hdr);
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

/**
 * Карточка диаграммы «МКД по статусам». Сначала рисунок небольшой (не растягивает ряд);
 * после вёрстки fitPie() подбирает его под высоту рамки, равную высоте показателей.
 */
export function renderPie(hdr) {
  const svg = pieSvg(hdr, 24);
  return svg ? `<section class="card card--pie"><h2>МКД по статусам</h2><div class="pie-box">${svg}</div></section>` : "";
}

/** Рисунок диаграммы под заданную высоту — тот же расчёт, что при вписывании в рамку (для проверок). */
export const pieSvgFor = (hdr, availMm) => pieSvg(hdr, availMm);

/** Вписывает диаграмму в рамку: высота рамки задана табличкой показателей (без неё — 38 мм). */
function fitPie(hdr) {
  const box = doc.querySelector(".card--pie .pie-box");
  if (!box || !hdr) return;
  const availMm = doc.querySelector(".card--kpi") ? box.getBoundingClientRect().height * 25.4 / 96 : 38;
  box.innerHTML = pieSvg(hdr, Math.max(availMm, 20));
}

function renderSummary(meta, hdr) {
  if (!meta.showSummary || !hdr) return "";
  // Показатели: подпись и значение — одна запись, между записями линия
  const kpi = hdr.k.length
    ? `<section class="card card--kpi"><h2>${esc(meta.sumTitle || "Показатели")}</h2><table class="kt"><tbody>${
        hdr.k.map(([l, v]) => `<tr><td><span class="kt__label">${esc(l)}</span><span class="kt__value">${esc(v)}</span></td></tr>`).join("")
      }</tbody></table></section>`
    : "";

  let st = "";
  if (hdr.s.length) {
    const periods = hdr.p;
    const heads = periods.concat(["Всего"]);
    // Столбец статусов — по самому длинному названию, остальные столбцы делят ширину поровну
    const plain = measurer(`7.5pt ${FONT_FAMILY}`);
    const bold = measurer(`bold 7.5pt ${FONT_FAMILY}`);
    const pad = 2.4 * MM + 2;                                   // поля 2 × 1,2 мм, рамка, запас
    const labelPx = Math.ceil(Math.max(bold("Статус"), ...hdr.s.map((r) => (r[0] ? plain : bold)(stripStatusPrefix(r[1])))) + pad);
    const valuePx = Math.max(bold("МКД"), bold("ВР"), ...hdr.s.flatMap((r) => r.slice(2).map((v) => bold(String(v))))) + pad;
    const pairPx = Math.max(...heads.map((p) => bold(p) + pad));
    const colMin = Math.max(valuePx, pairPx / 2);
    const minPx = Math.ceil(labelPx + colMin * 2 * heads.length + 3);

    const head1 = `<tr><th rowspan="2">Статус</th>${periods.map((p) => `<th colspan="2">${esc(p)}</th>`).join("")}<th colspan="2">Всего</th></tr>`;
    const head2 = `<tr>${heads.map(() => "<th>МКД</th><th>ВР</th>").join("")}</tr>`;
    const body = hdr.s.map((row) => {
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
    st = `<section class="card card--st" style="min-width:${minPx}px"><h2>Статусы по периодам КП</h2>` +
      `<table class="st"><colgroup><col style="width:${labelPx}px"><col span="${2 * heads.length}"></colgroup>` +
      `<thead>${head1}${head2}</thead><tbody>${body}</tbody></table></section>`;
  }

  // Виды работ: название и значение — одна строка, между строками линия
  const vr = hdr.w.length
    ? `<section class="card card--vr"><h2 class="split"><span>Виды работ</span><span class="note">МКД (ВР/лифт.)</span></h2><table class="kt kt--list"><tbody>${
        hdr.w.map(([l, v]) => `<tr><td>${esc(l)}</td><td class="kt__num">${esc(v)}</td></tr>`).join("")
      }</tbody></table></section>`
    : "";

  const pie = hdr.s.length ? renderPie(hdr) : "";
  const pair = kpi || pie ? `<div class="summary__pair">${kpi}${pie}</div>` : "";
  return `<div class="summary">${pair}${st}${vr}</div>`;
}

function renderFilters(meta, hdr) {
  if (!meta.showFilters || !hdr || !hdr.f.length) return "";
  const items = hdr.f.map(([l, v]) => `<b>${esc(l)}:</b> ${esc(v)}`).join(";&emsp;");
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
  doc.innerHTML =
    renderHead(p.meta, now) +
    renderSummary(p.meta, p.hdr) +
    renderFilters(p.meta, p.hdr) +
    renderTable(p, table);
  fitPie(p.hdr);
  document.title = `${p.meta.title || "Выгрузка"} — ${now.slice(0, 10)}`;
  current = { p, table, now };
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
    const { p, table, now } = current;
    const blob = await buildXlsx({
      title: p.meta.title, subtitle: p.meta.subtitle, created: now, rows: table.n,
      cols: table.cols, values: table.data, texts: table.texts, rowColors: table.colors, aligns: table.aligns,
      hdr: p.hdr, sumTitle: p.meta.sumTitle, showSummary: p.meta.showSummary, showFilters: p.meta.showFilters
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
