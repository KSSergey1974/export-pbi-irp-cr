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

/** Доли ширины: заданные в визуале или по длине содержимого (выборка до 400 строк). */
export function columnWeights(cols, texts) {
  return cols.map((c, j) => {
    if (c.w > 0) return c.w;
    const sample = texts[j].slice(0, 400);
    let sum = 0;
    let word = 0;    // самое длинное слово: его нельзя перенести, столбец должен его вмещать
    for (const s of sample) {
      sum += Math.min(s.length, 70);
      if (c.wr) { for (const w of s.split(/\s+/)) if (w.length > word) word = w.length; }
      else if (s.length > word) word = s.length;
    }
    const avg = sample.length ? sum / sample.length : 4;
    const headWord = Math.max(...c.h.split(/\s+/).map((w) => w.length), 3);
    return Math.min(Math.max(avg * 1.05, Math.min(word, 24) * 0.95, headWord * 0.9, 3.5), 34);
  });
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

function renderSummary(meta, hdr) {
  if (!meta.showSummary || !hdr) return "";
  const kpi = hdr.k.length
    ? `<section class="card"><h2>${esc(meta.sumTitle || "Показатели")}</h2><dl class="kv">${
        hdr.k.map(([l, v]) => `<dt>${esc(l)}</dt><dd>${esc(v)}</dd>`).join("")}</dl></section>`
    : "<div></div>";

  let st = "<div></div>";
  if (hdr.s.length) {
    const periods = hdr.p;
    const head1 = `<tr><th rowspan="2">Статус</th>${periods.map((p) => `<th colspan="2">${esc(p)}</th>`).join("")}<th colspan="2">Всего</th></tr>`;
    const head2 = `<tr>${periods.concat(["Всего"]).map(() => "<th>МКД</th><th>ВР</th>").join("")}</tr>`;
    const body = hdr.s.map((row) => {
      const [color, label, ...vals] = row;
      const total = !color && label === "Всего";
      const bg = color ? ` style="background:${esc(color)}"` : "";
      const cells = vals.map((v, i) => `<td class="${i >= vals.length - 2 ? "st__all" : ""}"${bg}>${esc(v)}</td>`).join("");
      return `<tr class="${total ? "st__total" : ""}"><td class="st__label"${bg}>${esc(stripStatusPrefix(label))}</td>${cells}</tr>`;
    }).join("");
    st = `<section class="card"><h2>Статусы по периодам КП</h2><table class="st"><thead>${head1}${head2}</thead><tbody>${body}</tbody></table></section>`;
  }

  const vr = hdr.w.length
    ? `<section class="card"><h2>Виды работ</h2><dl class="kv kv--list">${
        hdr.w.map(([l, v]) => `<dt>${esc(l)}</dt><dd>${esc(v)}</dd>`).join("")}</dl></section>`
    : "<div></div>";

  return `<div class="summary">${kpi}${st}${vr}</div>`;
}

function renderFilters(meta, hdr) {
  if (!meta.showFilters || !hdr || !hdr.f.length) return "";
  const items = hdr.f.map(([l, v]) => `<b>${esc(l)}:</b> ${esc(v)}`).join(";&emsp;");
  return `<section class="filters"><h2>Отбор</h2><p>${items}</p></section>`;
}

function renderTable(p) {
  const cols = p.cols;
  const data = p.data.map(decodeColumn);
  const colors = p.rc ? decodeColumn(p.rc) : [];
  const n = p.meta.rows;
  const texts = cols.map((c, j) => data[j].map((v) => formatCell(v, c)));

  const weights = columnWeights(cols, texts);
  const total = weights.reduce((a, b) => a + b, 0) || 1;
  const colgroup = weights.map((w) => `<col style="width:${(w / total * 100).toFixed(2)}%">`).join("");
  const thead = cols.map((c) => `<th>${esc(c.h)}</th>`).join("");
  const aligns = cols.map(autoAlign);

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
  applyPageRules(p.meta);
  doc.innerHTML =
    renderHead(p.meta, now) +
    renderSummary(p.meta, p.hdr) +
    renderFilters(p.meta, p.hdr) +
    renderTable(p);
  document.title = `${p.meta.title || "Выгрузка"} — ${now.slice(0, 10)}`;
}

// ---------------------------------------------------------------- запуск

function showError(title, hint) {
  bar.classList.add("bar--error");
  barTitle.textContent = title;
  barHint.textContent = hint;
  printBtn.disabled = true;
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
    showError("Страница открыта без данных", "Откройте её кнопкой «Выгрузить в PDF» в отчёте Power BI.");
    doc.innerHTML = "";
    return;
  }
  const p = decoded.payload;
  const t0 = performance.now();
  renderDocument(p);
  const ms = Math.round(performance.now() - t0);
  barTitle.textContent = "Документ готов";
  // Колонтитул с номерами страниц рисует сама страница (Chrome, Edge, Яндекс); Firefox его не поддерживает
  const tip = /Firefox\//.test(navigator.userAgent)
    ? "В окне печати выберите «Сохранить в PDF»; номера страниц включаются в «Колонтитулах» Firefox."
    : "В окне печати выберите «Сохранить как PDF» и снимите галочку «Колонтитулы» — номера страниц уже есть внизу листа.";
  barHint.textContent = `Строк: ${p.meta.rows.toLocaleString("ru-RU")}. ${tip}`;
  barHint.title = `Подготовка документа заняла ${ms} мс`;
  printBtn.disabled = false;
  document.body.dataset.ready = "1";

  const noPrint = new URLSearchParams(location.search).has("noprint");
  if (p.meta.autoPrint && !noPrint) {
    const fontsReady = document.fonts && document.fonts.ready ? document.fonts.ready : Promise.resolve();
    fontsReady.then(() => setTimeout(() => window.print(), 350));
  }
}

printBtn.addEventListener("click", () => window.print());
// Новая ссылка, вставленная в уже открытую вкладку, меняет только фрагмент — перерисовываем
window.addEventListener("hashchange", () => { if (location.hash) main(); });
main();
