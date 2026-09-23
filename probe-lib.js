// probe-lib.js — общая логика тестовых страниц пробника (ES-модуль, без зависимостей).
// Формат должен совпадать с visual/src/probeCore.ts.

export const TYPES = Object.freeze({
  PING: "PBI_PROBE_PING",
  PONG: "PBI_PROBE_PONG",
  DATA: "PBI_PROBE_DATA"
});

/** FNV-1a (32 бита) по кодам символов — та же функция, что в визуале. */
export function fnv1a(text) {
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return ("00000000" + (h >>> 0).toString(16)).slice(-8);
}

export function fromBase64Url(s) {
  let b64 = s.replace(/-/g, "+").replace(/_/g, "/");
  const rem = b64.length % 4;
  if (rem === 1) throw new Error("некорректная длина base64url");
  if (rem) b64 += "=".repeat(4 - rem);
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export async function gunzipText(bytes) {
  if (typeof DecompressionStream !== "function") {
    throw new Error("браузер не поддерживает DecompressionStream");
  }
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"));
  return await new Response(stream).text();
}

/**
 * Разбор фрагмента адреса:
 *   #t.<длина>.<fnv>.<данные>             — синтетический тест
 *   #d.<gz>.<байт JSON>.<fnv>.<пакет>     — реальный пакет
 */
export function parseHash(hash) {
  const body = (hash || "").replace(/^#/, "");
  if (!body) return { kind: "none" };
  const parts = body.split(".");
  if (parts[0] === "t" && parts.length === 4) {
    return { kind: "t", expected: Number(parts[1]), fnv: parts[2], data: parts[3] };
  }
  if (parts[0] === "d" && parts.length === 5) {
    return { kind: "d", gz: parts[1] === "1", rawBytes: Number(parts[2]), fnv: parts[3], packed: parts[4] };
  }
  return { kind: "bad", reason: "Формат фрагмента не распознан: адрес мог быть обрезан по пути." };
}

/** Проверка целостности и раскрытие пакета (base64url → gzip → JSON). */
export async function decodePacked({ gz, fnv, packed }) {
  const t0 = performance.now();
  if (typeof packed !== "string" || packed.length === 0) {
    return { ok: false, error: "Пакет пуст." };
  }
  const actual = fnv1a(packed);
  if (fnv && actual !== fnv) {
    return {
      ok: false,
      packedLength: packed.length,
      error: `Контрольная сумма ${actual}, ожидалась ${fnv}: данные обрезаны или искажены по пути.`
    };
  }
  try {
    const bytes = fromBase64Url(packed);
    const text = gz ? await gunzipText(bytes) : new TextDecoder().decode(bytes);
    const obj = JSON.parse(text);
    return {
      ok: true,
      obj,
      rows: Array.isArray(obj.rows) ? obj.rows.length : 0,
      cols: Array.isArray(obj.cols) ? obj.cols.length : 0,
      textBytes: new Blob([text]).size,
      packedLength: packed.length,
      ms: Math.round(performance.now() - t0)
    };
  } catch (err) {
    return {
      ok: false,
      packedLength: packed.length,
      error: "Пакет не раскрылся: " + (err && err.message ? err.message : String(err))
    };
  }
}

export function fmtInt(n) {
  return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, "\u00a0");
}

export function fmtBytes(n) {
  if (n < 1024) return fmtInt(n) + "\u00a0Б";
  if (n < 1048576) return (n / 1024).toFixed(1).replace(".", ",") + "\u00a0КБ";
  return (n / 1048576).toFixed(2).replace(".", ",") + "\u00a0МБ";
}

export function browserName(ua = navigator.userAgent) {
  const pick = (re, name) => {
    const m = ua.match(re);
    return m ? name + " " + m[1] : null;
  };
  return pick(/YaBrowser\/([\d.]+)/, "Яндекс Браузер")
    || pick(/Edg[A-Za-z]*\/([\d.]+)/, "Edge")
    || pick(/OPR\/([\d.]+)/, "Opera")
    || pick(/FxiOS\/([\d.]+)/, "Firefox iOS")
    || pick(/Firefox\/([\d.]+)/, "Firefox")
    || pick(/CriOS\/([\d.]+)/, "Chrome iOS")
    || pick(/Chrome\/([\d.]+)/, "Chrome")
    || (/Safari\//.test(ua) ? (pick(/Version\/([\d.]+)/, "Safari") || "Safari") : null)
    || "неизвестный браузер";
}

export function stamp(d = new Date()) {
  const p = (n) => String(n).padStart(2, "0");
  return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

/** Таблица первых строк пакета. Только textContent — данные не интерпретируются как HTML. */
export function renderPreview(host, obj, maxRows = 12) {
  host.replaceChildren();
  if (!obj || !Array.isArray(obj.cols) || !Array.isArray(obj.rows)) return;
  const table = document.createElement("table");
  const headRow = table.createTHead().insertRow();
  for (const c of obj.cols) {
    const th = document.createElement("th");
    th.textContent = String(c);
    headRow.appendChild(th);
  }
  const body = table.createTBody();
  for (const r of obj.rows.slice(0, maxRows)) {
    const tr = body.insertRow();
    for (const v of r) {
      const td = tr.insertCell();
      td.textContent = v === null || v === undefined ? "" : String(v);
      if (typeof v === "number") td.className = "num";
    }
  }
  host.appendChild(table);
}
