import { readXlsx } from "./xlsx.js";

export const SOURCE_COLUMN = "Файл-источник";

export function coerce(value) {
  const text = String(value).trim();
  if (!text) return null;

  const normalized = text.replace(/[\s ]/g, "").replace(",", ".");
  const digits = normalized.replace(/^-/, "");

  // Ведущий ноль — артикул или телефон, а не число: 007 должно остаться 007.
  if (digits.length > 1 && digits.startsWith("0") && !digits.startsWith("0.")) return text;

  if (/^-?\d+$/.test(normalized)) return Number(normalized);
  if (/^-?\d+\.\d+$/.test(normalized)) return Number(normalized);
  return text;
}

function detectDelimiter(line) {
  const counts = [
    [";", (line.match(/;/g) || []).length],
    ["\t", (line.match(/\t/g) || []).length],
    [",", (line.match(/,/g) || []).length],
  ];
  counts.sort((a, b) => b[1] - a[1]);
  return counts[0][1] > 0 ? counts[0][0] : ",";
}

function parseCsv(text, delimiter) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];

    if (quoted) {
      if (char === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else quoted = false;
      } else field += char;
      continue;
    }

    if (char === '"') quoted = true;
    else if (char === delimiter) { row.push(field); field = ""; }
    else if (char === "\n") { row.push(field); rows.push(row); row = []; field = ""; }
    else if (char !== "\r") field += char;
  }

  if (field || row.length) { row.push(field); rows.push(row); }
  return rows.filter((r) => r.some((cell) => cell.trim()));
}

async function readCsvFile(file) {
  const buffer = await file.arrayBuffer();
  let text = new TextDecoder("utf-8", { fatal: false }).decode(buffer);

  // Признак битой кириллицы — пробуем windows-1251.
  if (/�/.test(text)) {
    try { text = new TextDecoder("windows-1251").decode(buffer); } catch {}
  }
  text = text.replace(/^﻿/, "");

  const firstLine = text.slice(0, text.indexOf("\n") + 1 || text.length);
  const rows = parseCsv(text, detectDelimiter(firstLine));
  if (!rows.length) return { header: [], rows: [] };

  return {
    header: rows[0].map((cell) => cell.trim()),
    rows: rows.slice(1).map((row) => row.map(coerce)),
  };
}

async function readXlsxFile(file) {
  const { rows } = await readXlsx(await file.arrayBuffer());
  const filled = rows.filter((row) => row.some((cell) => cell !== null && String(cell).trim() !== ""));
  if (!filled.length) return { header: [], rows: [] };

  return {
    header: filled[0].map((cell) => (cell === null ? "" : String(cell).trim())),
    rows: filled.slice(1),
  };
}

export async function mergeFiles(files) {
  const tables = [];
  const skipped = [];

  for (const file of files) {
    if (file.name.startsWith("~$")) continue;
    try {
      const isCsv = /\.csv$/i.test(file.name);
      const table = isCsv ? await readCsvFile(file) : await readXlsxFile(file);
      if (!table.header.length) { skipped.push({ name: file.name, reason: "файл пустой" }); continue; }
      tables.push({ name: file.name, ...table });
    } catch (error) {
      skipped.push({ name: file.name, reason: error.message || "не удалось прочитать" });
    }
  }

  if (!tables.length) {
    const details = skipped.map((s) => `${s.name} — ${s.reason}`).join("; ");
    throw new Error(details ? `Ни один файл не удалось прочитать: ${details}` : "Нет подходящих файлов");
  }

  const columns = [];
  for (const table of tables) {
    for (const name of table.header) {
      if (name && !columns.includes(name)) columns.push(name);
    }
  }

  const merged = [];
  const perFile = [];
  for (const table of tables) {
    const indexByName = new Map();
    table.header.forEach((name, index) => { if (name && !indexByName.has(name)) indexByName.set(name, index); });

    for (const row of table.rows) {
      const ordered = columns.map((name) => {
        const index = indexByName.get(name);
        return index === undefined || index >= row.length ? null : row[index];
      });
      ordered.push(table.name);
      merged.push(ordered);
    }
    perFile.push({ name: table.name, rows: table.rows.length });
  }

  const allColumns = [...columns, SOURCE_COLUMN];

  const sums = new Map();
  const hits = new Map();
  for (const row of merged) {
    row.forEach((value, index) => {
      if (typeof value === "number" && Number.isFinite(value)) {
        const name = allColumns[index];
        sums.set(name, (sums.get(name) || 0) + value);
        hits.set(name, (hits.get(name) || 0) + 1);
      }
    });
  }

  // Колонка считается числовой, только если числа в большинстве строк.
  const threshold = Math.max(1, Math.floor(merged.length / 2));
  const totals = [...sums.entries()]
    .filter(([name]) => hits.get(name) >= threshold)
    .map(([name, sum]) => ({
      name,
      sum: Math.round(sum * 100) / 100,
      average: Math.round((sum / hits.get(name)) * 100) / 100,
    }));

  return { columns: allColumns, rows: merged, perFile, totals, skipped };
}

export function buildSheets(result) {
  const summary = [["Источник", "Строк"]];
  for (const item of result.perFile) summary.push([item.name, item.rows]);
  summary.push(["Всего", result.rows.length]);

  if (result.totals.length) {
    summary.push([]);
    summary.push(["Числовая колонка", "Сумма", "Среднее"]);
    for (const total of result.totals) summary.push([total.name, total.sum, total.average]);
  }

  return [
    { name: "Данные", rows: [result.columns, ...result.rows] },
    { name: "Сводка", rows: summary },
  ];
}
