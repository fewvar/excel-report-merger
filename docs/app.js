import { mergeFiles, buildSheets } from "./merge.js";
import { writeXlsx } from "./xlsx.js";

const PREVIEW_ROWS = 12;
const ACCEPTED = /\.(xlsx|xlsm|csv)$/i;

const dom = {
  drop: document.getElementById("drop"),
  picker: document.getElementById("picker"),
  choose: document.getElementById("choose"),
  demo: document.getElementById("demo"),
  error: document.getElementById("error"),
  files: document.getElementById("files"),
  filesList: document.getElementById("filesList"),
  clear: document.getElementById("clear"),
  run: document.getElementById("run"),
  result: document.getElementById("result"),
  stat: document.getElementById("stat"),
  download: document.getElementById("download"),
  previewData: document.getElementById("previewData"),
  previewSummary: document.getElementById("previewSummary"),
  previewNote: document.getElementById("previewNote"),
};

let selected = [];
let blob = null;

function showError(message) {
  dom.error.textContent = message;
  dom.error.hidden = !message;
}

function formatNumber(value) {
  return typeof value === "number" ? value.toLocaleString("ru-RU") : value;
}

function renderFiles() {
  dom.filesList.innerHTML = "";
  for (const file of selected) {
    const item = document.createElement("li");
    const name = document.createElement("span");
    name.textContent = file.name;
    const size = document.createElement("span");
    size.className = "files__size";
    size.textContent = `${(file.size / 1024).toFixed(0)} КБ`;
    item.append(name, size);
    dom.filesList.append(item);
  }
  dom.files.hidden = selected.length === 0;
  dom.run.textContent = `Собрать отчёт из ${selected.length} ${plural(selected.length)}`;
}

function plural(count) {
  const mod10 = count % 10;
  const mod100 = count % 100;
  if (mod10 === 1 && mod100 !== 11) return "файла";
  return "файлов";
}

function addFiles(list) {
  const incoming = [...list].filter((file) => ACCEPTED.test(file.name) && !file.name.startsWith("~$"));
  if (!incoming.length) {
    showError("Подходящих файлов нет. Нужны .xlsx, .xlsm или .csv");
    return;
  }
  showError("");
  const seen = new Set(selected.map((f) => f.name + f.size));
  for (const file of incoming) {
    if (!seen.has(file.name + file.size)) selected.push(file);
  }
  dom.result.hidden = true;
  renderFiles();
}

function renderTable(container, columns, rows) {
  const table = document.createElement("table");

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  for (const name of columns) {
    const th = document.createElement("th");
    th.textContent = name;
    headRow.append(th);
  }
  thead.append(headRow);
  table.append(thead);

  const tbody = document.createElement("tbody");
  for (const row of rows) {
    const tr = document.createElement("tr");
    for (const value of row) {
      const td = document.createElement("td");
      if (value === null || value === undefined || value === "") {
        td.textContent = "—";
        td.className = "empty";
      } else {
        td.textContent = formatNumber(value);
        if (typeof value === "number") td.className = "num";
      }
      tr.append(td);
    }
    tbody.append(tr);
  }
  table.append(tbody);

  container.innerHTML = "";
  container.append(table);
}

async function run() {
  if (!selected.length) return;
  dom.run.disabled = true;
  dom.run.textContent = "Собираю…";
  showError("");

  try {
    const result = await mergeFiles(selected);
    const sheets = buildSheets(result);
    blob = await writeXlsx(sheets);

    dom.stat.textContent = `${result.perFile.length} ${plural(result.perFile.length)} → ${result.rows.length} строк, колонок ${result.columns.length}`;

    renderTable(dom.previewData, result.columns, result.rows.slice(0, PREVIEW_ROWS));

    const summaryRows = [
      ...result.perFile.map((item) => [item.name, item.rows]),
      ["Всего", result.rows.length],
      ...(result.totals.length ? [[], ["Числовая колонка", "Сумма", "Среднее"]] : []),
      ...result.totals.map((total) => [total.name, total.sum, total.average]),
    ];
    renderTable(dom.previewSummary, ["Источник", "Строк", ""], summaryRows);

    const notes = [];
    if (result.rows.length > PREVIEW_ROWS) {
      notes.push(`показаны первые ${PREVIEW_ROWS} строк из ${result.rows.length} — в файле все`);
    }
    if (result.skipped.length) {
      notes.push(`пропущено: ${result.skipped.map((s) => `${s.name} (${s.reason})`).join(", ")}`);
    }
    dom.previewNote.textContent = notes.join(" · ");

    dom.result.hidden = false;
    dom.result.scrollIntoView({ behavior: "smooth", block: "nearest" });
  } catch (error) {
    showError(error.message || "Не удалось собрать отчёт");
  } finally {
    dom.run.disabled = false;
    renderFiles();
  }
}

async function loadDemo() {
  showError("");
  dom.demo.disabled = true;
  dom.demo.textContent = "Загружаю…";
  try {
    const names = ["январь", "февраль", "март", "апрель", "май"];
    const files = await Promise.all(
      names.map(async (month, index) => {
        const path = `demo/продажи-${String(index + 1).padStart(2, "0")}-${month}.xlsx`;
        const response = await fetch(path);
        if (!response.ok) throw new Error(`не найден ${path}`);
        return new File([await response.blob()], path.split("/").pop());
      })
    );
    selected = [];
    addFiles(files);
  } catch (error) {
    showError(`Не удалось загрузить пример: ${error.message}`);
  } finally {
    dom.demo.disabled = false;
    dom.demo.textContent = "Взять пример";
  }
}

dom.choose.addEventListener("click", () => dom.picker.click());
dom.picker.addEventListener("change", (event) => addFiles(event.target.files));
dom.demo.addEventListener("click", loadDemo);
dom.run.addEventListener("click", run);

dom.clear.addEventListener("click", () => {
  selected = [];
  blob = null;
  dom.result.hidden = true;
  showError("");
  renderFiles();
});

dom.download.addEventListener("click", () => {
  if (!blob) return;
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "сводный-отчёт.xlsx";
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
});

for (const event of ["dragenter", "dragover"]) {
  dom.drop.addEventListener(event, (e) => {
    e.preventDefault();
    dom.drop.classList.add("is-over");
  });
}
for (const event of ["dragleave", "drop"]) {
  dom.drop.addEventListener(event, (e) => {
    e.preventDefault();
    if (event === "dragleave" && dom.drop.contains(e.relatedTarget)) return;
    dom.drop.classList.remove("is-over");
  });
}
dom.drop.addEventListener("drop", (e) => addFiles(e.dataTransfer.files));
dom.drop.addEventListener("click", (e) => {
  if (e.target.closest("button")) return;
  dom.picker.click();
});

document.querySelectorAll(".tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("is-active", t === tab));
    const showData = tab.dataset.tab === "data";
    dom.previewData.hidden = !showData;
    dom.previewSummary.hidden = showData;
  });
});
