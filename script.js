// script.js — универсальный "солдат" для чтения разных 1С-документов
let items = [];
let currentIndex = 0;
let currentConfig = null;
let totalRowsInSheet = null;

document.getElementById('file-input').addEventListener('change', handleFile, false);

// Создаём индикатор типа документа (если в index.html его нет — создадим динамически)
if (!document.getElementById('doc-type')) {
  const el = document.createElement('div');
  el.id = 'doc-type';
  el.style.fontWeight = '600';
  el.style.marginTop = '6px';
  document.querySelector('h1')?.after(el);
}

function setDocTypeLabel(text) {
  const el = document.getElementById('doc-type');
  el.textContent = text ? `Тип распознанного документа: ${text}` : '';
}

// ---------- handleFile: загрузка и детект типа ----------
function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();

  reader.onload = function (ev) {
    const data = new Uint8Array(ev.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

    // Попробуем детектировать тип документа
    const detected = detectDocType(json);
    currentConfig = getConfigForType(detected);
    totalRowsInSheet = json.length;

    // Парсим лист по конфигу
    items = parseWithConfig(json, currentConfig);

    renderTable(totalRowsInSheet);
    currentIndex = 0;
    setDocTypeLabel(currentConfig ? currentConfig.label : 'Не определён');
  };

  reader.readAsArrayBuffer(file);
}

// ---------- Типы документов и их конфиги ----------
function detectDocType(json) {
  // Просматриваем первые 30 строк (или до конца) и ищем сигнатуры
  const upto = Math.min(30, json.length);
  for (let i = 0; i < upto; i++) {
    const row = (json[i] || []).join(' ').toString().toLowerCase();
    if (!row) continue;

    if (row.includes('расходная накладная') || row.includes('расходная')) return 'rashod';
    if (row.includes('счет-фактура') || row.includes('счёт-фактура')) return 'ufd';
    if (row.includes('счёт') || row.includes('счет') || row.includes('счет №')) {
      // Чтобы не спутать с 'расходная' — проверка выше пройдёт раньше
      return 'schet';
    }
    if (row.includes('перемещение')) return 'perem';
    if (row.includes('универсальный передаточный документ') || row.includes('упд')) return 'ufd';
  }

  // fallback — если ничего не найдено, считаем расходной (как прежняя логика)
  return 'rashod';
}

function getConfigForType(type) {
  // индексы: A=0, B=1, C=2, ... (нуль-индекс)
  const commonQty = [20, 22]; // U..W => 20..22
  switch (type) {
    case 'rashod':
      return {
        type: 'rashod',
        label: 'Расходная накладная',
        startRowIndex: 8,             // i = 8 -> 9-я строка в Excel
        articleCols: [3, 19],         // D..T
        qtyCols: commonQty,           // U..W
        numCols: null
      };
    case 'schet':
      return {
        type: 'schet',
        label: 'Счёт',
        startRowIndex: 20,            // 21-я строка
        articleCols: [3, 19],         // D..T
        qtyCols: commonQty,           // U..W
        numCols: null
      };
    case 'perem':
      return {
        type: 'perem',
        label: 'Перемещение',
        startRowIndex: 8,             // 9-я строка
        articleCols: [3, 19],         // D..T
        qtyCols: commonQty,           // U..W
        numCols: [1, 2]               // B..C
      };
    case 'ufd':
      return {
        type: 'ufd',
        label: 'Счёт-фактура / УПД',
        startRowIndex: 16,            // 17-я строка
        articleCols: [16, 43],        // Q..AR -> 16..43
        qtyCols: [68, 74],            // BQ..BW -> 68..74
        numCols: [11, 15]             // L..P -> 11..15
      };
    default:
      return {
        type: 'unknown',
        label: 'Неизвестный',
        startRowIndex: 8,
        articleCols: [3, 19],
        qtyCols: commonQty,
        numCols: null
      };
  }
}

// ---------- Разбор листа с учётом конфигурации ----------
function parseWithConfig(json, cfg) {
  const out = [];
  if (!cfg) return out;

  for (let i = cfg.startRowIndex; i < json.length; i++) {
    const row = json[i] || [];

    // собираем текст для артикула из диапазона articleCols
    const articleRangeText = collectRangeText(row, cfg.articleCols[0], cfg.articleCols[1]);

    // также попытаемся взять "основной" артикул из знакомой F (index 5) — если там что-то есть
    const primaryCell = row[5] || '';
    const candidateText = (articleRangeText || '') + ' ' + (primaryCell || '');
    const fullArticleText = (candidateText || '').toString().trim();

    // вычисляем количество как максимум из qtyCols
    let qty = 0;
    if (cfg.qtyCols && cfg.qtyCols.length === 2) {
      for (let c = cfg.qtyCols[0]; c <= cfg.qtyCols[1]; c++) {
        const val = parseFloat((row[c] !== undefined && row[c] !== null) ? String(row[c]).replace(',', '.') : '') || 0;
        qty = Math.max(qty, Math.floor(val));
      }
    }

    // пропуск пустых/нулевых строк
    if (!fullArticleText || qty <= 0) continue;

    // попытка извлечь артикула и префикс по паттерну
    const extracted = extractArticleFromTextRange(row, cfg.articleCols[0], cfg.articleCols[1]) 
                      || extractArticleFromAnyCell(row) 
                      || { article: fullArticleText, prefix: null, main: null, extra: null };

    out.push({
      article: extracted.article,
      prefix: extracted.prefix,
      main: extracted.main,
      extra: extracted.extra,
      qty,
      row,
      checked: false,
      type: cfg.type,
      sheetRowIndex: i // для отладки/индикации
    });
  }

  return out;
}

// вспомогательная: собрать текст из диапазона колонок (inclusively)
function collectRangeText(row, startCol, endCol) {
  if (!row) return '';
  const parts = [];
  for (let c = startCol; c <= endCol; c++) {
    const v = row[c];
    if (v !== undefined && v !== null && String(v).trim() !== '') parts.push(String(v).trim());
  }
  return parts.join(' ').replace(/\s+/g, ' ').trim();
}

// ищем артикулы KR/KU/KLT/PT в указанном диапазоне
function extractArticleFromTextRange(row, startCol, endCol) {
  const pattern = /(KR|KU|КР|КУ|KLT|PT|РТ)[-.\s–—]?([\w\d.-]+)/i;
  for (let c = startCol; c <= endCol; c++) {
    const cell = row[c];
    if (!cell || typeof cell !== 'string') continue;
    const match = cell.match(pattern);
    if (match) {
      const prefix = match[1];
      const main = match[2] || null;
      // PT special: return whole joined range if PT is found and that's desired behavior
      if (prefix.toUpperCase() === 'PT') {
        return { article: collectRangeText(row, startCol, endCol), prefix: 'PT', main: null, extra: null };
      }
      return { article: `${prefix}-${main}`, prefix, main, extra: null };
    }
  }
  return null;
}

// Более «жирный» поиск по всем ячейкам строки (fallback)
function extractArticleFromAnyCell(row) {
  const pattern = /(KR|KU|КР|КУ|KLT|PT|РТ)[-.\s–—]?([\w\d.-]+)/i;
  for (let cell of row) {
    if (!cell || typeof cell !== 'string') continue;
    const match = cell.match(pattern);
    if (match) {
      const prefix = match[1];
      const main = match[2] || null;
      if (prefix.toUpperCase() === 'PT') {
        return { article: row.filter(Boolean).join(', '), prefix: 'PT', main: null, extra: null };
      }
      return { article: `${prefix}-${main}`, prefix, main, extra: null };
    }
  }
  return null;
}

// ---------- Рендер таблицы ----------
function renderTable(totalRows = null) {
  const tbody = document.querySelector("#items-table tbody");
  tbody.innerHTML = "";

  items.forEach((item, idx) => {
    const rowEl = document.createElement("tr");
    if (idx === currentIndex) rowEl.classList.add("active-row");

    const td1 = document.createElement("td");
    td1.textContent = item.article;

    const td2 = document.createElement("td");
    td2.textContent = item.qty;

    const td3 = document.createElement("td");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = item.checked;
    checkbox.addEventListener("change", (e) => {
      items[idx].checked = e.target.checked;
      renderTable(totalRows);
    });
    td3.appendChild(checkbox);

    rowEl.appendChild(td1);
    rowEl.appendChild(td2);
    rowEl.appendChild(td3);

    rowEl.addEventListener("click", (e) => {
      if (e.target.tagName.toLowerCase() === "input") return;
      currentIndex = idx;
      speakCurrent();
    });

    tbody.appendChild(rowEl);
  });

  let text = `Загружено позиций: ${items.length}`;
  if (totalRows !== null && currentConfig) {
    const visibleRows = Math.max(0, totalRows - currentConfig.startRowIndex);
    text += ` / всего строк (лист): ${totalRows} (обрабатываемых с ${currentConfig.startRowIndex + 1}: ${visibleRows})`;
  }
  document.getElementById("count").textContent = text;
}

// ---------- Текст в речь ----------
const synth = window.speechSynthesis;
const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
let recognition;
let isListening = false;

function startReading() {
  if (!items.length) return;
  currentIndex = 0;
  speakCurrent();
  startListening();
}

function speakCurrent() {
  if (!items[currentIndex]) {
    speak("Нет позиций для обработки.");
    return;
  }

  const it = items[currentIndex];
  const { prefix, main, extra, qty } = it;
  let articleText;

  if (prefix && ["KR", "КР", "KU", "КУ", "KLT"].includes(String(prefix).toUpperCase())) {
    articleText = formatArticle(prefix, main, extra);
  } else {
    // прочитать соединённый текст артикула (range) — если есть
    const cfg = currentConfig;
    if (cfg) {
      const row = it.row || [];
      const textRange = collectRangeText(row, cfg.articleCols[0], cfg.articleCols[1]);
      articleText = textRange || it.article;
    } else {
      articleText = it.article;
    }
  }

  const qtyText = numberToWordsRu(qty);
  const qtyEnding = getQtySuffix(qty);
  const phrase = `${articleText} положить ${qtyText} ${qtyEnding}`;
  speak(phrase);
  renderTable(totalRowsInSheet);
}

function speak(text) {
  const utterance = new SpeechSynthesisUtterance(text);
  utterance.lang = 'ru-RU';
  synth.cancel();
  synth.speak(utterance);
}

// ---------- Числа в слова (исправлено) ----------
function numberToWordsRuNom(num) {
  num = parseInt(num);
  if (isNaN(num)) return String(num);

  const ones = ["ноль", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"];
  const teens = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"];
  const tens = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"];
  const hundreds = ["", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"];

  if (num < 10) return ones[num];
  if (num < 20) return teens[num - 10];
  if (num < 100) {
    const t = Math.floor(num / 10);
    const o = num % 10;
    return tens[t] + (o ? " " + ones[o] : "");
  }
  if (num < 1000) {
    const h = Math.floor(num / 100);
    const rem = num % 100;
    return hundreds[h] + (rem ? " " + numberToWordsRuNom(rem) : "");
  }

  return num.toString();
}

function numberToWordsRu(num) {
  num = parseInt(num);
  if (isNaN(num)) return String(num);

  const ones = ["ноль", "одну", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"];
  const teens = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"];
  const tens = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"];
  const hundreds = ["", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"];

  if (num < 10) return ones[num];
  if (num < 20) return teens[num - 10];
  if (num < 100) {
    const t = Math.floor(num / 10);
    const o = num % 10;
    return tens[t] + (o ? " " + ones[o] : "");
  }
  if (num < 1000) {
    const h = Math.floor(num / 100);
    const rem = num % 100;
    return hundreds[h] + (rem ? " " + numberToWordsRuNom(rem) : "");
  }

  return num.toString();
}

// ---------- Форматирование артикула для чтения ----------
function formatArticle(prefix, main, extra) {
  const upperPrefix = String(prefix || '').toUpperCase();
  const isKR = upperPrefix.includes("KR") || upperPrefix.includes("КР");
  const isKU = upperPrefix.includes("KU") || upperPrefix.includes("КУ");
  const isKLT = upperPrefix === "KLT";

  if (isKR) {
    const ruPrefix = "КаЭр";
    return `${ruPrefix} ${numberToWordsRuNom(main)}${extra ? ' дробь ' + numberToWordsRuNom(extra) : ''}`;
  }

  if (isKU) {
    const ruPrefix = "Кудо";
    const raw = String(main || '');
    let parts = [];

    if (raw.length === 4) {
      parts = [raw.slice(0, 2), raw.slice(2)];
    } else if (raw.length === 5) {
      parts = [raw.slice(0, 2), raw.slice(2)];
    } else if (raw.length === 6) {
      parts = [raw.slice(0, 2), raw.slice(2, 4), raw.slice(4)];
    } else {
      parts = [raw];
    }

    const spoken = parts.map(p => {
      if (p.length === 2 && p.startsWith("0")) {
        return "ноль " + numberToWordsRuNom(p[1]);
      } else {
        return numberToWordsRuNom(parseInt(p));
      }
    }).join(" ");

    if (isKLT) {
      return `КэЭлТэ ${numberToWordsRuNom(main)}${extra ? ' дробь ' + numberToWordsRuNom(extra) : ''}`;
    }

    return `${ruPrefix} ${spoken}${extra ? ' ' + extra : ''}`;
  }

  // fallback: просто соединяем
  return `${prefix}${main ? '-' + main : ''}${extra ? '-' + extra : ''}`.replace(/^-/, '');
}

// ---------- Суффикс для количества ----------
function getQtySuffix(num) {
  const rem10 = num % 10;
  const rem100 = num % 100;
  if (rem10 === 1 && rem100 !== 11) return "штуку";
  if ([2, 3, 4].includes(rem10) && ![12, 13, 14].includes(rem100)) return "штуки";
  return "штук";
}

// ---------- Голосовое распознавание ----------
function startListening() {
  if (isListening) return;

  recognition = new SpeechRecognition();
  recognition.lang = 'ru-RU';
  recognition.continuous = true;
  recognition.interimResults = false;

  recognition.onresult = function (event) {
    const transcript = event.results[event.results.length - 1][0].transcript.trim().toLowerCase();
    console.log("Распознано:", transcript);
    handleVoiceCommand(transcript);
  };

  recognition.onerror = function (event) {
    console.error("Ошибка распознавания:", event.error);
    if (event.error === "not-allowed" || event.error === "service-not-allowed") {
      isListening = false;
    }
  };

  recognition.onend = function () {
    console.log("Прослушка завершена");
    if (isListening) {
      setTimeout(() => {
        try { recognition.start(); } catch (e) { console.warn('restart fail', e); }
      }, 300);
    }
  };

  isListening = true;
  try {
    recognition.start();
    console.log("Прослушка запущена");
  } catch (e) {
    console.warn("Не удалось запустить распознавание:", e);
  }
}

function stopListening() {
  if (!isListening || !recognition) return;
  try {
    recognition.stop();
  } catch (e) {}
  isListening = false;
}

// ---------- Навигация по позициям ----------
function speakNextUnprocessed() {
  let next = currentIndex + 1;
  while (next < items.length && items[next].checked) next++;
  if (next < items.length) {
    currentIndex = next;
    speakCurrent();
  } else {
    speak("Больше неотмеченных позиций нет.");
  }
}

function handleVoiceCommand(cmd) {
  if (!cmd) return;
  // упрощённая нормализация коротких команд
  if (["готово", "положил", "ок"].some(k => cmd === k || cmd.includes(k))) {
    items[currentIndex].checked = true;
    speakNextUnprocessed();
  } else if (["дальше", "пропускаем", "некст", "следующий"].some(k => cmd === k || cmd.includes(k))) {
    speakNextUnprocessed();
  } else if (cmd.includes("назад")) {
    currentIndex = Math.max(0, currentIndex - 1);
    speakCurrent();
  } else if (["повтори", "щё раз", "ещё раз", "повторить"].some(k => cmd.includes(k))) {
    speakCurrent();
  } else if (cmd.includes("начни") && cmd.includes("пропущ")) {
    startFromSkipped();
  } else {
    console.log("Команда не распознана как управление:", cmd);
  }

  renderTable(totalRowsInSheet);
}

// ---------- Начать с пропущенных ----------
function startFromSkipped() {
  let next = items.findIndex(item => !item.checked);
  if (next !== -1) {
    currentIndex = next;
    speakCurrent();
  } else {
    speak("Все позиции уже обработаны.");
  }
}