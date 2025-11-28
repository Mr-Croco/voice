// script.js — универсальный "солдат" для чтения разных 1С-документов
// Полная версия — вставь целиком вместо старого script.js

let items = [];
let currentIndex = 0;
let currentConfig = null;
let totalRowsInSheet = null;

let recognition = null;
let isListening = false;

const synth = window.speechSynthesis;
const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition || null;

// Если нет индикатора типа — создаём
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

// ====== Mojibake fixer + utils ======
function fixWin1251Mojibake(str) {
  if (!str || typeof str !== 'string') return str;
  // Если в строке нет подозрительных символов — не трогаем
  if (!/[ÃÐÂ]/.test(str)) return str;
  try {
    return decodeURIComponent(escape(str));
  } catch (e) {
    return str;
  }
}

function arrayBufferToBinaryString(buf) {
  const bytes = new Uint8Array(buf);
  const chunkSize = 0x8000;
  let result = '';
  for (let i = 0; i < bytes.length; i += chunkSize) {
    result += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
  }
  return result;
}

// ====== File input ======
document.getElementById('file-input').addEventListener('change', handleFile, false);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();

  reader.onload = function (ev) {
    let binary;
    if (ev.target.result instanceof ArrayBuffer) {
      binary = arrayBufferToBinaryString(ev.target.result);
    } else {
      binary = ev.target.result;
    }

    let workbook;
    try {
      // Попытка с codepage — если сборка sheetjs поддерживает cptable
      workbook = XLSX.read(binary, { type: 'binary', codepage: 1251, raw: false });
    } catch (err) {
      // fallback
      try {
        workbook = XLSX.read(binary, { type: 'binary', raw: false });
      } catch (err2) {
        console.error('Ошибка чтения файла XLS/XLSX:', err2);
        alert('Не удалось прочитать файл: проверьте консоль (Ctrl+Shift+I).');
        return;
      }
    }

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

    const detected = detectDocType(json);
    currentConfig = getConfigForType(detected);
    totalRowsInSheet = json.length;

    items = parseWithConfig(json, currentConfig);

    console.log('detected document type:', detected, 'parsed items:', items.length);

    renderTable(totalRowsInSheet);
    currentIndex = 0;
    setDocTypeLabel(currentConfig ? currentConfig.label : 'Не определён');
  };

  reader.readAsArrayBuffer(file);
}

// ====== Detect + configs ======
function detectDocType(json) {
  const upto = Math.min(30, json.length);
  for (let i = 0; i < upto; i++) {
    const rowArr = (json[i] || []).map(cell => typeof cell === 'string' ? fixWin1251Mojibake(cell) : cell);
    const row = rowArr.join(' ').toString().toLowerCase();
    if (!row) continue;

    if (row.includes('расходная накладная') || row.includes('расходная')) return 'rashod';
    if (row.includes('счет-фактура') || row.includes('счёт-фактура')) return 'ufd';
    if (row.includes('универсальный передаточный документ') || row.includes('упд')) return 'ufd';
    if (row.includes('перемещение')) return 'perem';
    if (row.includes('счет') || row.includes('счёт') || row.includes('счет №')) return 'schet';
  }
  return 'rashod';
}

function getConfigForType(type) {
  const commonQty = [20, 22]; // U..W => 20..22
  switch (type) {
    case 'rashod':
      return { type: 'rashod', label: 'Расходная накладная', startRowIndex: 8, articleCols: [3, 19], qtyCols: commonQty, numCols: null };
    case 'schet':
      return { type: 'schet', label: 'Счёт', startRowIndex: 20, articleCols: [3, 19], qtyCols: commonQty, numCols: null };
    case 'perem':
      return { type: 'perem', label: 'Перемещение', startRowIndex: 8, articleCols: [3, 19], qtyCols: commonQty, numCols: [1,2] };
    case 'ufd':
      return { type: 'ufd', label: 'Счёт-фактура / УПД', startRowIndex: 16, articleCols: [16, 43], qtyCols: [68, 74], numCols: [11, 15] };
    default:
      return { type: 'unknown', label: 'Неизвестный', startRowIndex: 8, articleCols: [3, 19], qtyCols: commonQty, numCols: null };
  }
}

// ====== Parse sheet with config ======
function parseWithConfig(json, cfg) {
  const out = [];
  if (!cfg) return out;

  for (let i = cfg.startRowIndex; i < json.length; i++) {
    const row = json[i] || [];

    const articleRangeText = collectRangeText(row, cfg.articleCols[0], cfg.articleCols[1]);

    const primaryCellRaw = row[5] || '';
    const primaryCell = typeof primaryCellRaw === 'string' ? fixWin1251Mojibake(primaryCellRaw) : primaryCellRaw;

    const candidateText = (articleRangeText || '') + ' ' + (primaryCell || '');
    const fullArticleText = (candidateText || '').toString().trim();

    // qty
    let qty = 0;
    if (cfg.qtyCols && cfg.qtyCols.length === 2) {
      for (let c = cfg.qtyCols[0]; c <= cfg.qtyCols[1]; c++) {
        const raw = row[c];
        const rawStr = (raw !== undefined && raw !== null) ? String(raw).replace(',', '.') : '';
        const val = parseFloat(rawStr) || 0;
        qty = Math.max(qty, Math.floor(val));
      }
    }

    if (!fullArticleText || qty <= 0) continue;

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
      sheetRowIndex: i
    });
  }

  return out;
}

// ====== Helpers: collect text from range ======
function collectRangeText(row, startCol, endCol) {
  if (!row) return '';
  const parts = [];
  for (let c = startCol; c <= endCol; c++) {
    const v = row[c];
    if (v !== undefined && v !== null && String(v).trim() !== '') {
      let s = String(v).trim();
      s = fixWin1251Mojibake(s);
      parts.push(s);
    }
  }
  return parts.join(' ').replace(/\s+/g, ' ').trim();
}

// ====== Extract article by patterns ======
function extractArticleFromTextRange(row, startCol, endCol) {
  const pattern = /(KR|KU|КР|КУ|KLT|PT|РТ)[-.\s–—]?([\w\d.-]+)/i;
  for (let c = startCol; c <= endCol; c++) {
    let cell = row[c];
    if (cell === undefined || cell === null) continue;
    if (typeof cell !== 'string') cell = String(cell);
    cell = fixWin1251Mojibake(cell);
    const match = cell.match(pattern);
    if (match) {
      const prefix = match[1];
      const main = match[2] || null;
      if (prefix.toUpperCase() === 'PT') {
        return { article: collectRangeText(row, startCol, endCol), prefix: 'PT', main: null, extra: null };
      }
      return { article: `${prefix}-${main}`, prefix, main, extra: null };
    }
  }
  return null;
}

function extractArticleFromAnyCell(row) {
  const pattern = /(KR|KU|КР|КУ|KLT|PT|РТ)[-.\s–—]?([\w\d.-]+)/i;
  for (let cell of row) {
    if (cell === undefined || cell === null) continue;
    if (typeof cell !== 'string') cell = String(cell);
    cell = fixWin1251Mojibake(cell);
    const match = cell.match(pattern);
    if (match) {
      const prefix = match[1];
      const main = match[2] || null;
      if (prefix.toUpperCase() === 'PT') {
        const joined = row.map(c => (c === undefined || c === null) ? '' : fixWin1251Mojibake(String(c))).filter(Boolean).join(', ');
        return { article: joined, prefix: 'PT', main: null, extra: null };
      }
      return { article: `${prefix}-${main}`, prefix, main, extra: null };
    }
  }
  return null;
}

// ====== Render table ======
function renderTable(totalRows = null) {
  const tbody = document.querySelector("#items-table tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  items.forEach((item, idx) => {
    const rowEl = document.createElement("tr");
    if (idx === currentIndex) rowEl.classList.add("active-row");

    const td1 = document.createElement("td");
    td1.textContent = item.article || '';

    const td2 = document.createElement("td");
    td2.textContent = item.qty ?? '';

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
  const countEl = document.getElementById("count");
  if (countEl) countEl.textContent = text;
}

// ====== Speech & control ======
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
  if (!synth) return;
  const utterance = new SpeechSynthesisUtterance(text);
  utterance.lang = 'ru-RU';
  synth.cancel();
  synth.speak(utterance);
}

// ====== number -> words ======
function numberToWordsRuNom(num) {
  num = parseInt(num);
  if (isNaN(num)) return String(num);

  const ones = ["ноль","один","два","три","четыре","пять","шесть","семь","восемь","девять"];
  const teens = ["десять","одиннадцать","двенадцать","тринадцать","четырнадцать","пятнадцать","шестнадцать","семнадцать","восемнадцать","девятнадцать"];
  const tens = ["","","двадцать","тридцать","сорок","пятьдесят","шестьдесят","семьдесят","восемьдесят","девяносто"];
  const hundreds = ["","сто","двести","триста","четыреста","пятьсот","шестьсот","семьсот","восемьсот","девятьсот"];

  if (num < 10) return ones[num];
  if (num < 20) return teens[num-10];
  if (num < 100) {
    const t = Math.floor(num/10);
    const o = num%10;
    return tens[t] + (o ? " " + ones[o] : "");
  }
  if (num < 1000) {
    const h = Math.floor(num/100);
    const rem = num % 100;
    return hundreds[h] + (rem ? " " + numberToWordsRuNom(rem) : "");
  }
  return num.toString();
}

function numberToWordsRu(num) {
  num = parseInt(num);
  if (isNaN(num)) return String(num);

  const ones = ["ноль","одну","две","три","четыре","пять","шесть","семь","восемь","девять"];
  const teens = ["десять","одиннадцать","двенадцать","тринадцать","четырнадцать","пятнадцать","шестнадцать","семнадцать","восемнадцать","девятнадцать"];
  const tens = ["","","двадцать","тридцать","сорок","пятьдесят","шестьдесят","семьдесят","восемьдесят","девяносто"];
  const hundreds = ["","сто","двести","триста","четыреста","пятьсот","шестьсот","семьсот","восемьсот","девятьсот"];

  if (num < 10) return ones[num];
  if (num < 20) return teens[num-10];
  if (num < 100) {
    const t = Math.floor(num/10);
    const o = num%10;
    return tens[t] + (o ? " " + ones[o] : "");
  }
  if (num < 1000) {
    const h = Math.floor(num/100);
    const rem = num % 100;
    return hundreds[h] + (rem ? " " + numberToWordsRuNom(rem) : "");
  }
  return num.toString();
}

// ====== pronounce alphanumeric helper ======
function pronounceAlphanumeric(str) {
  if (!str) return '';

  const digitMap = {
    '0':'ноль','1':'один','2':'два','3':'три','4':'четыре','5':'пять','6':'шесть','7':'семь','8':'восемь','9':'девять'
  };
  const letterMap = {
    a:'эй', b:'би', c:'си', d:'ди', e:'и', f:'эф', g:'джи', h:'эйч', i:'ай', j:'джей', k:'кей', l:'эл', m:'эм',
    n:'эн', o:'о', p:'пи', q:'кью', r:'эр', s:'эс', t:'ти', u:'ю', v:'ви', w:'дабл-ю', x:'икс', y:'уай', z:'зед'
  };

  // если строка полностью цифр и не начинается с 0 — читаем как число
  if (/^\d+$/.test(str) && !str.startsWith('0')) {
    return numberToWordsRuNom(parseInt(str));
  }

  // иначе читаем по символам
  const parts = [];
  for (let i = 0; i < str.length; i++) {
    const ch = str[i];
    if (/\d/.test(ch)) {
      parts.push(digitMap[ch]);
    } else if (/[A-Za-z]/.test(ch)) {
      const low = ch.toLowerCase();
      parts.push(letterMap[low] || low);
    } else if (ch === '-' || ch === '–' || ch === '—' || ch === '.') {
      // пропускаем тире/точки в артикулах (в речи не нужны)
      continue;
    } else {
      parts.push(ch);
    }
  }
  return parts.join(' ').replace(/\s+/g, ' ').trim();
}

// ====== format article ======
function formatArticle(prefix, main, extra) {
  const upperPrefix = String(prefix || '').toUpperCase();
  const isKR = upperPrefix.includes("KR") || upperPrefix.includes("КР");
  const isKU = upperPrefix.includes("KU") || upperPrefix.includes("КУ");
  const isKLT = upperPrefix === "KLT";

  if (isKR) {
    const ruPrefix = "КаЭр";
    if (main && /[A-Za-z]/i.test(String(main))) {
      // буквы в main — читаем посимвольно
      const pronounced = pronounceAlphanumeric(String(main));
      return `${ruPrefix} ${pronounced}${extra ? ' дробь ' + pronounceAlphanumeric(String(extra)) : ''}`;
    }
    return `${ruPrefix} ${numberToWordsRuNom(main)}${extra ? ' дробь ' + numberToWordsRuNom(extra) : ''}`;
  }

  if (isKU) {
    const ruPrefix = "Кудо";
    if (!main) return ruPrefix;

    const raw = String(main);

    // Если main содержит буквы или содержит смешанные символы или начинается с 0 -> читаем посимвольно
    let pronouncedMain;
    if (/[A-Za-z]/.test(raw) || /[^0-9]/.test(raw) || raw.startsWith('0')) {
      pronouncedMain = pronounceAlphanumeric(raw);
    } else {
      // чистые цифры без ведущего нуля — пробуем привычное групповое чтение (оригинальная логика)
      let parts = [];
      if (raw.length === 4) parts = [raw.slice(0,2), raw.slice(2)];
      else if (raw.length === 5) parts = [raw.slice(0,2), raw.slice(2)];
      else if (raw.length === 6) parts = [raw.slice(0,2), raw.slice(2,4), raw.slice(4)];
      else parts = [raw];

      pronouncedMain = parts.map(p => {
        if (p.length === 2 && p.startsWith("0")) {
          // '01' -> "ноль один"
          const d = p[1];
          return 'ноль ' + numberToWordsRuNom(parseInt(d));
        }
        // безопасно парсим p как число (p должно быть только цифры здесь)
        const n = parseInt(p);
        return numberToWordsRuNom(isNaN(n) ? p : n);
      }).join(' ');
    }

    let pronouncedExtra = '';
    if (extra) pronouncedExtra = pronounceAlphanumeric(String(extra));

    return `${ruPrefix} ${pronouncedMain}${pronouncedExtra ? ' ' + pronouncedExtra : ''}`;
  }

  if (isKLT) {
    const ruPrefix = "КэЭлТэ";
    if (main && /[A-Za-z]/i.test(String(main))) {
      return `${ruPrefix} ${pronounceAlphanumeric(String(main))}${extra ? ' дробь ' + pronounceAlphanumeric(String(extra)) : ''}`;
    }
    return `${ruPrefix} ${numberToWordsRuNom(main)}${extra ? ' дробь ' + numberToWordsRuNom(extra) : ''}`;
  }

  // default fallback
  if (main && /[A-Za-z]/.test(String(main))) {
    return `${prefix} ${pronounceAlphanumeric(String(main))}${extra ? ' ' + pronounceAlphanumeric(String(extra)) : ''}`;
  }

  return `${prefix}${main ? '-' + main : ''}${extra ? '-' + extra : ''}`.replace(/^-/, '');
}

// ====== qty suffix ======
function getQtySuffix(num) {
  const rem10 = num % 10;
  const rem100 = num % 100;
  if (rem10 === 1 && rem100 !== 11) return "штуку";
  if ([2,3,4].includes(rem10) && ![12,13,14].includes(rem100)) return "штуки";
  return "штук";
}

// ====== speech recognition ======
function startListening() {
  if (!SpeechRecognition) {
    console.warn('SpeechRecognition не поддерживается в этом браузере');
    return;
  }
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
  try { recognition.stop(); } catch (e) {}
  isListening = false;
}

// ====== navigation & voice commands ======
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
  if (["готово","положил","ок"].some(k => cmd === k || cmd.includes(k))) {
    items[currentIndex].checked = true;
    speakNextUnprocessed();
  } else if (["дальше","пропускаем","некст","следующий"].some(k => cmd === k || cmd.includes(k))) {
    speakNextUnprocessed();
  } else if (cmd.includes("назад")) {
    currentIndex = Math.max(0, currentIndex - 1);
    speakCurrent();
  } else if (["повтори","щё раз","ещё раз","повторить"].some(k => cmd.includes(k))) {
    speakCurrent();
  } else if (cmd.includes("начни") && cmd.includes("пропущ")) {
    startFromSkipped();
  } else {
    console.log("Команда не распознана как управление:", cmd);
  }
  renderTable(totalRowsInSheet);
}

function startFromSkipped() {
  let next = items.findIndex(item => !item.checked);
  if (next !== -1) {
    currentIndex = next;
    speakCurrent();
  } else {
    speak("Все позиции уже обработаны.");
  }
}