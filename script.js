
let items = [];
let currentIndex = 0;

document.getElementById('file-input').addEventListener('change', handleFile, false);

function handleFile(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    items = [];
    for (let i = 10; i < sheetData.length; i++) {
  const row = sheetData[i];
  if (!row) continue;

  // Получаем текст товара из колонок F–T
  const fullRowText = row.slice(5, 20)
    .filter(cell => typeof cell === 'string' && cell.trim())
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();

  if (!fullRowText) continue;

  // Считаем количество из колонок U, V, W
  const qtyRaw = [row[20], row[21], row[22]]
    .filter(v => typeof v === 'number' || (typeof v === 'string' && v.trim()))
    .map(Number)
    .reduce((a, b) => a + b, 0);

  const qty = Math.round(qtyRaw);

  // Ищем артикула KU/KR/KLT
  const match = fullRowText.match(/(KU|KR|KLT)[-.\s]?(\d+)[-.]?(\d+)?/i);

  if (match) {
    items.push({
      article: match[0],
      prefix: match[1],
      main: match[2],
      extra: match[3] || null,
      qty,
      checked: false
    });
  } else {
    // Все прочие позиции читаем как есть
    items.push({
      article: fullRowText,
      prefix: null,
      main: fullRowText,
      extra: null,
      qty,
      checked: false
    });
  }
}

    renderTable();
  };

  reader.readAsArrayBuffer(file);
}

function renderTable() {
  const tbody = document.querySelector("#items-table tbody");
  tbody.innerHTML = "";

  items.forEach((item, idx) => {
    const row = document.createElement("tr");
    if (idx === currentIndex) row.classList.add("active-row");

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
      renderTable();
    });
    td3.appendChild(checkbox);

    row.appendChild(td1);
    row.appendChild(td2);
    row.appendChild(td3);
    row.addEventListener("click", (e) => {
      if (e.target.tagName.toLowerCase() === "input") return;
      currentIndex = idx;
      speakCurrent();
    });

    tbody.appendChild(row);
  });

  document.getElementById("count").textContent = `Загружено позиций: ${items.length}`;
}

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
  if (!items[currentIndex]) return;

  const { prefix, main, extra, qty } = items[currentIndex];
  let articleText;

  if (["KR", "КР", "KU", "КУ", "KLT"].includes((prefix || "").toUpperCase())) {
    articleText = formatArticle(prefix, main, extra);
  } else {
    articleText = items[currentIndex].article;
  }
  
  const qtyText = numberToWordsRu(qty);
  const qtyEnding = getQtySuffix(qty);
  const phrase = `${articleText} положить ${qtyText} ${qtyEnding}`;
  speak(phrase);
  renderTable();
}

function speak(text) {
  const utterance = new SpeechSynthesisUtterance(text);
  utterance.lang = 'ru-RU';
  synth.cancel();
  synth.speak(utterance);
}

function numberToWordsRuNom(num) {
  num = parseInt(num);
  const ones = ["ноль", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"];
  const teens = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"];
  const tens = ["", "", "двадцать", "тридцать", "сОрок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"];
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

function extractArticle(row) {
  const pattern = /(KR|KU|КР|КУ|KLT)[-–]?(\d+)(?:[-–.]?(\d+))?/i;

  for (let cell of row) {
    const match = typeof cell === 'string' && cell.match(pattern);
    if (match) {
      const prefix = match[1].toUpperCase();


      // Стандартная озвучка по префиксам
      return formatArticle(match[1], match[2], match[3]);
    }
  }

  return null;
}


function formatArticle(prefix, main, extra) {
  if (!prefix) {
  return main; // если нет префикса, то произносим как есть
}
  const upperPrefix = prefix.toUpperCase();
  const isKR = upperPrefix.includes("KR") || upperPrefix.includes("КР");
  const isKU = upperPrefix.includes("KU") || upperPrefix.includes("КУ");

  if (isKR) {
    const ruPrefix = "КаЭр";
    return `${ruPrefix} ${numberToWordsRuNom(main)}${extra ? ' дробь ' + numberToWordsRuNom(extra) : ''}`;
  }

  if (isKU) {
    const ruPrefix = "Кудо";
    const raw = main.toString();
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

    const isKLT = upperPrefix === "KLT";

    if (isKLT) {
  return `КаЭЛТЭ ${numberToWordsRuNom(main)}${extra ? ' дробь ' + numberToWordsRuNom(extra) : ''}`;
}
    
    return `${ruPrefix} ${spoken}${extra ? ' ' + extra : ''}`;
  }

  return `${prefix}-${main}${extra ? '-' + extra : ''}`;
}

function numberToWordsRu(num) {
  num = parseInt(num);
  const ones = ["ноль", "одну", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"];
  const teens = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"];
  const tens = ["", "", "двадцать", "тридцать", "сОрок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"];
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

function getQtySuffix(num) {
  const rem10 = num % 10;
  const rem100 = num % 100;
  if (rem10 === 1 && rem100 !== 11) return "штуку";
  if ([2, 3, 4].includes(rem10) && ![12, 13, 14].includes(rem100)) return "штуки";
  return "штук";
}

function startListening() {
  if (isListening) return;

  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
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
      setTimeout(() => recognition.start(), 300); // безопасный перезапуск
    }
  };

  isListening = true;
  recognition.start();
  console.log("Прослушка запущена");
}

function speakNextUnprocessed() {
  let next = currentIndex + 1;
  while (next < items.length && items[next].checked) {
    next++;
  }
  if (next < items.length) {
    currentIndex = next;
    speakCurrent();
  } else {
    speak("Больше неотмеченных позиций нет.");
  }
}

function handleVoiceCommand(cmd) {
  console.log("Распознано:", cmd);
  if (["готово", "положил", "ок"].includes(cmd)) {
    items[currentIndex].checked = true;
    speakNextUnprocessed();
  } else if (["дальше", "пропускаем", "некст"].includes(cmd)) {
    speakNextUnprocessed();
  } else if (cmd === "назад") {
    currentIndex = Math.max(0, currentIndex - 1);
    speakCurrent();
  } else if (["повтори", "ещё раз", "повторить"].includes(cmd)) {
    speakCurrent();
  } else if (cmd.includes("начни") && cmd.includes("пропущ")) {
    startFromSkipped();
  }

  renderTable();
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
