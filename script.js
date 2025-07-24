
let items = [];
let currentIndex = 0;

document.getElementById('file-input').addEventListener('change', handleFile, false);

function handleFile(event) {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = (e) => {
    const dataBinary = new Uint8Array(e.target.result);
    const workbook = XLSX.read(dataBinary, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    items = json
      .slice(10) // ⛔ Пропускаем первые 10 строк
      .map((row, index) => {
        const article = extractArticle(row);
        const qtyRaw = row.slice(20, 23).filter(x => !isNaN(x)).join("");
        const qty = parseInt(qtyRaw) || 1;

        if (article) {
          return {
            index: index + 10,
            article,
            qty,
            checked: false,
            skipped: false,
            rawRow: row
          };
        } else {
          return null;
        }
      })
      .filter(x => x !== null);

    console.log("Загружено позиций:", items.length);
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

  const { article, qty } = items[currentIndex];
  const qtyText = numberToWordsRu(qty);
  const qtyEnding = getQtySuffix(qty);
  const phrase = `${article} положить ${qtyText} ${qtyEnding}`;
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
  const pattern = /(KR|KU|КР|КУ|KLT|РТ|PT)[-–]?(\d{2,6})(?:[-–.]?(\d+))?/i;

  for (let cell of row) {
    if (typeof cell !== 'string') continue;
    const match = cell.match(pattern);
    if (match) {
      const prefix = match[1].toUpperCase();

      if (prefix === "PT") {
        return row.filter(Boolean).join(", ");
      }

      return formatArticle(match[1], match[2], match[3]);
    }
  }

  return null;
}

function formatArticle(prefix, main, extra) {
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

  recognition = new SpeechRecognition();
  recognition.lang = 'ru-RU';
  recognition.interimResults = false;
  recognition.continuous = true;

  recognition.onresult = (event) => {
    const transcript = event.results[event.results.length - 1][0].transcript.trim().toLowerCase();
    handleVoiceCommand(transcript);
  };

  recognition.onend = () => {
    if (isListening) recognition.start();
  };

  isListening = true;
  recognition.start();
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
