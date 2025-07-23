
let data = [];
let currentIndex = 0;
let isSpeaking = false;
let startFromSkipped = false;

const synth = window.speechSynthesis;
const recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();
recognition.lang = 'ru-RU';
recognition.continuous = true;

document.getElementById('excel-file').addEventListener('change', handleFile);
document.getElementById('start-btn').addEventListener('click', () => {
  currentIndex = 0;
  startFromSkipped = false;
  speakCurrent();
});
document.getElementById('start-skipped-btn').addEventListener('click', () => {
  currentIndex = 0;
  startFromSkipped = true;
  speakNextUnprocessed();
});

function handleFile(event) {
  const file = event.target.files[0];
  const reader = new FileReader();
  reader.onload = (e) => {
    const dataBinary = new Uint8Array(e.target.result);
    const workbook = XLSX.read(dataBinary, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    data = json.map((row, index) => ({ row, index, checked: false }));
    renderTable();
  };
  reader.readAsArrayBuffer(file);
}

function speak(text, callback) {
  if (synth.speaking) {
    synth.cancel();
  }
  const utter = new SpeechSynthesisUtterance(text);
  utter.lang = 'ru-RU';
  utter.onend = callback;
  synth.speak(utter);
}

function numberToWordsRu(num) {
  num = parseInt(num);
  const ones = ["ноль", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"];
  const teens = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"];
  const tens = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"];
  const hundreds = ["", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"];

  if (num < 10) return ones[num];
  if (num < 20) return teens[num - 10];
  if (num < 100) return tens[Math.floor(num / 10)] + (num % 10 ? " " + ones[num % 10] : "");
  if (num < 1000) return hundreds[Math.floor(num / 100)] + (num % 100 ? " " + numberToWordsRu(num % 100) : "");
  return num.toString();
}

function numberToWordsRuNom(num) {
  num = parseInt(num);
  const ones = ["ноль", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"];
  const teens = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"];
  const tens = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"];
  const hundreds = ["", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"];

  if (num < 10) return ones[num];
  if (num < 20) return teens[num - 10];
  if (num < 100) return tens[Math.floor(num / 10)] + (num % 10 ? " " + ones[num % 10] : "");
  if (num < 1000) return hundreds[Math.floor(num / 100)] + (num % 100 ? " " + numberToWordsRuNom(num % 100) : "");
  return num.toString();
}

function getQtySuffix(num) {
  const last = num % 10;
  const secondLast = Math.floor((num % 100) / 10);
  if (last === 1 && secondLast !== 1) return "штуку";
  if ([2, 3, 4].includes(last) && secondLast !== 1) return "штуки";
  return "штук";
}

function formatArticle(prefix, main, extra) {
  const upperPrefix = prefix.toUpperCase();
  const isKR = upperPrefix === "KR" || upperPrefix === "КР";
  const isKU = upperPrefix === "KU" || upperPrefix === "КУ";
  const isKLT = upperPrefix === "KLT";
  const isRT = upperPrefix === "РТ" || upperPrefix === "PT";

  if (isKR) {
    return `КаЭр ${numberToWordsRuNom(main)}${extra ? ' дробь ' + numberToWordsRuNom(extra) : ''}`;
  }

  if (isKU) {
    const raw = main.toString();
    const ruPrefix = "Кудо";
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
        return "ноль " + numberToWordsRu(p[1]);
      } else {
        return numberToWordsRuNom(parseInt(p));
      }
    }).join(" ");

    return `${ruPrefix} ${spoken}${extra ? ' ' + extra : ''}`;
  }

  if (isKLT) {
    return `КэЭлТэ ${numberToWordsRuNom(main)}`;
  }

  if (isRT) {
    return `Эртэ ${numberToWordsRuNom(main)}`;
  }

  if (main.toLowerCase().includes("маркер")) {
    return `Маркер`;
  }

  return `${prefix}-${main}${extra ? '-' + extra : ''}`;
}

function speakCurrent() {
  while (currentIndex < data.length && (!data[currentIndex] || data[currentIndex].checked)) {
    currentIndex++;
  }
  if (currentIndex >= data.length) return;
  const row = data[currentIndex].row;
  const text = row.filter(Boolean).join(", ");
  const qty = extractQuantity(row);
  const article = extractArticle(row);
  const spoken = article ? `${article} положить ${numberToWordsRu(qty)} ${getQtySuffix(qty)}` : text;
  isSpeaking = true;
  speak(spoken, () => { isSpeaking = false; });
}

function speakNextUnprocessed() {
  do {
    currentIndex++;
  } while (currentIndex < data.length && (data[currentIndex]?.checked));
  speakCurrent();
}

function extractArticle(row) {
  const pattern = /(KR|KU|КР|КУ|KLT|РТ|PT)[-–]?(\d+)(?:[-–.]?(\d+))?/i;
  for (let cell of row) {
    const match = typeof cell === 'string' && cell.match(pattern);
    if (match) {
      return formatArticle(match[1], match[2], match[3]);
    }
  }
  return null;
}

function extractQuantity(row) {
  return parseInt(row.slice(20, 23).filter(x => !isNaN(x)).join("")) || 1;
}

function renderTable() {
  const table = document.getElementById("table-body");
  table.innerHTML = "";
  data.forEach((item, index) => {
    const tr = document.createElement("tr");
    const tdIndex = document.createElement("td");
    tdIndex.textContent = index + 1;
    const tdData = document.createElement("td");
    tdData.textContent = item.row.filter(Boolean).join(", ");
    const tdCheck = document.createElement("td");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = item.checked;
    checkbox.addEventListener("change", () => {
      item.checked = checkbox.checked;
    });
    tdCheck.appendChild(checkbox);
    tr.appendChild(tdIndex);
    tr.appendChild(tdData);
    tr.appendChild(tdCheck);
    table.appendChild(tr);
  });
}

recognition.onresult = (event) => {
  const cmd = event.results[event.results.length - 1][0].transcript.trim().toLowerCase();
  if (["готово", "ок", "положил"].includes(cmd)) {
    data[currentIndex].checked = true;
    renderTable();
    currentIndex++;
    speakCurrent();
  } else if (["дальше", "пропускаем", "некст"].includes(cmd)) {
    speakNextUnprocessed();
  } else if (["назад"].includes(cmd)) {
    if (currentIndex > 0) {
      currentIndex--;
      speakCurrent();
    }
  } else if (["повтори", "ещё раз"].includes(cmd)) {
    speakCurrent();
  }
};

recognition.start();
