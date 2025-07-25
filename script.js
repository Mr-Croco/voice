
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
    for (let i = 8; i < json.length; i++) {
      const row = json[i];
      if (!row) continue;

      const rawArticle = row[5];
      const u = parseInt(row[20]) || 0;
      const v = parseInt(row[21]) || 0;
      const w = parseInt(row[22]) || 0;
      const qty = Math.max(u, v, w);

      if (typeof rawArticle === 'string' && /(KR|KU|РљР |РљРЈ|KLT|Р Рў|PT)[-.\s]?\d+/i.test(rawArticle)) {
        const match = rawArticle.match(/(KR|KU|РљР |РљРЈ|KLT|Р Рў|PT)[-.\s]?(\d+)[-.]?(\d+)?/i);
        if (match) {
          items.push({
            article: match[0],
            prefix: match[1],
            main: match[2],
            extra: match[3] || null,
            qty,
            row, // в†ђ СЃРѕС…СЂР°РЅСЏРµРј СЃС‚СЂРѕРєСѓ РґР»СЏ РѕР·РІСѓС‡РєРё РїРѕР»РЅРѕСЃС‚СЊСЋ
            checked: false
          });
        }
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

  document.getElementById("count").textContent = `Р—Р°РіСЂСѓР¶РµРЅРѕ РїРѕР·РёС†РёР№: ${items.length}`;
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

if (["KR", "РљР ", "KU", "РљРЈ", "KLT"].includes(prefix.toUpperCase())) {
  articleText = formatArticle(prefix, main, extra);
} else {
  // РџСЂРѕС‡РёС‚Р°С‚СЊ РІСЃСЋ СЃС‚СЂРѕРєСѓ, РµСЃР»Рё СЌС‚Рѕ РЅРµ KR, KU РёР»Рё KLT
  articleText = items[currentIndex].row
  .filter(cell => cell && typeof cell === 'string')
  .join(' ')
  .replace(/\s+/g, ' ')
  .trim();
   }
  
  const qtyText = numberToWordsRu(qty);
  const qtyEnding = getQtySuffix(qty);
  const phrase = `${articleText} РїРѕР»РѕР¶РёС‚СЊ ${qtyText} ${qtyEnding}`;
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
  const ones = ["РЅРѕР»СЊ", "РѕРґРёРЅ", "РґРІР°", "С‚СЂРё", "С‡РµС‚С‹СЂРµ", "РїСЏС‚СЊ", "С€РµСЃС‚СЊ", "СЃРµРјСЊ", "РІРѕСЃРµРјСЊ", "РґРµРІСЏС‚СЊ"];
  const teens = ["РґРµСЃСЏС‚СЊ", "РѕРґРёРЅРЅР°РґС†Р°С‚СЊ", "РґРІРµРЅР°РґС†Р°С‚СЊ", "С‚СЂРёРЅР°РґС†Р°С‚СЊ", "С‡РµС‚С‹СЂРЅР°РґС†Р°С‚СЊ", "РїСЏС‚РЅР°РґС†Р°С‚СЊ", "С€РµСЃС‚РЅР°РґС†Р°С‚СЊ", "СЃРµРјРЅР°РґС†Р°С‚СЊ", "РІРѕСЃРµРјРЅР°РґС†Р°С‚СЊ", "РґРµРІСЏС‚РЅР°РґС†Р°С‚СЊ"];
  const tens = ["", "", "РґРІР°РґС†Р°С‚СЊ", "С‚СЂРёРґС†Р°С‚СЊ", "СЃРћСЂРѕРє", "РїСЏС‚СЊРґРµСЃСЏС‚", "С€РµСЃС‚СЊРґРµСЃСЏС‚", "СЃРµРјСЊРґРµСЃСЏС‚", "РІРѕСЃРµРјСЊРґРµСЃСЏС‚", "РґРµРІСЏРЅРѕСЃС‚Рѕ"];
  const hundreds = ["", "СЃС‚Рѕ", "РґРІРµСЃС‚Рё", "С‚СЂРёСЃС‚Р°", "С‡РµС‚С‹СЂРµСЃС‚Р°", "РїСЏС‚СЊСЃРѕС‚", "С€РµСЃС‚СЊСЃРѕС‚", "СЃРµРјСЊСЃРѕС‚", "РІРѕСЃРµРјСЊСЃРѕС‚", "РґРµРІСЏС‚СЊСЃРѕС‚"];

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
  const pattern = /(KR|KU|РљР |РљРЈ|KLT|Р Рў|PT)[-вЂ“]?(\d+)(?:[-вЂ“.]?(\d+))?/i;

  for (let cell of row) {
    const match = typeof cell === 'string' && cell.match(pattern);
    if (match) {
      const prefix = match[1].toUpperCase();

      // рџЋЇ РћСЃРѕР±С‹Р№ СЃР»СѓС‡Р°Р№: РµСЃР»Рё РїСЂРµС„РёРєСЃ PT в†’ РѕР·РІСѓС‡РёРІР°РµРј РІСЃСЋ СЃС‚СЂРѕРєСѓ
      if (prefix === "PT") {
        return row.filter(Boolean).join(", ");
      }

      // РЎС‚Р°РЅРґР°СЂС‚РЅР°СЏ РѕР·РІСѓС‡РєР° РїРѕ РїСЂРµС„РёРєСЃР°Рј
      return formatArticle(match[1], match[2], match[3]);
    }
  }

  return null;
}


function formatArticle(prefix, main, extra) {
  const upperPrefix = prefix.toUpperCase();
  const isKR = upperPrefix.includes("KR") || upperPrefix.includes("РљР ");
  const isKU = upperPrefix.includes("KU") || upperPrefix.includes("РљРЈ");

  if (isKR) {
    const ruPrefix = "РљР°Р­СЂ";
    return `${ruPrefix} ${numberToWordsRuNom(main)}${extra ? ' РґСЂРѕР±СЊ ' + numberToWordsRuNom(extra) : ''}`;
  }

  if (isKU) {
    const ruPrefix = "РљСѓРґРѕ";
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
        return "РЅРѕР»СЊ " + numberToWordsRuNom(p[1]);
      } else {
        return numberToWordsRuNom(parseInt(p));
      }
    }).join(" ");

    const isKLT = upperPrefix === "KLT";

    if (isKLT) {
  return `РљСЌР­Р»РўСЌ ${numberToWordsRuNom(main)}${extra ? ' РґСЂРѕР±СЊ ' + numberToWordsRuNom(extra) : ''}`;
}
    
    return `${ruPrefix} ${spoken}${extra ? ' ' + extra : ''}`;
  }

  return `${prefix}-${main}${extra ? '-' + extra : ''}`;
}

function numberToWordsRu(num) {
  num = parseInt(num);
  const ones = ["РЅРѕР»СЊ", "РѕРґРЅСѓ", "РґРІРµ", "С‚СЂРё", "С‡РµС‚С‹СЂРµ", "РїСЏС‚СЊ", "С€РµСЃС‚СЊ", "СЃРµРјСЊ", "РІРѕСЃРµРјСЊ", "РґРµРІСЏС‚СЊ"];
  const teens = ["РґРµСЃСЏС‚СЊ", "РѕРґРёРЅРЅР°РґС†Р°С‚СЊ", "РґРІРµРЅР°РґС†Р°С‚СЊ", "С‚СЂРёРЅР°РґС†Р°С‚СЊ", "С‡РµС‚С‹СЂРЅР°РґС†Р°С‚СЊ", "РїСЏС‚РЅР°РґС†Р°С‚СЊ", "С€РµСЃС‚РЅР°РґС†Р°С‚СЊ", "СЃРµРјРЅР°РґС†Р°С‚СЊ", "РІРѕСЃРµРјРЅР°РґС†Р°С‚СЊ", "РґРµРІСЏС‚РЅР°РґС†Р°С‚СЊ"];
  const tens = ["", "", "РґРІР°РґС†Р°С‚СЊ", "С‚СЂРёРґС†Р°С‚СЊ", "СЃРћСЂРѕРє", "РїСЏС‚СЊРґРµСЃСЏС‚", "С€РµСЃС‚СЊРґРµСЃСЏС‚", "СЃРµРјСЊРґРµСЃСЏС‚", "РІРѕСЃРµРјСЊРґРµСЃСЏС‚", "РґРµРІСЏРЅРѕСЃС‚Рѕ"];
  const hundreds = ["", "СЃС‚Рѕ", "РґРІРµСЃС‚Рё", "С‚СЂРёСЃС‚Р°", "С‡РµС‚С‹СЂРµСЃС‚Р°", "РїСЏС‚СЊСЃРѕС‚", "С€РµСЃС‚СЊСЃРѕС‚", "СЃРµРјСЊСЃРѕС‚", "РІРѕСЃРµРјСЊСЃРѕС‚", "РґРµРІСЏС‚СЊСЃРѕС‚"];

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
  if (rem10 === 1 && rem100 !== 11) return "С€С‚СѓРєСѓ";
  if ([2, 3, 4].includes(rem10) && ![12, 13, 14].includes(rem100)) return "С€С‚СѓРєРё";
  return "С€С‚СѓРє";
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
    console.log("Р Р°СЃРїРѕР·РЅР°РЅРѕ:", transcript);
    handleVoiceCommand(transcript);
  };

  recognition.onerror = function (event) {
    console.error("РћС€РёР±РєР° СЂР°СЃРїРѕР·РЅР°РІР°РЅРёСЏ:", event.error);
    if (event.error === "not-allowed" || event.error === "service-not-allowed") {
      isListening = false;
    }
  };

  recognition.onend = function () {
    console.log("РџСЂРѕСЃР»СѓС€РєР° Р·Р°РІРµСЂС€РµРЅР°");
    if (isListening) {
      setTimeout(() => recognition.start(), 300); // Р±РµР·РѕРїР°СЃРЅС‹Р№ РїРµСЂРµР·Р°РїСѓСЃРє
    }
  };

  isListening = true;
  recognition.start();
  console.log("РџСЂРѕСЃР»СѓС€РєР° Р·Р°РїСѓС‰РµРЅР°");
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
    speak("Р‘РѕР»СЊС€Рµ РЅРµРѕС‚РјРµС‡РµРЅРЅС‹С… РїРѕР·РёС†РёР№ РЅРµС‚.");
  }
}

function handleVoiceCommand(cmd) {
  console.log("Р Р°СЃРїРѕР·РЅР°РЅРѕ:", cmd);
  if (["РіРѕС‚РѕРІРѕ", "РїРѕР»РѕР¶РёР»", "РѕРє"].includes(cmd)) {
    items[currentIndex].checked = true;
    speakNextUnprocessed();
  } else if (["РґР°Р»СЊС€Рµ", "РїСЂРѕРїСѓСЃРєР°РµРј", "РЅРµРєСЃС‚"].includes(cmd)) {
    speakNextUnprocessed();
  } else if (cmd === "РЅР°Р·Р°Рґ") {
    currentIndex = Math.max(0, currentIndex - 1);
    speakCurrent();
  } else if (["РїРѕРІС‚РѕСЂРё", "РµС‰С‘ СЂР°Р·", "РїРѕРІС‚РѕСЂРёС‚СЊ"].includes(cmd)) {
    speakCurrent();
  } else if (cmd.includes("РЅР°С‡РЅРё") && cmd.includes("РїСЂРѕРїСѓС‰")) {
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
    speak("Р’СЃРµ РїРѕР·РёС†РёРё СѓР¶Рµ РѕР±СЂР°Р±РѕС‚Р°РЅС‹.");
  }
}
