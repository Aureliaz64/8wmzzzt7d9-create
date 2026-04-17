// Color buckets based on team-level statistics discussed in S01.
// You can replace counts after final data cleaning.
const defaultColorStats = [
  { key: "red", zh: "红", en: "Red", count: 534 },
  { key: "black", zh: "黑", en: "Black", count: 327 },
  { key: "yellow", zh: "黄", en: "Yellow", count: 169 },
  { key: "gray", zh: "灰", en: "Gray", count: 46 },
  { key: "white", zh: "白", en: "White", count: 0 },
  { key: "green", zh: "绿", en: "Green", count: 0 },
  { key: "blue", zh: "蓝", en: "Blue", count: 0 },
];

// CN palette references aligned to traditional-color style naming.
// Ref source: https://vivacolor.art/colors#google_vignette
const cnPalette = [
  { colorZh: "朱砂", colorEn: "Cinnabar", hex: "#d83828", bucket: "red" },
  { colorZh: "绯红", colorEn: "Crimson Red", hex: "#c02c38", bucket: "red" },
  { colorZh: "玄青", colorEn: "Dark Cyan", hex: "#2f3b4a", bucket: "black" },
  { colorZh: "玄黑", colorEn: "Mystic Black", hex: "#1b1b1b", bucket: "black" },
  { colorZh: "月白", colorEn: "Moon White", hex: "#d6ecf0", bucket: "white" },
  { colorZh: "秋香", colorEn: "Autumn Beige", hex: "#d4b98c", bucket: "yellow" },
  { colorZh: "秋黄", colorEn: "Autumn Yellow", hex: "#ddb772", bucket: "yellow" },
  { colorZh: "竹青", colorEn: "Bamboo Green", hex: "#789262", bucket: "green" },
];

// Western classic color references in Pantone-style communication.
// Ref source: https://www.pantonecn.com/pantone-connect
const westernPalette = [
  { colorZh: "猩红", colorEn: "Scarlet", hex: "#c41e3a", bucket: "red" },
  { colorZh: "酒红", colorEn: "Burgundy", hex: "#800020", bucket: "red" },
  { colorZh: "乌木黑", colorEn: "Ebony", hex: "#1c1b1a", bucket: "black" },
  { colorZh: "炭灰", colorEn: "Charcoal Gray", hex: "#36454f", bucket: "gray" },
  { colorZh: "象牙白", colorEn: "Ivory", hex: "#fffff0", bucket: "white" },
  { colorZh: "柠檬黄", colorEn: "Lemon Yellow", hex: "#f7ea48", bucket: "yellow" },
  { colorZh: "橄榄绿", colorEn: "Olive", hex: "#708238", bucket: "green" },
  { colorZh: "群青", colorEn: "Ultramarine", hex: "#4166f5", bucket: "blue" },
];

// Example quote records (replace with your cleaned table export).
const defaultQuoteRecords = [
  {
    bucket: "red",
    quote: "我仿佛看见冬花开在雪野中，有许多蜜蜂们忙碌地飞着。",
    work: "《野草·雪》",
    year: "1925",
    source: "鲁迅全集（定稿时补页码）",
  },
  {
    bucket: "yellow",
    quote: "脸上瘦削不堪，黄中带灰，像是被年月风干了。",
    work: "《祝福》",
    year: "1924",
    source: "鲁迅全集（定稿时补页码）",
  },
  {
    bucket: "black",
    quote: "铁屋子一般，沉沉地压着人。",
    work: "《呐喊·自序》",
    year: "1922",
    source: "鲁迅全集（定稿时补页码）",
  },
  {
    bucket: "gray",
    quote: "天是青白的，云是灰白的，路上行人却都匆匆。",
    work: "《故乡》",
    year: "1921",
    source: "鲁迅全集（定稿时补页码）",
  },
];

let colorStats = structuredClone(defaultColorStats);
let quoteRecords = structuredClone(defaultQuoteRecords);
let bucketMap = Object.fromEntries(colorStats.map((i) => [i.key, i]));

/** 弹窗内引文分页 */
const QUOTE_PAGE_SIZE = 8;
const quoteIndexByBucket = new Map();
let modalQuoteState = {
  bucketKey: "",
  page: 0,
  colorZh: "",
  colorEn: "",
  hex: "",
  systemName: "",
};

function rebuildQuoteIndex() {
  quoteIndexByBucket.clear();
  quoteRecords.forEach((q) => {
    const b = q.bucket;
    if (!quoteIndexByBucket.has(b)) {
      quoteIndexByBucket.set(b, []);
    }
    quoteIndexByBucket.get(b).push(q);
  });
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function firstQuoteForBucket(bucketKey) {
  const arr = quoteIndexByBucket.get(bucketKey);
  return arr && arr.length ? arr[0] : null;
}

function applyPreloadedData() {
  const payload = window.PRELOADED_DATA;
  if (!payload || typeof payload !== "object") {
    return false;
  }
  if (Array.isArray(payload.stats)) {
    payload.stats.forEach((s) => {
      const hit = colorStats.find((c) => c.key === s.key);
      if (hit && s.count !== undefined) {
        hit.count = Number(s.count) || hit.count;
      }
    });
  }
  if (Array.isArray(payload.quotes) && payload.quotes.length) {
    quoteRecords = payload.quotes.map((q) => ({
      bucket: String(q.bucket || "gray"),
      quote: String(q.quote || ""),
      work: String(q.work || "待补作品名"),
      year: String(q.year || "待补年份"),
      source: String(q.source || "预载入数据"),
    }));
  }
  bucketMap = Object.fromEntries(colorStats.map((i) => [i.key, i]));
  return true;
}

const cnGrid = document.getElementById("cnGrid");
const westernGrid = document.getElementById("westernGrid");
const searchInput = document.getElementById("searchInput");
const systemFilter = document.getElementById("systemFilter");
const detailModal = document.getElementById("detailModal");
const modalContent = document.getElementById("modalContent");
const closeModal = document.getElementById("closeModal");
const dataFileInput = document.getElementById("dataFileInput");
const importStatus = document.getElementById("importStatus");
const loadSampleBtn = document.getElementById("loadSampleBtn");
const downloadTemplateBtn = document.getElementById("downloadTemplateBtn");
const sheetPickerRow = document.getElementById("sheetPickerRow");
const statsSheetSelect = document.getElementById("statsSheetSelect");
const quotesSheetSelect = document.getElementById("quotesSheetSelect");
const applySheetsBtn = document.getElementById("applySheetsBtn");
const statCards = document.getElementById("statCards");
const colorBars = document.getElementById("colorBars");
const yearBars = document.getElementById("yearBars");
const headlineInsight = document.getElementById("headlineInsight");
const storyTimeline = document.getElementById("storyTimeline");
const exportPosterBtn = document.getElementById("exportPosterBtn");
let currentWorkbook = null;

const bucketColorMap = {
  red: "#c6362a",
  black: "#373737",
  yellow: "#d6aa4f",
  gray: "#7c7e86",
  white: "#d7dce3",
  green: "#5e8d62",
  blue: "#4b74d8",
};

function textColorByBg(hex) {
  const c = hex.replace("#", "");
  const r = parseInt(c.substring(0, 2), 16);
  const g = parseInt(c.substring(2, 4), 16);
  const b = parseInt(c.substring(4, 6), 16);
  const yiq = (r * 299 + g * 587 + b * 114) / 1000;
  return yiq >= 140 ? "#111111" : "#f7f2ea";
}


function setStatus(text, isError = false) {
  importStatus.textContent = text;
  importStatus.style.color = isError ? "#f08d88" : "#e6b98f";
}

function normalizeBucket(raw) {
  const v = String(raw || "").trim().toLowerCase();
  const map = {
    red: "red",
    "红": "red",
    black: "black",
    "黑": "black",
    yellow: "yellow",
    "黄": "yellow",
    gray: "gray",
    grey: "gray",
    "灰": "gray",
    white: "white",
    "白": "white",
    green: "green",
    "绿": "green",
    blue: "blue",
    "蓝": "blue",
    "青": "blue",
  };
  return map[v] || null;
}

function inferBucketFromText(raw) {
  const v = String(raw || "");
  if (!v) {
    return null;
  }
  if (v.includes("红")) return "red";
  if (v.includes("黑")) return "black";
  if (v.includes("黄")) return "yellow";
  if (v.includes("灰")) return "gray";
  if (v.includes("白")) return "white";
  if (v.includes("绿")) return "green";
  if (v.includes("蓝") || v.includes("青")) return "blue";
  return normalizeBucket(v);
}

function getField(row, keys) {
  for (const key of keys) {
    if (row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== "") {
      return row[key];
    }
  }
  return "";
}

function isKeywordSheetName(name) {
  return /关键词/.test(String(name || ""));
}

function isLikelyS02StatsRows(rows) {
  if (!rows || !rows.length) {
    return false;
  }
  const first = rows[0];
  const keys = Object.keys(first);
  const yearLike = keys.filter((k) => /^\d{4}$/.test(String(k).trim())).length;
  return yearLike >= 8;
}

function parseCsv(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.trim());
  if (lines.length < 2) {
    throw new Error("CSV 至少需要标题行和一行数据。");
  }
  const headers = lines[0].split(",").map((h) => h.trim().toLowerCase());
  const rows = [];
  for (let i = 1; i < lines.length; i += 1) {
    const cells = lines[i].split(",").map((c) => c.trim());
    const row = {};
    headers.forEach((h, idx) => {
      row[h] = cells[idx] || "";
    });
    rows.push(row);
  }
  return rows;
}

function applyImportedRows(rows) {
  const statsClone = structuredClone(defaultColorStats);
  const quoteList = [];
  rows.forEach((row) => {
    const bucketRaw = getField(row, [
      "bucket",
      "color",
      "colorbucket",
      "颜色",
      "色类",
      "统计色类",
      "颜色类别",
    ]);
    const quoteRaw = getField(row, [
      "quote",
      "text",
      "citation",
      "原文",
      "引文",
      "例句",
      "句子",
    ]);
    const workRaw = getField(row, ["work", "title", "book", "作品", "篇目", "出处作品"]);
    const yearRaw = getField(row, ["year", "date", "年份", "创作年份", "年代"]);
    const sourceRaw = getField(row, ["source", "ref", "出处", "来源", "参考"]);
    const countRaw = getField(row, ["count", "freq", "频次", "次数", "统计数量"]);

    const bucket = normalizeBucket(bucketRaw);
    if (!bucket) {
      return;
    }
    const statItem = statsClone.find((s) => s.key === bucket);
    if (countRaw && statItem) {
      statItem.count = Number(countRaw) || statItem.count;
    }
    if (quoteRaw) {
      quoteList.push({
        bucket,
        quote: String(quoteRaw),
        work: String(workRaw || "待补作品名"),
        year: String(yearRaw || "待补年份"),
        source: String(sourceRaw || "待补出处"),
      });
    }
  });
  colorStats = statsClone;
  quoteRecords = quoteList.length ? quoteList : structuredClone(defaultQuoteRecords);
  bucketMap = Object.fromEntries(colorStats.map((i) => [i.key, i]));
  rebuildQuoteIndex();
}

function applyStatsRows(rows, sheetName = "") {
  const statsClone = structuredClone(defaultColorStats);
  rows.forEach((row) => {
    const bucketRaw = getField(row, [
      "bucket",
      "color",
      "colorbucket",
      "颜色",
      "色类",
      "统计色类",
      "颜色类别",
    ]);
    const countRaw = getField(row, ["count", "freq", "频次", "次数", "统计数量"]);
    let bucket = normalizeBucket(bucketRaw);
    if (!bucket) {
      bucket = inferBucketFromText(getField(row, Object.keys(row)));
    }
    if (!bucket) {
      return;
    }
    const statItem = statsClone.find((s) => s.key === bucket);
    if (!statItem) {
      return;
    }
    if (countRaw) {
      statItem.count = Number(countRaw) || statItem.count;
    } else {
      // S02-like wide table: pick the largest numeric cell as total.
      const numericValues = Object.values(row)
        .map((v) => Number(v))
        .filter((n) => Number.isFinite(n));
      if (numericValues.length) {
        statItem.count = Math.max(...numericValues);
      }
    }
  });
  if (!rows.length && sheetName) {
    const b = inferBucketFromText(sheetName);
    if (b) {
      const item = statsClone.find((s) => s.key === b);
      if (item && item.count === 0) item.count = 1;
    }
  }
  colorStats = statsClone;
  bucketMap = Object.fromEntries(colorStats.map((i) => [i.key, i]));
}

function applyQuoteRows(rows, sheetName = "") {
  const quoteList = [];
  const sheetBucket = inferBucketFromText(sheetName);
  rows.forEach((row) => {
    const bucketRaw = getField(row, [
      "bucket",
      "color",
      "colorbucket",
      "颜色",
      "色类",
      "统计色类",
      "颜色类别",
    ]);
    const quoteRaw = getField(row, [
      "quote",
      "text",
      "citation",
      "原文",
      "引文",
      "例句",
      "句子",
      "所在句子",
      "句段",
    ]);
    const workRaw = getField(row, [
      "work",
      "title",
      "book",
      "作品",
      "篇目",
      "出处作品",
      "所属篇章",
    ]);
    const yearRaw = getField(row, ["year", "date", "年份", "创作年份", "年代"]);
    const sourceRaw = getField(row, ["source", "ref", "出处", "来源", "参考"]);
    const bucket = normalizeBucket(bucketRaw) || sheetBucket || inferBucketFromText(quoteRaw);
    if (!bucket || !quoteRaw) {
      return;
    }
    quoteList.push({
      bucket,
      quote: String(quoteRaw),
      work: String(workRaw || "待补作品名"),
      year: String(yearRaw || "待补年份"),
      source: String(sourceRaw || "待补出处"),
    });
  });
  quoteRecords = quoteList.length ? quoteList : structuredClone(defaultQuoteRecords);
  rebuildQuoteIndex();
}

function applyQuoteRowsFromWorkbook(workbook, sheetNames) {
  const all = [];
  sheetNames.forEach((name) => {
    const ws = workbook.Sheets[name];
    const rows = window.XLSX.utils.sheet_to_json(ws, { defval: "" });
    rows.forEach((r) => {
      all.push({ ...r, __sheetName: name });
    });
  });
  const quoteList = [];
  all.forEach((row) => {
    const sheetName = row.__sheetName || "";
    const sheetBucket = inferBucketFromText(sheetName);
    const bucketRaw = getField(row, [
      "bucket",
      "color",
      "colorbucket",
      "颜色",
      "色类",
      "统计色类",
      "颜色类别",
      "关键词",
      "黑色",
    ]);
    const quoteRaw = getField(row, [
      "quote",
      "text",
      "citation",
      "原文",
      "引文",
      "例句",
      "句子",
      "所在句子",
      "句段",
    ]);
    const workRaw = getField(row, ["work", "title", "book", "作品", "篇目", "出处作品", "所属篇章"]);
    const yearRaw = getField(row, ["year", "date", "年份", "创作年份", "年代"]);
    const sourceRaw = getField(row, ["source", "ref", "出处", "来源", "参考"]);
    const bucket = normalizeBucket(bucketRaw) || sheetBucket || inferBucketFromText(quoteRaw);
    if (!bucket || !quoteRaw) return;
    quoteList.push({
      bucket,
      quote: String(quoteRaw),
      work: String(workRaw || "待补作品名"),
      year: String(yearRaw || "待补年份"),
      source: String(sourceRaw || "待补出处"),
    });
  });
  if (quoteList.length) {
    quoteRecords = quoteList;
    rebuildQuoteIndex();
  }
}

function fillSheetSelects(sheetNames) {
  statsSheetSelect.innerHTML = "";
  quotesSheetSelect.innerHTML = "";
  sheetNames.forEach((name, idx) => {
    const s1 = document.createElement("option");
    s1.value = name;
    s1.textContent = name;
    statsSheetSelect.appendChild(s1);

    const s2 = document.createElement("option");
    s2.value = name;
    s2.textContent = name;
    quotesSheetSelect.appendChild(s2);

    if (idx === 0) {
      statsSheetSelect.value = name;
    }
    if (idx === 1) {
      quotesSheetSelect.value = name;
    }
  });
  if (sheetNames.length === 1) {
    quotesSheetSelect.value = sheetNames[0];
  }
}

function scoreStatsSheet(name, rows) {
  let score = 0;
  const lowerName = String(name).toLowerCase();
  if (
    /s02|统计|count|频次|年份|year|annual|timeline|数量|总表|汇总/.test(lowerName)
  ) {
    score += 6;
  }
  const sample = rows.slice(0, 30);
  sample.forEach((row) => {
    const hasBucket =
      normalizeBucket(
        getField(row, ["bucket", "color", "colorbucket", "颜色", "色类", "统计色类", "颜色类别"])
      ) || inferBucketFromText(getField(row, Object.keys(row)));
    const countRaw = getField(row, ["count", "freq", "频次", "次数", "统计数量"]);
    if (hasBucket) {
      score += 1;
    }
    if (countRaw && !Number.isNaN(Number(countRaw))) {
      score += 2;
    }
  });
  return score;
}

function scoreQuoteSheet(name, rows) {
  let score = 0;
  const lowerName = String(name).toLowerCase();
  if (/s01|引文|语料|句|quote|citation|原文|出处|文本/.test(lowerName)) {
    score += 6;
  }
  const sample = rows.slice(0, 30);
  sample.forEach((row) => {
    const hasBucket =
      normalizeBucket(
        getField(row, ["bucket", "color", "colorbucket", "颜色", "色类", "统计色类", "颜色类别"])
      ) || inferBucketFromText(name);
    const quoteRaw = getField(row, [
      "quote",
      "text",
      "citation",
      "原文",
      "引文",
      "例句",
      "句子",
      "所在句子",
    ]);
    const workRaw = getField(row, ["work", "title", "book", "作品", "篇目", "出处作品", "所属篇章"]);
    if (hasBucket) {
      score += 1;
    }
    if (quoteRaw && String(quoteRaw).length >= 8) {
      score += 3;
    }
    if (workRaw) {
      score += 1;
    }
  });
  return score;
}

function guessSheetRoles(workbook) {
  const names = workbook.SheetNames || [];
  if (!names.length) {
    return { stats: "", quotes: "" };
  }
  const analyzed = names.map((name) => {
    const ws = workbook.Sheets[name];
    const rows = window.XLSX.utils.sheet_to_json(ws, { defval: "" });
    return {
      name,
      statsScore: scoreStatsSheet(name, rows),
      quoteScore: scoreQuoteSheet(name, rows),
    };
  });

  const statsBest = [...analyzed].sort((a, b) => b.statsScore - a.statsScore)[0];
  const quoteBest = [...analyzed].sort((a, b) => b.quoteScore - a.quoteScore)[0];

  let stats = statsBest ? statsBest.name : names[0];
  let quotes = quoteBest ? quoteBest.name : names[Math.min(1, names.length - 1)];

  const keywordSheets = names.filter((n) => /关键词/.test(String(n)));
  if (keywordSheets.length) {
    quotes = keywordSheets[0];
  }
  const s02Like = names.find((n) => /s02|统计|年份|sheet1/i.test(String(n)));
  if (s02Like) {
    stats = s02Like;
  }

  if (stats === quotes && names.length > 1) {
    const secondQuote = [...analyzed]
      .filter((i) => i.name !== stats)
      .sort((a, b) => b.quoteScore - a.quoteScore)[0];
    quotes = secondQuote ? secondQuote.name : names[1];
  }
  return { stats, quotes };
}

function applySelectedSheets() {
  if (!currentWorkbook) {
    throw new Error("请先上传 Excel 文件。");
  }
  const statsName = statsSheetSelect.value;
  const quotesName = quotesSheetSelect.value;
  if (!statsName || !quotesName) {
    throw new Error("请先选择统计与引文工作表。");
  }
  const statsWs = currentWorkbook.Sheets[statsName];
  const quotesWs = currentWorkbook.Sheets[quotesName];
  const statsRows = window.XLSX.utils.sheet_to_json(statsWs, { defval: "" });
  const quotesRows = window.XLSX.utils.sheet_to_json(quotesWs, { defval: "" });
  applyStatsRows(statsRows, statsName);
  applyQuoteRows(quotesRows, quotesName);
}

function handleJsonImport(text) {
  const obj = JSON.parse(text);
  if (!obj || !Array.isArray(obj.rows)) {
    throw new Error("JSON 需使用 { rows: [...] } 结构。");
  }
  applyImportedRows(obj.rows);
}

function handleCsvImport(text) {
  const rows = parseCsv(text);
  applyImportedRows(rows);
}

async function handleFileImport(file) {
  const lower = file.name.toLowerCase();
  if (lower.endsWith(".xlsx")) {
    if (!window.XLSX) {
      throw new Error("未加载 XLSX 解析库，请检查网络或刷新页面后重试。");
    }
    const buffer = await file.arrayBuffer();
    const wb = window.XLSX.read(buffer, { type: "array" });
    currentWorkbook = wb;
    const sheetNames = wb.SheetNames || [];
    if (!sheetNames.length) {
      throw new Error("Excel 文件中未找到工作表。");
    }
    fillSheetSelects(sheetNames);
    sheetPickerRow.style.display = "flex";
    const keywordSheets = sheetNames.filter((n) => isKeywordSheetName(n));
    if (keywordSheets.length === sheetNames.length && keywordSheets.length > 0) {
      applyQuoteRowsFromWorkbook(wb, keywordSheets);
      quotesSheetSelect.value = keywordSheets[0];
      setStatus(`已识别为 S01 引文工作簿，导入 ${keywordSheets.length} 个颜色分表引文。`);
      return;
    }
    if (sheetNames.length >= 2) {
      const guessed = guessSheetRoles(wb);
      if (guessed.stats) {
        statsSheetSelect.value = guessed.stats;
      }
      if (guessed.quotes) {
        quotesSheetSelect.value = guessed.quotes;
      }
      applySelectedSheets();
      setStatus(`已自动匹配工作表：统计(${statsSheetSelect.value}) + 引文(${quotesSheetSelect.value})。`);
    } else {
      const ws = wb.Sheets[sheetNames[0]];
      const rows = window.XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (/关键词/.test(sheetNames[0])) {
        applyQuoteRows(rows, sheetNames[0]);
      } else if (isLikelyS02StatsRows(rows)) {
        applyStatsRows(rows, sheetNames[0]);
      } else {
        applyImportedRows(rows);
      }
      setStatus(`已导入单工作表：${sheetNames[0]}。`);
    }
    return;
  }

  const text = await file.text();
  if (lower.endsWith(".json")) {
    handleJsonImport(text);
  } else if (lower.endsWith(".csv")) {
    handleCsvImport(text);
  } else {
    throw new Error("仅支持 .xlsx、.json 或 .csv 文件。");
  }
}

function downloadCsvTemplate() {
  const header = "bucket,quote,work,year,source,count";
  const sample = [
    "red,我仿佛看见冬花开在雪野中,《野草·雪》,1925,鲁迅全集（补页码）,534",
    "black,铁屋子一般，沉沉地压着人,《呐喊·自序》,1922,鲁迅全集（补页码）,327",
    "yellow,脸上瘦削不堪，黄中带灰,《祝福》,1924,鲁迅全集（补页码）,169",
  ].join("\n");
  const blob = new Blob([`${header}\n${sample}`], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "luxun_colors_template.csv";
  a.click();
  URL.revokeObjectURL(a.href);
}

function cardTemplate(item, systemName) {
  const bucket = bucketMap[item.bucket];
  const q = firstQuoteForBucket(item.bucket);
  const total = quoteIndexByBucket.get(item.bucket)?.length ?? 0;
  const quoteTip = q
    ? total > 1
      ? `点击查看原文（分页，共 ${total} 条）`
      : "点击查看原文与出处"
    : "点击查看色类统计";
  return `
    <div class="swatch" style="background:${item.hex};color:${textColorByBg(item.hex)}"></div>
    <div class="card-body">
      <div class="name-row">
        <div>
          <div class="name-zh">${item.colorZh}</div>
          <div class="name-en">${item.colorEn}</div>
        </div>
        <div class="muted">${item.hex}</div>
      </div>
      <div class="meta">
        统计色类：${bucket.zh} (${bucket.en}) · 频次：${bucket.count || "-"}
        <span class="source">${systemName} · ${quoteTip}</span>
      </div>
    </div>
  `;
}

function createCard(item, systemName) {
  const button = document.createElement("button");
  button.className = "card";
  button.innerHTML = cardTemplate(item, systemName);
  button.addEventListener("click", () => openModal(item, systemName));
  return button;
}

function renderModalQuotePage() {
  const { bucketKey, page, colorZh, colorEn, hex, systemName } = modalQuoteState;
  const bucket = bucketMap[bucketKey];
  if (!bucket) {
    return;
  }
  const list = quoteIndexByBucket.get(bucketKey) || [];
  const total = list.length;
  const totalPages = Math.max(1, Math.ceil(total / QUOTE_PAGE_SIZE) || 1);
  let safePage = page;
  if (safePage < 0) safePage = 0;
  if (safePage > totalPages - 1) safePage = totalPages - 1;
  modalQuoteState.page = safePage;
  const start = safePage * QUOTE_PAGE_SIZE;
  const slice = list.slice(start, start + QUOTE_PAGE_SIZE);

  const quoteBlocks =
    slice.length > 0
      ? slice
          .map(
            (q) => `
    <div class="quote-block">
      <p class="quote">${escapeHtml(q.quote)}</p>
      <p class="quote-meta"><strong>作品：</strong>${escapeHtml(q.work)} · <strong>年份：</strong>${escapeHtml(q.year)}</p>
      <p class="quote-meta muted"><strong>出处：</strong>${escapeHtml(q.source)}</p>
    </div>`
          )
          .join("")
      : `<p class="muted">该色类暂无引文，请导入数据补充。</p>`;

  const pager =
    total > QUOTE_PAGE_SIZE
      ? `<div class="modal-pagination">
      <button type="button" class="pager-btn" data-quote-nav="-1" ${safePage <= 0 ? "disabled" : ""}>上一页</button>
      <span class="pager-info">第 ${safePage + 1} / ${totalPages} 页（共 ${total} 条）</span>
      <button type="button" class="pager-btn" data-quote-nav="1" ${
        safePage >= totalPages - 1 ? "disabled" : ""
      }>下一页</button>
    </div>`
      : total > 0
        ? `<p class="pager-info single">本类共 ${total} 条引文</p>`
        : "";

  modalContent.innerHTML = `
    <h3>${escapeHtml(colorZh)} <small class="muted">(${escapeHtml(colorEn)})</small></h3>
    <p><strong>色卡体系：</strong>${escapeHtml(systemName)}</p>
    <p><strong>色值：</strong>${escapeHtml(hex)}</p>
    <p><strong>匹配统计色类：</strong>${escapeHtml(bucket.zh)} (${escapeHtml(bucket.en)})</p>
    <p><strong>该色类频次：</strong>${escapeHtml(String(bucket.count || "待填充"))}</p>
    <div id="modalQuoteList" class="modal-quote-list">${quoteBlocks}</div>
    ${pager}
  `;
}

function openModal(item, systemName) {
  modalQuoteState = {
    bucketKey: item.bucket,
    page: 0,
    colorZh: item.colorZh,
    colorEn: item.colorEn,
    hex: item.hex,
    systemName,
  };
  renderModalQuotePage();
  detailModal.showModal();
}

function passesSearch(item, q) {
  if (!q) {
    return true;
  }
  const bucket = bucketMap[item.bucket];
  const quotes = quoteIndexByBucket.get(item.bucket) || [];
  const quoteHay = quotes
    .map((qq) => [qq.quote, qq.work, qq.year, qq.source].join(" "))
    .join(" ");
  const joined = [
    item.colorZh,
    item.colorEn,
    item.hex,
    bucket.zh,
    bucket.en,
    quoteHay,
  ]
    .join(" ")
    .toLowerCase();
  return joined.includes(q);
}

function extractYear(value) {
  const m = String(value || "").match(/(19\d{2}|20\d{2})/);
  return m ? m[1] : "";
}

function renderDashboard() {
  const totalQuotes = quoteRecords.length;
  const topBucket = [...colorStats].sort((a, b) => (b.count || 0) - (a.count || 0))[0];
  const activeBuckets = colorStats.filter((c) => (c.count || 0) > 0).length;
  const years = quoteRecords.map((q) => extractYear(q.year || q.work)).filter(Boolean);
  const uniqueYears = new Set(years).size;

  statCards.innerHTML = `
    <article class="stat-card"><div class="stat-label">总引文数</div><div class="stat-value">${totalQuotes}</div></article>
    <article class="stat-card"><div class="stat-label">活跃色类</div><div class="stat-value">${activeBuckets}</div></article>
    <article class="stat-card"><div class="stat-label">主导色类</div><div class="stat-value">${topBucket ? topBucket.zh : "-"}</div></article>
    <article class="stat-card"><div class="stat-label">覆盖年份</div><div class="stat-value">${uniqueYears || "-"}</div></article>
  `;

  const maxCount = Math.max(...colorStats.map((c) => c.count || 0), 1);
  colorBars.innerHTML = colorStats
    .filter((c) => (c.count || 0) > 0)
    .sort((a, b) => (b.count || 0) - (a.count || 0))
    .map(
      (c) => `<div class="bar-row">
        <div class="bar-label">${c.zh} ${c.en}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${Math.round(((c.count || 0) / maxCount) * 100)}%;background:${bucketColorMap[c.key] || "#888"}"></div></div>
        <div class="bar-val">${c.count || 0}</div>
      </div>`
    )
    .join("");

  const yearMap = new Map();
  years.forEach((y) => yearMap.set(y, (yearMap.get(y) || 0) + 1));
  const yearRows = [...yearMap.entries()].sort((a, b) => a[0].localeCompare(b[0])).slice(-10);
  const yearMax = Math.max(...yearRows.map(([, v]) => v), 1);
  yearBars.innerHTML = yearRows.length
    ? yearRows
        .map(
          ([y, v]) => `<div class="bar-row">
          <div class="bar-label">${y}</div>
          <div class="bar-track"><div class="bar-fill" style="width:${Math.round((v / yearMax) * 100)}%;background:#c07a46"></div></div>
          <div class="bar-val">${v}</div>
        </div>`
        )
        .join("")
    : `<p class="muted">暂无年份数据</p>`;

  const topYear = yearRows.length ? [...yearRows].sort((a, b) => b[1] - a[1])[0][0] : "未知";
  headlineInsight.textContent = `结论提示：${topBucket ? topBucket.zh : "灰"}色意象最显著，${topYear}年前后引文密度更高，文本整体呈现“高对比色 + 社会情绪”传播特征。`;

  const topColors = [...colorStats]
    .filter((c) => (c.count || 0) > 0)
    .sort((a, b) => (b.count || 0) - (a.count || 0))
    .slice(0, 3);
  storyTimeline.innerHTML = `
    <article class="story-item">
      <div class="story-title">01 数据起点</div>
      <div class="story-desc">当前共收录 ${totalQuotes} 条引文，覆盖 ${activeBuckets} 个活跃色类，形成可传播的视觉样本池。</div>
    </article>
    <article class="story-item">
      <div class="story-title">02 色彩主轴</div>
      <div class="story-desc">主导色类为 ${topBucket ? topBucket.zh : "-"}，前三色为 ${topColors.map((c) => c.zh).join(" / ") || "-"}，体现鲁迅文本中的强烈对比结构。</div>
    </article>
    <article class="story-item">
      <div class="story-title">03 时间窗口</div>
      <div class="story-desc">从年份分布看，${topYear} 年附近引文出现更密集，可作为传播叙事的重点时段。</div>
    </article>
    <article class="story-item">
      <div class="story-title">04 传播建议</div>
      <div class="story-desc">建议以“颜色卡片 + 原文弹窗 + 时间分布图”三联结构发布，兼顾可视化吸引力与学术可信度。</div>
    </article>
  `;
}

async function exportPosterAsPng() {
  if (!window.html2canvas) {
    setStatus("导出失败：截图库未加载，请刷新后重试。", true);
    return;
  }
  const node = document.querySelector(".page");
  const canvas = await window.html2canvas(node, {
    backgroundColor: "#0f1116",
    scale: Math.min(2, window.devicePixelRatio || 1.5),
  });
  const a = document.createElement("a");
  a.href = canvas.toDataURL("image/png");
  a.download = `鲁迅颜色意象互动展_${new Date().toISOString().slice(0, 10)}.png`;
  a.click();
}

function render() {
  cnGrid.innerHTML = "";
  westernGrid.innerHTML = "";
  const q = searchInput.value.trim().toLowerCase();
  const system = systemFilter.value;

  if (system === "all" || system === "cn") {
    cnPalette.filter((i) => passesSearch(i, q)).forEach((i) => {
      cnGrid.appendChild(createCard(i, "中国传统色"));
    });
  }
  if (system === "all" || system === "western") {
    westernPalette.filter((i) => passesSearch(i, q)).forEach((i) => {
      westernGrid.appendChild(createCard(i, "西方经典色"));
    });
  }
  renderDashboard();
}

if (applyPreloadedData()) {
  setStatus("已加载预置数据（S01/S02），可直接浏览；仍可继续上传新文件覆盖。");
}
rebuildQuoteIndex();

dataFileInput.addEventListener("change", async (e) => {
  const file = e.target.files && e.target.files[0];
  if (!file) {
    return;
  }
  try {
    await handleFileImport(file);
    render();
    setStatus(`导入成功：${file.name}，已更新统计和引文。`);
  } catch (err) {
    const msg = err instanceof Error ? err.message : "未知错误";
    setStatus(`导入失败：${msg}`, true);
  }
});

loadSampleBtn.addEventListener("click", () => {
  colorStats = structuredClone(defaultColorStats);
  quoteRecords = structuredClone(defaultQuoteRecords);
  bucketMap = Object.fromEntries(colorStats.map((i) => [i.key, i]));
  rebuildQuoteIndex();
  currentWorkbook = null;
  sheetPickerRow.style.display = "none";
  render();
  setStatus("已恢复内置示例数据。");
});

downloadTemplateBtn.addEventListener("click", downloadCsvTemplate);
exportPosterBtn.addEventListener("click", async () => {
  try {
    await exportPosterAsPng();
    setStatus("已导出展示图 PNG。");
  } catch (err) {
    setStatus(`导出失败：${err instanceof Error ? err.message : "未知错误"}`, true);
  }
});
applySheetsBtn.addEventListener("click", () => {
  try {
    applySelectedSheets();
    render();
    setStatus(
      `已应用工作表：统计(${statsSheetSelect.value}) + 引文(${quotesSheetSelect.value})。`
    );
  } catch (err) {
    const msg = err instanceof Error ? err.message : "未知错误";
    setStatus(`应用失败：${msg}`, true);
  }
});
searchInput.addEventListener("input", render);
systemFilter.addEventListener("change", render);
closeModal.addEventListener("click", () => detailModal.close());
detailModal.addEventListener("click", (e) => {
  const nav = e.target.closest("[data-quote-nav]");
  if (nav && !nav.disabled) {
    e.stopPropagation();
    modalQuoteState.page += Number(nav.dataset.quoteNav);
    renderModalQuotePage();
    return;
  }
  if (e.target === detailModal) {
    detailModal.close();
  }
});

render();
