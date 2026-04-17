const STAGES = [
  {
    key: "childhood",
    name: "童年与少年",
    period: "1881.9.25-1898.5",
    startYear: 1881,
    endYear: 1898,
    place: "浙江绍兴东昌坊口新台门",
    route: "绍兴",
    works: [
      "接触《山海经》及历代野史杂记，奠定民间文学兴趣。"
    ]
  },
  {
    key: "new-education",
    name: "新式教育启蒙",
    period: "1898.5-1902.3",
    startYear: 1898,
    endYear: 1902,
    place: "南京江南水师学堂→矿务铁路学堂",
    route: "南京",
    works: [
      "1901年《戛剑生杂记》（现存最早个人短文）。"
    ]
  },
  {
    key: "turning",
    name: "弃医从文的转折",
    period: "1902.4-1909.8",
    startYear: 1902,
    endYear: 1909,
    place: "日本东京→仙台→东京",
    route: "日本",
    works: [
      "1903年译述《斯巴达之魂》《月界旅行》。",
      "1907年《人之历史》《摩罗诗力说》《文化偏至论》。",
      "1909年与周作人合译《域外小说集》第一、二集。"
    ]
  },
  {
    key: "return",
    name: "回国初期任教",
    period: "1909.8-1912.5",
    startYear: 1909,
    endYear: 1912,
    place: "杭州→绍兴→南京临时政府教育部",
    route: "杭州/绍兴/南京",
    works: [
      "1911年文言小说《怀旧》（第一篇小说）。"
    ]
  },
  {
    key: "beijing",
    name: "新文化运动旗手",
    period: "1912.5-1926.8",
    startYear: 1912,
    endYear: 1926,
    place: "北京（绍兴会馆→八道湾→西三条）",
    route: "北京",
    works: [
      "1918年《狂人日记》",
      "1919年《孔乙己》《药》",
      "1921年《阿Q正传》连载",
      "1923年《呐喊》出版",
      "1924-1926《野草》《朝花夕拾》",
      "1926年《彷徨》出版"
    ]
  },
  {
    key: "xiamen",
    name: "厦门短暂任教",
    period: "1926.9-1927.1",
    startYear: 1926,
    endYear: 1927,
    place: "福建厦门大学",
    route: "厦门",
    works: [
      "完成《朝花夕拾》后五篇。",
      "开始《故事新编》之《奔月》。"
    ]
  },
  {
    key: "guangzhou",
    name: "思想重大转变",
    period: "1927.1-1927.9",
    startYear: 1927,
    endYear: 1927,
    place: "广东广州中山大学",
    route: "广州",
    works: [
      "完成《铸剑》（原名《眉间尺》）。",
      "演讲《魏晋风度及文章与药及酒之关系》。"
    ]
  },
  {
    key: "shanghai",
    name: "最后十年战斗",
    period: "1927.10-1936.10.19",
    startYear: 1927,
    endYear: 1936,
    place: "上海（景云里→拉摩斯公寓→大陆新村）",
    route: "上海",
    works: [
      "1936年《故事新编》出版。",
      "杂文集：三闲集、二心集、南腔北调集等。",
      "翻译果戈理《死魂灵》并编校多种文献。"
    ]
  }
];

const ROUTE_TEXT = "绍兴（1881-1898）→南京（1898-1902）→日本（1902-1909）→杭州（1909-1910）→绍兴（1910-1912）→南京（1912.2-1912.5）→北京（1912-1926）→厦门（1926-1927）→广州（1927.1-1927.9）→上海（1927-1936）";

const COLOR_ORDER = ["red", "black", "white", "yellow", "green", "blue", "gray", "cyan", "pink", "purple", "orange"];
const COLOR_LABEL = {
  red: "红",
  black: "黑",
  white: "白",
  yellow: "黄",
  green: "绿",
  blue: "蓝",
  gray: "灰",
  cyan: "青",
  pink: "粉",
  purple: "紫",
  orange: "橙"
};

const QUOTE_ROWS = (window.PRELOADED_DATA && Array.isArray(window.PRELOADED_DATA.quotes))
  ? window.PRELOADED_DATA.quotes
  : [];

const STATS_ROWS = (window.PRELOADED_DATA && Array.isArray(window.PRELOADED_DATA.stats))
  ? window.PRELOADED_DATA.stats
  : [];

const state = {
  stageIdx: 0
};

const stageSelect = document.getElementById("stageSelect");
const playBtn = document.getElementById("playBtn");
const exportPosterBtn = document.getElementById("exportPosterBtn");
const playStatus = document.getElementById("playStatus");
const pathTrack = document.getElementById("pathTrack");
const stageMeta = document.getElementById("stageMeta");
const stageWorks = document.getElementById("stageWorks");
const yearChart = document.getElementById("yearChart");
const yearChartCaption = document.getElementById("yearChartCaption");
const colorChart = document.getElementById("colorChart");
const genreColorTable = document.getElementById("genreColorTable");

let autoplayTimer = null;

function parseYear(value) {
  const s = String(value || "");
  const m = s.match(/(18|19|20)\d{2}/);
  return m ? Number(m[0]) : null;
}

function normalizeBucket(raw) {
  const s = String(raw || "").toLowerCase().trim();
  if (!s) return "unknown";
  if (s.includes("red") || s.includes("红")) return "red";
  if (s.includes("black") || s.includes("黑")) return "black";
  if (s.includes("white") || s.includes("白")) return "white";
  if (s.includes("yellow") || s.includes("黄")) return "yellow";
  if (s.includes("green") || s.includes("绿")) return "green";
  if (s.includes("blue") || s.includes("蓝")) return "blue";
  if (s.includes("gray") || s.includes("grey") || s.includes("灰")) return "gray";
  if (s.includes("cyan") || s.includes("青")) return "cyan";
  if (s.includes("pink") || s.includes("粉")) return "pink";
  if (s.includes("purple") || s.includes("紫")) return "purple";
  if (s.includes("orange") || s.includes("橙")) return "orange";
  return "unknown";
}

function inferGenre(workRaw) {
  const w = String(workRaw || "");
  if (!w) return "未标注";
  if (w.includes("野草")) return "散文诗";
  if (w.includes("朝花夕拾")) return "散文";
  if (w.includes("故事新编") || w.includes("铸剑") || w.includes("奔月")) return "历史小说";
  if (
    w.includes("狂人日记") || w.includes("阿Q") || w.includes("孔乙己") || w.includes("药") ||
    w.includes("祝福") || w.includes("故乡") || w.includes("社戏") || w.includes("伤逝") ||
    w.includes("在酒楼上") || w.includes("孤独者") || w.includes("风波") || w.includes("离婚") ||
    w.includes("端午节") || w.includes("示众") || w.includes("明天") || w.includes("肥皂")
  ) {
    return "小说";
  }
  if (
    w.includes("热风") || w.includes("坟") || w.includes("华盖集") || w.includes("而已集") ||
    w.includes("三闲集") || w.includes("二心集") || w.includes("南腔北调集") || w.includes("且介亭")
  ) {
    return "杂文";
  }
  if (w.includes("诗")) return "诗歌";
  return "其他";
}

function stageQuotes(stage) {
  return QUOTE_ROWS.filter((row) => {
    const year = parseYear(row.year);
    return year && year >= stage.startYear && year <= stage.endYear;
  });
}

function countByYear(rows) {
  const map = new Map();
  rows.forEach((r) => {
    const y = parseYear(r.year);
    if (!y) return;
    map.set(y, (map.get(y) || 0) + 1);
  });
  return [...map.entries()].sort((a, b) => a[0] - b[0]);
}

function countByBucket(rows) {
  const map = new Map();
  rows.forEach((r) => {
    const b = normalizeBucket(r.bucket);
    map.set(b, (map.get(b) || 0) + 1);
  });
  return [...map.entries()].sort((a, b) => b[1] - a[1]);
}

function renderBars(target, data, labelFn = (x) => x) {
  target.innerHTML = "";
  if (!data.length) {
    target.innerHTML = "<p class='caption'>当前阶段无可用统计数据。</p>";
    return;
  }
  const max = Math.max(...data.map((d) => d[1]), 1);
  data.forEach(([name, val]) => {
    const row = document.createElement("div");
    row.className = "bar-row";
    row.innerHTML = `
      <span class="bar-name">${labelFn(name)}</span>
      <span class="bar-track"><span class="bar-fill" style="width:${(val / max) * 100}%"></span></span>
      <span class="bar-val">${val}</span>
    `;
    target.appendChild(row);
  });
}

function renderPath() {
  pathTrack.innerHTML = "";
  STAGES.forEach((stage, idx) => {
    const card = document.createElement("article");
    card.className = "path-node" + (idx === state.stageIdx ? " active" : "");
    card.innerHTML = `
      <div class="node-time">${stage.period}</div>
      <div class="node-place">${stage.route}</div>
      <div class="node-stage">${stage.name}</div>
    `;
    card.addEventListener("click", () => {
      state.stageIdx = idx;
      stageSelect.value = String(idx);
      stopAutoplay();
      render();
    });
    pathTrack.appendChild(card);
  });
}

function renderStageInfo(stage) {
  stageMeta.innerHTML = `
    <div><strong>阶段：</strong>${stage.name}</div>
    <div><strong>起止：</strong>${stage.period}</div>
    <div><strong>主要地点：</strong>${stage.place}</div>
    <div><strong>行迹总览：</strong>${ROUTE_TEXT}</div>
  `;
  stageWorks.innerHTML = "";
  stage.works.forEach((w) => {
    const line = document.createElement("div");
    line.className = "work-item";
    line.textContent = w;
    stageWorks.appendChild(line);
  });
}

function buildGenreColorTable(rows) {
  const genres = ["小说", "历史小说", "散文", "散文诗", "杂文", "诗歌", "其他", "未标注"];
  const grid = new Map();
  genres.forEach((g) => grid.set(g, new Map()));
  rows.forEach((r) => {
    const genre = inferGenre(r.work);
    const bucket = normalizeBucket(r.bucket);
    if (!COLOR_ORDER.includes(bucket)) return;
    if (!grid.has(genre)) grid.set(genre, new Map());
    const m = grid.get(genre);
    m.set(bucket, (m.get(bucket) || 0) + 1);
  });

  let maxVal = 1;
  grid.forEach((m) => {
    m.forEach((v) => {
      if (v > maxVal) maxVal = v;
    });
  });

  const thead = `<thead><tr><th>体裁</th>${COLOR_ORDER.map((b) => `<th>${COLOR_LABEL[b]}</th>`).join("")}</tr></thead>`;
  const bodyRows = genres.map((genre) => {
    const m = grid.get(genre) || new Map();
    const tds = COLOR_ORDER.map((b) => {
      const v = m.get(b) || 0;
      const alpha = v === 0 ? 0.04 : 0.12 + (v / maxVal) * 0.6;
      return `<td style="background:rgba(114,146,255,${alpha.toFixed(3)})">${v || "-"}</td>`;
    }).join("");
    return `<tr><td class="genre-cell">${genre}</td>${tds}</tr>`;
  }).join("");
  genreColorTable.innerHTML = `${thead}<tbody>${bodyRows}</tbody>`;
}

function renderStageSelect() {
  stageSelect.innerHTML = STAGES.map((s, idx) => (
    `<option value="${idx}">${s.name}（${s.period}）</option>`
  )).join("");
  stageSelect.value = String(state.stageIdx);
  stageSelect.addEventListener("change", (e) => {
    state.stageIdx = Number(e.target.value || 0);
    stopAutoplay();
    render();
  });
}

function updatePlayUi() {
  const running = Boolean(autoplayTimer);
  playBtn.textContent = running ? "暂停播放" : "播放行迹";
  playStatus.textContent = running
    ? "正在自动播放：每 1.8 秒切换到下一阶段。"
    : "可手动切换阶段，或点击“播放行迹”自动演示。";
}

function stopAutoplay() {
  if (!autoplayTimer) return;
  clearInterval(autoplayTimer);
  autoplayTimer = null;
  updatePlayUi();
}

function startAutoplay() {
  if (autoplayTimer) return;
  autoplayTimer = setInterval(() => {
    state.stageIdx = (state.stageIdx + 1) % STAGES.length;
    stageSelect.value = String(state.stageIdx);
    render();
  }, 1800);
  updatePlayUi();
}

function toggleAutoplay() {
  if (autoplayTimer) {
    stopAutoplay();
  } else {
    startAutoplay();
  }
}

async function exportPosterAsPng() {
  if (!window.html2canvas) {
    alert("导出失败：未加载海报导出组件。");
    return;
  }
  const target = document.querySelector(".page");
  if (!target) return;
  const prevText = exportPosterBtn.textContent;
  exportPosterBtn.disabled = true;
  exportPosterBtn.textContent = "导出中...";
  try {
    const canvas = await window.html2canvas(target, {
      backgroundColor: "#10131a",
      scale: Math.min(window.devicePixelRatio || 1.5, 2)
    });
    const url = canvas.toDataURL("image/png");
    const a = document.createElement("a");
    a.href = url;
    a.download = "鲁迅行迹_创作与颜色意象_海报版.png";
    a.click();
  } catch (err) {
    alert("导出失败，请重试。");
  } finally {
    exportPosterBtn.disabled = false;
    exportPosterBtn.textContent = prevText;
  }
}

function render() {
  const stage = STAGES[state.stageIdx];
  const rows = stageQuotes(stage);
  const yearData = countByYear(rows);
  const colorData = countByBucket(rows).filter(([k]) => COLOR_ORDER.includes(k));

  renderPath();
  renderStageInfo(stage);
  renderBars(yearChart, yearData, (name) => `${name}`);
  renderBars(colorChart, colorData, (k) => COLOR_LABEL[k] || k);
  buildGenreColorTable(rows);

  yearChartCaption.textContent = `当前阶段 ${stage.period} 共匹配到 ${rows.length} 条引文记录（来自已加载数据）。`;
}

function bootstrap() {
  // If quote data is unavailable, fallback to stats rows by color.
  if (!QUOTE_ROWS.length && STATS_ROWS.length) {
    yearChartCaption.textContent = "未检测到引文数据，仅检测到颜色总量统计。";
  }
  renderStageSelect();
  playBtn.addEventListener("click", toggleAutoplay);
  exportPosterBtn.addEventListener("click", exportPosterAsPng);
  updatePlayUi();
  render();
}

bootstrap();
