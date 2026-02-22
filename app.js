const TEMPLATE_FILE = "260121_Spielberichtsbogen_DE-ENG_ausgefüllt (3).xlsx";
const MATCH_ROWS = [12, 13, 14, 15, 16, 17, 18, 19];
const MATCH_NAMES = Array.from({ length: 8 }, (_, i) => `Match ${i + 1}`);

const sheetSelect = document.getElementById("sheetName");
const matchesWrap = document.getElementById("matches");
const summary = document.getElementById("summary");
const statusEl = document.getElementById("status");
const inputs = [];

for (let i = 1; i <= 5; i += 1) {
  const name = `${i}.Begegnung`;
  const opt = document.createElement("option");
  opt.value = name;
  opt.textContent = name;
  sheetSelect.appendChild(opt);
}

const tpl = document.getElementById("matchTemplate");
MATCH_NAMES.forEach((name) => {
  const node = tpl.content.firstElementChild.cloneNode(true);
  node.querySelector(".name").textContent = name;
  matchesWrap.appendChild(node);
  inputs.push({
    home1: node.querySelector(".home1"),
    guest1: node.querySelector(".guest1"),
    home2: node.querySelector(".home2"),
    guest2: node.querySelector(".guest2"),
    home3: node.querySelector(".home3"),
    guest3: node.querySelector(".guest3"),
  });
});

function parseNum(v) {
  if (v === "" || v == null) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function calcMatch(entry) {
  const sets = [
    [parseNum(entry.home1.value), parseNum(entry.guest1.value)],
    [parseNum(entry.home2.value), parseNum(entry.guest2.value)],
    [parseNum(entry.home3.value), parseNum(entry.guest3.value)],
  ];
  let homeSets = 0;
  let guestSets = 0;
  let homePoints = 0;
  let guestPoints = 0;
  sets.forEach(([h, g]) => {
    if (h == null || g == null) return;
    homePoints += h;
    guestPoints += g;
    if (h > g) homeSets += 1;
    if (g > h) guestSets += 1;
  });
  return {
    sets,
    homePoints,
    guestPoints,
    homeSets,
    guestSets,
    homeMatch: homeSets > guestSets ? 1 : 0,
    guestMatch: guestSets > homeSets ? 1 : 0,
  };
}

function renderSummary() {
  const totals = {
    homePoints: 0,
    guestPoints: 0,
    homeSets: 0,
    guestSets: 0,
    homeMatches: 0,
    guestMatches: 0,
  };
  inputs.forEach((entry) => {
    const r = calcMatch(entry);
    totals.homePoints += r.homePoints;
    totals.guestPoints += r.guestPoints;
    totals.homeSets += r.homeSets;
    totals.guestSets += r.guestSets;
    totals.homeMatches += r.homeMatch;
    totals.guestMatches += r.guestMatch;
  });

  summary.textContent = `Points H:G ${totals.homePoints}:${totals.guestPoints} · Sets H:G ${totals.homeSets}:${totals.guestSets} · Matches H:G ${totals.homeMatches}:${totals.guestMatches}`;
  return totals;
}

inputs.forEach((entry) => Object.values(entry).forEach((i) => i.addEventListener("input", renderSummary)));
renderSummary();

async function loadWorkbook() {
  const resp = await fetch(TEMPLATE_FILE);
  if (!resp.ok) throw new Error(`Could not load template file: ${TEMPLATE_FILE}`);
  const ab = await resp.arrayBuffer();
  return XLSX.read(ab, { type: "array" });
}

function setCell(ws, ref, value) {
  if (value instanceof Date) ws[ref] = { t: "d", v: value };
  else ws[ref] = { t: typeof value === "number" ? "n" : "s", v: value };
  if (!ws["!ref"]) ws["!ref"] = "A1:T70";
}

document.getElementById("downloadBtn").addEventListener("click", async () => {
  try {
    const wb = await loadWorkbook();
    const ws = wb.Sheets[sheetSelect.value];
    if (!ws) throw new Error("Selected sheet not found in template.");

    const group = document.getElementById("group").value.trim();
    const homeTeam = document.getElementById("homeTeam").value.trim();
    const guestTeam = document.getElementById("guestTeam").value.trim();
    const venue = document.getElementById("venue").value.trim();
    const date = document.getElementById("date").value;

    if (group) setCell(ws, "M5", group);
    if (homeTeam) setCell(ws, "D5", homeTeam);
    if (guestTeam) setCell(ws, "D7", guestTeam);
    if (venue) setCell(ws, "M7", venue);
    if (date) setCell(ws, "E27", new Date(date));

    let homePoints = 0;
    let guestPoints = 0;
    let homeSets = 0;
    let guestSets = 0;
    let homeMatches = 0;
    let guestMatches = 0;

    MATCH_ROWS.forEach((row, idx) => {
      const r = calcMatch(inputs[idx]);
      const cols = ["I", "J", "K", "L", "M", "N"];
      const flat = [r.sets[0][0], r.sets[0][1], r.sets[1][0], r.sets[1][1], r.sets[2][0], r.sets[2][1]];
      flat.forEach((val, cIdx) => {
        if (val != null) setCell(ws, `${cols[cIdx]}${row}`, val);
      });

      setCell(ws, `O${row}`, r.homePoints);
      setCell(ws, `P${row}`, r.guestPoints);
      setCell(ws, `Q${row}`, r.homeSets);
      setCell(ws, `R${row}`, r.guestSets);
      setCell(ws, `S${row}`, r.homeMatch);
      setCell(ws, `T${row}`, r.guestMatch);

      homePoints += r.homePoints;
      guestPoints += r.guestPoints;
      homeSets += r.homeSets;
      guestSets += r.guestSets;
      homeMatches += r.homeMatch;
      guestMatches += r.guestMatch;
    });

    setCell(ws, "O21", homePoints);
    setCell(ws, "P21", guestPoints);
    setCell(ws, "Q21", homeSets);
    setCell(ws, "R21", guestSets);
    setCell(ws, "S21", homeMatches);
    setCell(ws, "T21", guestMatches);

    const out = XLSX.write(wb, { type: "array", bookType: "xlsx", cellDates: true });
    const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `updated_${sheetSelect.value.replace('.', '_')}.xlsx`;
    a.click();
    URL.revokeObjectURL(a.href);
    statusEl.textContent = "Workbook generated successfully.";
  } catch (err) {
    statusEl.textContent = err.message;
  }
});
