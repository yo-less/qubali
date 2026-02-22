const TEMPLATE_FILE = "260121_Spielberichtsbogen_DE-ENG_ausgefüllt (3).xlsx";

const TEAM_COUNT = 6;
const MAX_PLAYERS = 15;
const MIN_PLAYERS = 4;
const ENCOUNTER_COUNT = 5;
const MATCHES = [
  { key: "d1", label: "1. Doppel", row: 12, doubles: true },
  { key: "d2", label: "2. Doppel", row: 14, doubles: true },
  { key: "d3", label: "3. Doppel", row: 16, doubles: true },
  { key: "s1", label: "1. Einzel", row: 18, doubles: false },
  { key: "s2", label: "2. Einzel", row: 19, doubles: false },
];

const state = {
  allgemeines: {
    group: "",
    coordinator: "",
    venue: "",
    place: "",
    date: "",
    teams: Array.from({ length: TEAM_COUNT }, (_, i) => ({ name: `Team ${i + 1}`, captain: "" })),
  },
  setzlisten: Array.from({ length: TEAM_COUNT }, () => Array.from({ length: MAX_PLAYERS }, () => "")),
  begegnungen: Array.from({ length: ENCOUNTER_COUNT }, () => ({
    homeTeam: "",
    guestTeam: "",
    startTime: "",
    endTime: "",
    matches: MATCHES.map((m) => ({
      key: m.key,
      home1: "",
      home2: "",
      guest1: "",
      guest2: "",
      sets: [
        { h: "", g: "" },
        { h: "", g: "" },
        { h: "", g: "" },
      ],
    })),
  })),
};

function setCell(ws, ref, value) {
  if (value == null || value === "") return;
  if (value instanceof Date) ws[ref] = { t: "d", v: value };
  else ws[ref] = { t: typeof value === "number" ? "n" : "s", v: value };
}
function parseNum(v) {
  if (v === "" || v == null) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function buildTabs() {
  const tabs = document.getElementById("tabs");
  [
    ["allgemeines", "Allgemeines"],
    ["setzlisten", "Setzlisten"],
    ["begegnungen", "1.–5. Begegnung"],
  ].forEach(([id, label], idx) => {
    const b = document.createElement("button");
    b.className = `tab-btn ${idx === 0 ? "active" : ""}`;
    b.textContent = label;
    b.addEventListener("click", () => showTab(id, b));
    tabs.appendChild(b);
  });
}

function showTab(id, btn) {
  document.querySelectorAll(".tab-panel").forEach((p) => p.classList.add("hidden"));
  document.getElementById(`tab-${id}`).classList.remove("hidden");
  document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
  btn.classList.add("active");
}

function renderAllgemeines() {
  const el = document.getElementById("tab-allgemeines");
  el.innerHTML = `
    <h2>Allgemeines</h2>
    <div class="grid meta-grid">
      <label>Gruppe<input id="group" /></label>
      <label>Gruppenkoordinator:in<input id="coordinator" /></label>
      <label>Austragungsort<input id="venue" /></label>
      <label>Ort<input id="place" /></label>
      <label>Datum<input id="date" type="date" /></label>
    </div>
    <h3>Teams</h3>
    <div id="teamMeta"></div>
  `;

  ["group", "coordinator", "venue", "place", "date"].forEach((k) => {
    const i = document.getElementById(k);
    i.value = state.allgemeines[k];
    i.addEventListener("input", () => {
      state.allgemeines[k] = i.value;
      renderBegegnungen();
    });
  });

  const teamMeta = document.getElementById("teamMeta");
  state.allgemeines.teams.forEach((t, idx) => {
    const row = document.createElement("div");
    row.className = "grid meta-grid";
    row.innerHTML = `
      <label>Team ${idx + 1} Name<input class="team-name" data-i="${idx}" value="${t.name}" /></label>
      <label>Team ${idx + 1} Kapitän:in<input class="captain" data-i="${idx}" value="${t.captain}" /></label>
    `;
    teamMeta.appendChild(row);
  });
  teamMeta.querySelectorAll(".team-name").forEach((i) => i.addEventListener("input", () => {
    const idx = Number(i.dataset.i);
    state.allgemeines.teams[idx].name = i.value;
    renderSetzlisten();
    renderBegegnungen();
  }));
  teamMeta.querySelectorAll(".captain").forEach((i) => i.addEventListener("input", () => {
    state.allgemeines.teams[Number(i.dataset.i)].captain = i.value;
    renderBegegnungen();
  }));
}

function renderSetzlisten() {
  const el = document.getElementById("tab-setzlisten");
  el.innerHTML = "<h2>Setzlisten</h2>";
  const tpl = document.getElementById("playerRowTemplate");

  state.allgemeines.teams.forEach((team, teamIdx) => {
    const block = document.createElement("div");
    block.className = "team-block";
    block.innerHTML = `<h3>${team.name || `Team ${teamIdx + 1}`}</h3><p class="muted">Mindestens 4, maximal 15 Spieler:innen. Reihenfolge nach Spielstärke.</p>`;
    for (let p = 0; p < MAX_PLAYERS; p += 1) {
      const node = tpl.content.firstElementChild.cloneNode(true);
      node.querySelector(".rank").textContent = `#${p + 1}`;
      const input = node.querySelector(".player-name");
      input.value = state.setzlisten[teamIdx][p];
      input.addEventListener("input", () => {
        state.setzlisten[teamIdx][p] = input.value.trim();
        renderBegegnungen();
      });
      block.appendChild(node);
    }
    el.appendChild(block);
  });
}

function playersForTeam(teamName) {
  const idx = state.allgemeines.teams.findIndex((t) => t.name === teamName);
  if (idx < 0) return [];
  return state.setzlisten[idx].filter(Boolean);
}

function renderBegegnungen() {
  const el = document.getElementById("tab-begegnungen");
  el.innerHTML = "<h2>1.–5. Begegnung</h2>";
  const teamNames = state.allgemeines.teams.map((t) => t.name).filter(Boolean);

  state.begegnungen.forEach((enc, encIdx) => {
    const card = document.createElement("div");
    card.className = "encounter-card";
    card.innerHTML = `
      <h3>${encIdx + 1}. Begegnung</h3>
      <div class="grid meta-grid">
        <label>Heimteam<select class="home"></select></label>
        <label>Gastteam<select class="guest"></select></label>
        <label>Startzeit (optional)<input class="start" type="time" /></label>
        <label>Endzeit (optional)<input class="end" type="time" /></label>
      </div>
      <div class="match-grid"></div>
      <div class="summary"></div>
    `;

    const fillSelect = (s, current) => {
      s.innerHTML = '<option value="">Bitte wählen</option>';
      teamNames.forEach((name) => {
        const o = document.createElement("option");
        o.value = name;
        o.textContent = name;
        if (name === current) o.selected = true;
        s.appendChild(o);
      });
    };

    const homeSel = card.querySelector(".home");
    const guestSel = card.querySelector(".guest");
    fillSelect(homeSel, enc.homeTeam);
    fillSelect(guestSel, enc.guestTeam);
    homeSel.addEventListener("change", () => { enc.homeTeam = homeSel.value; renderBegegnungen(); });
    guestSel.addEventListener("change", () => { enc.guestTeam = guestSel.value; renderBegegnungen(); });

    card.querySelector(".start").value = enc.startTime;
    card.querySelector(".end").value = enc.endTime;
    card.querySelector(".start").addEventListener("input", (e) => { enc.startTime = e.target.value; });
    card.querySelector(".end").addEventListener("input", (e) => { enc.endTime = e.target.value; });

    const homePlayers = playersForTeam(enc.homeTeam);
    const guestPlayers = playersForTeam(enc.guestTeam);

    const matchGrid = card.querySelector(".match-grid");
    MATCHES.forEach((m, idx) => {
      const match = enc.matches[idx];
      const row = document.createElement("div");
      row.className = "match-row";
      row.innerHTML = `<strong>${m.label}</strong>`;
      const mkPlayerSelect = (list, val, cb) => {
        const s = document.createElement("select");
        s.innerHTML = '<option value="">-</option>';
        list.forEach((p) => {
          const o = document.createElement("option"); o.value = p; o.textContent = p; if (p === val) o.selected = true; s.appendChild(o);
        });
        s.addEventListener("change", () => cb(s.value));
        return s;
      };
      const lineup = document.createElement("div");
      lineup.className = "set-row";
      lineup.append(
        mkPlayerSelect(homePlayers, match.home1, (v) => { match.home1 = v; }),
        m.doubles ? mkPlayerSelect(homePlayers, match.home2, (v) => { match.home2 = v; }) : document.createElement("span"),
        mkPlayerSelect(guestPlayers, match.guest1, (v) => { match.guest1 = v; }),
        m.doubles ? mkPlayerSelect(guestPlayers, match.guest2, (v) => { match.guest2 = v; }) : document.createElement("span"),
      );
      row.appendChild(lineup);

      for (let s = 0; s < 3; s += 1) {
        const sets = document.createElement("div");
        sets.className = "set-row";
        const h = document.createElement("input"); h.type = "number"; h.min = "0"; h.placeholder = `S${s + 1} Heim`; h.value = match.sets[s].h;
        const g = document.createElement("input"); g.type = "number"; g.min = "0"; g.placeholder = `S${s + 1} Gast`; g.value = match.sets[s].g;
        h.addEventListener("input", () => { match.sets[s].h = h.value; renderBegegnungen(); });
        g.addEventListener("input", () => { match.sets[s].g = g.value; renderBegegnungen(); });
        sets.append(h, g);
        row.appendChild(sets);
      }
      matchGrid.appendChild(row);
    });

    const totals = calcEncounter(enc);
    card.querySelector(".summary").textContent = `Punkte ${totals.homePoints}:${totals.guestPoints} · Sätze ${totals.homeSets}:${totals.guestSets} · Spiele ${totals.homeMatches}:${totals.guestMatches}`;

    el.appendChild(card);
  });
}

function calcEncounter(enc) {
  const totals = { homePoints: 0, guestPoints: 0, homeSets: 0, guestSets: 0, homeMatches: 0, guestMatches: 0 };
  enc.matches.forEach((m) => {
    let hs = 0; let gs = 0; let hp = 0; let gp = 0;
    m.sets.forEach((s) => {
      const h = parseNum(s.h); const g = parseNum(s.g);
      if (h == null || g == null) return;
      hp += h; gp += g;
      if (h > g) hs += 1;
      if (g > h) gs += 1;
    });
    totals.homePoints += hp; totals.guestPoints += gp;
    totals.homeSets += hs; totals.guestSets += gs;
    totals.homeMatches += hs > gs ? 1 : 0;
    totals.guestMatches += gs > hs ? 1 : 0;
  });
  return totals;
}

function validateData() {
  for (let i = 0; i < TEAM_COUNT; i += 1) {
    const count = state.setzlisten[i].filter(Boolean).length;
    if (count < MIN_PLAYERS || count > MAX_PLAYERS) {
      throw new Error(`${state.allgemeines.teams[i].name || `Team ${i + 1}`} benötigt 4 bis 15 Spieler:innen (aktuell ${count}).`);
    }
  }
}

async function loadWorkbook() {
  const resp = await fetch(TEMPLATE_FILE);
  if (!resp.ok) throw new Error(`Template nicht gefunden: ${TEMPLATE_FILE}`);
  const ab = await resp.arrayBuffer();
  return XLSX.read(ab, { type: "array", cellDates: true });
}

function fillAllgemeines(wb) {
  const ws = wb.Sheets["Allgemeines"];
  const a = state.allgemeines;
  setCell(ws, "B1", a.group);
  setCell(ws, "D1", a.coordinator);
  setCell(ws, "B2", a.venue);
  setCell(ws, "B3", a.place);
  if (a.date) setCell(ws, "B4", new Date(a.date));
  a.teams.forEach((team, i) => {
    const row = 7 + i;
    setCell(ws, `B${row}`, team.name);
    setCell(ws, `D${row}`, team.captain);
  });
}

function fillSetzlisten(wb) {
  const ws = wb.Sheets["Setzlisten"];
  state.allgemeines.teams.forEach((team, i) => {
    const head = 1 + i * 18;
    const start = 3 + i * 18;
    setCell(ws, `B${head}`, team.name);
    for (let p = 0; p < MAX_PLAYERS; p += 1) {
      const row = start + p;
      const name = state.setzlisten[i][p] || "";
      setCell(ws, `A${row}`, `${team.name || `Team ${i + 1}`} ${p + 1}`);
      setCell(ws, `C${row}`, name);
      setCell(ws, `D${row}`, "");
      setCell(ws, `E${row}`, name);
    }
  });
}

function fillBegegnungen(wb) {
  state.begegnungen.forEach((enc, i) => {
    const ws = wb.Sheets[`${i + 1}.Begegnung`];
    const a = state.allgemeines;
    setCell(ws, "M5", a.group);
    setCell(ws, "D5", enc.homeTeam);
    setCell(ws, "D7", enc.guestTeam);
    setCell(ws, "M7", a.venue);
    setCell(ws, "A27", a.place);
    if (a.date) setCell(ws, "E27", new Date(a.date));
    setCell(ws, "G29", a.coordinator);
    const homeCaptain = a.teams.find((t) => t.name === enc.homeTeam)?.captain || "";
    const guestCaptain = a.teams.find((t) => t.name === enc.guestTeam)?.captain || "";
    setCell(ws, "J29", homeCaptain);
    setCell(ws, "P29", guestCaptain);
    if (enc.startTime) setCell(ws, "A10", enc.startTime);
    if (enc.endTime) setCell(ws, "A22", enc.endTime);

    const totals = calcEncounter(enc);

    MATCHES.forEach((m, idx) => {
      const row = m.row;
      const match = enc.matches[idx];
      setCell(ws, `D${row}`, match.home1);
      setCell(ws, `G${row}`, match.guest1);
      if (m.doubles) {
        setCell(ws, `D${row + 1}`, match.home2);
        setCell(ws, `G${row + 1}`, match.guest2);
      }
      const flat = [match.sets[0].h, match.sets[0].g, match.sets[1].h, match.sets[1].g, match.sets[2].h, match.sets[2].g];
      ["I", "J", "K", "L", "M", "N"].forEach((col, cIdx) => {
        const val = parseNum(flat[cIdx]);
        if (val != null) setCell(ws, `${col}${row}`, val);
      });

      let hs = 0; let gs = 0; let hp = 0; let gp = 0;
      match.sets.forEach((s) => {
        const h = parseNum(s.h); const g = parseNum(s.g);
        if (h == null || g == null) return;
        hp += h; gp += g;
        if (h > g) hs += 1;
        if (g > h) gs += 1;
      });
      setCell(ws, `O${row}`, hp);
      setCell(ws, `P${row}`, gp);
      setCell(ws, `Q${row}`, hs);
      setCell(ws, `R${row}`, gs);
      setCell(ws, `S${row}`, hs > gs ? 1 : 0);
      setCell(ws, `T${row}`, gs > hs ? 1 : 0);
    });

    setCell(ws, "O21", totals.homePoints);
    setCell(ws, "P21", totals.guestPoints);
    setCell(ws, "Q21", totals.homeSets);
    setCell(ws, "R21", totals.guestSets);
    setCell(ws, "S21", totals.homeMatches);
    setCell(ws, "T21", totals.guestMatches);
    setCell(ws, "C21", totals.homeMatches > totals.guestMatches ? enc.homeTeam : totals.guestMatches > totals.homeMatches ? enc.guestTeam : "Unentschieden");
    setCell(ws, "I21", `${totals.homeMatches}:${totals.guestMatches}`);
  });
}

document.getElementById("downloadBtn").addEventListener("click", async () => {
  const status = document.getElementById("status");
  try {
    validateData();
    const wb = await loadWorkbook();
    fillAllgemeines(wb);
    fillSetzlisten(wb);
    fillBegegnungen(wb);

    const out = XLSX.write(wb, { type: "array", bookType: "xlsx", cellDates: true });
    const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "spielbericht_ausgefuellt.xlsx";
    a.click();
    URL.revokeObjectURL(a.href);
    status.textContent = "XLSX erfolgreich erstellt.";
  } catch (e) {
    status.textContent = e.message;
  }
});

buildTabs();
renderAllgemeines();
renderSetzlisten();
renderBegegnungen();
