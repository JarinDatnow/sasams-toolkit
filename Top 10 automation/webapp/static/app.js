const dbPathEl = document.getElementById("dbPath");
const dbPasswordEl = document.getElementById("dbPassword");
const termEl = document.getElementById("term");
const yearEl = document.getElementById("year");
const outputFolderEl = document.getElementById("outputFolder");
const generateBtn = document.getElementById("generateBtn");
const resultBanner = document.getElementById("resultBanner");
const logSection = document.getElementById("logSection");
const logOutput = document.getElementById("logOutput");

const modeTopTenBtn = document.getElementById("modeTopTen");
const modeAnnexureKBtn = document.getElementById("modeAnnexureK");
const annexureKSection = document.getElementById("annexureKSection");
const outputRuleNumber = document.getElementById("outputRuleNumber");
const tagline = document.getElementById("tagline");
const rule1Heading = document.getElementById("rule1Heading");
const rule2Heading = document.getElementById("rule2Heading");

const grouping8to12El = document.getElementById("grouping8to12");
const groupingR7El = document.getElementById("groupingR7");
const template8to12Field = document.getElementById("template8to12Field");
const templateR7Field = document.getElementById("templateR7Field");
const template8to12El = document.getElementById("template8to12");
const templateR7El = document.getElementById("templateR7");

let mode = "topten";
let lastOutputFolder = "";
const savedDefaults = {}; // per-mode defaults cache, so switching modes doesn't refetch

const MODE_COPY = {
  topten: {
    tagline: "You are not your test average.",
    rule1: "The first rule of Top 10 Club",
    rule2: "The second rule of Top 10 Club",
  },
  annexurek: {
    tagline: "Seven levels. No exceptions.",
    rule1: "The first rule of Annexure K",
    rule2: "The second rule of Annexure K",
  },
};

function setMode(newMode) {
  mode = newMode;
  modeTopTenBtn.classList.toggle("active", mode === "topten");
  modeAnnexureKBtn.classList.toggle("active", mode === "annexurek");
  annexureKSection.classList.toggle("hidden", mode !== "annexurek");
  outputRuleNumber.textContent = mode === "annexurek" ? "04" : "03";

  const copy = MODE_COPY[mode];
  tagline.textContent = copy.tagline;
  rule1Heading.textContent = copy.rule1;
  rule2Heading.textContent = copy.rule2;

  resultBanner.classList.add("hidden");
  logSection.classList.add("hidden");

  loadDefaults();
}

modeTopTenBtn.addEventListener("click", () => setMode("topten"));
modeAnnexureKBtn.addEventListener("click", () => setMode("annexurek"));

function updateGroupingFieldVisibility() {
  template8to12Field.classList.toggle("hidden", !grouping8to12El.checked);
  templateR7Field.classList.toggle("hidden", !groupingR7El.checked);
}
grouping8to12El.addEventListener("change", updateGroupingFieldVisibility);
groupingR7El.addEventListener("change", updateGroupingFieldVisibility);

async function loadDefaults() {
  if (savedDefaults[mode]) {
    applyDefaults(savedDefaults[mode]);
    return;
  }
  try {
    const res = await fetch(`/api/defaults?mode=${mode}`);
    const data = await res.json();
    savedDefaults[mode] = data;
    applyDefaults(data);
  } catch (err) {
    // server not reachable yet on first paint - ignore, fields stay blank
  }
}

function applyDefaults(data) {
  dbPathEl.value = data.db_path || "";
  termEl.value = String(data.term || 3);
  yearEl.value = data.year || new Date().getFullYear();
  outputFolderEl.value = data.output_folder || "";
  lastOutputFolder = data.output_folder || "";

  if (mode === "annexurek") {
    // Groupings are intentionally left unchecked by default - pick
    // explicitly each run rather than inheriting a "generate everything"
    // default, so GENERATE only ever does what's checked on screen.
    const templates = data.templates || {};
    template8to12El.value = templates["8-12"] || "";
    templateR7El.value = templates["R-7"] || "";
    updateGroupingFieldVisibility();
  }
}

document.getElementById("browseDb").addEventListener("click", async () => {
  const res = await fetch("/api/browse-file", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ kind: "db" }),
  });
  const data = await res.json();
  if (data.path) dbPathEl.value = data.path;
});

document.getElementById("browseTemplate8to12").addEventListener("click", async () => {
  const res = await fetch("/api/browse-file", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ kind: "template" }),
  });
  const data = await res.json();
  if (data.path) template8to12El.value = data.path;
});

document.getElementById("browseTemplateR7").addEventListener("click", async () => {
  const res = await fetch("/api/browse-file", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ kind: "template" }),
  });
  const data = await res.json();
  if (data.path) templateR7El.value = data.path;
});

document.getElementById("browseFolder").addEventListener("click", async () => {
  const res = await fetch("/api/browse-folder", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ initial: outputFolderEl.value || "" }),
  });
  const data = await res.json();
  if (data.path) outputFolderEl.value = data.path;
});

document.getElementById("togglePw").addEventListener("click", (e) => {
  const showing = dbPasswordEl.type === "text";
  dbPasswordEl.type = showing ? "password" : "text";
  e.target.textContent = showing ? "SHOW" : "HIDE";
});

document.getElementById("openFolderBtn")?.addEventListener("click", openLastFolder);

async function openLastFolder() {
  if (!lastOutputFolder) return;
  await fetch("/api/open-folder", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path: lastOutputFolder }),
  });
}

function showResult(kind, title, message, showOpenBtn) {
  resultBanner.className = "result " + kind;
  resultBanner.innerHTML =
    `<span class="result-title">${title}</span>${message}` +
    (showOpenBtn
      ? `<div class="result-actions"><button type="button" id="openFolderBtn" class="btn-ghost">OPEN FOLDER</button></div>`
      : "");
  resultBanner.classList.remove("hidden");
  if (showOpenBtn) {
    document.getElementById("openFolderBtn").addEventListener("click", openLastFolder);
  }
}

function buildPayload() {
  if (mode === "annexurek") {
    const groupings = [];
    if (grouping8to12El.checked) groupings.push("8-12");
    if (groupingR7El.checked) groupings.push("R-7");
    return {
      mode: "annexurek",
      db_path: dbPathEl.value,
      db_password: dbPasswordEl.value,
      term: termEl.value,
      year: yearEl.value,
      groupings,
      templates: { "8-12": template8to12El.value, "R-7": templateR7El.value },
      output_folder: outputFolderEl.value,
    };
  }
  return {
    mode: "topten",
    db_path: dbPathEl.value,
    db_password: dbPasswordEl.value,
    term: termEl.value,
    year: yearEl.value,
    output_folder: outputFolderEl.value,
  };
}

generateBtn.addEventListener("click", async () => {
  resultBanner.classList.add("hidden");
  logSection.classList.add("hidden");
  logOutput.textContent = "";

  generateBtn.disabled = true;
  generateBtn.querySelector(".btn-fight-text").textContent = "FIGHTING...";

  try {
    const res = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(buildPayload()),
    });
    const data = await res.json();

    if (data.log) {
      logOutput.textContent = data.log;
      logSection.classList.remove("hidden");
    }

    if (data.ok) {
      lastOutputFolder = outputFolderEl.value || lastOutputFolder;
      const paths = data.out_paths || (data.out_path ? [data.out_path] : []);
      showResult(
        "ok",
        "SELF IMPROVEMENT IS MASTURBATION.",
        `Workbook${paths.length > 1 ? "s" : ""} saved:<br>${paths.map(escapeHtml).join("<br>")}`,
        true
      );
    } else {
      showResult("error", "THIS IS YOUR LIFE, AND IT'S ENDING ONE ERROR AT A TIME.",
        escapeHtml(data.error || "Something went wrong."), false);
    }
  } catch (err) {
    showResult("error", "CONNECTION LOST", escapeHtml(String(err)), false);
  } finally {
    generateBtn.disabled = false;
    generateBtn.querySelector(".btn-fight-text").textContent = "GENERATE";
  }
});

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

setMode("topten");
