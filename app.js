import { createExcelEngine } from "./engine.js";

const DATA_BASE_URL =
  "https://raw.githubusercontent.com/geekr2013/univ-admission-possibility/main/data";
const DATA_URL = `${DATA_BASE_URL}/admissions.json`;
const AUDIT_URL = `${DATA_BASE_URL}/audit.json`;
const INTEGRITY_URL = `${DATA_BASE_URL}/integrity.json`;

const tiers = [
  { id: "all", label: "모든 대학", schools: [] },
  { id: "sky", label: "SKY", schools: ["서울대학교", "고려대학교", "연세대학교"] },
  { id: "seoseonghan", label: "서성한", schools: ["서강대학교", "성균관대학교", "한양대학교"] },
  { id: "jungkyung", label: "중경외시이", schools: ["중앙대학교", "경희대학교", "한국외국어대학교", "서울시립대학교", "이화여자대학교"] },
  { id: "geondong", label: "건동홍아숙", schools: ["건국대학교", "동국대학교", "홍익대학교", "아주대학교", "숙명여자대학교"] },
  { id: "guksoong", label: "국숭세단인", schools: ["국민대학교", "숭실대학교", "세종대학교", "단국대학교", "인하대학교", "인천대학교"] },
  { id: "gwangmyung", label: "광명상가", schools: ["광운대학교", "명지대학교", "상명대학교", "가톨릭대학교", "가천대학교"] },
  { id: "regional", label: "지거국", schools: ["강원대학교", "경북대학교", "경상국립대학교", "부산대학교", "전남대학교", "전북대학교", "충남대학교", "충북대학교"] },
];

const subjectRows = [
  { key: "korean", label: "국어", defaultSubject: "국어", standard: 127, percentile: 95, grade: 2 },
  { key: "math", label: "수학", defaultSubject: "미적분", standard: 135, percentile: 98, grade: 1 },
  { key: "english", label: "영어", defaultSubject: "", standard: "", percentile: "", grade: 3 },
  { key: "history", label: "한국사", defaultSubject: "", standard: "", percentile: "", grade: 1 },
  { key: "explore1", label: "탐구1", defaultSubject: "화학Ⅰ", standard: 64, percentile: 91, grade: 2 },
  { key: "explore2", label: "탐구2", defaultSubject: "물리학Ⅱ", standard: 71, percentile: 98, grade: 1 },
  { key: "language2", label: "제2외국어/한문", defaultSubject: "", standard: "", percentile: "", grade: "" },
];

const state = {
  data: null,
  scores: {},
  gradeIndex: new Map(),
  conversionIndex: {},
  audit: null,
  auditRows: new Map(),
  integrity: null,
  engine: null,
};

const $ = (id) => document.getElementById(id);
const numberFormat = new Intl.NumberFormat("ko-KR");

function safeNumber(value, fallback = 0) {
  const num = Number(value);
  return Number.isFinite(num) ? num : fallback;
}

function ratioNumber(value, item, key) {
  if (typeof value === "string") {
    const match = value.match(/^\s*([0-9.]+)(?:\(([0-9.]+)\))?/);
    if (!match) return 0;
    const outside = safeNumber(match[1], 0);
    const inside = match[2] != null ? safeNumber(match[2], outside) : outside;
    if (match[2] == null) return outside;

    const math = state.scores.math?.subject || "";
    const explores = [state.scores.explore1?.subject, state.scores.explore2?.subject];
    const scienceCount = explores.filter(isScienceSubject).length;
    const socialCount = explores.filter(isSocialSubject).length;
    const majorText = normalizeText(`${item?.major || ""} ${item?.requirements || ""}`);
    const naturalMajor = /공학|자연|의예|약학|간호\(자연\)|수학|물리|화학|생명|소프트웨어|컴퓨터|전자|전기|반도체|융합|데이터|인공지능|건축|통계/.test(majorText);
    const naturalChoice = math.includes("미적분") || math.includes("기하") || scienceCount > socialCount;

    if (key === "korean") return naturalMajor || naturalChoice ? outside : inside;
    if (key === "math") return naturalChoice ? Math.max(outside, inside) : Math.min(outside, inside);
    if (key === "explore1" || key === "explore2") return naturalMajor || scienceCount >= socialCount ? outside : inside;
    return outside;
  }
  return safeNumber(value, 0);
}

function cleanSchoolName(name) {
  return String(name || "").replace(/_\S+$/, "");
}

function rowId(row) {
  return [row.university, row.group, row.track, row.major].map((part) => String(part || "")).join("||");
}

function includesSchool(university, schools) {
  const clean = cleanSchoolName(university);
  return schools.some((school) => clean.includes(school));
}

function gradeToPercentile(grade) {
  const table = [100, 96, 89, 77, 60, 40, 23, 11, 4];
  return table[Math.max(1, Math.min(9, safeNumber(grade, 9))) - 1];
}

function isScienceSubject(subject) {
  return /물리|화학|생명과학|지구과학|과학탐구/.test(String(subject || ""));
}

function isSocialSubject(subject) {
  return /생활과 윤리|윤리와 사상|한국지리|세계지리|동아시아사|세계사|경제|정치와 법|사회.?문화|사회탐구/.test(String(subject || ""));
}

function subjectForRange(key, subject) {
  if (key === "korean") return "국어";
  if (key === "math") return "수학";
  if (key === "english") return "영어";
  if (key === "history") return "한국사";
  return subject || "";
}

function normalizeText(value) {
  return String(value || "").replace(/[\s·ㆍ・]/g, "").toLowerCase();
}

function examKey(year, grade, month, subject) {
  return [year, grade, month, normalizeText(subject)].join("|");
}

function buildGradeIndex() {
  state.gradeIndex.clear();
  for (const item of state.data.gradeRanges) {
    const key = examKey(item.year, item.gradeLevel, item.month, item.subject);
    if (!state.gradeIndex.has(key)) state.gradeIndex.set(key, []);
    state.gradeIndex.get(key).push(item);
  }
  for (const ranges of state.gradeIndex.values()) {
    ranges.sort((a, b) => safeNumber(a.grade) - safeNumber(b.grade));
  }
}

function selectedExam() {
  return {
    year: safeNumber($("mockYear").value),
    grade: $("mockGrade").value,
    month: $("mockMonth").value,
  };
}

function isAuditInputActive() {
  if (!state.audit?.exam || !state.audit?.scores) return false;
  const exam = selectedExam();
  const auditExam = state.audit.exam;
  if (safeNumber(auditExam.year) !== exam.year || auditExam.grade !== exam.grade || auditExam.month !== exam.month) return false;

  return ["korean", "math", "english", "history", "explore1", "explore2"].every((key) => {
    const current = state.scores[key] || {};
    const audit = state.audit.scores[key] || {};
    return String(current.subject || "") === String(audit.subject || "")
      && safeNumber(current.standard, null) === safeNumber(audit.standard, null)
      && safeNumber(current.percentile, null) === safeNumber(audit.percentile, null)
      && safeNumber(current.grade, null) === safeNumber(audit.grade, null);
  });
}

function rangeForScore(key, score) {
  const exam = selectedExam();
  const subject = subjectForRange(key, score?.subject);
  return state.gradeIndex.get(examKey(exam.year, exam.grade, exam.month, subject)) || [];
}

function percentileFromStandard(key, score) {
  const standard = safeNumber(score?.standard, NaN);
  if (!Number.isFinite(standard)) return null;
  const ranges = rangeForScore(key, score);
  if (!ranges.length) return null;

  const first = ranges[0];
  const last = ranges[ranges.length - 1];
  if (standard > safeNumber(first.standardMax)) {
    return { percentile: 99, grade: 1, status: "상한 초과" };
  }
  if (standard < safeNumber(last.standardMin)) {
    return { percentile: 1, grade: 9, status: "하한 미만" };
  }

  const found = ranges.find((range) => standard >= safeNumber(range.standardMin) && standard <= safeNumber(range.standardMax));
  if (!found) return null;

  const grade = safeNumber(found.grade);
  const min = safeNumber(found.standardMin);
  const max = safeNumber(found.standardMax);
  const position = max === min ? 0.5 : (standard - min) / (max - min);
  const upper = gradeToPercentile(grade);
  const lower = gradeToPercentile(Math.min(9, grade + 1));
  const percentile = lower + (upper - lower) * position;
  return { percentile, grade, status: `${grade}등급 구간` };
}

function effectivePercentile(key, score) {
  const fromStandard = percentileFromStandard(key, score);
  if (fromStandard) return fromStandard.percentile;
  const percentile = safeNumber(score?.percentile, NaN);
  if (Number.isFinite(percentile) && percentile > 0) return percentile;
  return gradeToPercentile(score?.grade);
}

function gradePointFromTable(points, grade) {
  const index = Math.max(1, Math.min(9, safeNumber(grade, 9))) - 1;
  const value = safeNumber(points?.[index], NaN);
  if (!Number.isFinite(value)) return gradeToPercentile(grade);
  const max = Math.max(...(points || []).map((point) => safeNumber(point, 0)));
  if (max > 100) return (value / max) * 100;
  return value;
}

function additiveGradePoint(points, grade) {
  const numeric = (points || []).map((point) => safeNumber(point, NaN)).filter(Number.isFinite);
  if (!numeric.length) return 0;
  const index = Math.max(1, Math.min(9, safeNumber(grade, 9))) - 1;
  return safeNumber(points?.[index], 0);
}

function effectiveSubjectScore(key, score, item) {
  if (key === "english") return gradePointFromTable(item.englishPoints, score?.grade);
  if (key === "history") return gradePointFromTable(item.historyPoints, score?.grade);
  return effectivePercentile(key, score);
}

function admissionYear(item) {
  const match = String(item.year || "").match(/20\d{2}/);
  return match ? match[0] : "2026";
}

function conversionColumn(item) {
  const table = state.conversionIndex[admissionYear(item)] || {};
  const university = item.university || "";
  const type = item.conversionType;
  const explores = [state.scores.explore1?.subject, state.scores.explore2?.subject];
  const scienceCount = explores.filter(isScienceSubject).length;
  const socialCount = explores.filter(isSocialSubject).length;
  const candidates = [];

  if (type) candidates.push(`${university}-${type}`);
  if (scienceCount >= socialCount) candidates.push(`${university}-과탐`);
  if (socialCount > scienceCount) candidates.push(`${university}-사탐`);
  candidates.push(university);

  const normalizedCandidates = candidates.map(normalizeText);
  return Object.keys(table).find((column) => normalizedCandidates.includes(normalizeText(column)))
    || Object.keys(table).find((column) => normalizeText(column).startsWith(normalizeText(university)));
}

function convertedScore(item, percentile) {
  const table = state.conversionIndex[admissionYear(item)] || {};
  const column = conversionColumn(item);
  const rounded = Math.max(0, Math.min(100, Math.round(safeNumber(percentile, 0))));
  const values = column ? table[column] : null;
  if (!values) return safeNumber(percentile, 0);
  if (values[String(rounded)] != null) return safeNumber(values[String(rounded)], percentile);

  const lower = Math.floor(rounded);
  const upper = Math.ceil(rounded);
  if (values[String(lower)] != null && values[String(upper)] != null && upper !== lower) {
    const lowValue = safeNumber(values[String(lower)]);
    const highValue = safeNumber(values[String(upper)]);
    return lowValue + (highValue - lowValue) * (percentile - lower);
  }
  return safeNumber(percentile, 0);
}

function academicSubjectValue(key, score, item) {
  const metric = String(item.metric || "");
  if (metric.includes("변환표준점수")) return convertedScore(item, effectivePercentile(key, score));
  if (metric.includes("표준점수") && !metric.includes("변환")) return safeNumber(score?.standard, effectivePercentile(key, score));
  if (metric.includes("등급")) return gradeToPercentile(score?.grade);
  return effectivePercentile(key, score);
}

function mathBonus(item) {
  const subject = state.scores.math?.subject || "";
  const bonuses = item.bonuses || {};
  if (subject.includes("확률") || subject.includes("통계")) return safeNumber(bonuses.probabilityStats, 0);
  if (subject.includes("미적분")) return safeNumber(bonuses.calculus, 0);
  if (subject.includes("기하")) return safeNumber(bonuses.geometry, 0);
  return 0;
}

function exploreBonus(item) {
  const bonuses = item.bonuses || {};
  const subjects = [state.scores.explore1?.subject, state.scores.explore2?.subject];
  const scienceCount = subjects.filter(isScienceSubject).length;
  const socialCount = subjects.filter(isSocialSubject).length;
  const scienceBonus = safeNumber(bonuses.science, 0);
  const socialBonus = safeNumber(bonuses.social, 0);
  const exploreAverage = (effectivePercentile("explore1", state.scores.explore1 || {}) + effectivePercentile("explore2", state.scores.explore2 || {})) / 2;

  let bonus = 0;
  if (scienceCount > 0 && scienceBonus) bonus += scienceBonus <= 1 ? exploreAverage * scienceBonus : scienceBonus;
  if (socialCount > 0 && socialBonus) bonus += socialBonus <= 1 ? exploreAverage * socialBonus : socialBonus;
  return bonus;
}

function requirementStatus(item) {
  const requirement = normalizeText(item.requirements);
  if (!requirement) return { eligible: true, message: "" };

  const math = state.scores.math?.subject || "";
  const explores = [state.scores.explore1?.subject, state.scores.explore2?.subject];
  const scienceCount = explores.filter(isScienceSubject).length;
  const problems = [];

  if (requirement.includes("미적분") || requirement.includes("기하")) {
    if (!(math.includes("미적분") || math.includes("기하"))) problems.push("수학 미적분/기하 필요");
  }
  if (requirement.includes("과탐2") || requirement.includes("과탐2과목") || requirement.includes("과탐2목")) {
    if (scienceCount < 2) problems.push("과탐 2과목 필요");
  } else if (requirement.includes("과탐")) {
    if (scienceCount < 1) problems.push("과탐 필요");
  }
  if (requirement.includes("생명과학필수")) {
    if (!explores.some((subject) => String(subject || "").includes("생명과학"))) problems.push("생명과학 필요");
  }

  return {
    eligible: problems.length === 0,
    message: problems.join(", "),
  };
}

function weightedScore(item) {
  const weights = item.weights || {};
  const fullScore = safeNumber(item.fullScore, 100);
  const scale = fullScore > 100 ? fullScore / 100 : 1;
  const exploreAverage = (academicSubjectValue("explore1", state.scores.explore1 || {}, item) + academicSubjectValue("explore2", state.scores.explore2 || {}, item)) / 2;

  let totalWeight = 0;
  let weighted = 0;
  const academicParts = [
    ["korean", ratioNumber(weights.korean, item, "korean"), academicSubjectValue("korean", state.scores.korean || {}, item)],
    ["math", ratioNumber(weights.math, item, "math"), academicSubjectValue("math", state.scores.math || {}, item)],
  ];

  const explore1Weight = ratioNumber(weights.explore1, item, "explore1");
  const explore2Weight = ratioNumber(weights.explore2, item, "explore2");
  if (explore1Weight > 0) academicParts.push(["explore1", explore1Weight, academicSubjectValue("explore1", state.scores.explore1 || {}, item)]);
  if (explore2Weight > 0) {
    academicParts.push(["explore2", explore2Weight, explore1Weight > 0 ? academicSubjectValue("explore2", state.scores.explore2 || {}, item) : exploreAverage]);
  }

  for (const [, rawWeight, value] of academicParts) {
    const weight = safeNumber(rawWeight, 0);
    if (weight <= 0) continue;
    weighted += value * weight;
    totalWeight += weight;
  }

  let baseScore;
  if (totalWeight === 0) {
    const core = ["korean", "math", "explore1", "explore2"].map((key) => effectivePercentile(key, state.scores[key] || {}));
    baseScore = core.reduce((sum, value) => sum + value, 0) / core.length * scale;
  } else {
    baseScore = (weighted / 100) * scale;
  }

  const englishWeight = ratioNumber(weights.english, item, "english");
  const historyWeight = ratioNumber(weights.history, item, "history");
  const english = englishWeight > 0
    ? effectiveSubjectScore("english", state.scores.english || {}, item) * englishWeight / 100 * scale
    : additiveGradePoint(item.englishPoints, state.scores.english?.grade);
  const history = historyWeight > 0
    ? effectiveSubjectScore("history", state.scores.history || {}, item) * historyWeight / 100 * scale
    : additiveGradePoint(item.historyPoints, state.scores.history?.grade);

  const total = baseScore + safeNumber(english, 0) + safeNumber(history, 0) + mathBonus(item) * scale + exploreBonus(item) * scale;
  return Math.min(fullScore, total);
}

function normalizeCut(item, key) {
  const cut = safeNumber(item[key], NaN);
  if (!Number.isFinite(cut)) return null;
  return cut;
}

function estimateChance(item) {
  const requirement = requirementStatus(item);
  const myScore = weightedScore(item);
  const cut70 = normalizeCut(item, "cut70");
  const cut50 = normalizeCut(item, "cut50");
  const baseCut = cut70 || cut50 || 85;
  const diff = myScore - baseCut;
  const fullScore = safeNumber(item.fullScore, 100);
  const defaultSpread = Math.max(1.8, fullScore * 0.018);
  const cutSpread = cut50 && cut70 ? Math.max(defaultSpread * 0.5, Math.abs(cut50 - cut70)) : defaultSpread;
  const probability = Math.round(100 / (1 + Math.exp(-(diff / Math.max(defaultSpread, cutSpread * 1.2)))));
  const adjusted = Math.max(5, Math.min(95, probability));
  return {
    myScore,
    cut70: baseCut,
    cut50,
    diff,
    chance: requirement.eligible ? adjusted : 0,
    requirement,
  };
}

function chanceLabel(chance, exactLabel = "") {
  if (exactLabel) return exactLabel;
  if (chance <= 0) return "조건 불일치";
  if (chance >= 90) return "90% 이상";
  if (chance >= 80) return "80% 이상";
  if (chance >= 70) return "70% 이상";
  if (chance >= 50) return "50% 이상";
  if (chance >= 30) return "30% 이상";
  return "30% 미만";
}

function chanceClass(chance) {
  if (chance >= 70) return "good";
  if (chance >= 30) return "watch";
  return "risk";
}

function getFilteredResults() {
  const tier = tiers.find((item) => item.id === $("tierFilter").value) || tiers[0];
  const university = $("universityFilter").value;
  const majorQuery = $("majorSearch").value.trim().toLowerCase();
  const minChance = safeNumber($("minChance").value, 0);
  const auditFilter = $("auditFilter").value;

  const targetYear = $("targetYear").value;
  const inputs = { exam: selectedExam(), ...state.scores };
  let rows = state.data.admissions
    .filter((item) => String(item.year).includes(targetYear))
    .filter((item) => !tier.schools.length || includesSchool(item.university, tier.schools))
    .filter((item) => !university || item.university === university)
    .filter((item) => !majorQuery || `${item.major} ${item.university}`.toLowerCase().includes(majorQuery))
    .map((item) => {
      const auditRow = isAuditInputActive() ? state.auditRows.get(rowId(item)) : null;
      const exact = state.engine.calculate(item, inputs);
      const cut70 = safeNumber(item.cut70, safeNumber(item.cut50, 0));
      const estimate = {
        myScore: exact.myScore,
        cut70,
        cut50: safeNumber(item.cut50, 0),
        diff: exact.myScore - cut70,
        chance: exact.chance,
        chanceLabel: exact.chanceLabel,
        requirement: { eligible: true, message: item.requirements || "" },
      };
      return { ...item, estimate, auditRow };
    })
    .filter((item) => auditFilter === "all" || auditStatusForFilter(item.auditRow) === auditFilter)
    .filter((item) => item.estimate.chance >= minChance);

  const sortMode = $("sortMode").value;
  rows.sort((a, b) => {
    if (sortMode === "reach") return a.estimate.chance - b.estimate.chance;
    if (sortMode === "university") return `${a.university}${a.major}`.localeCompare(`${b.university}${b.major}`, "ko");
    if (sortMode === "competition") return safeNumber(a.competition, 99) - safeNumber(b.competition, 99);
    return b.estimate.chance - a.estimate.chance;
  });

  return rows;
}

function auditStatusForFilter(auditRow) {
  if (!state.audit || !isAuditInputActive() || !auditRow) return "unknown";
  return auditRow.status || "unknown";
}

function chanceFromLabel(label, fallback) {
  const text = String(label || "");
  if (text.includes("90")) return 95;
  if (text.includes("80")) return 85;
  if (text.includes("70")) return 75;
  if (text.includes("50")) return 55;
  if (text.includes("30% 미만")) return 20;
  if (text.includes("30")) return 35;
  return fallback;
}

function updateMockOptions() {
  const options = state.data.mockOptions;
  const years = [...new Set(options.map((item) => item.year))].sort((a, b) => b - a);
  $("mockYear").innerHTML = years.map((year) => `<option>${year}</option>`).join("");
  $("mockYear").value = years[0];

  function syncGrades() {
    const year = safeNumber($("mockYear").value);
    const grades = [...new Set(options.filter((item) => item.year === year).map((item) => item.grade))];
    $("mockGrade").innerHTML = grades.map((grade) => `<option>${grade}</option>`).join("");
    if (grades.includes("고3")) $("mockGrade").value = "고3";
    syncMonths();
  }

  function syncMonths() {
    const year = safeNumber($("mockYear").value);
    const grade = $("mockGrade").value;
    const found = options.find((item) => item.year === year && item.grade === grade);
    const months = found?.months?.length ? found.months : ["6월"];
    $("mockMonth").innerHTML = months.map((month) => `<option>${month}</option>`).join("");
    if (state.audit?.exam && safeNumber(state.audit.exam.year) === year && state.audit.exam.grade === grade && months.includes(state.audit.exam.month)) {
      $("mockMonth").value = state.audit.exam.month;
    }
  }

  $("mockYear").addEventListener("change", () => {
    syncGrades();
    render();
  });
  $("mockGrade").addEventListener("change", () => {
    syncMonths();
    render();
  });
  $("mockMonth").addEventListener("change", render);
  syncGrades();
}

function updateFilters() {
  $("tierFilter").innerHTML = tiers.map((tier) => `<option value="${tier.id}">${tier.label}</option>`).join("");
  const years = [...new Set(state.data.admissions.map((item) => String(item.year).match(/20\d{2}/)?.[0]).filter(Boolean))].sort((a, b) => b - a);
  $("targetYear").innerHTML = years.map((year) => `<option value="${year}">${year}학년도 정시</option>`).join("");
  function syncUniversities() {
    const year = $("targetYear").value;
    const universities = [...new Set(state.data.admissions.filter((item) => String(item.year).includes(year)).map((item) => item.university))].sort((a, b) => a.localeCompare(b, "ko"));
    $("universityFilter").innerHTML = `<option value="">전체</option>${universities.map((name) => `<option>${name}</option>`).join("")}`;
  }
  $("targetYear").addEventListener("change", () => { syncUniversities(); render(); });
  syncUniversities();
}

function buildScoreInputs() {
  const template = $("scoreTemplate");
  const grid = $("scoreGrid");
  const subjectOptions = {
    korean: state.data.subjects.korean,
    math: state.data.subjects.math,
    english: [""],
    history: [""],
    explore1: state.data.subjects.explore,
    explore2: state.data.subjects.explore,
    language2: ["", ...(state.data.subjects.language2 || [])],
  };

  grid.innerHTML = "";
  subjectRows.forEach((row) => {
    state.scores[row.key] = { subject: row.defaultSubject, standard: row.standard, percentile: row.percentile, grade: row.grade };
    const node = template.content.firstElementChild.cloneNode(true);
    node.dataset.key = row.key;
    node.querySelector(".subject-label").textContent = row.label;

    const select = node.querySelector(".subject-select");
    select.innerHTML = subjectOptions[row.key].map((subject) => `<option>${subject}</option>`).join("");
    select.value = subjectOptions[row.key].includes(row.defaultSubject) ? row.defaultSubject : subjectOptions[row.key][0] || "";
    select.disabled = row.key === "english" || row.key === "history";

    node.querySelector(".standard-input").value = row.standard;
    node.querySelector(".percentile-input").value = row.percentile;
    node.querySelector(".grade-input").value = row.grade;
    if (["english", "history", "language2"].includes(row.key)) {
      node.querySelector(".standard-input").disabled = true;
      node.querySelector(".percentile-input").disabled = true;
    }

    node.addEventListener("input", () => {
      state.scores[row.key] = {
        subject: select.value,
        standard: node.querySelector(".standard-input").value,
        percentile: node.querySelector(".percentile-input").value,
        grade: node.querySelector(".grade-input").value,
      };
      render();
    });
    grid.appendChild(node);
  });
}

function renderRangeHints() {
  document.querySelectorAll(".score-item").forEach((node) => {
    const key = node.dataset.key;
    const score = state.scores[key] || {};
    const hint = node.querySelector(".range-hint");
    if (!hint) return;
    if (key === "english" || key === "history" || key === "language2") {
      hint.textContent = "등급을 기준으로 반영합니다.";
      hint.className = "range-hint";
      return;
    }
    const ranges = rangeForScore(key, score);
    const estimate = percentileFromStandard(key, score);
    if (!ranges.length) {
      hint.textContent = "선택한 시험의 표준점수 구간이 없어 입력 백분위를 사용합니다.";
      hint.className = "range-hint warning";
      return;
    }
    const min = ranges[ranges.length - 1].standardMin;
    const max = ranges[0].standardMax;
    if (estimate) {
      hint.textContent = `표준점수 기준 ${estimate.status}, 추정 백분위 ${estimate.percentile.toFixed(1)} 적용`;
      hint.className = "range-hint";
    } else {
      hint.textContent = `유효 표준점수 ${min}~${max}. 표준점수를 넣으면 해당 시험 기준으로 보정됩니다.`;
      hint.className = "range-hint";
    }
  });
}

function renderSourceStatus() {
  const source = state.data.source || {};
  const generated = source.generatedAt ? new Date(source.generatedAt) : null;
  const generatedText = generated && !Number.isNaN(generated.getTime())
    ? generated.toLocaleString("ko-KR", { timeZone: "Asia/Seoul" })
    : "확인 불가";
  const validUntil = generated && !Number.isNaN(generated.getTime()) ? new Date(generated.getTime() + 7 * 24 * 60 * 60 * 1000) : null;
  const validText = validUntil ? validUntil.toLocaleDateString("ko-KR", { timeZone: "Asia/Seoul" }) : "확인 불가";
  const options = state.data.mockOptions || [];
  const latest = options[0];
  const auditText = state.audit?.summary
    ? ` · 원본 일치 ${state.audit.summary.verifiedWithinPoint5}/${state.audit.summary.matched}`
    : "";
  $("updatedText").textContent = `데이터 생성: ${generatedText} · 확인 권장 기한: ${validText} · 최신 시험: ${latest ? `${latest.year} ${latest.grade} ${latest.months.join("/")}` : "-"}${auditText}`;
}

function shortHash(value) {
  return String(value || "").slice(0, 10);
}

function renderValidationPanel() {
  const audit = state.audit?.summary || {};
  const integrity = state.integrity || {};
  const macros = integrity.macros || {};
  const workbook = integrity.workbook || {};
  const coverage = integrity.coverage || {};
  const source = integrity.source || state.data.source || {};
  const exact = safeNumber(audit.verifiedWithinPoint5, 0);
  const closeTotal = safeNumber(audit.closeWithin3, 0);
  const closeOnly = Math.max(0, closeTotal - exact);
  const mismatch = safeNumber(audit.mismatch, 0);
  const matched = safeNumber(audit.matched, 0);
  const admissions = coverage.admissions?.totalRows || state.data.admissions.length;
  const parsedSheets = workbook.parsedSheets?.length || 0;
  const sheetCount = workbook.sheetCount || 0;
  const macroText = workbook.hasVbaProject
    ? (macros.totalLineCount ? `${numberFormat.format(macros.totalLineCount)}줄 VBA` : "VBA 포함")
    : "매크로 없음";

  $("validationVerdict").textContent = integrity.verdict?.isExcelEquivalent ? "100% 검증 완료" : "부분 검증 완료";
  $("validationAudit").textContent = matched
    ? `일치 ${exact} · 근접 ${closeOnly} · 오차 큼 ${mismatch}`
    : "대조값 없음";
  $("validationMacro").textContent = macroText;
  $("validationCoverage").textContent = `${numberFormat.format(admissions)}개 모집단위 · ${parsedSheets}/${sheetCount}개 시트`;
  $("validationSummary").textContent = source.sha256
    ? `원본 ${source.file || "-"} · SHA ${shortHash(source.sha256)} · 매일 새벽 최신 원본 기준으로 재생성됩니다.`
    : "원본 파일 정보와 검증 리포트를 확인하는 중입니다.";
  $("validationPolicy").textContent = isAuditInputActive()
    ? "현재 입력은 최신 원본에 저장된 검증 입력과 같습니다. 저장된 카드에는 엑셀과 웹 엔진의 완전 일치 여부를 표시합니다."
    : "현재 카드는 VBA 번역 엔진 결과입니다. 원본에 저장된 동일 입력 결과가 없는 카드는 독립 실행 검증 대기로 표시됩니다.";
}

function renderStats(rows) {
  $("statCount").textContent = numberFormat.format(rows.length);
  $("statSafe").textContent = numberFormat.format(rows.filter((row) => row.estimate.chance >= 90).length);
  $("statGood").textContent = numberFormat.format(rows.filter((row) => row.estimate.chance >= 70).length);
  $("statUniversities").textContent = numberFormat.format(new Set(rows.map((row) => row.university)).size);

  const buckets = [
    ["90%+", rows.filter((row) => row.estimate.chance >= 90).length],
    ["70-89%", rows.filter((row) => row.estimate.chance >= 70 && row.estimate.chance < 90).length],
    ["30-69%", rows.filter((row) => row.estimate.chance >= 30 && row.estimate.chance < 70).length],
    ["~29%", rows.filter((row) => row.estimate.chance > 0 && row.estimate.chance < 30).length],
    ["조건 불일치", rows.filter((row) => row.estimate.chance <= 0).length],
  ];
  const max = Math.max(1, ...buckets.map(([, count]) => count));
  $("chanceBars").innerHTML = buckets
    .map(([label, count]) => `<div class="bar-row"><span>${label}</span><div class="bar-track"><div class="bar-fill" style="width:${(count / max) * 100}%"></div></div><span>${count}</span></div>`)
    .join("");

  const safe = rows.filter((row) => row.estimate.chance >= 90).slice(0, 1)[0];
  const reach = rows.filter((row) => row.estimate.chance >= 30 && row.estimate.chance < 70).slice(0, 1)[0];
  const blocked = rows.filter((row) => row.estimate.chance <= 0).length;
  const topUni = rows[0]?.university || "조건에 맞는 대학";
  const auditSummary = state.audit?.summary;
  const visibleAudit = rows.reduce((acc, row) => {
    const status = auditStatusForFilter(row.auditRow);
    acc[status] = (acc[status] || 0) + 1;
    return acc;
  }, {});
  const auditMessage = auditSummary && isAuditInputActive()
    ? `현재 결과 검증: 일치 ${visibleAudit.verified || 0}개, 근접 ${visibleAudit.close || 0}개, 오차 큼 ${visibleAudit.mismatch || 0}개, 검증 불가 ${visibleAudit.unknown || 0}개`
    : "현재 입력은 원본 저장 검증 기준과 달라 카드별 검증 배지는 참고용입니다.";
  $("insightList").innerHTML = [
    safe ? `${cleanSchoolName(safe.university)} ${safe.major}는 안정권 후보입니다.` : "90% 이상 조건이 적으면 최소 합격률을 낮춰 보세요.",
    reach ? `${cleanSchoolName(reach.university)} ${reach.major}는 도전권으로 비교해볼 만합니다.` : "상향 지원 후보는 정렬을 '상향 지원 순'으로 바꾸면 찾기 쉽습니다.",
    auditMessage,
    blocked ? `응시조건이 맞지 않는 학과 ${blocked}개는 합격률을 0%로 처리했습니다.` : `${topUni} 기준으로 유사 학과를 함께 비교해보세요.`,
  ].map((text) => `<li>${text}</li>`).join("");
}

function renderCards(rows) {
  const target = $("results");
  if (!rows.length) {
    target.innerHTML = `<div class="empty">조건에 맞는 학과가 없습니다. 최소 합격률이나 대학권역을 낮춰보세요.</div>`;
    return;
  }
  target.innerHTML = rows.slice(0, 120).map((row) => {
    const chance = row.estimate.chance;
    const cut = row.estimate.cut70;
    const my = row.estimate.myScore;
    const diff = row.estimate.diff;
    const audit = auditBadge(row.auditRow);
    const auditDetail = auditDetailBlock(row, audit);
    const scoreLabel = "VBA 환산점수";
    return `<article class="result-card">
      <div class="card-top">
        <div>
          <h3>${row.major}</h3>
          <p>${cleanSchoolName(row.university)} · ${row.group || "-"} · ${row.track || "-"}</p>
        </div>
        <div class="chance-badge ${chanceClass(chance)}">${chanceLabel(chance, row.estimate.chanceLabel)}</div>
      </div>
      <div class="score-line">
        <div class="score-line-top"><span>${scoreLabel} ${my.toFixed(1)}</span><span>컷 기준 ${cut ? cut.toFixed(1) : "-"}</span></div>
        <div class="meter"><span style="width:${Math.max(5, Math.min(100, chance))}%"></span></div>
      </div>
      <div class="meta-grid">
        <div class="meta">차이<strong>${diff >= 0 ? "+" : ""}${diff.toFixed(1)}</strong></div>
        <div class="meta">모집<strong>${row.seats || "-"}</strong></div>
        <div class="meta">경쟁률<strong>${row.competition || "-"}</strong></div>
      </div>
      ${audit}
      ${auditDetail}
      ${row.requirements ? `<div class="requirement">${row.estimate.requirement.eligible ? "응시조건: " : "조건 불일치: "}${row.estimate.requirement.message || row.requirements}</div>` : ""}
    </article>`;
  }).join("");
}

function auditBadge(auditRow) {
  if (!state.audit) return `<div class="audit-badge unknown">검증 데이터 없음</div>`;
  if (!isAuditInputActive()) return `<div class="audit-badge unknown">현재 입력 검증 불가</div>`;
  if (!auditRow) return `<div class="audit-badge unknown">원본 대조 없음</div>`;
  if (auditRow.status === "verified") return `<div class="audit-badge verified">원본 완전 일치</div>`;
  if (auditRow.status === "close") return `<div class="audit-badge close">원본 0.5점 이내</div>`;
  return `<div class="audit-badge mismatch">원본과 불일치</div>`;
}

function auditDetailBlock(row) {
  if (!state.audit || !isAuditInputActive() || !row.auditRow) return "";
  const audit = row.auditRow;
  const cause = auditCause(row);
  return `<div class="audit-detail">
    <span>엑셀 ${safeNumber(audit.excelScore).toFixed(2)}</span>
    <span>웹 엔진 ${safeNumber(audit.webScore).toFixed(2)}</span>
    <span>오차 ${audit.diff > 0 ? "+" : ""}${audit.diff}</span>
    <span>원본판정 ${audit.excelChance || "-"}</span>
    ${audit.status === "mismatch" ? `<strong>${cause}</strong>` : ""}
  </div>`;
}

function auditCause(row) {
  const reasons = [];
  if (String(row.metric || "").includes("변환표준점수")) reasons.push("변환표준점수 환산식");
  if (Object.values(row.weights || {}).some((value) => String(value || "").includes("("))) reasons.push("계열별 괄호 반영비율");
  if (Object.values(row.bonuses || {}).some((value) => value !== null && value !== "" && safeNumber(value, 0) !== 0)) reasons.push("가산점/감점");
  if (row.requirements) reasons.push("응시과목 조건");
  if (!reasons.length) reasons.push("숨김 수식/VBA");
  return `오차 원인 추정: ${reasons.join(", ")}`;
}

function render() {
  if (!state.data) return;
  renderRangeHints();
  renderValidationPanel();
  $("minChanceText").textContent = `${$("minChance").value}% 이상`;
  const rows = getFilteredResults();
  const tierLabel = tiers.find((tier) => tier.id === $("tierFilter").value)?.label || "모든 대학";
  const resultLabel = $("universityFilter").value || tierLabel;
  $("resultTitle").textContent = `${cleanSchoolName(resultLabel)} 기준 ${numberFormat.format(rows.length)}개 학과`;
  renderStats(rows);
  renderCards(rows);
}

function bindEvents() {
  ["tierFilter", "universityFilter", "majorSearch", "minChance", "sortMode", "auditFilter"].forEach((id) => {
    $(id).addEventListener("input", render);
  });
  $("resetButton").addEventListener("click", () => {
    $("tierFilter").value = "all";
    $("universityFilter").value = "";
    $("majorSearch").value = "";
    $("minChance").value = "0";
    $("sortMode").value = "chance";
    $("auditFilter").value = "all";
    render();
  });
  $("chanceTabs").addEventListener("click", (event) => {
    const button = event.target.closest("button");
    if (!button) return;
    [...$("chanceTabs").querySelectorAll("button")].forEach((item) => item.classList.toggle("active", item === button));
    $("minChance").value = button.dataset.min;
    render();
  });
  $("refreshDataButton").addEventListener("click", () => {
    window.location.reload();
  });
}

async function loadData() {
  const response = await fetch(DATA_URL);
  state.data = await response.json();
  try {
    const auditResponse = await fetch(AUDIT_URL);
    state.audit = auditResponse.ok ? await auditResponse.json() : null;
    state.auditRows = new Map((state.audit?.rows || []).map((row) => [row.id, row]));
  } catch {
    state.audit = null;
    state.auditRows = new Map();
  }
  try {
    const integrityResponse = await fetch(INTEGRITY_URL);
    state.integrity = integrityResponse.ok ? await integrityResponse.json() : null;
  } catch {
    state.integrity = null;
  }
  buildGradeIndex();
  state.conversionIndex = state.data.conversionTables || {};
  state.engine = createExcelEngine(state.data);
}

async function init() {
  await loadData();
  updateMockOptions();
  updateFilters();
  buildScoreInputs();
  bindEvents();
  renderSourceStatus();
  render();
}

init().catch((error) => {
  console.error(error);
  $("resultTitle").textContent = "데이터를 불러오지 못했습니다";
  $("results").innerHTML = `<div class="empty">브라우저를 새로고침해 주세요.</div>`;
});

