import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createExcelEngine } from "../engine.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const dataPath = path.join(root, "data", "admissions.json");
const auditPath = path.join(root, "data", "audit.json");
const data = JSON.parse(fs.readFileSync(dataPath, "utf8"));
const cached = JSON.parse(fs.readFileSync(auditPath, "utf8"));
const engine = createExcelEngine(data);

const key = (row) => [row.university, row.group, row.track, row.major].map((value) => String(value || "")).join("||");
const admissions = new Map(data.admissions.filter((row) => String(row.year).includes("2026")).map((row) => [key(row), row]));
const inputs = {
  exam: cached.exam,
  korean: cached.scores.korean,
  math: cached.scores.math,
  english: cached.scores.english,
  history: cached.scores.history,
  explore1: cached.scores.explore1,
  explore2: cached.scores.explore2,
  language2: cached.scores.language2 || { subject: "", grade: 9 },
};

const rows = [];
const missingRows = [];
for (const sourceRow of cached.rows || []) {
  const item = admissions.get(sourceRow.id || key(sourceRow));
  if (!item) {
    missingRows.push(sourceRow);
    continue;
  }
  const result = engine.calculate(item, inputs);
  const diff = Math.round((result.myScore - Number(sourceRow.excelScore)) * 1000) / 1000;
  const exact = Math.abs(diff) <= 0.005 && result.chanceLabel === sourceRow.excelChance;
  rows.push({
    ...sourceRow,
    webScore: result.myScore,
    webChance: result.chanceLabel,
    diff,
    status: exact ? "verified" : Math.abs(diff) <= 0.5 ? "close" : "mismatch",
  });
}

rows.sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));
const exact = rows.filter((row) => row.status === "verified").length;
const close = rows.filter((row) => row.status === "close").length;
const mismatch = rows.filter((row) => row.status === "mismatch").length;
const payload = {
  ...cached,
  generatedAt: new Date().toISOString(),
  sourceSha256: data.source.sha256,
  scope: "원본 XLSB에 저장된 최신 대학별검색 결과와 웹 VBA 변환 엔진을 점수 0.005점 및 합격가능성 문구 기준으로 대조합니다.",
  statusGuide: {
    verified: "환산점수 오차 0.005점 이하이며 합격가능성 문구도 동일",
    close: "환산점수 오차 0.5점 이하",
    mismatch: "환산점수 오차 0.5점 초과 또는 합격가능성 문구 불일치",
    unknown: "원본 저장 결과가 없어 직접 대조 불가",
  },
  summary: {
    matched: rows.length,
    missing: missingRows.length,
    exact,
    verifiedWithinPoint5: exact,
    closeWithinPoint5: close,
    mismatch,
  },
  rows,
  missingRows,
};

fs.writeFileSync(auditPath, JSON.stringify(payload), "utf8");
console.log(`engine audit: exact=${exact}, close=${close}, mismatch=${mismatch}, missing=${missingRows.length}`);

