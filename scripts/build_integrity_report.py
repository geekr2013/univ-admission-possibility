import hashlib
import json
import math
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyxlsb import open_workbook


ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIR = ROOT / ".source-data"
DATA = ROOT / "data" / "admissions.json"
AUDIT = ROOT / "data" / "audit.json"
OUT = ROOT / "data" / "integrity.json"

PARSED_SHEETS = {
    "대학별검색",
    "26정시",
    "25정시",
    "24정시",
    "26정시 변환표준점수",
    "25정시 변환표준점수",
    "24정시 변환표준점수",
    "모의고사학생입력성적유효범위",
    "정시 수능등급기준",
    "정시 상세과목명",
}


def safe_number(value, default=0):
    try:
        if value is None or value == "":
            return default
        number = float(value)
        return number if math.isfinite(number) else default
    except (TypeError, ValueError):
        return default


def source_file(data):
    source_name = (data.get("source") or {}).get("file")
    if source_name and (SOURCE_DIR / source_name).exists():
        return SOURCE_DIR / source_name
    files = sorted(SOURCE_DIR.glob("*.xlsb"), key=lambda path: path.stat().st_mtime, reverse=True)
    if files:
        return files[0]
    raise FileNotFoundError("No .xlsb source file found under .source-data")


def sha256(path):
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def workbook_manifest(source):
    sheets = []
    with open_workbook(source) as wb:
        for sheet_name in wb.sheets:
            rows = 0
            nonempty_rows = 0
            max_cols = 0
            with wb.get_sheet(sheet_name) as sheet:
                for row in sheet.rows():
                    rows += 1
                    values = [cell.v for cell in row]
                    max_cols = max(max_cols, len(values))
                    if any(value not in (None, "") for value in values):
                        nonempty_rows += 1
            sheets.append(
                {
                    "name": sheet_name,
                    "rows": rows,
                    "nonemptyRows": nonempty_rows,
                    "maxColumns": max_cols,
                    "parsedByWebApp": sheet_name in PARSED_SHEETS,
                }
            )
    return sheets


def has_vba_project(source):
    with zipfile.ZipFile(source) as archive:
        names = archive.namelist()
    return "xl/vbaProject.bin" in names


def vba_summary(source):
    local_tools = ROOT / ".codex" / "pytools"
    if local_tools.exists():
        sys.path.insert(0, str(local_tools))
    try:
        from oletools.olevba import VBA_Parser
    except Exception as exc:
        return {
            "toolAvailable": False,
            "error": str(exc),
            "note": "oletools가 없어서 VBA 모듈명/함수명까지는 추출하지 못했습니다.",
        }

    modules = []
    definitions = []
    parser = VBA_Parser(str(source))
    try:
        contains = bool(parser.detect_vba_macros())
        for _filename, _stream_path, vba_filename, code in parser.extract_macros():
            lines = code.splitlines()
            found = re.findall(r"(?im)^\s*(?:Public\s+|Private\s+)?(Sub|Function)\s+([^\(\n]+)", code)
            module_defs = [{"kind": kind, "name": name.strip()} for kind, name in found]
            modules.append({"name": vba_filename, "lineCount": len(lines), "procedureCount": len(module_defs)})
            for item in module_defs:
                definitions.append({"module": vba_filename, **item})
    finally:
        parser.close()

    names = [item["name"] for item in definitions]
    core_names = [
        "CalculateAndDisplay",
        "InitializeScoreVariables",
        "FindMockDBRow",
        "ValidateKoreanMathScores",
        "ValidateInquiryScores",
        "ProcessKoreanMathScores",
        "ProcessEnglishScore",
        "ProcessKoreanHistoryScore",
        "ProcessInquiryScores",
        "ProcessBonus",
        "ProcessWeightRatio",
        "ProcessTargetUniversityScore",
        "ProcessUniversityPossibility",
        "DisplayConvertScoreResult",
    ]
    return {
        "toolAvailable": True,
        "containsMacros": contains,
        "moduleCount": len(modules),
        "totalLineCount": sum(item["lineCount"] for item in modules),
        "procedureCount": len(definitions),
        "universityProcedureCount": sum(1 for name in names if name.startswith("Univ_")),
        "coreProceduresFound": [name for name in core_names if name in names],
        "largestModules": sorted(modules, key=lambda item: item["lineCount"], reverse=True)[:5],
        "sampleUniversityProcedures": [name for name in names if name.startswith("Univ_")][:20],
    }


def admissions_by_year(admissions):
    counts = {}
    for row in admissions:
        match = re.search(r"20\d{2}", str(row.get("year") or ""))
        year = match.group(0) if match else str(row.get("year") or "unknown")
        counts[year] = counts.get(year, 0) + 1
    return dict(sorted(counts.items(), reverse=True))


def conversion_summary(tables):
    return {
        year: {
            "columns": len(columns),
            "percentileRowsMax": max((len(values) for values in columns.values()), default=0),
        }
        for year, columns in sorted(tables.items(), reverse=True)
    }


def audit_by_university(audit):
    stats = {}
    for row in audit.get("rows", []):
        university = row.get("university") or "unknown"
        item = stats.setdefault(
            university,
            {"matched": 0, "verified": 0, "close": 0, "mismatch": 0, "maxAbsDiff": 0, "avgAbsDiff": 0},
        )
        status = row.get("status") or "unknown"
        diff = abs(safe_number(row.get("diff")))
        item["matched"] += 1
        if status in item:
            item[status] += 1
        item["maxAbsDiff"] = max(item["maxAbsDiff"], round(diff, 2))
        item["avgAbsDiff"] += diff
    for item in stats.values():
        item["avgAbsDiff"] = round(item["avgAbsDiff"] / item["matched"], 2) if item["matched"] else 0
    return dict(sorted(stats.items(), key=lambda pair: pair[1]["maxAbsDiff"], reverse=True))


def main():
    data = json.loads(DATA.read_text(encoding="utf-8"))
    audit = json.loads(AUDIT.read_text(encoding="utf-8")) if AUDIT.exists() else {}
    source = source_file(data)
    sheets = workbook_manifest(source)
    parsed = [sheet["name"] for sheet in sheets if sheet["parsedByWebApp"]]
    unparsed = [sheet["name"] for sheet in sheets if not sheet["parsedByWebApp"]]
    audit_summary = audit.get("summary") or {}
    matched = safe_number(audit_summary.get("matched"))
    close = safe_number(audit_summary.get("closeWithinPoint5"))
    exact = safe_number(audit_summary.get("verifiedWithinPoint5"))
    mismatch = safe_number(audit_summary.get("mismatch"))

    payload = {
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "source": {
            "repo": (data.get("source") or {}).get("repo"),
            "file": source.name,
            "sizeBytes": source.stat().st_size,
            "sha256": sha256(source),
        },
        "workbook": {
            "sheetCount": len(sheets),
            "hasVbaProject": has_vba_project(source),
            "sheets": sheets,
            "parsedSheets": parsed,
            "unparsedSheets": unparsed,
        },
        "coverage": {
            "admissions": {
                "totalRows": len(data.get("admissions", [])),
                "byYear": admissions_by_year(data.get("admissions", [])),
            },
            "gradeRanges": len(data.get("gradeRanges", [])),
            "csatGradeRanges": len(data.get("csatGradeRanges", [])),
            "mockOptions": data.get("mockOptions", []),
            "subjects": {key: len(value) for key, value in (data.get("subjects") or {}).items()},
            "conversionTables": conversion_summary(data.get("conversionTables") or {}),
            "engineFields": {
                "ratioColumnsPerAdmission": 27,
                "bonusColumnsPerAdmission": 14,
                "englishGrades": 9,
                "historyGrades": 9,
                "secondLanguageGrades": 9,
            },
        },
        "macros": vba_summary(source),
        "audit": {
            "summary": audit_summary,
            "scope": audit.get("scope"),
            "exam": audit.get("exam"),
            "scores": audit.get("scores"),
            "byUniversity": audit_by_university(audit),
        },
        "verdict": {
            "isExcelEquivalent": False,
            "displayPolicy": "모든 카드는 VBA에서 번역한 웹 계산 엔진 결과를 표시하며, 원본 저장 결과가 있는 카드는 일치 여부를 별도로 표시합니다.",
            "reason": (
                "원본의 점수 환산, 가산점, 27개 반영비율, 대학별 계산, 합격가능성 판정을 웹 엔진으로 옮겼습니다. "
                f"최신 원본 저장 결과 {int(matched)}건 중 완전 일치 {int(exact)}건, "
                f"0.5점 이내 {int(close)}건, 불일치 {int(mismatch)}건입니다."
            ),
            "nextAccuracyStep": "최신 원본은 성균관대 결과만 저장되어 있어, 나머지 대학은 별도 Excel 실행 기준값이 추가되기 전까지 구조 검증 완료·독립 실행 검증 대기로 표시합니다.",
        },
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"wrote {OUT}")
    print(payload["verdict"]["reason"])


if __name__ == "__main__":
    main()

