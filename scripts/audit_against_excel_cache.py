import json
import math
from datetime import datetime, timezone
from pathlib import Path

from pyxlsb import open_workbook


ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIR = ROOT / ".source-data"
DATA = ROOT / "data" / "admissions.json"
OUT = ROOT / "data" / "audit.json"


SCORES = {
    "korean": {"subject": "국어", "standard": 127, "percentile": 95, "grade": 2},
    "math": {"subject": "수학", "standard": 135, "percentile": 98, "grade": 1},
    "english": {"subject": "", "standard": None, "percentile": None, "grade": 3},
    "history": {"subject": "", "standard": None, "percentile": None, "grade": 1},
    "explore1": {"subject": "화학Ⅰ", "standard": 64, "percentile": 91, "grade": 2},
    "explore2": {"subject": "물리학Ⅱ", "standard": 71, "percentile": 98, "grade": 1},
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


def normalize(value):
    return str(value or "").replace(" ", "").replace("·", "").replace("ㆍ", "").replace("・", "").lower()


def grade_to_percentile(grade):
    table = [100, 96, 89, 77, 60, 40, 23, 11, 4]
    index = max(1, min(9, int(safe_number(grade, 9)))) - 1
    return table[index]


def is_science(subject):
    return any(token in str(subject or "") for token in ["물리", "화학", "생명과학", "지구과학", "과학탐구"])


def is_social(subject):
    return any(token in str(subject or "") for token in ["생활과 윤리", "윤리와 사상", "한국지리", "세계지리", "동아시아사", "세계사", "경제", "정치와 법", "사회·문화", "사회탐구"])


def ratio_number(value, item, key):
    if isinstance(value, str):
        import re

        match = re.match(r"^\s*([0-9.]+)(?:\(([0-9.]+)\))?", value)
        if not match:
            return 0
        outside = safe_number(match.group(1))
        inside = safe_number(match.group(2), outside) if match.group(2) is not None else outside
        if match.group(2) is None:
            return outside

        math_subject = SCORES["math"]["subject"]
        explores = [SCORES["explore1"]["subject"], SCORES["explore2"]["subject"]]
        science_count = sum(1 for subject in explores if is_science(subject))
        social_count = sum(1 for subject in explores if is_social(subject))
        major_text = normalize(f"{item.get('major')} {item.get('requirements')}")
        natural_major = any(token in major_text for token in ["공학", "자연", "의예", "약학", "수학", "물리", "화학", "생명", "소프트웨어", "컴퓨터", "전자", "전기", "반도체", "데이터", "인공지능", "건축", "통계"])
        natural_choice = "미적분" in math_subject or "기하" in math_subject or science_count > social_count
        if key == "korean":
            return outside if natural_major or natural_choice else inside
        if key == "math":
            return max(outside, inside) if natural_choice else min(outside, inside)
        if key in ("explore1", "explore2"):
            return outside if natural_major or science_count >= social_count else inside
        return outside
    return safe_number(value)


def subject_for_range(key, subject):
    return {"korean": "국어", "math": "수학", "english": "영어", "history": "한국사"}.get(key, subject or "")


def build_grade_index(data):
    index = {}
    for item in data["gradeRanges"]:
        key = (item["year"], item["gradeLevel"], item["month"], normalize(item["subject"]))
        index.setdefault(key, []).append(item)
    for ranges in index.values():
        ranges.sort(key=lambda row: safe_number(row["grade"]))
    return index


def percentile_from_standard(key, score, grade_index):
    standard = safe_number(score.get("standard"), None)
    if standard is None:
        return None
    subject = subject_for_range(key, score.get("subject"))
    ranges = grade_index.get((2026, "고3", "6월", normalize(subject)), [])
    if not ranges:
        return None
    first, last = ranges[0], ranges[-1]
    if standard > safe_number(first["standardMax"]):
        return 99
    if standard < safe_number(last["standardMin"]):
        return 1
    for row in ranges:
        low = safe_number(row["standardMin"])
        high = safe_number(row["standardMax"])
        if low <= standard <= high:
            grade = safe_number(row["grade"])
            position = 0.5 if high == low else (standard - low) / (high - low)
            upper = grade_to_percentile(grade)
            lower_percentile = grade_to_percentile(min(9, grade + 1))
            return lower_percentile + (upper - lower_percentile) * position
    return None


def effective_percentile(key, score, grade_index):
    from_standard = percentile_from_standard(key, score, grade_index)
    if from_standard is not None:
        return from_standard
    percentile = safe_number(score.get("percentile"), None)
    if percentile is not None and percentile > 0:
        return percentile
    return grade_to_percentile(score.get("grade"))


def grade_point(points, grade):
    values = [safe_number(point, None) for point in points or []]
    values = [value for value in values if value is not None]
    if not values:
        return grade_to_percentile(grade)
    index = max(1, min(9, int(safe_number(grade, 9)))) - 1
    value = safe_number((points or [])[index], 0)
    max_value = max(values)
    return (value / max_value) * 100 if max_value > 100 else value


def additive_grade_point(points, grade):
    values = [safe_number(point, None) for point in points or []]
    values = [value for value in values if value is not None]
    if not values:
        return 0
    index = max(1, min(9, int(safe_number(grade, 9)))) - 1
    return safe_number((points or [])[index], 0)


def converted_score(item, percentile, data):
    table = data.get("conversionTables", {}).get("2026", {})
    university = item.get("university") or ""
    type_code = item.get("conversionType")
    explores = [SCORES["explore1"]["subject"], SCORES["explore2"]["subject"]]
    science_count = sum(1 for subject in explores if is_science(subject))
    social_count = sum(1 for subject in explores if is_social(subject))
    candidates = []
    if type_code:
        candidates.append(f"{university}-{type_code}")
    candidates.append(f"{university}-과탐" if science_count >= social_count else f"{university}-사탐")
    candidates.append(university)
    column = None
    normalized = [normalize(candidate) for candidate in candidates]
    for name in table:
        if normalize(name) in normalized:
            column = name
            break
    if column is None:
        for name in table:
            if normalize(name).startswith(normalize(university)):
                column = name
                break
    values = table.get(column or "", {})
    rounded = str(round(max(0, min(100, percentile))))
    return safe_number(values.get(rounded), percentile)


def academic_value(key, item, data, grade_index):
    score = SCORES[key]
    metric = str(item.get("metric") or "")
    if "변환표준점수" in metric:
        return converted_score(item, effective_percentile(key, score, grade_index), data)
    if "표준점수" in metric and "변환" not in metric:
        return safe_number(score.get("standard"), effective_percentile(key, score, grade_index))
    if "등급" in metric:
        return grade_to_percentile(score.get("grade"))
    return effective_percentile(key, score, grade_index)


def weighted_score(item, data, grade_index):
    weights = item.get("weights") or {}
    full_score = safe_number(item.get("fullScore"), 100)
    scale = full_score / 100 if full_score > 100 else 1
    explore_average = (academic_value("explore1", item, data, grade_index) + academic_value("explore2", item, data, grade_index)) / 2
    parts = [
        ("korean", ratio_number(weights.get("korean"), item, "korean"), academic_value("korean", item, data, grade_index)),
        ("math", ratio_number(weights.get("math"), item, "math"), academic_value("math", item, data, grade_index)),
    ]
    explore1_weight = ratio_number(weights.get("explore1"), item, "explore1")
    explore2_weight = ratio_number(weights.get("explore2"), item, "explore2")
    if explore1_weight > 0:
        parts.append(("explore1", explore1_weight, academic_value("explore1", item, data, grade_index)))
    if explore2_weight > 0:
        parts.append(("explore2", explore2_weight, academic_value("explore2", item, data, grade_index) if explore1_weight > 0 else explore_average))
    weighted = sum(value * weight for _, weight, value in parts if weight > 0)
    base = (weighted / 100) * scale if weighted else sum(effective_percentile(key, SCORES[key], grade_index) for key in ["korean", "math", "explore1", "explore2"]) / 4 * scale
    english_weight = ratio_number(weights.get("english"), item, "english")
    history_weight = ratio_number(weights.get("history"), item, "history")
    english = grade_point(item.get("englishPoints"), SCORES["english"]["grade"]) * english_weight / 100 * scale if english_weight > 0 else additive_grade_point(item.get("englishPoints"), SCORES["english"]["grade"])
    history = grade_point(item.get("historyPoints"), SCORES["history"]["grade"]) * history_weight / 100 * scale if history_weight > 0 else additive_grade_point(item.get("historyPoints"), SCORES["history"]["grade"])
    total = base + english + history
    return min(full_score, total)


def excel_cached_rows(source):
    rows = []
    exam = {}
    scores = {}
    with open_workbook(source) as wb:
        with wb.get_sheet("대학별검색") as sheet:
            for index, row in enumerate(sheet.rows(), start=1):
                values = [cell.v for cell in row]
                if index == 3:
                    exam = {"year": values[2], "grade": values[3], "month": values[4]}
                if 5 <= index <= 10:
                    key_map = {5: "korean", 6: "math", 7: "english", 8: "history", 9: "explore1", 10: "explore2"}
                    scores[key_map[index]] = {
                        "label": values[0],
                        "subject": values[1],
                        "standard": values[2],
                        "percentile": values[3],
                        "grade": values[4],
                    }
                if index >= 17 and len(values) > 11 and values[0] and values[10] not in (None, ""):
                    rows.append(
                        {
                            "university": values[0],
                            "group": values[1],
                            "track": values[2],
                            "major": values[3],
                            "excelScore": safe_number(values[10]),
                            "excelChance": values[11],
                        }
                    )
    return exam, scores, rows


def key(row):
    return (row.get("university"), row.get("group"), row.get("track"), row.get("major"))


def row_id(row):
    return "||".join(str(part or "") for part in key(row))


def status_for(diff):
    absolute = abs(diff)
    if absolute <= 0.5:
        return "verified"
    if absolute <= 3:
        return "close"
    return "mismatch"


def main():
    data = json.loads(DATA.read_text(encoding="utf-8"))
    source = source_file(data)
    grade_index = build_grade_index(data)
    admissions = {key(row): row for row in data["admissions"] if "2026" in str(row.get("year"))}
    diffs = []
    missing = []
    exam, scores, cached_rows = excel_cached_rows(source)
    for cached in cached_rows:
        item = admissions.get(key(cached))
        if not item:
            missing.append(cached)
            continue
        web_score = weighted_score(item, data, grade_index)
        diff = round(web_score - cached["excelScore"], 2)
        diffs.append(
            {
                **cached,
                "id": row_id(cached),
                "webScore": round(web_score, 2),
                "diff": diff,
                "status": status_for(diff),
            }
        )
    diffs.sort(key=lambda row: abs(row["diff"]), reverse=True)
    exact = sum(1 for row in diffs if abs(row["diff"]) <= 0.5)
    close = sum(1 for row in diffs if abs(row["diff"]) <= 3)
    payload = {
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "sourceFile": source.name,
        "source": "대학별검색 시트에 저장된 계산 결과",
        "scope": "원본 엑셀 파일을 열었을 때 대학별검색 시트에 캐시되어 있던 한 가지 입력값/검색결과만 대조합니다. Excel 계산 엔진이나 VBA를 새로 실행하지는 않습니다.",
        "statusGuide": {
            "verified": "웹 계산값과 엑셀 저장값 차이가 0.5점 이하",
            "close": "웹 계산값과 엑셀 저장값 차이가 3점 이하",
            "mismatch": "웹 계산값과 엑셀 저장값 차이가 3점을 초과",
            "unknown": "해당 카드가 엑셀 저장 검색결과에 없어 직접 대조 불가",
        },
        "exam": exam,
        "scores": scores,
        "summary": {
            "matched": len(diffs),
            "missing": len(missing),
            "verifiedWithinPoint5": exact,
            "closeWithin3": close,
            "mismatch": len(diffs) - close,
        },
        "rows": diffs,
        "missingRows": missing,
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"matched={len(diffs)} missing={len(missing)} exact_0.5={exact} close_3={close}")
    print(f"wrote {OUT}")
    print("worst differences:")
    for row in diffs[:30]:
        print(json.dumps(row, ensure_ascii=True))


if __name__ == "__main__":
    main()

