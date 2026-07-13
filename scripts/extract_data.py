import json
import hashlib
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote, urlsplit, urlunsplit
from urllib.request import urlopen, urlretrieve

from pyxlsb import open_workbook


ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIR = ROOT / ".source-data"
DEFAULT_SOURCE_FILENAME = "27학년도 정시 지원가능 대학 및 학과 검색.xlsb"
REPO_CONTENTS_API = "https://api.github.com/repos/YMABI/SearchUnivMajorPossibility/contents"
DEFAULT_SOURCE_URL = f"https://raw.githubusercontent.com/YMABI/SearchUnivMajorPossibility/main/{quote(DEFAULT_SOURCE_FILENAME)}"
OUT_DIR = ROOT / "data"
OUT_FILE = OUT_DIR / "admissions.json"


YEAR_SHEETS = ["26정시", "25정시", "24정시"]
CONVERSION_SHEETS = ["26정시 변환표준점수", "25정시 변환표준점수", "24정시 변환표준점수"]
MOCK_SHEET = "모의고사 시행 목록"
SUBJECT_SHEET = "정시 상세과목명"
GRADE_SHEET = "모의고사학생입력성적유효범위"
CSAT_GRADE_SHEET = "정시 수능등급기준"


def local_source_files():
    if not SOURCE_DIR.exists():
        return []
    return sorted(SOURCE_DIR.glob("*.xlsb"), key=lambda path: path.stat().st_mtime, reverse=True)


def latest_remote_source():
    with urlopen(REPO_CONTENTS_API, timeout=30) as response:
        contents = json.loads(response.read().decode("utf-8"))
    files = [item for item in contents if item.get("type") == "file" and item.get("name", "").lower().endswith(".xlsb")]
    if not files:
        raise RuntimeError("No .xlsb source file found in YMABI repository")
    files.sort(key=lambda item: item.get("name", ""), reverse=True)
    selected = files[0]
    return selected["name"], selected.get("download_url") or f"https://raw.githubusercontent.com/YMABI/SearchUnivMajorPossibility/main/{quote(selected['name'])}"


def safe_url(url):
    parts = urlsplit(url)
    return urlunsplit((parts.scheme, parts.netloc, quote(parts.path), parts.query, parts.fragment))


def ensure_source():
    SOURCE_DIR.mkdir(exist_ok=True)
    try:
        filename, url = latest_remote_source()
    except Exception as exc:
        print(f"remote source discovery failed: {exc}")
        filename, url = DEFAULT_SOURCE_FILENAME, DEFAULT_SOURCE_URL

    target = SOURCE_DIR / filename
    try:
        download_url = safe_url(url)
        print(f"downloading {download_url}")
        urlretrieve(download_url, target)
        return target
    except Exception as exc:
        print(f"source download failed: {exc}")
        local_files = local_source_files()
        if local_files:
            print(f"using local source {local_files[0]}")
            return local_files[0]
        raise


def clean(value):
    if value is None:
        return None
    if isinstance(value, str):
        value = value.strip()
        return value if value else None
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return round(value, 4)
    return value


def row_values(row):
    return [clean(cell.v) for cell in row]


def parse_admissions(wb, sheet_name):
    with wb.get_sheet(sheet_name) as sheet:
        rows = sheet.rows()
        header = row_values(next(rows))
        items = []
        for row in rows:
            values = row_values(row)
            if not values or not values[1] or not values[4]:
                continue
            item = dict(zip(header, values))
            year = str(item.get("구분") or sheet_name)
            prefix = "".join(ch for ch in sheet_name if ch.isdigit())[:2]
            items.append(
                {
                    "year": year,
                    "university": item.get("대학교명"),
                    "group": item.get("지원군"),
                    "track": item.get("모집전형"),
                    "major": item.get("모집단위"),
                    "seats": item.get("모집인원"),
                    "competition": item.get("경쟁률"),
                    "additional": item.get("충원합격"),
                    "fullScore": item.get("만점"),
                    "cut50": item.get(f"{prefix}수능(50% cut)"),
                    "cut70": item.get(f"{prefix}수능(70% cut)"),
                    "cutPercentile50": item.get(f"{prefix}수능백분위(50%cut)"),
                    "cutPercentile70": item.get(f"{prefix}수능백분위(70%cut)"),
                    "metric": item.get("백분위/표준점수/변환표준점수/등급"),
                    "conversionType": item.get("변환점수적용구분"),
                    "weights": {
                        "korean": item.get("국어"),
                        "math": item.get("수학"),
                        "english": item.get("영어"),
                        "explore1": item.get("탐1"),
                        "explore2": item.get("탐2"),
                        "history": item.get("한국사"),
                    },
                    "ratioColumns": values[30:57],
                    "englishPoints": [item.get(f"영{i}") for i in range(1, 10)],
                    "historyPoints": [item.get(f"한{i}") for i in range(1, 10)],
                    "secondLanguagePoints": [item.get(f"제2한{i}") for i in range(1, 10)],
                    "bonuses": {
                        "probabilityStats": item.get(f"{prefix}확률과통계"),
                        "calculus": item.get(f"{prefix}미적분"),
                        "geometry": item.get(f"{prefix}기하"),
                        "science": item.get("과탐가산"),
                        "social": item.get("사탐가산"),
                    },
                    "bonusColumns": values[89:103],
                    "requirements": item.get("응시과목지정"),
                }
            )
        return items


def parse_mock_options(grade_ranges):
    by_year_grade = {}
    for item in grade_ranges:
        year = item.get("year")
        grade = item.get("gradeLevel")
        month = item.get("month")
        if not isinstance(year, (int, float)) or not str(grade).startswith("고") or not month:
            continue
        by_year_grade.setdefault((year, grade), [])
        if month not in by_year_grade[(year, grade)]:
            by_year_grade[(year, grade)].append(month)

    def month_key(month):
        digits = "".join(ch for ch in str(month) if ch.isdigit())
        return int(digits or 99)

    def year_key(year):
        digits = "".join(ch for ch in str(year) if ch.isdigit())
        return int(digits[:4] or 0)

    options = []
    for (year, grade), months in sorted(by_year_grade.items(), key=lambda item: (-year_key(item[0][0]), item[0][1])):
        options.append({"year": year, "grade": grade, "months": sorted(months, key=month_key)})
    return options


def parse_subjects(wb):
    subjects = {"korean": set(), "math": set(), "explore": set(), "language2": set()}
    with wb.get_sheet(SUBJECT_SHEET) as sheet:
        for index, row in enumerate(sheet.rows()):
            values = row_values(row)
            if index == 0:
                continue
            if len(values) > 1 and values[1]:
                subjects["korean"].add(values[1])
            if len(values) > 2 and values[2]:
                subjects["math"].add(values[2])
            if len(values) > 3 and values[3]:
                subjects["explore"].add(values[3])
            if len(values) > 4 and values[4]:
                subjects["language2"].add(values[4])
    with wb.get_sheet(GRADE_SHEET) as sheet:
        for index, row in enumerate(sheet.rows()):
            values = row_values(row)
            if index == 0 or len(values) < 4:
                continue
            subject = values[3]
            if not subject:
                continue
            if subject in ("국어", "수학", "영어", "한국사"):
                continue
            subjects["explore"].add(subject)
    ignored = {"과목", "과목명", "상세과목명"}
    return {key: sorted(item for item in value if item not in ignored) for key, value in subjects.items()}


def parse_grade_ranges(wb):
    ranges = []
    with wb.get_sheet(GRADE_SHEET) as sheet:
        header = row_values(next(sheet.rows()))
        for row in sheet.rows():
            item = dict(zip(header, row_values(row)))
            if not item.get("년도"):
                continue
            ranges.append(
                {
                    "year": item.get("년도"),
                    "gradeLevel": item.get("학년"),
                    "month": item.get("회차/월"),
                    "subject": item.get("과목"),
                    "grade": item.get("등급"),
                    "standardMin": item.get("표준점수 (최저)"),
                    "standardMax": item.get("표준점수 (최고)"),
                }
            )
    return ranges


def parse_csat_grade_ranges(wb):
    ranges = []
    with wb.get_sheet(CSAT_GRADE_SHEET) as sheet:
        rows = sheet.rows()
        header = row_values(next(rows))
        for row in rows:
            item = dict(zip(header, row_values(row)))
            if not item.get("구분") or not item.get("과목") or not item.get("등급"):
                continue
            ranges.append(
                {
                    "examYear": item.get("구분"),
                    "subject": item.get("과목"),
                    "grade": item.get("등급"),
                    "standardMin": item.get("표준점수 (최저)"),
                    "standardMax": item.get("표준점수 (최고)"),
                }
            )
    return ranges


def parse_conversion_tables(wb):
    tables = {}
    for sheet_name in CONVERSION_SHEETS:
        if sheet_name not in wb.sheets:
            continue
        year = "".join(ch for ch in sheet_name if ch.isdigit())[:2]
        year_key = f"20{year}"
        tables[year_key] = {}
        with wb.get_sheet(sheet_name) as sheet:
            rows = sheet.rows()
            header = row_values(next(rows))
            for row in rows:
                values = row_values(row)
                if len(values) < 3 or values[2] is None:
                    continue
                percentile = str(values[2])
                for index, column in enumerate(header[3:], start=3):
                    if not column or index >= len(values) or values[index] is None:
                        continue
                    tables[year_key].setdefault(column, {})[percentile] = values[index]
    return tables


def main():
    OUT_DIR.mkdir(exist_ok=True)
    source = ensure_source()
    with open_workbook(source) as wb:
        admissions = []
        for sheet_name in YEAR_SHEETS:
            admissions.extend(parse_admissions(wb, sheet_name))
        grade_ranges = parse_grade_ranges(wb)
        payload = {
            "source": {
                "repo": "https://github.com/YMABI/SearchUnivMajorPossibility",
                "file": source.name,
                "generatedAt": datetime.now(timezone.utc).isoformat(),
                "sizeBytes": source.stat().st_size,
                "sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
                "note": "원본 xlsb의 정시 모집 데이터와 성적 선택지를 웹앱용 JSON으로 추출했습니다.",
            },
            "admissions": admissions,
            "mockOptions": parse_mock_options(grade_ranges),
            "subjects": parse_subjects(wb),
            "gradeRanges": grade_ranges,
            "csatGradeRanges": parse_csat_grade_ranges(wb),
            "conversionTables": parse_conversion_tables(wb),
        }
    OUT_FILE.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"wrote {OUT_FILE} with {len(admissions)} admission rows")


if __name__ == "__main__":
    main()

