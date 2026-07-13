import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIR = ROOT / ".source-data"
OUT_DIR = ROOT / ".analysis" / "vba"
MANIFEST = OUT_DIR / "manifest.json"


def load_vba_parser():
    try:
        from oletools.olevba import VBA_Parser

        return VBA_Parser
    except ImportError:
        local_tools = ROOT / ".codex" / "pytools"
        if local_tools.exists():
            sys.path.insert(0, str(local_tools))
        from oletools.olevba import VBA_Parser

        return VBA_Parser


def source_file():
    files = sorted(SOURCE_DIR.glob("*.xlsb"), key=lambda path: path.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError("No .xlsb source file found under .source-data")
    return files[0]


def safe_filename(name):
    cleaned = re.sub(r"[\\/:*?\"<>|]", "_", name).strip()
    return cleaned or "module.bas"


def procedure_names(code):
    pattern = re.compile(
        r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z0-9_\u3131-\uD79D]+)",
        re.IGNORECASE | re.MULTILINE,
    )
    return pattern.findall(code)


def main():
    parser_class = load_vba_parser()
    source = source_file()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for old_file in OUT_DIR.glob("*.bas"):
        old_file.unlink()
    for old_file in OUT_DIR.glob("*.cls"):
        old_file.unlink()

    modules = []
    parser = parser_class(str(source))
    try:
        if not parser.detect_vba_macros():
            raise RuntimeError("No VBA project found in source workbook")
        for _, stream_path, vba_filename, vba_code in parser.extract_macros():
            module_name = safe_filename(vba_filename or Path(stream_path).name)
            target = OUT_DIR / module_name
            vba_code = vba_code.replace("\r\r\n", "\n").replace("\r\n", "\n").replace("\r", "\n")
            target.write_text(vba_code, encoding="utf-8")
            procedures = procedure_names(vba_code)
            modules.append(
                {
                    "name": module_name,
                    "stream": stream_path,
                    "lineCount": len(vba_code.splitlines()),
                    "procedures": procedures,
                }
            )
    finally:
        parser.close()

    payload = {
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "sourceFile": source.name,
        "modules": sorted(modules, key=lambda item: item["name"]),
        "summary": {
            "moduleCount": len(modules),
            "lineCount": sum(item["lineCount"] for item in modules),
            "procedureCount": sum(len(item["procedures"]) for item in modules),
        },
    }
    MANIFEST.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload["summary"], ensure_ascii=False))


if __name__ == "__main__":
    main()

