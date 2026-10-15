"""Extract reviewable VBA source without opening Excel or running macros."""
import argparse
import hashlib
import json
import sys
from pathlib import Path
from oletools.olevba import VBA_Parser

def snapshot(workbook):
    files = {}
    components = []
    parser = VBA_Parser(str(workbook))
    try:
        for _, stream, name, code in parser.extract_macros():
            name = Path(name).name
            suffix = Path(name).suffix.lower()
            group = {".bas": "modules", ".frm": "forms", ".cls": "host"}.get(suffix)
            if group is None:
                raise ValueError("Unexpected VBA component type: " + name)
            # Extracted forms lack the designer and .frx resources.
            output_name = name if suffix == ".bas" else name + ".txt"
            relative = group + "/" + output_name
            if relative in files:
                raise ValueError("Duplicate component path: " + relative)
            text = code.replace("\r\n", "\n").replace("\r", "\n").rstrip() + "\n"
            files[relative] = text
            components.append({"component": name, "stream": stream, "path": relative,
                               "sha256": hashlib.sha256(text.encode("utf-8")).hexdigest()})
    finally:
        parser.close()
    if not components:
        raise ValueError("No VBA components were extracted")
    manifest = {"workbook": workbook.name,
                "workbook_sha256": hashlib.sha256(workbook.read_bytes()).hexdigest(),
                "format": "VBA source text; forms and host classes are not full designer exports",
                "components": sorted(components, key=lambda item: item["path"])}
    files["manifest.json"] = json.dumps(manifest, indent=2) + "\n"
    return files

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("workbook", type=Path)
    ap.add_argument("--output", type=Path, default=Path("src"))
    ap.add_argument("--check", action="store_true", help="Compare snapshots without writing")
    args = ap.parse_args()
    files = snapshot(args.workbook)
    expected = set(files)
    actual = set()
    for folder in ("modules", "forms", "host"):
        base = args.output / folder
        if base.exists():
            actual.update(p.relative_to(args.output).as_posix() for p in base.rglob("*") if p.is_file())
    if (args.output / "manifest.json").exists():
        actual.add("manifest.json")
    if args.check:
        differences = sorted(expected ^ actual)
        for relative, content in files.items():
            path = args.output / relative
            if path.exists() and path.read_bytes() != content.encode("utf-8"):
                differences.append(relative + " (content differs)")
        if differences:
            print("Snapshot differs:\n" + "\n".join(differences), file=sys.stderr)
            return 1
        print("Verified " + str(len(files) - 1) + " VBA components and workbook hash")
        return 0
    stale = actual - expected
    if stale:
        raise ValueError("Unexpected existing files; review before exporting: " + ", ".join(sorted(stale)))
    for relative, content in files.items():
        path = args.output / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(content.encode("utf-8"))
    print("Extracted " + str(len(files) - 1) + " VBA components")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
