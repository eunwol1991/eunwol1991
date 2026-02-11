import os
import re
import sys
import importlib.util
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent

SKIP_DIRS = {".venv", "venv", "__pycache__", ".git", ".idea", ".vscode", "build", "dist"}
SKIP_FILES_RE = re.compile(r"(^test_|_test\.py$)")

def read_text_best_effort(path: Path) -> str | None:
    # Try a few common encodings, then give up.
    for enc in ("utf-8", "utf-8-sig", "cp950", "gbk", "latin-1"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
        except Exception:
            return None
    return None

IMPORT_RE = re.compile(r"^\s*(?:from\s+([a-zA-Z0-9_\.]+)\s+import|import\s+([a-zA-Z0-9_\.]+))", re.M)

def top_level(mod: str) -> str:
    return mod.split(".")[0]

def is_stdlib(name: str) -> bool:
    # Quick heuristic: if it's in stdlib modules list (py3.11 has sys.stdlib_module_names)
    std = getattr(sys, "stdlib_module_names", set())
    return name in std

missing = {}
skipped_files = []

for path in PROJECT_ROOT.rglob("*.py"):
    # skip dirs
    parts = set(path.parts)
    if any(d in parts for d in SKIP_DIRS):
        continue
    if SKIP_FILES_RE.search(path.name):
        continue

    text = read_text_best_effort(path)
    if text is None:
        skipped_files.append(str(path))
        continue

    mods = set()
    for a, b in IMPORT_RE.findall(text):
        m = a or b
        if not m:
            continue
        mods.add(top_level(m))

    for m in sorted(mods):
        if m in {"__future__", "typing"}:
            continue
        if is_stdlib(m):
            continue
        if importlib.util.find_spec(m) is None:
            missing.setdefault(m, []).append(str(path))

print("=== Missing (not found in current environment) ===")
for m in sorted(missing):
    print(f"{m}  (referenced in {len(missing[m])} file(s))")

print("\n=== Files skipped due to unreadable encoding/other issues ===")
for f in skipped_files[:50]:
    print(f)
if len(skipped_files) > 50:
    print(f"... and {len(skipped_files)-50} more")
