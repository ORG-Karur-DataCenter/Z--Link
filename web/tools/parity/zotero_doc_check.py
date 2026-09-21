"""Check the non-citation half of a Zotero document: preferences and packaging.

The citation fields have their own harnesses. This one covers what surrounds
them — the ZOTERO_PREF custom properties that tell Zotero which style to render
in, and the content-type and relationship entries that make those properties
readable at all. None of it is visible in the manuscript, and all of it only
matters on a machine that is not the author's.

Run:  python web/tools/parity/zotero_doc_check.py
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

RUNNER = Path(__file__).with_name("zotero_doc_runner.mjs")


def main() -> int:
    proc = subprocess.run(
        ["node", str(RUNNER)], capture_output=True, text=True, encoding="utf-8",
    )
    if proc.returncode != 0:
        sys.stderr.write(proc.stderr)
        print("FAIL  zotero doc: the runner did not complete")
        return 1

    checks = json.loads(proc.stdout)
    failed = [(name, detail) for name, ok, detail in checks if not ok]
    if not failed:
        print(f"PASS  zotero doc: {len(checks)} checks, preferences and packaging intact")
        return 0

    print(f"FAIL  zotero doc: {len(failed)}/{len(checks)} checks")
    for name, detail in failed:
        print(f"  {name}" + (f"\n      {detail}" if detail else ""))
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
