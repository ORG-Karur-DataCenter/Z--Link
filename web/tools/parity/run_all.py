"""Run every parity harness, then verify the harnesses themselves can fail.

This is the command that backs the claim that the browser engine behaves
identically to the Python one. It runs in two phases, and both matter:

  1. Parity — each ported module is compared against its Python original over a
     corpus, field by field, with floats compared as exact IEEE-754 bit
     patterns rather than within a tolerance. The scorer's gates are hard
     thresholds, so a last-bit difference is a behaviour difference.

  2. Mutation — known port mistakes are injected into the JavaScript one at a
     time and each must be caught. Phase 1 passing means nothing on its own: a
     harness with a gap in its corpus reports a clean pass for a port that is
     provably wrong, which is exactly what happened twice while this was being
     built.

Usage:
    python web/tools/parity/run_all.py            # both phases
    python web/tools/parity/run_all.py --quick    # parity only
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]

HARNESSES = [
    ("fuzz", ["fuzz_parity.py", "--cases", "20000"]),
    ("utils", ["utils_parity.py"]),
    ("extractor", ["extractor_parity.py"]),
    ("scorer", ["scorer_parity.py"]),
    ("detached", ["detached_parity.py"]),
    # Not a parity harness: the document preferences and packaging exist only in
    # the browser engine. It runs here because this is the command that gates the
    # deploy, and a document Zotero cannot set up is not a shippable one.
    ("zotero doc", ["zotero_doc_check.py"]),
]


def run(argv: list[str]) -> tuple[int, str]:
    proc = subprocess.run(
        [sys.executable, str(HERE / argv[0]), *argv[1:]],
        capture_output=True, text=True, encoding="utf-8", errors="replace", cwd=ROOT,
    )
    return proc.returncode, (proc.stdout or "") + (proc.stderr or "")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--quick", action="store_true", help="skip the mutation phase")
    args = ap.parse_args()

    print("=" * 72)
    print("phase 1 — parity: JS port vs Python original")
    print("=" * 72)
    failed = []
    for name, argv in HARNESSES:
        code, out = run(argv)
        line = next((l for l in out.splitlines() if l.startswith(("PASS", "FAIL"))), "(no verdict)")
        print(f"  {name:<12} {line}")
        if code != 0:
            failed.append(name)
            print(out)

    if args.quick:
        print(f"\n{'FAILED: ' + ', '.join(failed) if failed else 'all parity harnesses pass'}")
        return 1 if failed else 0

    print()
    print("=" * 72)
    print("phase 2 — mutation: can the harnesses actually fail?")
    print("=" * 72)
    code, out = run(["mutation_check.py"])
    for line in out.splitlines():
        if line.startswith("mutation score") or "SURVIVED" in line or "STALE" in line:
            print("  " + line.strip())
    if code != 0:
        failed.append("mutation")

    print()
    if failed:
        print(f"FAILED: {', '.join(failed)}")
        return 1
    print("All parity harnesses pass, and every injected port mistake was caught.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
