"""Entry point for the excel reader. Modes:

  parse    - read the workbook(s) and print per-family counts. Writes nothing.
  dry-run  - parse, then write excel/out/report.json (counts, per file) and
             excel/out/review.json (the items to check, with cell values; the
             folder is gitignored).
  commit   - TODO: not wired yet. Will run visits.match and visits.pipeline.

Workbooks are only read. Usage (from importer/):
  .venv/bin/python -m excel.cli --mode dry-run "../mock_visits/excel/*.xlsx"
"""
import argparse
import dataclasses
import glob
import json
import os
import sys

from .run import read_paths, summarize

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")


def _expand(patterns):
    paths = []
    for pattern in patterns:
        matches = sorted(glob.glob(os.path.expanduser(pattern)))
        if not matches:
            print(f"no file matches {pattern}", file=sys.stderr)
        paths += matches
    return [p for p in paths if os.path.isfile(p) and not os.path.basename(p).startswith("~$")]


def _print(summary, result):
    cols = ("files", "sheets", "read_sheets", "visits", "auto_accepted", "queued", "skipped_files")
    print(f"{'family':<12}" + "".join(f"{c:>15}" for c in cols))
    for name, f in sorted(summary.items()):
        print(f"{name:<12}" + "".join(f"{f[c]:>15}" for c in cols))
        if f["queued_by_reason"]:
            reasons = ", ".join(f"{k} {v}" for k, v in sorted(f["queued_by_reason"].items()))
            print(f"{'':<12}queued by reason: {reasons}")
        for err in f["errors"]:
            print(f"{'':<12}file with error: {err}")
    stale = sum(1 for s in result.per_file_stats if s.stale_copy)
    print(f"visits total {len(result.parsed.visits)}, weeks {len(result.parsed.weeks)}, "
          f"review items {len(result.review)}, stale copies {stale}")


def _write(summary, result):
    os.makedirs(OUT_DIR, exist_ok=True)
    report = {
        "families": summary,
        "files": [{"file": os.path.basename(s.path), "family": s.family, "sheets": s.sheets,
                   "read_sheets": s.read_sheets, "visits": s.visits, "stale_copy": s.stale_copy,
                   "error": s.error} for s in result.per_file_stats],
        "skipped": [{"file": os.path.basename(x["file"]), "error": x["error"]} for x in result.skipped_files],
    }
    with open(os.path.join(OUT_DIR, "report.json"), "w") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)
    with open(os.path.join(OUT_DIR, "review.json"), "w") as f:
        json.dump([dataclasses.asdict(item) for item in result.review], f, indent=2, ensure_ascii=False)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("excel", nargs="+", help="workbook path(s) or glob(s)")
    ap.add_argument("--mode", choices=["parse", "dry-run", "commit"], default="parse")
    args = ap.parse_args(argv)
    if args.mode == "commit":
        raise NotImplementedError("commit is a TODO: wire visits.match and visits.pipeline")
    result = read_paths(_expand(args.excel))
    summary = summarize(result)
    _print(summary, result)
    if args.mode == "dry-run":
        _write(summary, result)
        print(f"wrote {OUT_DIR}/report.json and review.json")


if __name__ == "__main__":
    main()
