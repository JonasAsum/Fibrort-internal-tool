"""Print a table of the violations of a quality test module: rule x header label x file.

    .venv/bin/python -m tests.quality.report_family accounts      # or contacts
"""
import collections
import importlib
import sys

from tests.conftest import SAMPLE_DIR


def main(family):
    mod = importlib.import_module(f"tests.quality.test_quality_{family}")
    world = mod.World(SAMPLE_DIR)
    viol = mod.collect_violations(world)
    by = collections.Counter((v["rule"], v["field"], v["file"][:22]) for v in viol)
    for (rule, field, f), n in sorted(by.items()):
        ex = next(v for v in viol if (v["rule"], v["field"], v["file"][:22]) == (rule, field, f))
        print(f"{n:5} {rule:22} {field[:26]:26} {f:22} e.g. {ex['row']} {ex['source'][:36]!r} -> {str(ex['got'])[:36]!r}")
    print("total", len(viol))


if __name__ == "__main__":
    main(sys.argv[1])
