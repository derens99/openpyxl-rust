# tests/test_compare_framework.py
"""Runs every comparison-framework workload with both libraries and asserts
the outputs are equivalent. Performance numbers are recorded (not asserted,
apart from a generous regression backstop) since CI timing is noisy.

Full-scale performance report: uv run python benchmarks/compare_framework.py
"""

import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "benchmarks"))

from compare_framework import (
    WORKLOADS,
    LibraryAdapter,
    run_comparison,
    verify_equivalence,
)


@pytest.mark.parametrize("workload", WORKLOADS, ids=lambda w: w.name)
def test_output_equivalent_to_openpyxl(workload, tmp_path):
    """Each workload must produce output equivalent to real openpyxl's."""
    ref_path = str(tmp_path / "ref.xlsx")
    rust_path = str(tmp_path / "rust.xlsx")
    workload.build(LibraryAdapter("openpyxl"), ref_path, scale=1)
    workload.build(LibraryAdapter("openpyxl_rust"), rust_path, scale=1)
    diffs = verify_equivalence(ref_path, rust_path, workload.verify_formatting)
    assert not diffs, "Outputs differ from openpyxl:\n" + "\n".join(diffs)


def test_performance_comparison_runs(capsys):
    """Smoke-run the timed comparison at small scale and report speedups.

    Asserts only a generous backstop: openpyxl_rust must not be dramatically
    slower than openpyxl on the bulk-write workload.
    """
    bulk = [w for w in WORKLOADS if w.name == "bulk_values"]
    results = run_comparison(workloads=bulk, scale=1, repeats=1, verify=False)
    r = results[0]
    with capsys.disabled():
        print(
            f"\n[perf] {r['workload']}: openpyxl {r['openpyxl_s']:.3f}s, "
            f"openpyxl_rust {r['openpyxl_rust_s']:.3f}s ({r['speedup']:.2f}x)"
        )
    assert r["openpyxl_rust_s"] < r["openpyxl_s"] * 2, (
        "openpyxl_rust is more than 2x slower than openpyxl on bulk writes - "
        "this indicates a severe performance regression"
    )
