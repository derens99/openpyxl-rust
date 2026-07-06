"""Comparison framework: openpyxl vs openpyxl_rust.

Each workload is defined ONCE against the shared openpyxl-compatible API and
executed with both libraries. The framework then:

1. verifies the two outputs are equivalent (cell values, formatting, and
   structure are diffed by loading both files back with real openpyxl), and
2. measures the wall-clock time of each library and reports the speedup.

Usage:
    uv run python benchmarks/compare_framework.py                 # full scale
    uv run python benchmarks/compare_framework.py --scale small   # quick run
    uv run python benchmarks/compare_framework.py --report out.md

The same workloads power `tests/test_compare_framework.py`, which asserts
output equivalence at small scale on every CI run.
"""

import argparse
import importlib
import os
import statistics
import tempfile
import time
from datetime import date, datetime


class LibraryAdapter:
    """Uniform handle to either library's Workbook/styles/load_workbook."""

    def __init__(self, module_name):
        self.name = module_name
        root = importlib.import_module(module_name)
        self.Workbook = root.Workbook
        self.load_workbook = root.load_workbook
        self.styles = importlib.import_module(f"{module_name}.styles")


class Workload:
    """A named workload built once against the shared API.

    build(adapter, path, scale) must create a workbook through
    adapter.Workbook / adapter.styles and save it to path.
    """

    def __init__(self, name, description, build, verify_formatting=True):
        self.name = name
        self.description = description
        self.build = build
        self.verify_formatting = verify_formatting


# ---------------------------------------------------------------------------
# Workload definitions (scale = row multiplier; small=1, full=20)
# ---------------------------------------------------------------------------


def _wl_bulk_values(adapter, path, scale):
    wb = adapter.Workbook()
    ws = wb.active
    ws.title = "Data"
    for r in range(1, 5000 * scale + 1):
        for c in range(1, 9):
            if c % 3 == 0:
                ws.cell(row=r, column=c, value=f"str_{r}_{c}")
            elif c % 3 == 1:
                ws.cell(row=r, column=c, value=r * c * 1.1)
            else:
                ws.cell(row=r, column=c, value=r % 2 == 0)
    wb.save(path)


def _wl_batch_append(adapter, path, scale):
    wb = adapter.Workbook()
    ws = wb.active
    for r in range(5000 * scale):
        ws.append([r, r * 2.5, f"row{r}", r % 2 == 0])
    wb.save(path)


def _wl_formatted_report(adapter, path, scale):
    st = adapter.styles
    wb = adapter.Workbook()
    ws = wb.active
    ws.title = "Report"
    header_font = st.Font(bold=True, size=12, color="FFFFFF")
    header_fill = st.PatternFill(fill_type="solid", start_color="4472C4")
    center = st.Alignment(horizontal="center")
    thin = st.Side(style="thin", color="999999")
    box = st.Border(left=thin, right=thin, top=thin, bottom=thin)
    for c, h in enumerate(["Name", "Revenue", "Cost", "Profit", "Margin"], 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = box
    for r in range(2, 1000 * scale + 2):
        ws.cell(row=r, column=1, value=f"Item {r}")
        for c in range(2, 5):
            cell = ws.cell(row=r, column=c, value=r * c * 10.5)
            cell.number_format = "$#,##0.00"
        cell = ws.cell(row=r, column=5, value=(r % 40) / 100)
        cell.number_format = "0.0%"
    wb.save(path)


def _wl_named_styles(adapter, path, scale):
    st = adapter.styles
    wb = adapter.Workbook()
    ws = wb.active
    highlight = st.NamedStyle(
        name="highlight",
        font=st.Font(bold=True, color="CC0000"),
        fill=st.PatternFill(fill_type="solid", start_color="FFF2CC"),
        number_format="0.00",
    )
    wb.add_named_style(highlight)
    for r in range(1, 1000 * scale + 1):
        ws.cell(row=r, column=1, value=r * 1.5)
        if r % 5 == 0:
            ws.cell(row=r, column=1).style = "highlight"
    wb.save(path)


def _wl_formulas(adapter, path, scale):
    wb = adapter.Workbook()
    ws = wb.active
    for r in range(1, 2000 * scale + 1):
        ws.cell(row=r, column=1, value=r)
        ws.cell(row=r, column=2, value=r * 2)
        ws.cell(row=r, column=3, value=f"=A{r}+B{r}")
    ws.cell(row=1, column=5, value=f"=SUM(A1:A{2000 * scale})")
    wb.save(path)


def _wl_multisheet(adapter, path, scale):
    wb = adapter.Workbook()
    wb.active.title = "First"
    for s in range(4):
        ws = wb.create_sheet(f"Sheet_{s}")
        for r in range(1, 1000 * scale + 1):
            ws.cell(row=r, column=1, value=f"s{s}r{r}")
            ws.cell(row=r, column=2, value=r * (s + 1))
    wb.active["A1"] = "index"
    wb.save(path)


def _wl_datetimes(adapter, path, scale):
    wb = adapter.Workbook()
    ws = wb.active
    for r in range(1, 2000 * scale + 1):
        ws.cell(row=r, column=1, value=datetime(2024, 1 + (r % 12), 1 + (r % 28), r % 24, r % 60))
        ws.cell(row=r, column=2, value=date(2023, 1 + (r % 12), 1 + (r % 28)))
    wb.save(path)


def _wl_structure(adapter, path, scale):
    wb = adapter.Workbook()
    ws = wb.active
    for r in range(1, 200 * scale + 1):
        for c in range(1, 6):
            ws.cell(row=r, column=c, value=f"{r}:{c}")
    ws.merge_cells("A1:E1")
    ws.freeze_panes = "A3"
    ws.column_dimensions["A"].width = 25
    ws.row_dimensions[1].height = 30
    ws2 = wb.create_sheet("Second", index=0)
    ws2["A1"] = "inserted first"
    wb.save(path)


WORKLOADS = [
    Workload("bulk_values", "Mixed-type cell writes (str/float/bool)", _wl_bulk_values, verify_formatting=False),
    Workload("batch_append", "Row-wise append()", _wl_batch_append, verify_formatting=False),
    Workload("formatted_report", "Fonts, fills, borders, number formats", _wl_formatted_report),
    Workload("named_styles", "NamedStyle registration and application", _wl_named_styles),
    Workload("formulas", "Formula cells", _wl_formulas, verify_formatting=False),
    Workload("multisheet", "5 sheets incl. create_sheet(index=...)", _wl_multisheet, verify_formatting=False),
    Workload("datetimes", "datetime and date cells", _wl_datetimes, verify_formatting=False),
    Workload("structure", "Merges, freeze panes, dimensions", _wl_structure),
]


# ---------------------------------------------------------------------------
# Equivalence verification
# ---------------------------------------------------------------------------


def _norm_color(color):
    """Normalize an openpyxl color to its RGB hex (alpha stripped) or None."""
    if color is None:
        return None
    rgb = getattr(color, "rgb", None)
    if not isinstance(rgb, str):
        return None
    return rgb[-6:]


def _cell_diffs(ref, a, b, check_formatting):
    """Compare two openpyxl cells; yield human-readable difference strings."""
    va, vb = a.value, b.value
    both_numeric = (
        isinstance(va, (int, float))
        and isinstance(vb, (int, float))
        and not isinstance(va, bool)
        and not isinstance(vb, bool)
    )
    if both_numeric:
        if abs(float(va) - float(vb)) > 1e-9 * max(1.0, abs(float(va))):
            yield f"{ref}: value {va!r} != {vb!r}"
    elif va != vb:
        yield f"{ref}: value {va!r} != {vb!r}"

    if not check_formatting:
        return
    fa, fb = a.font, b.font
    for attr in ("bold", "italic", "strikethrough"):
        if bool(getattr(fa, attr)) != bool(getattr(fb, attr)):
            yield f"{ref}: font.{attr} {getattr(fa, attr)!r} != {getattr(fb, attr)!r}"
    if (fa.size or 11) != (fb.size or 11):
        yield f"{ref}: font.size {fa.size!r} != {fb.size!r}"
    if _norm_color(fa.color) != _norm_color(fb.color):
        yield f"{ref}: font.color {_norm_color(fa.color)!r} != {_norm_color(fb.color)!r}"
    pa = a.fill.patternType if a.fill else None
    pb = b.fill.patternType if b.fill else None
    if pa != pb:
        yield f"{ref}: fill.patternType {pa!r} != {pb!r}"
    elif pa is not None and _norm_color(a.fill.fgColor) != _norm_color(b.fill.fgColor):
        yield f"{ref}: fill.fgColor differs"
    for side in ("left", "right", "top", "bottom"):
        side_a = getattr(a.border, side, None) if a.border else None
        side_b = getattr(b.border, side, None) if b.border else None
        sa = side_a.style if side_a is not None else None
        sb = side_b.style if side_b is not None else None
        if sa != sb:
            yield f"{ref}: border.{side} {sa!r} != {sb!r}"
    if (a.alignment.horizontal, a.alignment.vertical) != (b.alignment.horizontal, b.alignment.vertical):
        yield f"{ref}: alignment differs"
    if not isinstance(va, (datetime, date)) and a.number_format != b.number_format:
        yield f"{ref}: number_format {a.number_format!r} != {b.number_format!r}"


def verify_equivalence(path_a, path_b, check_formatting=True, max_diffs=20):
    """Diff two xlsx files by loading both with real openpyxl.

    Returns a list of difference strings (empty = equivalent).
    """
    import openpyxl

    wa = openpyxl.load_workbook(path_a)
    wb = openpyxl.load_workbook(path_b)
    diffs = []

    if wa.sheetnames != wb.sheetnames:
        return [f"sheetnames {wa.sheetnames} != {wb.sheetnames}"]

    for sheet_name in wa.sheetnames:
        sa, sb = wa[sheet_name], wb[sheet_name]
        max_row = max(sa.max_row, sb.max_row)
        max_col = max(sa.max_column, sb.max_column)
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                ca, cb = sa.cell(row=r, column=c), sb.cell(row=r, column=c)
                ref = f"{sheet_name}!{ca.coordinate}"
                for diff in _cell_diffs(ref, ca, cb, check_formatting):
                    diffs.append(diff)
                    if len(diffs) >= max_diffs:
                        return diffs
        merges_a = sorted(str(m) for m in sa.merged_cells.ranges)
        merges_b = sorted(str(m) for m in sb.merged_cells.ranges)
        if merges_a != merges_b:
            diffs.append(f"{sheet_name}: merged ranges {merges_a} != {merges_b}")
        if sa.freeze_panes != sb.freeze_panes:
            diffs.append(f"{sheet_name}: freeze_panes {sa.freeze_panes!r} != {sb.freeze_panes!r}")
    return diffs


# ---------------------------------------------------------------------------
# Timing and reporting
# ---------------------------------------------------------------------------


def _time_workload(workload, adapter, scale, repeats):
    """Run a workload `repeats` times; return (median_seconds, last_output_path)."""
    times = []
    path = None
    for _ in range(repeats):
        if path is not None:
            os.unlink(path)
        fd, path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        start = time.perf_counter()
        workload.build(adapter, path, scale)
        times.append(time.perf_counter() - start)
    return statistics.median(times), path


def run_comparison(workloads=None, scale=1, repeats=3, verify=True):
    """Run all workloads with both libraries. Returns a list of result dicts."""
    rust = LibraryAdapter("openpyxl_rust")
    ref = LibraryAdapter("openpyxl")
    results = []
    for workload in workloads or WORKLOADS:
        ref_time, ref_path = _time_workload(workload, ref, scale, repeats)
        rust_time, rust_path = _time_workload(workload, rust, scale, repeats)
        diffs = []
        if verify:
            diffs = verify_equivalence(ref_path, rust_path, workload.verify_formatting)
        os.unlink(ref_path)
        os.unlink(rust_path)
        results.append(
            {
                "workload": workload.name,
                "description": workload.description,
                "openpyxl_s": ref_time,
                "openpyxl_rust_s": rust_time,
                "speedup": ref_time / rust_time if rust_time > 0 else float("inf"),
                "equivalent": not diffs,
                "diffs": diffs,
            }
        )
    return results


def format_report(results, scale):
    lines = [
        "# openpyxl vs openpyxl_rust — comparison report",
        "",
        f"Scale factor: {scale}. Times are medians; speedup = openpyxl / openpyxl_rust.",
        "",
        "| Workload | Description | openpyxl | openpyxl_rust | Speedup | Output equivalent |",
        "|---|---|---:|---:|---:|:---:|",
    ]
    for r in results:
        eq = "yes" if r["equivalent"] else "NO"
        lines.append(
            f"| {r['workload']} | {r['description']} | {r['openpyxl_s']:.3f}s "
            f"| {r['openpyxl_rust_s']:.3f}s | {r['speedup']:.2f}x | {eq} |"
        )
    speedups = [r["speedup"] for r in results]
    lines += ["", f"**Geometric-mean speedup: {statistics.geometric_mean(speedups):.2f}x**"]
    failed = [r for r in results if not r["equivalent"]]
    if failed:
        lines += ["", "## Differences"]
        for r in failed:
            lines.append(f"### {r['workload']}")
            lines += [f"- {d}" for d in r["diffs"]]
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--scale", choices=["small", "full"], default="full")
    parser.add_argument("--repeats", type=int, default=3)
    parser.add_argument("--report", help="write markdown report to this path")
    parser.add_argument("--no-verify", action="store_true", help="skip equivalence checks")
    args = parser.parse_args()

    scale = 1 if args.scale == "small" else 20
    results = run_comparison(scale=scale, repeats=args.repeats, verify=not args.no_verify)
    report = format_report(results, scale)
    print(report)
    if args.report:
        with open(args.report, "w") as f:
            f.write(report + "\n")
        print(f"\nReport written to {args.report}")


if __name__ == "__main__":
    main()
