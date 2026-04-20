#!/usr/bin/env python3
"""
commission_annotator.py  –  Commission PDF Annotation Tool
===========================================================
Automates the manual annotation process for commission calculations.

INPUTS
  --estimate   Estimate Details PDF  (has Unit Cost / Ext Cost columns)
  --commission Commission Sheet PDF  (has Commission Estimate Details table)
  --ff         Finance Fee amount    (optional, default 0)
  --ff-name    Finance Fee lender    (optional, e.g. "Lendhi")
  --output-dir Output directory      (optional, default = same folder as inputs)

OUTPUTS
  Two annotated PDFs saved to --output-dir:
  · Estimate Details  – contractor Ext Cost highlighted + summary block at bottom
  · Commission Sheet  – top-right note + corrected Greenline / % GL / Est Commission

USAGE EXAMPLE
  python3 commission_annotator.py \\
      --estimate  "Valenzuela Estimate Details.pdf" \\
      --commission "Valenzuela Commission Sheet.pdf" \\
      --ff 1799.13 --ff-name "Lendhi" \\
      --output-dir "./output"
"""

from __future__ import annotations
import fitz          # pip install pymupdf
import re, sys, argparse
from pathlib import Path

# ── Colour palette ─────────────────────────────────────────────────────────────
RED    = (0.80, 0.00, 0.00)   # deep red  – all overlay text / lines
YELLOW = (1.00, 1.00, 0.00)   # yellow    – contractor Ext Cost highlight

CONTRACTOR_KEYWORDS  = ["outside partner", "outsource", "outsourced", "contractor task", "contractor tasks"]
DESIGN_FEE_KEYWORDS  = ["bath design fee", "design fee"]

# ── Formatting helpers ──────────────────────────────────────────────────────────
def fmt_usd(v: float) -> str:  return f"${v:,.2f}"
def fmt_pct(v: float) -> str:  return f"{round(v * 100)}%"
def parse_usd(s: str)-> float: return float(s.replace("$","").replace(",","").strip())
def parse_pct(s: str)-> float: return float(s.replace("%","").strip()) / 100.0

# ── Span utilities ──────────────────────────────────────────────────────────────
def get_spans(page) -> list[dict]:
    """Return every non-empty text span on the page with x0,y0,x1,y1."""
    result = []
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                t = span["text"].strip()
                if t:
                    b = span["bbox"]
                    result.append(dict(text=t, x0=b[0], y0=b[1], x1=b[2], y1=b[3]))
    return result

def spans_on_row(spans: list[dict], y_ref: float, tol: float = 4) -> list[dict]:
    """Return spans whose y0 is within ±tol of y_ref."""
    return [s for s in spans if abs(s["y0"] - y_ref) <= tol]

def find_label(spans: list[dict], label: str) -> dict | None:
    """Case-insensitive exact label match; returns first matching span or None."""
    for s in spans:
        if s["text"].strip().lower() == label.lower():
            return s
    return None

def find_label_startswith(spans: list[dict], prefix: str) -> dict | None:
    for s in spans:
        if s["text"].strip().lower().startswith(prefix.lower()):
            return s
    return None

def rightmost_number_on_row(spans: list[dict], y_ref: float, min_val=10, tol=4) -> dict | None:
    """
    Among spans on the same row as y_ref, return the rightmost one
    that parses as a number >= min_val.
    """
    row = spans_on_row(spans, y_ref, tol)
    candidates = []
    for s in row:
        t = s["text"].replace(",","").replace("$","").replace("ea","").strip()
        try:
            v = float(t)
            if v >= min_val:
                candidates.append(s)
        except ValueError:
            pass
    if not candidates:
        return None
    return max(candidates, key=lambda s: s["x0"])

# ── Data extraction ─────────────────────────────────────────────────────────────

def extract_estimate_data(pdf_path: str) -> dict:
    """
    From the Estimate Details PDF extract:
      invoice_amount  – total Ext Price  (bottom of items table)
      estimate_cost   – total Ext Cost   (bottom of items table)
      contractor_rows – list of {page, ext_cost, rect} for contractor items
      rep_name        – design specialist name
    """
    doc = fitz.open(pdf_path)
    rep_name        = None
    invoice_amount  = None
    estimate_cost   = None
    contractor_rows = []
    design_fee_rows = []

    for pg_num, page in enumerate(doc):
        spans = get_spans(page)

        # ── Rep name (first page only) ─────────────────────────────────────────
        if rep_name is None:
            ds = find_label(spans, "Design Specialist")
            if ds:
                below = [s for s in spans
                         if ds["y0"] < s["y0"] <= ds["y0"] + 25
                         and not any(k in s["text"].lower()
                                     for k in ["phone:", "email:", "@", "date"])
                         and len(s["text"]) > 5]
                if below:
                    rep_name = min(below, key=lambda s: s["y0"])["text"].strip()

        # ── Contractor rows ────────────────────────────────────────────────────
        for s in spans:
            if any(kw in s["text"].lower() for kw in CONTRACTOR_KEYWORDS):
                y_ref = s["y0"]
                num = rightmost_number_on_row(spans, y_ref, min_val=10, tol=3)
                if num and num["x0"] > 400:
                    contractor_rows.append({
                        "page":     pg_num,
                        "ext_cost": parse_usd(num["text"]),
                        "rect":     fitz.Rect(num["x0"], num["y0"], num["x1"], num["y1"]),
                    })

        # ── Design Fee rows (auto-detected, always deducted) ───────────────────
        for s in spans:
            if any(kw in s["text"].lower() for kw in DESIGN_FEE_KEYWORDS):
                y_ref = s["y0"]
                num = rightmost_number_on_row(spans, y_ref, min_val=1, tol=3)
                if num and num["x0"] > 400:
                    design_fee_rows.append({
                        "page":     pg_num,
                        "ext_cost": parse_usd(num["text"]),
                        "rect":     fitz.Rect(num["x0"], num["y0"], num["x1"], num["y1"]),
                    })

        # ── Totals row: two $ amounts side-by-side at bottom of items ──────────
        for s in spans:
            if "$" not in s["text"]:
                continue
            row = spans_on_row(spans, s["y0"], tol=2)
            amounts = []
            for r in row:
                if "$" in r["text"]:
                    try:
                        amounts.append((r["x0"], parse_usd(r["text"])))
                    except ValueError:
                        pass
            if len(amounts) >= 2:
                amounts.sort()
                a1 = amounts[-2][1]
                a2 = amounts[-1][1]
                if a1 > 5000 and 500 < a2 < a1:
                    invoice_amount = a1
                    estimate_cost  = a2

    doc.close()

    design_fee_total  = sum(r["ext_cost"] for r in design_fee_rows)
    adjusted_cost     = (estimate_cost or 0.0) - design_fee_total

    return dict(
        rep_name          = rep_name or "Unknown",
        invoice_amount    = invoice_amount or 0.0,
        estimate_cost     = adjusted_cost,
        estimate_cost_raw = estimate_cost or 0.0,
        design_fee_total  = design_fee_total,
        design_fee_rows   = design_fee_rows,
        contractor_task   = sum(r["ext_cost"] for r in contractor_rows),
        contractor_rows   = contractor_rows,
    )


def extract_commission_table(pdf_path: str) -> dict:
    """
    From the Commission Sheet PDF (page 1), extract the current values
    and their rects from the Commission Estimate Details table.
    Returns dict with keys: greenline, pct_greenline, commission_pct, est_commission
    Each value: { text, rect }
    """
    doc    = fitz.open(pdf_path)
    page   = doc[0]
    spans  = get_spans(page)
    doc.close()

    result = {}

    # Label → key mapping (check most-specific first)
    targets = [
        ("Estimated Commission", "est_commission"),
        ("% Greenline",          "pct_greenline"),
        ("Greenline",            "greenline"),
        ("Commission",           "commission_pct"),
    ]

    for label, key in targets:
        lspan = find_label(spans, label)
        if lspan is None:
            # Try prefix match for edge cases
            lspan = find_label_startswith(spans, label)
        if lspan is None:
            continue
        # Find value span on the same row (right side)
        row = spans_on_row(spans, lspan["y0"], tol=3)
        val_spans = [s for s in row if ("$" in s["text"] or "%" in s["text"])
                     and s["x0"] > lspan["x1"]]
        if val_spans:
            best = max(val_spans, key=lambda s: s["x0"])
            result[key] = dict(
                text = best["text"].strip(),
                rect = fitz.Rect(best["x0"], best["y0"], best["x1"], best["y1"]),
            )

    return result


# ── Commission calculations ─────────────────────────────────────────────────────

def calculate(invoice_amount, estimate_cost, contractor_task, finance_fee,
              commission_pct_system, rep_name, job_date=None) -> dict:
    """
    Apply all commission rules and return a results dict.
    Formulas verified against sample output:
      PM  = (Invoice − EC − FF) / Invoice
      GL  = ((EC − CT) × 2.5 / 0.8) + (FF + CT)
      %GL = Invoice / GL
      Est = (Invoice − FF − CT) × Comm%

    Date-based tiers (standard reps only):
      For jobs dated March 31, 2026 or earlier:
        ≥100% → 10%, ≥90% → 9%, ≥80% → 8%, ≥60% → 6%, <60% → 4%
      For jobs dated April 1, 2026 or later:
        ≥100% → 10%, ≥90% → 9%, ≥80% → 8%, <80% → 4%
    """
    from datetime import date as _date
    ia = invoice_amount
    ec = estimate_cost
    ct = contractor_task
    ff = finance_fee or 0.0

    profit_margin = (ia - ec - ff) / ia if ia else 0.0
    greenline     = ((ec - ct) * 2.5 / 0.8) + (ff + ct)
    pct_greenline = ia / greenline if greenline else 0.0
    pct_gl_int    = round(pct_greenline * 100)

    # Determine if old tiers apply (job date ≤ March 31, 2026)
    OLD_TIER_CUTOFF = _date(2026, 3, 31)
    use_old_tiers = (job_date is not None) and (job_date <= OLD_TIER_CUTOFF)

    # Commission % tier — always computed from table, never from PDF system value.
    is_sean = "sean" in (rep_name or "").lower()
    if is_sean:
        if   pct_gl_int >= 100: commission_pct = 0.12
        elif pct_gl_int >= 90:  commission_pct = 0.11
        elif pct_gl_int >= 80:  commission_pct = 0.10
        elif pct_gl_int >= 75:  commission_pct = 0.09
        else:                   commission_pct = 0.05
    elif use_old_tiers:
        if   pct_gl_int >= 100: commission_pct = 0.10
        elif pct_gl_int >= 90:  commission_pct = 0.09
        elif pct_gl_int >= 80:  commission_pct = 0.08
        elif pct_gl_int >= 60:  commission_pct = 0.06
        else:                   commission_pct = 0.04
    else:
        if   pct_gl_int >= 100: commission_pct = 0.10
        elif pct_gl_int >= 90:  commission_pct = 0.09
        elif pct_gl_int >= 80:  commission_pct = 0.08
        else:                   commission_pct = 0.04

    est_commission = (ia - ff - ct) * commission_pct

    # Alert remarks
    remarks = []
    if round(profit_margin * 100) < 60:
        remarks.append(f"Profit Margin: {round(profit_margin*100)}% – Please inform Anne.")
    if pct_gl_int < 80:
        remarks.append(f"Greenline: {pct_gl_int}% – Please inform Anne.")

    return dict(
        profit_margin  = profit_margin,
        greenline      = greenline,
        pct_greenline  = pct_greenline,
        commission_pct = commission_pct,
        est_commission = est_commission,
        remarks        = remarks,
    )


# ── PDF annotation helpers ──────────────────────────────────────────────────────

def red_text(page, x: float, y: float, text: str, size: float = 11):
    """Insert bold red text. y = baseline (bottom of text)."""
    page.insert_text(fitz.Point(x, y), text,
                     fontname="hebo", fontsize=size, color=RED)


def strikethrough_and_replace(page, value_rect: fitz.Rect,
                               new_text: str, font_size: float = 10):
    """Draw a red strikethrough over value_rect, then insert new_text in red to its right."""
    y_mid = (value_rect.y0 + value_rect.y1) / 2
    shape = page.new_shape()
    shape.draw_line(fitz.Point(value_rect.x0, y_mid),
                    fitz.Point(value_rect.x1, y_mid))
    shape.finish(color=RED, width=1.4)
    shape.commit()
    # New value right of old one, baseline aligned
    red_text(page, value_rect.x1 + 8, value_rect.y1 - 1, new_text, size=font_size)


# ── Main annotation routines ────────────────────────────────────────────────────

def annotate_estimate_details(pdf_path: str, output_path: str,
                               data: dict, results: dict,
                               finance_fee: float, ff_name: str):
    doc = fitz.open(pdf_path)
    ff  = finance_fee or 0.0
    ct  = data["contractor_task"]

    # 1 ── Highlight contractor Ext Cost value in yellow on each contractor row ─
    for row in data["contractor_rows"]:
        page = doc[row["page"]]
        rect = row["rect"]
        bg = fitz.Rect(rect.x0 - 2, rect.y0 - 1, rect.x1 + 2, rect.y1 + 1)
        shape = page.new_shape()
        shape.draw_rect(bg)
        shape.finish(fill=YELLOW, color=YELLOW, fill_opacity=0.50)
        shape.commit()

    # 2 ── Strike through Design Fee Ext Cost in red ────────────────────────────
    for row in data.get("design_fee_rows", []):
        page  = doc[row["page"]]
        rect  = row["rect"]
        y_mid = (rect.y0 + rect.y1) / 2
        shape = page.new_shape()
        shape.draw_line(fitz.Point(rect.x0, y_mid), fitz.Point(rect.x1, y_mid))
        shape.finish(color=RED, width=1.4)
        shape.commit()

    # 3 ── Summary annotation block ─────────────────────────────────────────────
    totals_pg_num = len(doc) - 1
    last_totals_y = None
    ia_str = fmt_usd(data["invoice_amount"])
    for pg_num, page in enumerate(doc):
        rects = page.search_for(ia_str)
        if rects:
            totals_pg_num = pg_num
            last_totals_y = rects[0].y0
            break

    if last_totals_y is None:
        last_totals_y = doc[totals_pg_num].rect.height - 100

    ann_lines = ["No Contractor Task" if ct == 0 else f"Contractor Task {fmt_usd(ct)}"]
    if ff > 0 or ff_name:
        lender = f"{ff_name} " if ff_name else ""
        ann_lines.append(f"{lender}{fmt_usd(ff)}")
    else:
        ann_lines.append("No Finance Fee Cash Sale")
    ann_lines.append(f"Profit Margin {fmt_pct(results['profit_margin'])}")

    line_h   = 15
    gl_text  = f"GL {fmt_usd(results['greenline'])}"

    # If the totals row is near the top of its page, place the annotation block
    # at the bottom of the previous page instead (avoids overlap with content).
    TOP_THRESHOLD = 150
    if last_totals_y < TOP_THRESHOLD and totals_pg_num > 0:
        ann_page = doc[totals_pg_num - 1]
        page_w   = ann_page.rect.width
        gl_y     = ann_page.rect.height - 40
        ann_y    = gl_y - len(ann_lines) * line_h - 10
    else:
        ann_page = doc[totals_pg_num]
        page_w   = ann_page.rect.width
        # All lines (including GL) placed above the totals row
        block_h  = len(ann_lines) * line_h + 18
        ann_y    = max(last_totals_y - block_h - 8, 40)
        gl_y     = ann_y + len(ann_lines) * line_h + 4

    ann_x = page_w * 0.52

    for i, line in enumerate(ann_lines):
        red_text(ann_page, ann_x, ann_y + i * line_h, line, size=11)

    red_text(ann_page, ann_x, gl_y, gl_text, size=14)

    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
    print(f"  ✅  Estimate Details  →  {output_path}")


def annotate_commission_sheet(pdf_path: str, output_path: str,
                               data: dict, results: dict,
                               comm_table: dict,
                               finance_fee: float, ff_name: str):
    doc  = fitz.open(pdf_path)
    ff   = finance_fee or 0.0
    ct   = data["contractor_task"]
    page = doc[0]

    # 1 ── Top-right floating note: Contractor Task + Finance Fee ───────────────
    # Anchor: right of the "Commission Estimate Details" header
    spans = get_spans(page)
    header = find_label(spans, "Commission Estimate Details")
    if header:
        note_x = header["x1"] + 18
        note_y = header["y1"] - 1
    else:
        note_x = page.rect.width * 0.60
        note_y = 240

    ct_label = "No Contractor Task" if ct == 0 else f"Contractor Task {fmt_usd(ct)}"
    red_text(page, note_x, note_y, ct_label, size=11)
    if ff > 0 or ff_name:
        lender = f"{ff_name} " if ff_name else ""
        red_text(page, note_x, note_y + 15, f"{lender}{fmt_usd(ff)}", size=11)
    else:
        red_text(page, note_x, note_y + 15, "No Finance Fee Cash Sale", size=11)

    # 2 ── Strike through old Greenline → insert corrected value ───────────────
    if "greenline" in comm_table:
        strikethrough_and_replace(
            page,
            comm_table["greenline"]["rect"],
            fmt_usd(results["greenline"]),
        )

    # 3 ── Strike through old % Greenline → insert corrected % ─────────────────
    if "pct_greenline" in comm_table:
        strikethrough_and_replace(
            page,
            comm_table["pct_greenline"]["rect"],
            fmt_pct(results["pct_greenline"]),
        )

    # 4 ── Strike through Commission % if calculated value differs from system ───
    if "commission_pct" in comm_table:
        try:
            system_pct = parse_pct(comm_table["commission_pct"]["text"])
            calc_pct   = results["commission_pct"]
            # Strike and replace if they differ (rounding to 2 decimals)
            if abs(system_pct - calc_pct) > 0.001:
                strikethrough_and_replace(
                    page,
                    comm_table["commission_pct"]["rect"],
                    fmt_pct(calc_pct),
                )
        except Exception:
            pass

    # 5 ── Strike through old Estimated Commission → insert corrected value ─────
    if "est_commission" in comm_table:
        strikethrough_and_replace(
            page,
            comm_table["est_commission"]["rect"],
            fmt_usd(results["est_commission"]),
        )

    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
    print(f"  ✅  Commission Sheet  →  {output_path}")


# ── CLI entry point ─────────────────────────────────────────────────────────────

def main():
    p = argparse.ArgumentParser(
        description="Commission PDF Annotation Tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("--estimate",    required=True,
                   help="Estimate Details PDF (has Unit Cost / Ext Cost columns)")
    p.add_argument("--commission",  required=True,
                   help="Commission Sheet PDF (has Commission Estimate Details table)")
    p.add_argument("--ff",          type=float, default=0.0,
                   help="Finance Fee amount (default 0)")
    p.add_argument("--ff-name",     default="",
                   help="Finance Fee lender name  e.g. 'Lendhi'")
    p.add_argument("--output-dir",  default=None,
                   help="Output directory (default: same folder as --estimate)")
    args = p.parse_args()

    # Resolve output directory
    out_dir = Path(args.output_dir) if args.output_dir else Path(args.estimate).parent
    out_dir.mkdir(parents=True, exist_ok=True)

    # ── Step 1: Extract ─────────────────────────────────────────────────────────
    print("\n📄  Reading PDFs …")
    est_data   = extract_estimate_data(args.estimate)
    comm_table = extract_commission_table(args.commission)

    commission_pct_system = None
    if "commission_pct" in comm_table:
        try:
            commission_pct_system = parse_pct(comm_table["commission_pct"]["text"])
        except Exception:
            pass

    # ── Step 2: Calculate ───────────────────────────────────────────────────────
    results = calculate(
        invoice_amount       = est_data["invoice_amount"],
        estimate_cost        = est_data["estimate_cost"],
        contractor_task      = est_data["contractor_task"],
        finance_fee          = args.ff,
        commission_pct_system= commission_pct_system,
        rep_name             = est_data["rep_name"],
    )

    # ── Step 3: Print summary ────────────────────────────────────────────────────
    ff_label = f"{args.ff_name} " if args.ff_name else ""
    design_fee_note = (f"  (Design Fee deducted: {fmt_usd(est_data['design_fee_total'])})"
                       if est_data['design_fee_total'] > 0 else "")
    print(f"""
╔══════════════════════════════════════════════════════╗
║         COMMISSION CALCULATION SUMMARY               ║
╠══════════════════════════════════════════════════════╣
  Rep              :  {est_data['rep_name']}
  Invoice Amount   :  {fmt_usd(est_data['invoice_amount'])}
  Estimate Cost    :  {fmt_usd(est_data['estimate_cost'])}{design_fee_note}
  Contractor Task  :  {fmt_usd(est_data['contractor_task'])}  ← Ext Cost only
  Finance Fee      :  {ff_label}{fmt_usd(args.ff)}
  System Comm %    :  {f"{commission_pct_system*100:.2f}%" if commission_pct_system else "N/A"}
╠══════════════════════════════════════════════════════╣
  Profit Margin    :  {fmt_pct(results['profit_margin'])}
  Greenline        :  {fmt_usd(results['greenline'])}
  % Greenline      :  {fmt_pct(results['pct_greenline'])}
  Commission %     :  {fmt_pct(results['commission_pct'])}
  Est. Commission  :  {fmt_usd(results['est_commission'])}""")

    if results["remarks"]:
        print(f"\n  ⚠️   ALERTS:")
        for r in results["remarks"]:
            print(f"       {r}")
    else:
        print(f"\n  ✅  No alerts.")

    print("╚══════════════════════════════════════════════════════╝")

    # ── Step 4: Annotate PDFs ───────────────────────────────────────────────────
    # Derive client name from the estimate filename.
    # Expected pattern: "Last , First_ Kit - XXXXXXX" or "Last , First - ..."
    # Output names: "Last , First - Estimate Details.pdf"
    #               "Last , First - Commission Sheet.pdf"
    est_stem    = Path(args.estimate).stem
    if "_" in est_stem:
        client_name = est_stem.split("_")[0].strip()
    elif " - " in est_stem:
        client_name = est_stem.split(" - ")[0].strip()
    else:
        client_name = est_stem.strip()

    est_out  = out_dir / f"{client_name} - Estimate Details.pdf"
    comm_out = out_dir / f"{client_name} - Commission Sheet.pdf"

    print("\n🖊   Annotating PDFs …")
    annotate_estimate_details(
        args.estimate, str(est_out), est_data, results, args.ff, args.ff_name
    )
    annotate_commission_sheet(
        args.commission, str(comm_out), est_data, results,
        comm_table, args.ff, args.ff_name
    )

    print(f"\n📁  Saved to: {out_dir}/\n")


if __name__ == "__main__":
    main()
