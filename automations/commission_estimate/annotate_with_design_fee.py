#!/usr/bin/env python3
"""
annotate_with_design_fee.py
───────────────────────────
Same as commission_annotator but deducts any "Bath Design Fee" / "Design Fee"
row's Ext Cost from estimate_cost before calculating, and strikes through
that Ext Cost value in red on the Estimate Details PDF.
"""

import sys
sys.path.insert(0, "/sessions/charming-sweet-carson/mnt/Automation/Commission Estimate Automation")

import fitz
import re
from pathlib import Path
from commission_annotator import (
    get_spans, find_label, find_label_startswith,
    spans_on_row, rightmost_number_on_row,
    parse_usd, parse_pct, fmt_usd, fmt_pct,
    extract_commission_table,
    calculate,
    annotate_estimate_details,
    annotate_commission_sheet,
    RED, YELLOW,
    CONTRACTOR_KEYWORDS,
)

DESIGN_FEE_KEYWORDS = ["bath design fee", "design fee"]


def extract_estimate_data_with_design_fee(pdf_path: str) -> dict:
    """
    Like extract_estimate_data but also:
    - Finds 'Bath Design Fee' rows and records their Ext Cost + rect
    - Subtracts Design Fee Ext Cost from estimate_cost
    """
    doc = fitz.open(pdf_path)
    rep_name        = None
    invoice_amount  = None
    estimate_cost   = None
    contractor_rows = []
    design_fee_rows = []   # NEW

    for pg_num, page in enumerate(doc):
        spans = get_spans(page)

        # ── Rep name ──────────────────────────────────────────────────────────
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

        # ── Contractor rows ───────────────────────────────────────────────────
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

        # ── Design Fee rows (NEW) ─────────────────────────────────────────────
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

        # ── Totals row ────────────────────────────────────────────────────────
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

    design_fee_total = sum(r["ext_cost"] for r in design_fee_rows)
    adjusted_cost    = (estimate_cost or 0.0) - design_fee_total

    return dict(
        rep_name          = rep_name or "Unknown",
        invoice_amount    = invoice_amount or 0.0,
        estimate_cost     = adjusted_cost,          # ← deducted
        estimate_cost_raw = estimate_cost or 0.0,   # original
        design_fee_total  = design_fee_total,
        design_fee_rows   = design_fee_rows,
        contractor_task   = sum(r["ext_cost"] for r in contractor_rows),
        contractor_rows   = contractor_rows,
    )


def annotate_estimate_with_design_fee(pdf_path: str, output_path: str,
                                       data: dict, results: dict,
                                       finance_fee: float, ff_name: str):
    """
    Wraps the standard annotate_estimate_details but also strikes through
    the Design Fee Ext Cost cell in red.
    """
    doc = fitz.open(pdf_path)
    ff  = finance_fee or 0.0
    ct  = data["contractor_task"]

    # 1 ── Highlight contractor rows in yellow ─────────────────────────────────
    for row in data["contractor_rows"]:
        page = doc[row["page"]]
        rect = row["rect"]
        bg = fitz.Rect(rect.x0 - 2, rect.y0 - 1, rect.x1 + 2, rect.y1 + 1)
        shape = page.new_shape()
        shape.draw_rect(bg)
        shape.finish(fill=YELLOW, color=YELLOW, fill_opacity=0.50)
        shape.commit()

    # 2 ── Strike through Design Fee Ext Cost in red (NEW) ────────────────────
    for row in data["design_fee_rows"]:
        page = doc[row["page"]]
        rect = row["rect"]
        y_mid = (rect.y0 + rect.y1) / 2
        shape = page.new_shape()
        shape.draw_line(fitz.Point(rect.x0, y_mid), fitz.Point(rect.x1, y_mid))
        shape.finish(color=RED, width=1.4)
        shape.commit()

    # 3 ── Summary annotation block ────────────────────────────────────────────
    # Find which page contains the totals row (invoice amount)
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

    line_h = 15

    # If the totals row is near the top of its page (not enough room above),
    # fall back to the bottom white space of the previous page instead.
    TOP_THRESHOLD = 150
    if last_totals_y < TOP_THRESHOLD and totals_pg_num > 0:
        ann_page = doc[totals_pg_num - 1]
        page_w   = ann_page.rect.width
        page_h   = ann_page.rect.height
        gl_y     = page_h - 40
        ann_y    = gl_y - len(ann_lines) * line_h - 10
    else:
        ann_page = doc[totals_pg_num]
        page_w   = ann_page.rect.width
        ann_y    = last_totals_y - len(ann_lines) * line_h - 14
        gl_y     = last_totals_y + 30

    ann_x = page_w * 0.52

    for i, line in enumerate(ann_lines):
        ann_page.insert_text(
            fitz.Point(ann_x, ann_y + i * line_h), line,
            fontname="hebo", fontsize=11, color=RED
        )

    ann_page.insert_text(
        fitz.Point(ann_x, gl_y), f"GL {fmt_usd(results['greenline'])}",
        fontname="hebo", fontsize=14, color=RED
    )

    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
    print(f"  ✅  Estimate Details  →  {output_path}")


# ── Main ────────────────────────────────────────────────────────────────────────

ESTIMATE_PDF    = "/sessions/charming-sweet-carson/mnt/uploads/Burkhart, Robert_ Bath - 2725155.pdf"
COMMISSION_PDF  = "/sessions/charming-sweet-carson/mnt/uploads/estimate_2725155_commission_sheet.pdf"
FINANCE_FEE     = 308.10
FF_NAME         = "Partial Cash / Partial Dividend"
OUTPUT_DIR      = "/sessions/charming-sweet-carson/mnt/Automation/Burkhart, Robert - Bath"

print("\n📄  Reading PDFs …")
est_data   = extract_estimate_data_with_design_fee(ESTIMATE_PDF)
comm_table = extract_commission_table(COMMISSION_PDF)

commission_pct_system = None
if "commission_pct" in comm_table:
    try:
        commission_pct_system = parse_pct(comm_table["commission_pct"]["text"])
    except Exception:
        pass

# Extract job date from commission sheet (format MM/DD/YYYY, top of page)
import re as _re
from datetime import datetime as _dt, date as _date
job_date = None
try:
    import fitz as _fitz
    _doc = _fitz.open(COMMISSION_PDF)
    _text = "".join(page.get_text() for page in _doc)
    _doc.close()
    _match = _re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', _text)
    if _match:
        job_date = _date(int(_match.group(3)), int(_match.group(1)), int(_match.group(2)))
        print(f"  📅  Job Date: {job_date.strftime('%B %d, %Y')}")
except Exception:
    pass

results = calculate(
    invoice_amount        = est_data["invoice_amount"],
    estimate_cost         = est_data["estimate_cost"],   # adjusted
    contractor_task       = est_data["contractor_task"],
    finance_fee           = FINANCE_FEE,
    commission_pct_system = commission_pct_system,
    rep_name              = est_data["rep_name"],
    job_date              = job_date,
)

print(f"""
╔══════════════════════════════════════════════════════╗
║     COMMISSION CALCULATION SUMMARY (w/ Design Fee)   ║
╠══════════════════════════════════════════════════════╣
  Rep              :  {est_data['rep_name']}
  Invoice Amount   :  {fmt_usd(est_data['invoice_amount'])}
  Estimate Cost    :  {fmt_usd(est_data['estimate_cost'])}  (deducted Design Fee {fmt_usd(est_data['design_fee_total'])})
  Contractor Task  :  {fmt_usd(est_data['contractor_task'])}
  Finance Fee      :  {fmt_usd(FINANCE_FEE)}
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

import tempfile, shutil
out_dir     = Path(OUTPUT_DIR)
client_name = "Burkhart, Robert"
est_out_final  = out_dir / f"{client_name} - Estimate Details.pdf"
comm_out_final = out_dir / f"{client_name} - Commission Sheet.pdf"
# Write to temp files first to avoid "cannot remove" lock issues
est_out  = Path(tempfile.mktemp(suffix=".pdf"))
comm_out = Path(tempfile.mktemp(suffix=".pdf"))

print("\n🖊   Annotating PDFs …")
annotate_estimate_with_design_fee(
    ESTIMATE_PDF, str(est_out), est_data, results, FINANCE_FEE, FF_NAME
)
annotate_commission_sheet(
    COMMISSION_PDF, str(comm_out), est_data, results,
    comm_table, FINANCE_FEE, FF_NAME
)

# Move temp files to final destinations
shutil.move(str(est_out),  str(est_out_final))
shutil.move(str(comm_out), str(comm_out_final))
print(f"\n📁  Saved to: {OUTPUT_DIR}/\n")
