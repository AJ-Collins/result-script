"""
TUM ICI – Results Proof / Validation Engine
============================================
Verifies arithmetic consistency of the extracted JSON data AND the
aggregated values that will appear in the Word document.

Arithmetic rules checked
------------------------

  For each RECORD (raw extracted data):
  [R1]  sat_exam = total - absent
  [R2]  pass + fail <= sat_exam
  [R3]  pass + fail + absent <= total
  [R4]  total = pass + fail + absent + incomplete  (incomplete >= 0)

  For pass_fail records – SP/SUP internal consistency:
  [S1]  sp_sat_exam  = sp_registered  - sp_absent
  [S2]  sup_sat_exam = sup_registered - sup_absent
  [S3]  sp_pass  + sp_fail  <= sp_sat_exam
  [S4]  sup_pass + sup_fail <= sup_sat_exam

  For pass_fail records – SP+SUP vs PDF SUMMARY total:
  [C1]  sp_registered  + sup_registered  >= total   (double-counted students OK)
  [C2]  sp_pass        + sup_pass        >= pass
  [C3]  sp_fail        + sup_fail        >= fail

  For competency records:
  [CP1] mty + pfy + cpt + nyc + abs == total
  [CP2] pass   == mty + pfy + cpt
  [CP3] fail   == nyc
  [CP4] absent == abs
  [CP5] sat_exam == total - absent

  For classification records:
  [CL1] total >= 1
  [CL2] sum(classification label counts) == total  (when present)

  For each AGGREGATED row (programme-year):
  [A1]  sat_exam = registered - absent
  [A2]  pass + fail <= sat_exam
  [A3]  pass + fail + absent <= total
  [A4]  incomplete = total - pass - fail - absent >= 0
  [AS1] sp_sat_exam  = sp_registered  - sp_absent    (when sp_registered > 0)
  [AS2] sup_sat_exam = sup_registered - sup_absent   (when sup_registered > 0)
  [AS3] sp_pass  + sp_fail  <= sp_sat_exam
  [AS4] sup_pass + sup_fail <= sup_sat_exam
  [AC1] sp_registered + sup_registered >= registered (when has SP/SUP)
  [AC2] sp_pass + sup_pass >= pass               (when has SP/SUP)
  [AC3] sp_fail + sup_fail >= fail               (when has SP/SUP)

  Grand total cross-checks (across all aggregated rows):
  [T1]  grand sat_exam = grand registered - grand absent
  [T2]  grand pass + fail <= grand sat_exam
  [T3]  SP grand sat_exam = SP grand registered - SP grand absent
  [T4]  SUP grand sat_exam = SUP grand registered - SUP grand absent
"""

import json
import sys
from pathlib import Path
from collections import defaultdict


# ─────────────────────────────────────────────────────────────────────────────
# Aggregation (mirrors generate.py)
# ─────────────────────────────────────────────────────────────────────────────

def _zero():
    return {
        "registered": 0, "absent": 0, "sat_exam": 0,
        "pass": 0, "fail": 0, "total": 0, "classified": 0,
        "sp_registered": 0, "sp_absent": 0, "sp_sat_exam": 0,
        "sp_pass": 0, "sp_fail": 0,
        "sup_registered": 0, "sup_absent": 0, "sup_sat_exam": 0,
        "sup_pass": 0, "sup_fail": 0,
    }


SP_SUP_KEYS = (
    "sp_registered", "sp_absent", "sp_sat_exam", "sp_pass", "sp_fail",
    "sup_registered", "sup_absent", "sup_sat_exam", "sup_pass", "sup_fail",
)


def _aggregate(records):
    agg = defaultdict(_zero)
    for r in records:
        key = (r["programme_name"], r["year_of_study"])
        a = agg[key]
        if r["summary_type"] in ("pass_fail", "competency"):
            a["registered"] += r["total"]
            a["absent"]     += r["absent"]
            a["sat_exam"]   += r["sat_exam"]
            a["pass"]       += r["pass"]
            a["fail"]       += r["fail"]
            a["total"]      += r["total"]
            for k in SP_SUP_KEYS:
                a[k] += r.get(k, 0)
        elif r["summary_type"] == "classification":
            a["registered"] += r["total"]
            a["sat_exam"]   += r["total"]
            a["pass"]       += r["total"]
            a["total"]      += r["total"]
            a["classified"] += r["total"]
    return agg


# ─────────────────────────────────────────────────────────────────────────────
# Record-level checks
# ─────────────────────────────────────────────────────────────────────────────

def _check_record(rec, idx, errors, warnings, infos):
    pg   = rec.get("source_page", "?")
    prog = rec.get("programme_name", "?")[:40]
    yr   = rec.get("year_of_study", "?")
    st   = rec["summary_type"]
    ctx  = f"[rec#{idx+1} p{pg} {st} {prog} yr{yr}]"

    if st == "pass_fail":
        p, f, a, t = rec["pass"], rec["fail"], rec["absent"], rec["total"]
        sat = rec.get("sat_exam", t - a)

        # Core arithmetic
        if sat != t - a:
            errors.append(f"{ctx} [R1] sat_exam={sat} != total({t})-absent({a})={t-a}")
        if p + f > sat:
            errors.append(f"{ctx} [R2] pass({p})+fail({f})={p+f} > sat_exam({sat})")
        if p + f + a > t:
            errors.append(f"{ctx} [R3] pass({p})+fail({f})+absent({a})={p+f+a} > total({t})")
        inc = t - p - f - a
        if inc < 0:
            errors.append(f"{ctx} [R4] incomplete={inc} < 0")

        # SP internal
        sp_r = rec.get("sp_registered",  0)
        sp_a = rec.get("sp_absent",      0)
        sp_s = rec.get("sp_sat_exam",    0)
        sp_p = rec.get("sp_pass",        0)
        sp_f = rec.get("sp_fail",        0)
        if sp_r > 0:
            if sp_s != sp_r - sp_a:
                errors.append(f"{ctx} [S1] sp_sat_exam={sp_s} != sp_registered({sp_r})-sp_absent({sp_a})={sp_r-sp_a}")
            if sp_p + sp_f > sp_s:
                errors.append(f"{ctx} [S2] sp_pass({sp_p})+sp_fail({sp_f})={sp_p+sp_f} > sp_sat_exam({sp_s})")

        # SUP internal
        su_r = rec.get("sup_registered",  0)
        su_a = rec.get("sup_absent",      0)
        su_s = rec.get("sup_sat_exam",    0)
        su_p = rec.get("sup_pass",        0)
        su_f = rec.get("sup_fail",        0)
        if su_r > 0:
            if su_s != su_r - su_a:
                errors.append(f"{ctx} [S2] sup_sat_exam={su_s} != sup_registered({su_r})-sup_absent({su_a})={su_r-su_a}")
            if su_p + su_f > su_s:
                errors.append(f"{ctx} [S4] sup_pass({su_p})+sup_fail({su_f})={su_p+su_f} > sup_sat_exam({su_s})")

        # Combined vs PDF total
        if sp_r + su_r > 0:
            if sp_r + su_r < t:
                errors.append(f"{ctx} [C1] sp_reg({sp_r})+sup_reg({su_r})={sp_r+su_r} < total({t})")
            if sp_p + su_p < p:
                # Can be < p if some students had undetermined outcome (WARN not ERROR)
                warnings.append(f"{ctx} [C2] sp_pass({sp_p})+sup_pass({su_p})={sp_p+su_p} < pass({p}) "
                                f"(possible unknown-outcome students — check WARN lines in extract output)")
            if sp_f + su_f < f:
                warnings.append(f"{ctx} [C3] sp_fail({sp_f})+sup_fail({su_f})={sp_f+su_f} < fail({f}) "
                                f"(possible unknown-outcome students)")
            if sp_r + su_r > t:
                infos.append(f"{ctx} [INFO] sp_reg+sup_reg={sp_r+su_r} > total({t}) "
                             f"(students in both SP and SUP — expected double-count)")


    elif st == "competency":
        mty = rec["mty"]; pfy = rec["pfy"]; cpt = rec["cpt"]
        nyc = rec["nyc"]; ab  = rec["abs"]; tot = rec["total"]
        if mty + pfy + cpt + nyc + ab != tot:
            errors.append(f"{ctx} [CP1] mty+pfy+cpt+nyc+abs={mty+pfy+cpt+nyc+ab} != total({tot})")
        if rec["pass"] != mty + pfy + cpt:
            errors.append(f"{ctx} [CP2] pass={rec['pass']} != mty+pfy+cpt={mty+pfy+cpt}")
        if rec["fail"] != nyc:
            errors.append(f"{ctx} [CP3] fail={rec['fail']} != nyc={nyc}")
        if rec["absent"] != ab:
            errors.append(f"{ctx} [CP4] absent={rec['absent']} != abs={ab}")
        sat = rec.get("sat_exam", tot - ab)
        if sat != tot - ab:
            errors.append(f"{ctx} [CP5] sat_exam={sat} != total({tot})-abs({ab})={tot-ab}")

    elif st == "classification":
        tot = rec["total"]
        if tot < 1:
            errors.append(f"{ctx} [CL1] total={tot} < 1")
        cls = rec.get("classification", {})
        if cls:
            label_sum = sum(v for v in cls.values() if isinstance(v, int))
            if label_sum != tot:
                errors.append(f"{ctx} [CL2] classification sum={label_sum} != total({tot}) | {cls}")


# ─────────────────────────────────────────────────────────────────────────────
# Aggregated-row checks
# ─────────────────────────────────────────────────────────────────────────────

def _check_agg_row(prog, year, v, errors, warnings, infos):
    ctx = f"[agg '{prog[:40]}' yr{year}]"
    r  = v["registered"]; a  = v["absent"]
    s  = v["sat_exam"];   p  = v["pass"]
    f  = v["fail"];       t  = v["total"]

    if s != r - a:
        errors.append(f"{ctx} [A1] sat_exam={s} != registered({r})-absent({a})={r-a}")
    if p + f > s:
        errors.append(f"{ctx} [A2] pass({p})+fail({f})={p+f} > sat_exam({s})")
    if p + f + a > t:
        errors.append(f"{ctx} [A3] pass({p})+fail({f})+absent({a})={p+f+a} > total({t})")
    inc = t - p - f - a
    if inc < 0:
        errors.append(f"{ctx} [A4] incomplete={inc} < 0")

    sp_r  = v["sp_registered"];  sp_a  = v["sp_absent"]
    sp_s  = v["sp_sat_exam"];    sp_p  = v["sp_pass"];  sp_f = v["sp_fail"]
    su_r  = v["sup_registered"]; su_a  = v["sup_absent"]
    su_s  = v["sup_sat_exam"];   su_p  = v["sup_pass"]; su_f = v["sup_fail"]

    if sp_r > 0:
        if sp_s != sp_r - sp_a:
            errors.append(f"{ctx} [AS1] sp_sat_exam={sp_s} != sp_reg({sp_r})-sp_abs({sp_a})={sp_r-sp_a}")
        if sp_p + sp_f > sp_s:
            errors.append(f"{ctx} [AS3] sp_pass+sp_fail={sp_p+sp_f} > sp_sat_exam({sp_s})")
    if su_r > 0:
        if su_s != su_r - su_a:
            errors.append(f"{ctx} [AS2] sup_sat_exam={su_s} != sup_reg({su_r})-sup_abs({su_a})={su_r-su_a}")
        if su_p + su_f > su_s:
            errors.append(f"{ctx} [AS4] sup_pass+sup_fail={su_p+su_f} > sup_sat_exam({su_s})")

    if sp_r + su_r > 0:
        classified = v.get("classified", 0)
        # [AC] checks compare SP+SUP counts against the combined totals.
        # Classification records contribute to registered/pass but NOT to
        # sp_registered/sup_registered.  When classified > 0, the combined
        # registered / pass includes those classification-only students, so we
        # must subtract them before comparing, or skip entirely.
        # We compare SP+SUP against the non-classified portion of totals.
        pf_registered = r - classified   # students from pass_fail / competency sheets only
        pf_pass       = p - classified   # same reasoning for pass
        if pf_registered > 0 and sp_r + su_r < pf_registered:
            errors.append(
                f"{ctx} [AC1] sp_reg({sp_r})+sup_reg({su_r})={sp_r+su_r} "
                f"< pf_registered({pf_registered}) [registered({r})-classified({classified})]"
            )
        if pf_pass > 0 and sp_p + su_p < pf_pass:
            errors.append(
                f"{ctx} [AC2] sp_pass({sp_p})+sup_pass({su_p})={sp_p+su_p} "
                f"< pf_pass({pf_pass}) [pass({p})-classified({classified})]"
            )
        if sp_f + su_f < f:
            errors.append(f"{ctx} [AC3] sp_fail({sp_f})+sup_fail({su_f})={sp_f+su_f} < fail({f})")
        if sp_r + su_r > pf_registered:
            infos.append(f"{ctx} [INFO] sp_reg+sup_reg={sp_r+su_r} > pf_registered({pf_registered}) (double-counted — expected)")


# ─────────────────────────────────────────────────────────────────────────────
# Grand-total cross-checks
# ─────────────────────────────────────────────────────────────────────────────

def _check_grand_totals(agg, errors, warnings):
    grand = defaultdict(int)
    for (prog, year), v in agg.items():
        for k in _zero().keys():
            grand[k] += v[k]

    ctx = "[grand totals]"

    if grand["sat_exam"] != grand["registered"] - grand["absent"]:
        errors.append(
            f"{ctx} [T1] grand sat_exam={grand['sat_exam']} != "
            f"registered({grand['registered']})-absent({grand['absent']})="
            f"{grand['registered']-grand['absent']}"
        )
    if grand["pass"] + grand["fail"] > grand["sat_exam"]:
        errors.append(
            f"{ctx} [T2] grand pass({grand['pass']})+fail({grand['fail']})="
            f"{grand['pass']+grand['fail']} > sat_exam={grand['sat_exam']}"
        )
    sp_r = grand["sp_registered"]; sp_a = grand["sp_absent"]; sp_s = grand["sp_sat_exam"]
    if sp_r > 0 and sp_s != sp_r - sp_a:
        errors.append(f"{ctx} [T3] SP sat_exam={sp_s} != SP reg({sp_r})-abs({sp_a})={sp_r-sp_a}")
    su_r = grand["sup_registered"]; su_a = grand["sup_absent"]; su_s = grand["sup_sat_exam"]
    if su_r > 0 and su_s != su_r - su_a:
        errors.append(f"{ctx} [T4] SUP sat_exam={su_s} != SUP reg({su_r})-abs({su_a})={su_r-su_a}")

    return dict(grand)


# ─────────────────────────────────────────────────────────────────────────────
# Main entry point
# ─────────────────────────────────────────────────────────────────────────────

def run_proof(json_path: Path) -> dict:
    data    = json.loads(json_path.read_text())
    records = data["records"]
    errors: list  = []
    warnings: list = []
    infos: list   = []

    SEP = "=" * 65
    print(f"\n{SEP}")
    print(f"PROOF REPORT  –  {json_path.name}")
    print(SEP)
    print(f"Total records in JSON: {len(records)}\n")

    # ── Per-record checks ──────────────────────────────────────────────────────
    for idx, rec in enumerate(records):
        _check_record(rec, idx, errors, warnings, infos)

    # ── Aggregation checks ─────────────────────────────────────────────────────
    agg = _aggregate(records)
    for (prog, year), v in sorted(agg.items()):
        _check_agg_row(prog, year, v, errors, warnings, infos)

    # ── Grand-total checks ─────────────────────────────────────────────────────
    grand = _check_grand_totals(agg, errors, warnings)

    # ── Print findings ─────────────────────────────────────────────────────────
    if errors:
        print(f"ERRORS ({len(errors)}):")
        for e in errors:
            print(f"  ERR  {e}")
    else:
        print("No errors found.")

    if warnings:
        print(f"\nWARNINGS ({len(warnings)}):")
        for w in warnings:
            print(f"  WARN {w}")

    if infos:
        print(f"\nINFO ({len(infos)}):")
        for info in infos:
            print(f"  INFO {info}")

    # ── Grand summary table ────────────────────────────────────────────────────
    print(f"\n{'-'*65}")
    print("GRAND TOTALS across all aggregated programme-year rows:")
    print(f"  {'Metric':<20}  {'SP':>6}  {'SUP':>6}  {'Combined':>8}")
    print(f"  {'-'*20}  {'-'*6}  {'-'*6}  {'-'*8}")
    for label, sp_k, sup_k, comb_k in [
        ("Registered",  "sp_registered",  "sup_registered",  "registered"),
        ("Absent",      "sp_absent",       "sup_absent",       "absent"),
        ("Sat Exam",    "sp_sat_exam",     "sup_sat_exam",     "sat_exam"),
        ("Pass",        "sp_pass",         "sup_pass",         "pass"),
        ("Fail",        "sp_fail",         "sup_fail",         "fail"),
        ("Total",       "sp_registered",   "sup_registered",   "total"),
        ("Classified",  None,              None,               "classified"),
    ]:
        sp  = grand.get(sp_k,   "-") if sp_k  else "-"
        sup = grand.get(sup_k,  "-") if sup_k else "-"
        comb = grand.get(comb_k, 0)
        print(f"  {label:<20}  {str(sp):>6}  {str(sup):>6}  {comb:>8}")

    inc_comb = max(0, grand["total"] - grand["pass"] - grand["fail"] - grand["absent"])
    sp_inc   = max(0, grand["sp_registered"]  - grand["sp_sat_exam"]  - grand["sp_absent"])
    sup_inc  = max(0, grand["sup_registered"] - grand["sup_sat_exam"] - grand["sup_absent"])
    print(f"  {'Incomplete':<20}  {sp_inc:>6}  {sup_inc:>6}  {inc_comb:>8}")
    print(f"{'-'*65}")

    # ── Per-programme-year detail ──────────────────────────────────────────────
    print(f"\nDETAIL – Per programme-year aggregated rows ({len(agg)} rows):")
    print(f"  {'Programme (truncated)':<42} {'Yr':>3}  {'Reg':>4}  {'Abs':>4}  "
          f"{'Sat':>4}  {'P':>4}  {'F':>4}  {'Tot':>4}  "
          f"{'SP_R':>5}  {'SP_P':>5}  {'SP_F':>5}  "
          f"{'SU_R':>5}  {'SU_P':>5}  {'SU_F':>5}  {'Cls':>4}")
    print("  " + "-" * 130)
    for (prog, year), v in sorted(agg.items(), key=lambda x: x[0]):
        ok = ""
        # Quick sanity flag
        sat_ok = v["sat_exam"] == v["registered"] - v["absent"]
        pf_ok  = v["pass"] + v["fail"] <= v["sat_exam"]
        if not sat_ok or not pf_ok:
            ok = " <-- FAIL"
        print(
            f"  {prog[:42]:<42} {year:>3}  "
            f"{v['registered']:>4}  {v['absent']:>4}  {v['sat_exam']:>4}  "
            f"{v['pass']:>4}  {v['fail']:>4}  {v['total']:>4}  "
            f"{v['sp_registered']:>5}  {v['sp_pass']:>5}  {v['sp_fail']:>5}  "
            f"{v['sup_registered']:>5}  {v['sup_pass']:>5}  {v['sup_fail']:>5}  "
            f"{v['classified']:>4}{ok}"
        )

    status = "PASSED ✓" if not errors else f"FAILED ✗ ({len(errors)} errors)"
    print(f"\nOVERALL: {status}\n")

    return {
        "error_count":   len(errors),
        "warning_count": len(warnings),
        "info_count":    len(infos),
        "errors":        errors,
        "warnings":      warnings,
        "infos":         infos,
        "grand_totals":  grand,
        "passed":        len(errors) == 0,
    }


if __name__ == "__main__":
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("output_extracted.json")
    result = run_proof(path)
    sys.exit(0 if result["passed"] else 1)
