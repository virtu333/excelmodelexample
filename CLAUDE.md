# Arctura Revenue Model v1 — Build Specification

## Project Overview

Excel financial model (`arctura_5partner_gtm_model_v1.xlsx`) for Arctura, an autonomous warehouse robotics startup. Projects quarterly revenue, costs, and margin across 5 anchor distribution partner deployments and a post-anchor GTM expansion pipeline.

**Build script**: `build_model.py` (~2,760 lines) — generates the entire workbook from scratch
**Output file**: `arctura_5partner_gtm_model_v1.xlsx` (11 sheets, ~2,850 formulas)

## Technical Stack

- **openpyxl** for all Excel creation (formulas, formatting, structure)
- **recalc_win.py** for LibreOffice formula recalculation after build
- **validate_*.py** — post-build validation scripts (EBITDA, non-GAAP, sensitivity)
- All revenue calculations use **Excel formulas**, never Python-computed hardcoded values
- See `skills/SKILL.md` for detailed xlsx creation best practices

## Financial Model Conventions

- **Blue text + yellow fill** (`#0000FF` font, `#F9E2A1` fill): Editable inputs / scenario-adjustable values
- **Black text** (`#000000`): All formulas, calculations, and cross-sheet links
- **Gray text**: Helper rows (quarter indices, lookups)
- **Title Case** headers throughout — no ALL CAPS
- **Currency**: `$#,##0;($#,##0);"-"` (zeros as dash)
- **Percentages**: `0.0%`
- **Integers**: `0` (no decimals)
- Font: Arial throughout

## Timeline

- 20 quarters: Q1'27 through Q4'31
- Quarter data starts in **column C** on all quarterly sheets
- Q1'27 = column C, Q2'27 = column D, ..., Q4'31 = column V

---

## Sheets (11 total)

| # | Sheet | Builder Function | Purpose |
|---|-------|-----------------|---------|
| 1 | Assumptions | `build_assumptions()` | All inputs: pricing, timing, volumes, ramp curves, costs, OpEx, valuation |
| 2 | Pipeline Assumptions | `build_pipeline_assumptions()` | Post-anchor new partner pipeline (20 quarters, standalone) |
| 3 | Anchor Revenue | `build_anchor()` | Revenue by partner × type (implementation, licensing, operations, maintenance, inspection) |
| 4 | Expansion Pipeline | `build_expansion()` | Pipeline deployments and revenue from new partners |
| 5 | Combined Summary | `build_combined()` | Annual rollup of all revenue + zone deployment timeline |
| 6 | Cost Forecast | `build_cost_forecast()` | Site deployment costs (NRC + recurring) |
| 7 | P&L Summary | `build_pnl()` | Revenue − Costs = GAAP/Non-GAAP margin |
| 8 | Valuation - Anchor | `build_valuation_anchor()` | DCF + EV/Revenue + EBITDA valuation (anchor only) |
| 9 | Valuation - Total | `build_valuation_total()` | DCF + EV/Revenue + EBITDA valuation (total business) |
| 10 | Sources | `build_sources()` | Trimmed reference for key assumption sources |
| 11 | Sensitivity Analysis | `build_sensitivity()` | Revenue sensitivity to key inputs |

---

## Sheet 1: Assumptions (row map from `build_model.py` constants)

### Anchor Distribution Partners (rows 8-26)
- Row 10 (`A_ZONES`): Number of Zones — blue inputs: 6, 4, 5, 3, 2 (columns C:G)
- Row 11 (`A_PLAN_Q`): Planning Start Quarter — blue inputs
- Row 12 (`A_GOLIVE_Q`): Go-Live Quarter — blue inputs
- Row 13 (`A_LICENSE_Q`): License Activation Quarter — blue inputs
- Row 14 (`A_IMPL_MSRP`): Implementation MSRP — formula: `=zones × per_zone_msrp`
- Row 15 (`A_LIC_FEE`): Annual Licensing Fee — formula: `=zones × per_zone_license`
- Row 17 (`A_PERZONE_MSRP`): Per-Zone Implementation MSRP: $350,000
- Row 18 (`A_PERZONE_LIC`): Per-Zone Annual License: $85,000
- Rows 24-26: Payment split 40/35/25 (signing/go-live/month 12)

### Zone Go-Live Ramp (rows 28-36, detailed in v1)
- Per-partner zone activation schedule across Q1'27-Q4'28
- Row 35 (`A_ZONE_RAMP_TOTAL`): Total zones going live per quarter
- Row 36 (`A_ZONE_RAMP_CUM`): Cumulative zones live

### Quarter Index Lookups (rows 38-41)
- Rows 39-41 (`A_PLAN_IDX`, `A_GOLIVE_IDX`, `A_LICENSE_IDX`): MATCH-based index lookups

### Expansion Customer Pricing (rows 43-48)
- Row 44 (`A_EXP_IMPL_FEE`): Implementation Fee per zone: $30,000
- Row 45 (`A_EXP_CONSULT`): Program Consulting Fee per partner: $75,000
- Row 46 (`A_EXP_SUB`): Annual Subscription per zone: $95,000
- Row 48 (`A_EXP_IMPL_Q`): Implementation to Go-Live (quarters): formula `=ROUNDUP(C47/3,0)`

### Pipeline Redirect (row 50)
Row 50: Gray italic note "See 'Pipeline Assumptions' tab" — pipeline data moved to separate sheet

### Utilization Pricing (rows 59-63)
- Row 60 (`A_TIER1`): Tier 1 $/minute: $3.50
- Row 61 (`A_TIER2`): Tier 2 $/minute: $0.75
- Row 62 (`A_MAINT_PRICE`): Predictive Maintenance $/robot/month: $40.00
- Row 63 (`A_INSP_PRICE`): Vision Inspection $/scan: $12.00

### Use Case Definitions (rows 65-74)
Row 67+ (`A_UC_START`): 5 operation use cases + 2 fixed (maintenance, inspection). All blue inputs.

### Hockey Stick Ramp (rows 76-84)
Rows 77-84 (`A_RAMP_START:A_RAMP_END`): Q1-Q8 ramp: 12%, 18%, 28%, 42%, 58%, 75%, 88%, 100%

### Site Volume Decay (rows 86-95)
- Rows 87-92: Per-zone decay (100%→65%)
- Rows 93-95 (`A_DECAY_6/5/4/3/2`): Blended averages for 6/5/4/3/2 zones

### Predictive Maintenance (rows 97-100)
- Row 98 (`A_ROBOTS`): Robots per Zone: 80
- Row 99 (`A_MAINT_RAMP_START`): Ramp Start: 40%
- Row 100 (`A_MAINT_RAMP_Q`): Quarters to Full: 5

### Vision Quality Inspection (rows 102-103)
- Row 103 (`A_SCANS`): Scans per Zone per Month: 250

### Scenario Volumes (rows 105-112)
Bear/Base/Bull/Evidence Range for 5 operation use cases

### Site Cost Assumptions (rows 114-124)
- Row 115 (`A_NRC_FIRST`): NRC first zone
- Row 116 (`A_MRC`): Monthly recurring cost
- Row 117 (`A_DISC`): Additional zone discount: 25%
- Rows 123-124 (`A_BLEND_NRC/QRC`): Blended rates referencing `'Pipeline Assumptions'!C7`

### Operating Expense & Valuation (rows 126-130)
- Row 127 (`A_GA_PCT`): G&A OpEx %: 12%
- Row 128 (`A_RD_PCT`): R&D OpEx %: 18%
- Row 130 (`A_EBITDA_MULT`): EBITDA Terminal Multiple: 12x

### Sources & Notes (row 132+)
Reference table with Ref | Parameter | Value | Sources | Notes columns

---

## Sheet 2: Pipeline Assumptions (standalone tab)

20-quarter layout (Q1'27-Q4'31). Q1'27-Q4'28 = zeros; Q1'29+ = pipeline ramp.

| Row | Constant | Label | Type |
|-----|----------|-------|------|
| 1 | `PA_HDR` | Title | Header |
| 3 | `PA_QHDR` | Quarter headers | Q1'27-Q4'31 |
| 4 | `PA_QIDX` | Quarter index | 1-20 |
| 6 | `PA_NEWSYS` | New Partners Signed | Blue input + yellow fill |
| 7 | `PA_AVGZONES` | Avg Zones per Partner | Blue input (3) in col C |
| 8 | `PA_NEWZONES` | New Zones This Quarter | Formula: `=C6*$C$7` |
| 9 | `PA_CUMZONES` | Cumulative New Zones | Running sum |
| 10 | `PA_GOLIVE` | Go-Live Zones This Quarter | INDEX/OFFSET with lag from `Assumptions!$C$48` |
| 11 | `PA_CUMGOLIVE` | Cumulative Go-Live Zones | Running sum |

Cross-sheet references from other sheets use `'Pipeline Assumptions'!` prefix (quoted because of space in name).

---

## Sheets 3-11: Summary

Detailed formula patterns for Anchor Revenue, Expansion Pipeline, Combined Summary, Cost Forecast, P&L, Valuations, Sources, and Sensitivity are documented in the builder functions in `build_model.py`. Key patterns:

- **Literal quarter indices** in formulas (1-20), not COLUMN()-based — enables copy-paste extensibility
- **Cohort-by-cohort ramp summation** for expansion utilization revenue — looks back through prior go-live quarters
- **Cross-sheet references** use black font
- **Annual summaries** sum 4 quarterly columns per year

---

## Build & Verify

```bash
python build_model.py                                          # Generate xlsx
python recalc_win.py arctura_5partner_gtm_model_v1.xlsx        # Recalculate formulas (LibreOffice)
python validate_ebitda.py                                       # Optional: check EBITDA calcs
python validate_nongaap.py                                      # Optional: check non-GAAP metrics
python validate_sensitivity.py                                  # Optional: check sensitivity ranges
```

### Validation Checks
- Total Anchor Zones (Assumptions!C19) = 20
- Partner 1 Implementation MSRP (Assumptions!C14) = $2,100,000
- Total Implementation Revenue (Assumptions!C20) = $7,000,000
- Pipeline Assumptions tab: 20 quarters, Q1-Q8 show 0 partners, Q9+ show pipeline values
- Pipeline-derived rows (zones, go-live) calculate correctly from formulas
- Zero formula errors in recalc output

## Implementation Notes

- **One script generates everything** — `build_model.py` creates all 11 sheets in sequence
- The script is ~2,760 lines. That's expected for this model complexity
- Helper functions: `set_cell()`, `write_quarter_headers()`, `write_quarter_indices()`, `qcol()`
- Row constants (`A_*`, `PA_*`) defined at top of script — use these, don't hardcode row numbers
- Expansion utilization cohort-by-cohort ramp is the most complex formula section
- All blue-font cells must have `fill=YELLOW_FILL` (`#F9E2A1`)
- All headers must be Title Case (preserve abbreviations: EBITDA, GAAP, GTM, NRC, QRC, MSRP, DCF, EV, COGS, G&A, R&D)
