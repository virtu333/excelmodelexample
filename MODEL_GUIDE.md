# Arctura Revenue Model v1 — User Guide

## What This Model Does

Projects quarterly revenue, site deployment costs, and margin for Arctura's autonomous warehouse robotics platform across two phases:

1. **Anchor** — 5 named distribution partner deployments (20 total zones), Q2'27–Q4'29
2. **Expansion** — GTM expansion pipeline (additional zones via new partners), Q1'29–Q4'31

Timeline: 20 quarters, Q1'27 through Q4'31.

---

## Sheet Overview

| Sheet | Purpose | View |
|-------|---------|------|
| **Assumptions** | All inputs — pricing, timing, volumes, ramp curves, costs, OpEx, valuation | Quarterly + static |
| **Pipeline Assumptions** | Post-anchor new system pipeline (standalone, all 20 quarters) | Quarterly |
| **Anchor Revenue** | Revenue by partner × type (implementation, licensing, operations, maintenance, inspection) | Quarterly |
| **Expansion Pipeline** | Pipeline deployments and revenue from new partners signed | Quarterly |
| **Combined Summary** | Annual rollup of all revenue + zone deployment timeline | Annual |
| **Cost Forecast** | Site deployment costs (NRC + recurring) for anchor and expansion zones | Quarterly + Annual |
| **P&L Summary** | Revenue − Costs = GAAP and Non-GAAP Margin, with cumulative view | Annual |
| **Valuation - Anchor** | DCF, EV/Revenue, and EBITDA-based valuation (anchor only) | Annual |
| **Valuation - Total** | DCF, EV/Revenue, and EBITDA-based valuation (total business) | Annual |
| **Sources** | Key assumption sources and documentation | Static |
| **Sensitivity Analysis** | Revenue sensitivity to key input changes | Static |

---

## How the Sheets Connect

```
Assumptions ──────────────────────┐
      │                           │
Pipeline Assumptions ──┐          │
      │                ▼          ▼
      └──► Expansion Pipeline    Anchor Revenue
                │                      │
                └──────┬───────────────┘
                       ▼
                Combined Summary ──► P&L Summary
                       ▲                  │
                       │                  ▼
                Cost Forecast        Valuations (Anchor + Total)
```

- **Blue text + yellow fill** = editable inputs you can change for scenario analysis
- **Black text** = formulas and cross-sheet links (don't overwrite)

---

## Revenue Streams

### Anchor (5 partners)
- **Implementation** — One-time fees paid 40/35/25 at signing, go-live, and month 12
- **Licensing** — Annual per-zone fee ($85K/zone/yr), activates after stabilization period
- **Operations** — Per-minute utilization across 5 use cases, ramping over 8 quarters
- **Predictive Maintenance** — Per-robot/month fee, linear ramp from 40% to 100% over 5 quarters
- **Vision Quality Inspection** — Per-scan fee, starts 1 quarter after go-live with hockey stick ramp

### Expansion (pipeline)
- **Consulting + Implementation** — One-time fees at contract signing
- **Annual Subscription** — $95K/zone/yr from signing onward
- **Utilization** — Same operations/maintenance/inspection model, applied cohort-by-cohort with ramp

---

## Cost Structure

Costs represent **site deployment and operations** — network infrastructure, robotic equipment, training, platform specialist support.

- **Non-recurring** (NRC): ~$62K for the first zone at a partner; additional zones at 25% discount
- **Recurring**: ~$18K/month per site (first zone full price, additional zones 25% off)
- NRC hits at go-live; recurring runs from go-live onward

These are Arctura's costs to deliver, not customer-facing prices.

---

## Key Assumptions to Adjust

For scenario analysis, the most impactful blue-text inputs are:

| Input | Location | Impact |
|-------|----------|--------|
| Operation volumes (ops/zone/wk) | Assumptions rows 67–71, column F | Drives bulk of utilization revenue |
| New Partners Signed pipeline | Pipeline Assumptions row 6 | Controls expansion growth trajectory |
| Per-Zone Implementation MSRP | Assumptions C17 | Anchor implementation revenue |
| Annual Subscription rate | Assumptions C46 | Expansion recurring revenue |
| Hockey stick ramp curve | Assumptions C77:C84 | How fast revenue materializes after go-live |
| NRC / MRC per site | Assumptions C114–C115 | Cost base for margin analysis |

---

## Rebuilding the Model

The xlsx is generated from `build_model.py`. All formulas are Excel-native — values recalculate when you change inputs directly in the spreadsheet. To regenerate from scratch:

```bash
python build_model.py
python recalc_win.py arctura_5partner_gtm_model_v1.xlsx
```

---

## Color Legend

| Style | Meaning |
|-------|---------|
| **Blue text + yellow fill** | Editable input — change for scenarios |
| **Black text** | Formula or cross-sheet link (don't overwrite) |
| **Yellow fill** (`#F9E2A1`) | Marks all editable input cells |
| **Gray text** | Helper row / lookup (can ignore) |
