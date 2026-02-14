# Building Financial Models in Excel with Claude

## The Core Insight

An `.xlsx` file is XML in a ZIP archive — a structured binary format no different from code, JSON, or any other machine-readable output. If an agent can generate a config file or a web page, it can generate a spreadsheet.

So you can build a live financial model — dynamic formulas, scenario toggles, industry-standard formatting, sourced assumptions, etc. with agentic coding.

## The Workflow

Three phases: design the model, specify it precisely, build it programmatically. The tools matter less than the separation between phases.

### 1. Model Design

Before writing any formulas, develop and stress-test the model's logic:

- Gather source data (10-Ks, investor decks, industry reports, raw datasets)
- Develop assumptions and debate them — growth rates, margin profiles, ramp curves, TAM sizing
- Define model architecture: sheet structure, driver logic, linkages between sections
- Establish scenario definitions and sensitivity ranges

This is the intellectual work. You're deciding that revenue should be `volume × ASP × utilization` with a ramp curve — not writing Excel syntax.

Both the Claude app and Claude Code work for this phase. The Claude app's project feature is useful for accumulating context across sessions and uploading source documents. Claude Code's plan mode works if you prefer staying in the terminal. Either way, the goal is the same: arrive at a defensible model design before generating any spreadsheet code.

### 2. Model Approach Document

The bridge between design and execution. This document specifies:

- **Sheet structure** — tabs, contents, data flow between them
- **Row/column layout** — time periods, line items, section breaks
- **Driver logic** — relationships in plain language (not Excel syntax)
- **Assumption architecture** — hardcoded inputs (blue font), calculations (black), cross-sheet references (green)
- **Formatting conventions** — number formats, units, sign conventions
- **Sources** — provenance for every hardcoded number

Precision here determines output quality. Ambiguity in this document becomes errors in the spreadsheet.

From the model approach, generate project scaffolding:

- **`CLAUDE.md`** — project context file Claude Code reads on startup (model structure, conventions, constraints)
- **Skill files** — `SKILL.md` and `recalc.py` from Anthropic's [skills repo](https://github.com/anthropics/skills/tree/main/skills/xlsx) (details below)
- **Data files** — cleaned source data
- **Directory structure**

### 3. Build and Validate

```bash
cd your-project-directory
claude
```

Start in plan mode. Review the project documentation, confirm the build approach, catch misalignments before any code runs. Then execute:

- Construct the workbook via `openpyxl` — formulas, formatting, structure
- Apply color coding (blue for inputs, black for formulas, green for cross-sheet links)
- Recalculate all formulas via `recalc.py`
- Validate for zero formula errors
- Iterate until clean

Open the output `.xlsx` in Excel or Sheets. Spot-check references, toggle scenarios, verify sensitivity behavior. Feed findings back into the design phase as needed.

## Under the Hood

Claude Code doesn't operate inside Excel. No GUI is involved. It writes Python that *constructs* an Excel file programmatically.

The pipeline:

1. **Python + openpyxl** builds the workbook: creates sheets, writes values and formula strings, applies formatting. Output is a valid `.xlsx`, but formulas are unevaluated text — `=SUM(B2:B9)` exists as a string, not a computed result.

2. **`recalc.py` + LibreOffice** evaluates the formulas. LibreOffice runs headless (no GUI), recalculates every cell, writes computed values back to the file, and scans for errors (#REF!, #DIV/0!, #VALUE!, etc.). This step is necessary because `openpyxl` has no formula engine.

The result is a fully resolved `.xlsx` — formulas and their calculated values — that behaves identically to a hand-built model.

## Key Infrastructure

Two files from Anthropic's open-source [skills repo](https://github.com/anthropics/skills/tree/main/skills/xlsx):

| File | Role |
|------|------|
| [`SKILL.md`](https://github.com/anthropics/skills/blob/main/skills/xlsx/SKILL.md) | Excel best practices: formula construction, formatting standards, color coding, error prevention. Claude Code references this to produce correctly structured workbooks. |
| [`recalc.py`](https://github.com/anthropics/skills/blob/main/skills/xlsx/recalc.py) | Formula recalculation via LibreOffice headless mode. Evaluates all formulas, writes computed values, returns error diagnostics as JSON. |

Drop both into your project directory or reference them in `CLAUDE.md`.

## Why This Works

**Excel files are code.** Cell references are pointers. Formulas are functions. Named ranges are variables. Sheets are modules. Generating a spreadsheet programmatically is the same class of problem as generating any other structured output.

**Design is the hard part, not construction.** Deciding that revenue = `volume × ASP × utilization` with a 24-month ramp is the analytical work. Translating that into `=B15*B16*B17` across 20 columns with proper formatting is rote. This workflow keeps the analyst on the former and the agent on the latter.

**Context compounds.** The planning phase builds up model-specific context — assumptions, sources, strategic rationale — across multiple sessions. That context flows through the model approach document into the build phase. A thin spec produces a thin model.

## Beyond Excel: Presentations

Same principle applies to PowerPoint. A `.pptx` is the same kind of artifact, and `python-pptx` is the `openpyxl` equivalent.

The workflow:

1. **Define the deck spec** — slide count, layout per slide, content outline, visual style. Be specific: "Slide 3: two-column layout, left = bullet summary of TAM sizing, right = bar chart from model output, source citation in footer."

2. **Provide a slide template** — drop a branded `.pptx` into the project directory. Claude Code applies your existing slide masters, color palettes, and fonts rather than starting from defaults.

3. **Build with `python-pptx`** — Claude Code constructs each slide programmatically: titles, text boxes, tables, charts, images, speaker notes. No recalculation step needed, so the output is immediately usable.

A deck spec might look like:

```
Slide 1: Title slide — company name, subtitle, date
Slide 2: Executive summary — 3-4 bullet key takeaways
Slide 3: Market opportunity — TAM/SAM/SOM with bar chart
Slide 4: Revenue model summary — table pulled from Excel output
Slide 5: Unit economics — two-column: metrics left, waterfall chart right
Slide 6: Competitive positioning — 2x2 matrix
Slide 7: Financial projections — 5-year P&L summary table
Slide 8: Appendix — detailed assumptions, source citations
```

Pairs naturally with the Excel workflow — build the model first, then generate a deck that pulls key outputs from it.

```bash
pip install python-pptx
```

## Example Model

This repo includes a working example: a 5-year quarterly revenue model for a robotics startup (Arctura) — 11 sheets covering anchor partner revenue, expansion pipeline, cost forecasting, P&L, DCF valuation, and sensitivity analysis.

### Project Files

| File | Role in Workflow |
|------|-----------------|
| `CLAUDE.md` | Build specification — the model approach document that Claude Code reads on startup. Contains complete row maps, sheet structure, formula conventions, and cross-sheet reference patterns. |
| `MODEL_GUIDE.md` | Human-readable overview of the model's revenue streams, pricing, and assumptions. |
| `build_model.py` | The build script (~2,760 lines). Generates the entire 11-sheet `.xlsx` workbook from scratch using `openpyxl`. |
| `v3_sources_notes.json` | Source documentation data — loaded by the build script into the Assumptions sheet. |
| `recalc_win.py` | Windows-specific formula recalculation via LibreOffice (see `skills/recalc.py` for macOS/Linux). |
| `validate_*.py` | Post-build validation scripts that verify EBITDA calculations, non-GAAP metrics, and sensitivity ranges. |
| `skills/SKILL.md` | Excel creation best practices reference. |
| `skills/recalc.py` | Cross-platform formula recalculation script. |

### Build and Validate

```bash
pip install openpyxl
python build_model.py
python recalc_win.py arctura_5partner_gtm_model_v1.xlsx   # Windows
# or: python skills/recalc.py arctura_5partner_gtm_model_v1.xlsx  # macOS/Linux
```

Optional validation:

```bash
python validate_ebitda.py
python validate_nongaap.py
python validate_sensitivity.py
```

## Getting Started

1. Clone this repo
2. Install [LibreOffice](https://www.libreoffice.org/) (required for formula recalculation)
3. Install Python dependencies: `pip install openpyxl pandas`
4. Review `SKILL.md` for formatting conventions
5. Study the example: read `CLAUDE.md` to see how a model approach document translates into a build script, then review `build_model.py` to see the patterns
6. Start your own: create a `CLAUDE.md` describing your model's structure, drop in `skills/`, and build
