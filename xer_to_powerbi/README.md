# XER → Power BI Excel Exporter

Parses a **Primavera P6 `.XER`** file and produces a structured **5-sheet Excel workbook** ready for Power BI, focused on water network infrastructure projects (NWC / Saudi Arabia standards).

---

## Quick Start

```bash
pip install -r requirements.txt

# Generate sample data + export
python main.py --sample --output dashboard_data.xlsx

# Parse a real XER file
python main.py --input project.xer --output dashboard_data.xlsx
```

---

## Output Sheets

| Sheet | Rows | Description |
|-------|------|-------------|
| `Activities` | 1 per activity | Schedule + progress + EVM per task |
| `WBS` | 1 per WBS node | Hierarchy (3 levels) + planned/actual cost |
| `Resources` | 1 per resource assignment | Quantities, costs, EV, CV |
| `Project_Info` | 1 row | Summary KPIs: SPI, CPI, EAC, VAC |
| `SCurve` | 1 per month | Cumulative planned vs actual cost/% |

---

## Key Calculations

- **Planned %** per activity: date-based progress vs `data_date`, clamped 0–100
- **Weight** per activity: `activity_target_cost / BAC` (normalised, Σ = 1)
- **Overall Planned %**: `Σ(planned_pct × weight)`
- **Overall Actual %**: `Σ(phys_complete_pct × weight)`
- **EV** = `Σ(target_cost × phys_complete_pct / 100)`
- **SPI** = `EV / PV` · **CPI** = `EV / AC` · **EAC** = `BAC / CPI`

---

## File Structure

```
xer_to_powerbi/
├── main.py              ← CLI entry point
├── xer_parser.py        ← XER → raw DataFrames
├── data_model.py        ← business logic & EVM calculations
├── excel_exporter.py    ← 5-sheet Excel writer with formatting
├── sample_generator.py  ← realistic 91-activity sample XER
└── requirements.txt
```

---

## Sample Project

- **Al-Narjis Water Network – Phase 2**, Riyadh
- 18-month schedule (Oct 2024 → Apr 2026)
- 8 Zones · 5 Disciplines · ~91 activities · 8 material resources
- Project is **100% planned** / **~52% actual** at data date (delayed)
- SPI ≈ 0.53 · CPI ≈ 1.00 · BAC ≈ 25.3M SAR
