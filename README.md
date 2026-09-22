# Audi Finance Management

Finance management tools and Claude Code skills for the Audi program.

## Skills

### `vacation-hours-analyzer`

Analyzes a vacation/leave report and calculates net working hours per employee for any period, excluding weekends, public holidays, PTO, and vacation days.

**Invoke with:** `/vacation-hours-analyzer` in Claude Code

**What it does:**
- Reads a CSV or Excel leave report (periods or daily format)
- Subtracts weekends, Brazilian or German public holidays, and leave days
- Returns hours worked per employee + utilization %
- Generates a clean HTML report

**Quick start:**
```bash
pip install pandas openpyxl

python scripts/vacation_analyzer.py sample/sample_vacation_report.csv \
  --start 2026-09-01 \
  --end 2026-09-30 \
  --html output.html
```

See [`skills/vacation-hours-analyzer/skill.md`](skills/vacation-hours-analyzer/skill.md) for full documentation.

## Scripts

| Script | Description |
|---|---|
| `scripts/vacation_analyzer.py` | Core analyzer: reads leave report, calculates net worked hours |

## Sample data

`sample/sample_vacation_report.csv` contains example data in the periods format to test the analyzer.

## Requirements

```bash
pip install pandas openpyxl
```

Python 3.8+
