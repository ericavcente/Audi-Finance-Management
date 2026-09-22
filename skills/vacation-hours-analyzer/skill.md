# Vacation Hours Analyzer

Analyzes a vacation/leave report and calculates how many **real working hours** each person worked in a given period, excluding weekends, public holidays, PTO, and vacation days.

## When to invoke
- User says "analyze vacation report", "calculate working hours", "how many hours did they work"
- User provides a leave/absence spreadsheet and wants net worked hours

## Script location
`scripts/vacation_analyzer.py` in the `ericavcente/Audi-Finance-Management` repo.
Clone or pull the repo before running if not local:
```
gh repo clone ericavcente/Audi-Finance-Management
cd Audi-Finance-Management
pip install pandas openpyxl
```

## Input file format

The script accepts **CSV or Excel (.xlsx)** with either layout:

### Option A: Periods format (one row per leave block)
| employee | start_date | end_date | leave_type |
|---|---|---|---|
| Ana Costa | 2026-09-01 | 2026-09-03 | Vacation |
| Bruno Lima | 2026-09-10 | 2026-09-10 | PTO |

### Option B: Daily format (one row per absent day)
| employee | date | leave_type |
|---|---|---|
| Ana Costa | 2026-09-01 | Vacation |
| Ana Costa | 2026-09-02 | Vacation |

Column headers are flexible: the script auto-detects common Portuguese and English variants (nome, colaborador, inicio, fim, data, etc.).

## How to run

### Basic (terminal table)
```bash
python scripts/vacation_analyzer.py leave_report.csv \
  --start 2026-09-01 \
  --end 2026-09-30
```

### With HTML report
```bash
python scripts/vacation_analyzer.py leave_report.xlsx \
  --start 2026-09-01 \
  --end 2026-09-30 \
  --html working_hours_sep2026.html
```

### Full options
```bash
python scripts/vacation_analyzer.py REPORT_FILE \
  --start YYYY-MM-DD \
  --end YYYY-MM-DD \
  --country BR        # BR (default) or DE for German/Bavarian holidays
  --hours 8           # working hours per day (default: 8)
  --format table      # table | json | csv
  --html OUTPUT.html  # optional HTML report
```

## What it calculates per employee

1. **Total working days** in the period (weekdays minus public holidays)
2. **Absent days** = vacation + PTO + any leave day that falls on a working day
3. **Net worked days** = total working days - absent days
4. **Net worked hours** = net worked days x hours/day
5. **Utilization %** = net worked days / total working days

## Public holidays included

**Brazil (--country BR, default):** Ano Novo, Carnaval, Sexta-feira Santa, Tiradentes, Dia do Trabalho, Corpus Christi, Independencia, N.S. Aparecida, Finados, Proclamacao da Republica, Consciencia Negra, Natal. Covers 2025-2026.

**Germany/Bavaria (--country DE):** Neujahr, Heilige Drei Koenige, Karfreitag, Ostermontag, Tag der Arbeit, Christi Himmelfahrt, Pfingstmontag, Fronleichnam, Maria Himmelfahrt, Tag der Deutschen Einheit, Allerheiligen, 1. & 2. Weihnachtstag. Covers 2025-2026.

## Steps Claude should follow

1. **Get the report file.** Ask the user for the path to their vacation/leave report if not provided. Confirm the file exists.

2. **Get the analysis period.** Ask for start and end dates if not clear. Default to the current month if the user says "this month".

3. **Confirm the country** for public holidays. Default to BR. Ask only if the context is ambiguous (e.g., team is mixed or user mentions Germany).

4. **Run the script** with `--html` to generate a report AND `--format json` to capture data:
   ```bash
   python scripts/vacation_analyzer.py FILE \
     --start START --end END \
     --country COUNTRY \
     --html vacation_hours_report.html \
     --format json
   ```

5. **Parse the JSON output** and present a clean summary table in the terminal showing: employee name, days absent, days worked, hours worked, utilization %.

6. **Open or share the HTML report.** Tell the user the file path. Optionally publish it as a Claude artifact if they want to share it.

7. **Flag anomalies:**
   - Utilization below 60%: flag as high absence
   - Anyone with 0 absent days: confirm data was parsed correctly (might mean the employee wasn't in the report)

## Output example
```
Working Hours Analysis
Period : 2026-09-01 to 2026-09-30
Country: BR | Working days: 21 | Public holidays: 1 (Independencia, Sep 7)

Employee                     Absent  Worked Days    Hours  Util%
--------------------------------------------------------------------
Ana Costa                         3           18     144h   85.7%
Bruno Lima                        5           16     128h   76.2%
Carlos Souza                      0           21     168h  100.0%
```
