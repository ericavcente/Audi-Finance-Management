#!/usr/bin/env python3
"""
Vacation Hours Analyzer
Calculates net working hours per employee by removing weekends, public holidays,
and leave days (vacation, PTO, sick leave, etc.) from a given period.

Usage:
    python vacation_analyzer.py report.csv --start 2026-01-01 --end 2026-01-31
    python vacation_analyzer.py report.xlsx --start 2026-09-01 --end 2026-09-30 --country DE
    python vacation_analyzer.py report.csv --start 2026-01-01 --end 2026-03-31 --html report.html
"""

import sys
import argparse
import json
import base64
from datetime import date, timedelta, datetime
from typing import Optional

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas is required. Run: pip install pandas openpyxl")
    sys.exit(1)


# ─────────────────────────────────────────────────────────────────
# Public Holidays
# ─────────────────────────────────────────────────────────────────

HOLIDAYS = {
    "BR": {
        # 2025 - Brazil national holidays
        date(2025, 1, 1): "Ano Novo",
        date(2025, 3, 3): "Carnaval (segunda)",
        date(2025, 3, 4): "Carnaval (terca)",
        date(2025, 4, 18): "Sexta-feira Santa",
        date(2025, 4, 21): "Tiradentes",
        date(2025, 5, 1): "Dia do Trabalho",
        date(2025, 6, 19): "Corpus Christi",
        date(2025, 9, 7): "Independencia",
        date(2025, 10, 12): "N.S. Aparecida",
        date(2025, 11, 2): "Finados",
        date(2025, 11, 15): "Proclamacao da Republica",
        date(2025, 11, 20): "Consciencia Negra",
        date(2025, 12, 25): "Natal",
        # 2026 - Brazil national holidays
        date(2026, 1, 1): "Ano Novo",
        date(2026, 2, 16): "Carnaval (segunda)",
        date(2026, 2, 17): "Carnaval (terca)",
        date(2026, 4, 3): "Sexta-feira Santa",
        date(2026, 4, 21): "Tiradentes",
        date(2026, 5, 1): "Dia do Trabalho",
        date(2026, 6, 4): "Corpus Christi",
        date(2026, 9, 7): "Independencia",
        date(2026, 10, 12): "N.S. Aparecida",
        date(2026, 11, 2): "Finados",
        date(2026, 11, 15): "Proclamacao da Republica",
        date(2026, 11, 20): "Consciencia Negra",
        date(2026, 12, 25): "Natal",
    },
    "DE": {
        # 2025 - Germany (Bavaria) public holidays
        date(2025, 1, 1): "Neujahr",
        date(2025, 1, 6): "Heilige Drei Koenige",
        date(2025, 4, 18): "Karfreitag",
        date(2025, 4, 21): "Ostermontag",
        date(2025, 5, 1): "Tag der Arbeit",
        date(2025, 5, 29): "Christi Himmelfahrt",
        date(2025, 6, 9): "Pfingstmontag",
        date(2025, 6, 19): "Fronleichnam",
        date(2025, 8, 15): "Maria Himmelfahrt",
        date(2025, 10, 3): "Tag der Deutschen Einheit",
        date(2025, 11, 1): "Allerheiligen",
        date(2025, 12, 25): "1. Weihnachtstag",
        date(2025, 12, 26): "2. Weihnachtstag",
        # 2026 - Germany (Bavaria) public holidays
        date(2026, 1, 1): "Neujahr",
        date(2026, 1, 6): "Heilige Drei Koenige",
        date(2026, 4, 3): "Karfreitag",
        date(2026, 4, 6): "Ostermontag",
        date(2026, 5, 1): "Tag der Arbeit",
        date(2026, 5, 14): "Christi Himmelfahrt",
        date(2026, 5, 25): "Pfingstmontag",
        date(2026, 6, 4): "Fronleichnam",
        date(2026, 8, 15): "Maria Himmelfahrt",
        date(2026, 10, 3): "Tag der Deutschen Einheit",
        date(2026, 11, 1): "Allerheiligen",
        date(2026, 12, 25): "1. Weihnachtstag",
        date(2026, 12, 26): "2. Weihnachtstag",
    },
}


# ─────────────────────────────────────────────────────────────────
# Core Calculation Logic
# ─────────────────────────────────────────────────────────────────

def get_working_days(start: date, end: date, holidays: dict) -> set:
    days = set()
    current = start
    while current <= end:
        if current.weekday() < 5 and current not in holidays:
            days.add(current)
        current += timedelta(days=1)
    return days


def parse_date(val) -> Optional[date]:
    if val is None or (hasattr(val, '__class__') and val.__class__.__name__ == 'float' and str(val) == 'nan'):
        return None
    try:
        import math
        if isinstance(val, float) and math.isnan(val):
            return None
    except Exception:
        pass
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y', '%Y/%m/%d'):
        try:
            return datetime.strptime(str(val).strip(), fmt).date()
        except ValueError:
            continue
    return None


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {}
    for col in df.columns:
        lower = col.lower().strip()
        if mapping.get('employee') is None and any(
            k in lower for k in ['employee', 'name', 'nome', 'colaborador', 'person', 'member', 'membro']
        ):
            mapping[col] = 'employee'
        elif mapping.get('start_date') is None and any(
            k in lower for k in ['start', 'inicio', 'início', 'from', ' de ', 'begin', 'data_inicio', 'data inicio']
        ):
            mapping[col] = 'start_date'
        elif mapping.get('end_date') is None and any(
            k in lower for k in ['end', 'fim', 'to', 'ate', 'até', 'finish', 'data_fim', 'data fim']
        ):
            mapping[col] = 'end_date'
        elif mapping.get('date') is None and any(
            k in lower for k in ['date', 'data', 'dia', 'day']
        ):
            mapping[col] = 'date'
        elif mapping.get('leave_type') is None and any(
            k in lower for k in ['type', 'tipo', 'reason', 'motivo', 'leave', 'category', 'absence', 'ausencia']
        ):
            mapping[col] = 'leave_type'
    return df.rename(columns=mapping)


def detect_format(df: pd.DataFrame) -> str:
    cols = [c.lower() for c in df.columns]
    has_end = any('end' in c or 'fim' in c or 'ate' in c for c in cols)
    has_start = any('start' in c or 'inicio' in c or 'begin' in c for c in cols)
    if has_start and has_end:
        return 'periods'
    return 'daily'


def collect_leave_days(df: pd.DataFrame, fmt: str, period_start: date, period_end: date, working_days: set) -> dict:
    leave_by_person: dict = {}

    for _, row in df.iterrows():
        person = str(row.get('employee', '')).strip()
        if not person or person.lower() in ('nan', 'none', ''):
            continue

        if fmt == 'periods':
            s = parse_date(row.get('start_date'))
            e = parse_date(row.get('end_date', row.get('start_date')))
        else:
            d = parse_date(row.get('date'))
            s = e = d

        if not s or not e:
            continue

        s = max(s, period_start)
        e = min(e, period_end)
        if s > e:
            continue

        if person not in leave_by_person:
            leave_by_person[person] = set()

        current = s
        while current <= e:
            if current in working_days:
                leave_by_person[person].add(current)
            current += timedelta(days=1)

    return leave_by_person


def analyze(
    file_path: str,
    start_date: date,
    end_date: date,
    country: str = "BR",
    hours_per_day: int = 8,
) -> dict:
    holidays = HOLIDAYS.get(country.upper(), HOLIDAYS["BR"])
    working_days = get_working_days(start_date, end_date, holidays)
    holidays_in_period = {d for d in holidays if start_date <= d <= end_date}

    if file_path.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(file_path)
    else:
        df = pd.read_csv(file_path)

    df = normalize_columns(df)
    fmt = detect_format(df)
    leave_by_person = collect_leave_days(df, fmt, start_date, end_date, working_days)

    rows = []
    for person in sorted(leave_by_person.keys()):
        absent = len(leave_by_person[person])
        net_days = len(working_days) - absent
        rows.append({
            "employee": person,
            "total_working_days": len(working_days),
            "absent_days": absent,
            "net_worked_days": max(0, net_days),
            "net_worked_hours": max(0, net_days) * hours_per_day,
            "utilization_pct": round(max(0, net_days) / len(working_days) * 100, 1) if working_days else 0,
        })

    return {
        "period": {"start": str(start_date), "end": str(end_date)},
        "country": country.upper(),
        "total_working_days": len(working_days),
        "public_holidays_in_period": len(holidays_in_period),
        "holiday_names": {str(k): v for k, v in holidays.items() if start_date <= k <= end_date},
        "hours_per_day": hours_per_day,
        "employees": rows,
    }


# ─────────────────────────────────────────────────────────────────
# HTML Report Generator
# ─────────────────────────────────────────────────────────────────

def generate_html(data: dict) -> str:
    employees = data["employees"]
    period = data["period"]
    total_days = data["total_working_days"]
    holidays = data["public_holidays_in_period"]
    country = data["country"]
    holidays_named = data.get("holiday_names", {})

    holiday_rows = ""
    for d, name in sorted(holidays_named.items()):
        dt = datetime.strptime(d, "%Y-%m-%d")
        holiday_rows += f'<tr><td>{dt.strftime("%b %d, %Y")}</td><td>{name}</td></tr>\n'

    employee_rows = ""
    for emp in employees:
        util = emp["utilization_pct"]
        bar_color = "#2B6CB0" if util >= 80 else "#C05621" if util >= 60 else "#9B2C2C"
        employee_rows += f"""
        <tr>
          <td class="emp-name">{emp['employee']}</td>
          <td class="num">{emp['absent_days']}</td>
          <td class="num">{emp['net_worked_days']}</td>
          <td class="num strong">{emp['net_worked_hours']}h</td>
          <td class="bar-cell">
            <div class="bar-wrap">
              <div class="bar" style="width:{util}%; background:{bar_color}"></div>
              <span class="bar-label">{util}%</span>
            </div>
          </td>
        </tr>"""

    total_hours = sum(e["net_worked_hours"] for e in employees)
    avg_util = round(sum(e["utilization_pct"] for e in employees) / len(employees), 1) if employees else 0

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Working Hours Analysis</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  :root {{
    --bg: #F7F8FC;
    --surface: #FFFFFF;
    --surface-alt: #F0F2F8;
    --border: #DDE2EF;
    --text: #1A1F36;
    --text-muted: #6B7594;
    --accent: #2B6CB0;
    --accent-light: #EBF4FF;
    --success: #276749;
    --warn: #C05621;
    --mono: 'IBM Plex Mono', monospace;
    --sans: 'Inter', system-ui, sans-serif;
  }}
  @media (prefers-color-scheme: dark) {{
    :root:not([data-theme="light"]) {{
      --bg: #0F1117;
      --surface: #1A1F2E;
      --surface-alt: #242938;
      --border: #2E3550;
      --text: #E8EAF6;
      --text-muted: #8892B0;
      --accent: #63B3ED;
      --accent-light: #1A2744;
      --success: #68D391;
      --warn: #F6AD55;
      --mono: 'IBM Plex Mono', monospace;
    }}
  }}
  :root[data-theme="dark"] {{
    --bg: #0F1117;
    --surface: #1A1F2E;
    --surface-alt: #242938;
    --border: #2E3550;
    --text: #E8EAF6;
    --text-muted: #8892B0;
    --accent: #63B3ED;
    --accent-light: #1A2744;
    --success: #68D391;
    --warn: #F6AD55;
  }}
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: var(--sans);
    background: var(--bg);
    color: var(--text);
    font-size: 14px;
    line-height: 1.6;
    padding: 32px 24px;
  }}
  .page {{ max-width: 900px; margin: 0 auto; }}
  header {{ margin-bottom: 32px; }}
  .label {{
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 6px;
  }}
  h1 {{
    font-size: 26px;
    font-weight: 700;
    color: var(--text);
    text-wrap: balance;
    margin-bottom: 6px;
  }}
  .period-tag {{
    display: inline-block;
    font-family: var(--mono);
    font-size: 12px;
    background: var(--accent-light);
    color: var(--accent);
    padding: 3px 10px;
    border-radius: 4px;
  }}
  .kpi-row {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 16px;
    margin-bottom: 32px;
  }}
  .kpi {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 18px 20px;
  }}
  .kpi-value {{
    font-size: 28px;
    font-weight: 700;
    font-family: var(--mono);
    color: var(--text);
    line-height: 1;
    margin-bottom: 4px;
    font-variant-numeric: tabular-nums;
  }}
  .kpi-desc {{ font-size: 12px; color: var(--text-muted); }}
  .section-title {{
    font-size: 12px;
    font-weight: 600;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    color: var(--text-muted);
    margin-bottom: 12px;
    padding-bottom: 8px;
    border-bottom: 1px solid var(--border);
  }}
  .card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    overflow: hidden;
    margin-bottom: 24px;
  }}
  .card-header {{
    padding: 14px 20px;
    border-bottom: 1px solid var(--border);
    background: var(--surface-alt);
  }}
  table {{ width: 100%; border-collapse: collapse; }}
  th {{
    padding: 10px 16px;
    text-align: left;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.05em;
    text-transform: uppercase;
    color: var(--text-muted);
    background: var(--surface-alt);
    border-bottom: 1px solid var(--border);
  }}
  th.num, td.num {{ text-align: right; }}
  td {{
    padding: 11px 16px;
    border-bottom: 1px solid var(--border);
    font-size: 13px;
  }}
  tr:last-child td {{ border-bottom: none; }}
  tr:hover td {{ background: var(--surface-alt); }}
  .emp-name {{ font-weight: 500; }}
  .strong {{ font-weight: 600; font-family: var(--mono); font-variant-numeric: tabular-nums; }}
  .bar-cell {{ min-width: 140px; }}
  .bar-wrap {{
    display: flex;
    align-items: center;
    gap: 8px;
  }}
  .bar {{
    height: 6px;
    border-radius: 3px;
    flex-shrink: 0;
    transition: width 0.3s ease;
  }}
  .bar-label {{
    font-family: var(--mono);
    font-size: 11px;
    color: var(--text-muted);
    white-space: nowrap;
  }}
  .holidays-table td {{ font-size: 12px; }}
  .holidays-table td:first-child {{
    font-family: var(--mono);
    color: var(--text-muted);
    width: 140px;
  }}
  .footer {{
    margin-top: 24px;
    font-size: 11px;
    color: var(--text-muted);
    text-align: right;
  }}
  @media (max-width: 600px) {{
    body {{ padding: 20px 16px; }}
    h1 {{ font-size: 20px; }}
    .kpi-value {{ font-size: 22px; }}
    th, td {{ padding: 9px 10px; }}
  }}
</style>
</head>
<body>
<div class="page">
  <header>
    <div class="label">Finance Management</div>
    <h1>Working Hours Analysis</h1>
    <span class="period-tag">{period['start']} &rarr; {period['end']}</span>
  </header>

  <div class="kpi-row">
    <div class="kpi">
      <div class="kpi-value">{total_days}</div>
      <div class="kpi-desc">Working days in period</div>
    </div>
    <div class="kpi">
      <div class="kpi-value">{holidays}</div>
      <div class="kpi-desc">Public holidays ({country})</div>
    </div>
    <div class="kpi">
      <div class="kpi-value">{len(employees)}</div>
      <div class="kpi-desc">Employees analyzed</div>
    </div>
    <div class="kpi">
      <div class="kpi-value">{total_hours}h</div>
      <div class="kpi-desc">Total net worked hours</div>
    </div>
    <div class="kpi">
      <div class="kpi-value">{avg_util}%</div>
      <div class="kpi-desc">Avg utilization</div>
    </div>
  </div>

  <div class="section-title">Employee Breakdown</div>
  <div class="card">
    <table>
      <thead>
        <tr>
          <th>Employee</th>
          <th class="num">Days Absent</th>
          <th class="num">Days Worked</th>
          <th class="num">Hours Worked</th>
          <th>Utilization</th>
        </tr>
      </thead>
      <tbody>
        {employee_rows}
      </tbody>
    </table>
  </div>

  {"<div class='section-title'>Public Holidays in Period</div><div class='card'><table class='holidays-table'><thead><tr><th>Date</th><th>Holiday</th></tr></thead><tbody>" + holiday_rows + "</tbody></table></div>" if holiday_rows else ""}

  <div class="footer">
    Generated {datetime.now().strftime('%B %d, %Y')} &middot; {data['hours_per_day']}h/day assumption
  </div>
</div>
</body>
</html>"""


# ─────────────────────────────────────────────────────────────────
# CLI Entry Point
# ─────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Analyze vacation reports and calculate net working hours.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("file", help="Path to the vacation report (CSV or Excel)")
    parser.add_argument("--start", required=True, help="Analysis start date (YYYY-MM-DD)")
    parser.add_argument("--end", required=True, help="Analysis end date (YYYY-MM-DD)")
    parser.add_argument("--country", default="BR", choices=["BR", "DE"],
                        help="Country public holiday calendar to use (default: BR)")
    parser.add_argument("--hours", type=int, default=8,
                        help="Working hours per day (default: 8)")
    parser.add_argument("--format", choices=["table", "json", "csv"], default="table",
                        help="Output format (default: table)")
    parser.add_argument("--html", metavar="OUTPUT_FILE",
                        help="Generate HTML report and save to this path")

    args = parser.parse_args()

    try:
        start = datetime.strptime(args.start, "%Y-%m-%d").date()
        end = datetime.strptime(args.end, "%Y-%m-%d").date()
    except ValueError as e:
        print(f"ERROR: Invalid date format. Use YYYY-MM-DD. ({e})")
        sys.exit(1)

    if start > end:
        print("ERROR: Start date must be before end date.")
        sys.exit(1)

    data = analyze(args.file, start, end, args.country, args.hours)

    if args.html:
        html = generate_html(data)
        with open(args.html, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"HTML report saved to: {args.html}")

    if args.format == "json":
        print(json.dumps(data, indent=2))
        return

    if args.format == "csv":
        df = pd.DataFrame(data["employees"])
        print(df.to_csv(index=False))
        return

    # Default: table
    employees = data["employees"]
    total_days = data["total_working_days"]
    holidays = data["public_holidays_in_period"]

    print(f"\n{'='*72}")
    print(f"  Working Hours Analysis")
    print(f"  Period : {start} to {end}")
    print(f"  Country: {args.country} public holiday calendar")
    print(f"  Working days in period: {total_days}  |  Public holidays: {holidays}")
    print(f"{'='*72}")
    print(f"  {'Employee':<28} {'Absent':>8} {'Worked Days':>12} {'Hours':>8} {'Util%':>7}")
    print(f"  {'-'*67}")

    for emp in employees:
        print(
            f"  {emp['employee']:<28}"
            f" {emp['absent_days']:>8}"
            f" {emp['net_worked_days']:>12}"
            f" {emp['net_worked_hours']:>7}h"
            f" {emp['utilization_pct']:>6.1f}%"
        )

    print(f"  {'-'*67}")
    total_hours = sum(e["net_worked_hours"] for e in employees)
    avg_util = round(sum(e["utilization_pct"] for e in employees) / len(employees), 1) if employees else 0
    print(f"\n  Employees: {len(employees)}   Total hours: {total_hours}h   Avg utilization: {avg_util}%")
    print()


if __name__ == "__main__":
    main()
