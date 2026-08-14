"""
Finance Hours Reconciliation
Runs every Wednesday — compares vacation leave with Allocation sheet and updates discrepancies.
"""

import os
import json
import base64
import tempfile
import requests
from datetime import datetime, timedelta, date
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── CONFIG ────────────────────────────────────────────────────────────────────

ALLOCATION_SHEET_ID  = "1LIbxO1gpQVX-QDRQXUfS497eX0aFT8t1YtxYCs4lx9k"
ALLOCATION_TAB       = "Allocation- changes PROHIBITED!!!"
ALLOCATION_SHEET_NUM = 2120159953
GITHUB_REPO          = "samciandt/Finance"
GITHUB_FILE          = "2026 Biz hours.xlsx"
TARGET_YEAR          = 2026
VALID_STATUS         = {"approved", "requested", "intention"}

TEAMS = [
    {
        "name":         "Shopping_Tools_Transactions",
        "alloc_key":    "Audi of America.Shopping_Tools_Transactions-High_Value",
        "leave_sheet":  "1JgQG1vZ6k0G9LxiwBlTkBVZnziqlnYt0W3E069dgmSE",
        "leave_tab":    "2026_LEAVE",
    },
    {
        "name":         "Charging",
        "alloc_key":    "Audi of America.Charging-High_Value",
        "leave_sheet":  "11WN88HrRW_GAyrAMY-3p3PHL2im0yu9BUjDX2eUFnCs",
        "leave_tab":    "2026_LEAVE",
    },
    {
        "name":         "MyAudi",
        "alloc_key":    "Audi of America.myAudi-High_Value",
        "leave_sheet":  "1rc5DelFVBhNNIboxzD52UyZ56ZMeAmICCWlTbjhBCJU",
        "leave_tab":    "2026_LEAVE",
    },
    {
        "name":         "Digital_Products",
        "alloc_key":    "Audi of America.eCommerce-Portfolio",
        "leave_sheet":  "1WqfsdAGI3YSBW6KrxaOOqEHL1gL20-CW9Us7cPl5jPE",
        "leave_tab":    "2026_LEAVE",
    },
    {
        "name":         "Prod_Support",
        "alloc_key":    "Audi of America.Prod_Support-High_Value",
        "leave_sheet":  "1gFe1uK21Abr7JiHdxLA_P-vSttPUbE85SztaGmWb4kY",
        "leave_tab":    "2026_LEAVE",
    },
    {
        "name":         "Enablement",
        "alloc_key":    "Audi of America.Leadership-High_Value",
        "leave_sheet":  "1DZKwHngg7AavyBQ_8SLQMso8Lc4Jpgxf8Fy7R9ReuIo",
        "leave_tab":    "2026_LEAVE",
    },
]

LOC_MAP = {
    "CPS": "CPS - Campinas", "SP": "SP - Sao Paulo",
    "BH": "BH - Belo Horizonte", "US": "US",
    "CO": "COL - Medellin", "COL": "COL - Medellin",
    "CA": "US", "SEA": "SEA - Seattle", "MNL": "MNL - Manila",
    "LIS": "LIS - Lisbon", "TOR": "TOR - Toronto",
}

MONTH_COL = {1:23,2:24,3:25,4:26,5:27,6:28,7:29,8:30,9:31,10:32,11:33,12:34}

ORANGE = {"red": 1.0, "green": round(109/255, 6), "blue": round(1/255, 6)}


# ── AUTH ──────────────────────────────────────────────────────────────────────

def get_sheets_service():
    key_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    creds = service_account.Credentials.from_service_account_info(
        json.loads(key_json),
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def get_github_token():
    return os.environ["GITHUB_PAT"]


def get_gchat_webhook():
    return os.environ["GCHAT_WEBHOOK"]


# ── BIZ DAYS ─────────────────────────────────────────────────────────────────

def download_biz_xlsx(github_token):
    import openpyxl
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE.replace(' ', '%20')}"
    resp = requests.get(url, headers={"Authorization": f"Bearer {github_token}"})
    resp.raise_for_status()
    content = resp.json()["content"].replace("\n", "").replace("\r", "").replace(" ", "")
    data = base64.b64decode(content)
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.write(data)
    tmp.close()
    return openpyxl.load_workbook(tmp.name, data_only=True)


def parse_biz_data(wb):
    biz_days = {}
    holidays = {}
    for loc in wb.sheetnames:
        ws = wb[loc]
        hols, in_summary = [], False
        biz_days[loc] = {}
        month_map = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
                     "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}
        for row in ws.iter_rows(values_only=True):
            if row[0] == "Monthly Business Hours Summary":
                in_summary = True
                continue
            if in_summary and row[0] in month_map:
                biz_days[loc][month_map[row[0]]] = row[3]
            if not in_summary and isinstance(row[0], str) and "/" in row[0]:
                try:
                    d = datetime.strptime(row[0], "%m/%d/%Y")
                    hols.append(d.strftime("%Y-%m-%d"))
                except:
                    pass
        holidays[loc] = hols
    return biz_days, holidays


# ── LEAVE CALCULATION ────────────────────────────────────────────────────────

def count_leave_days(start_str, end_str, month, hols):
    try:
        start = datetime.strptime(start_str.strip(), "%m/%d/%Y").date()
        end   = datetime.strptime(end_str.strip(),   "%m/%d/%Y").date()
    except:
        return 0
    count, cursor = 0, start
    while cursor <= end:
        if (cursor.year == TARGET_YEAR
                and cursor.month == month
                and cursor.weekday() < 5
                and cursor.strftime("%Y-%m-%d") not in hols):
            count += 1
        cursor += timedelta(days=1)
    return count


def calc_person_hours(name, loc, leave_rows, biz_days, holidays, months, factor=1.0):
    hols = holidays.get(loc, [])
    result = {}
    for m in months:
        bd = biz_days.get(loc, {}).get(m, 0)
        leave_d = 0
        for row in leave_rows[1:]:
            if len(row) < 8: continue
            emp, start, end, ltype, status = (
                row[0].strip(), row[3].strip(), row[4].strip(),
                row[5].strip(), row[7].strip().lower()
            )
            if emp.lower() != name.lower(): continue
            if status not in VALID_STATUS: continue
            if ltype.lower() == "compensation": continue
            leave_d += count_leave_days(start, end, m, hols)
        result[m] = round((bd - leave_d) * 8 * factor)
    return result


# ── SHEETS HELPERS ───────────────────────────────────────────────────────────

def read_sheet(service, sheet_id, tab):
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=tab
    ).execute()
    return result.get("values", [])


def read_allocation(service):
    return read_sheet(service, ALLOCATION_SHEET_ID, ALLOCATION_TAB)


def get_person_loc(name, leave_rows, default_loc):
    for row in leave_rows[1:]:
        if len(row) > 11 and row[0].strip().lower() == name.lower():
            code = row[11].strip().upper() if row[11] else ""
            if code in LOC_MAP:
                return LOC_MAP[code]
    return default_loc


# ── RECONCILE ────────────────────────────────────────────────────────────────

def reconcile_team(team, alloc_rows, leave_rows, biz_days, holidays, months):
    """Returns list of updates: {row, col_idx, col_letter, login, month, current, new}"""
    updates = []
    alloc_key = team["alloc_key"]

    for i, row in enumerate(alloc_rows):
        team_val = row[1] if len(row) > 1 else ""
        if alloc_key not in team_val:
            continue
        login = row[7].strip().lower() if len(row) > 7 else ""
        if not login:
            continue

        # Find this person's name in leave sheet by login match
        person_name = None
        for lr in leave_rows[1:]:
            if not lr:
                continue
            # Try to match by comparing against leave sheet PEOPLE tab would be ideal,
            # but we match login via allocation col H vs leave sheet name heuristically
            # We'll resolve name from the leave rows directly
            pass

        # Get person name from allocation col — not available directly.
        # We match leave rows by checking if any leave row's name produces this login
        # via the roster. For now skip rows where we can't resolve name.
        # Instead, collect names from leave sheet and cross-match.
        for lr in leave_rows[1:]:
            if not lr or len(lr) < 8:
                continue
            emp_name = lr[0].strip()
            # Use location from leave sheet col L if available
            loc_code = lr[11].strip().upper() if len(lr) > 11 and lr[11] else ""
            loc = LOC_MAP.get(loc_code, "US")

            # Check if this person matches this allocation login
            # We do this by computing hours and checking against current values
            # We'll resolve below using a pre-built name→login map per team
            pass

        # Simpler: build name→login map from leave rows × known roster
        # This is resolved in main() with the full roster per team
        break

    return updates


def col_letter(n):
    """0-based column index to letter (e.g. 30 → AE)"""
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def reconcile_and_update(service, team, alloc_rows, leave_rows, biz_days, holidays, months, roster):
    """
    roster: [{name, login, loc, factor}]
    Returns list of changes applied.
    """
    alloc_key = team["alloc_key"]
    changes = []
    value_data = []
    format_requests = []

    # Build login→person map
    login_map = {p["login"].lower(): p for p in roster}

    for i, row in enumerate(alloc_rows):
        team_val = row[1] if len(row) > 1 else ""
        if alloc_key not in team_val:
            continue
        login = row[7].strip().lower() if len(row) > 7 else ""
        if login not in login_map:
            continue

        person = login_map[login]
        row_num = i + 1  # 1-based

        hours_by_month = calc_person_hours(
            person["name"], person["loc"], leave_rows,
            biz_days, holidays, months, person.get("factor", 1.0)
        )

        for m in months:
            col_idx = MONTH_COL[m]
            current = int(row[col_idx]) if len(row) > col_idx and str(row[col_idx]).lstrip("-").isdigit() else 0
            calc_val = hours_by_month[m]
            if current == calc_val:
                continue

            cell_range = f"'{ALLOCATION_TAB}'!{col_letter(col_idx)}{row_num}"
            value_data.append({
                "range": cell_range,
                "values": [[calc_val]]
            })
            format_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": ALLOCATION_SHEET_NUM,
                        "startRowIndex": i,
                        "endRowIndex": i + 1,
                        "startColumnIndex": col_idx,
                        "endColumnIndex": col_idx + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {"bold": True, "foregroundColor": ORANGE}
                        }
                    },
                    "fields": "userEnteredFormat(textFormat)",
                }
            })
            changes.append({
                "team": team["name"], "login": login,
                "row": row_num, "month": m,
                "col": col_letter(col_idx),
                "old": current, "new": calc_val,
            })

    if value_data:
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=ALLOCATION_SHEET_ID,
            body={"valueInputOption": "RAW", "data": value_data}
        ).execute()

    if format_requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=ALLOCATION_SHEET_ID,
            body={"requests": format_requests}
        ).execute()

    return changes


# ── ROSTER BUILDER ────────────────────────────────────────────────────────────

def build_roster_from_leave(leave_rows, alloc_rows, alloc_key):
    """Build roster by cross-referencing leave sheet names with allocation logins."""
    # Get all logins for this team from allocation
    team_logins = {}
    for row in alloc_rows:
        if len(row) > 7 and alloc_key in (row[1] if len(row) > 1 else ""):
            login = row[7].strip().lower()
            role  = row[2].strip() if len(row) > 2 else ""
            if login and login not in team_logins:
                team_logins[login] = role

    # Get all people names + locations from leave sheet
    leave_people = {}
    for row in leave_rows[1:]:
        if not row or len(row) < 8:
            continue
        name = row[0].strip()
        loc_code = row[11].strip().upper() if len(row) > 11 and row[11] else ""
        loc = LOC_MAP.get(loc_code, None)
        if name and name not in leave_people and loc:
            leave_people[name] = loc

    # Match names to logins: exact login substring match or manual override
    # We use a fuzzy match: if login appears as substring of name (lowercased, no spaces/dots)
    roster = []
    matched_logins = set()

    for name, loc in leave_people.items():
        name_slug = name.lower().replace(" ", "").replace(".", "")
        best_login = None
        for login in team_logins:
            login_slug = login.replace(".", "").replace(" ", "")
            if login_slug in name_slug or name_slug.startswith(login_slug[:4]):
                best_login = login
                break
        if best_login and best_login not in matched_logins:
            matched_logins.add(best_login)
            roster.append({
                "name": name, "login": best_login, "loc": loc, "factor": 1.0
            })

    return roster


# ── GCHAT ────────────────────────────────────────────────────────────────────

def send_gchat(webhook_url, all_changes, run_date):
    if not all_changes:
        msg = f"✅ *Finance Hours — {run_date}*\nNenhuma divergência encontrada em todos os times."
    else:
        lines = [f"📊 *Finance Hours Reconciliation — {run_date}*\n"]
        by_team = {}
        for c in all_changes:
            by_team.setdefault(c["team"], []).append(c)

        month_names = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
                       7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}

        for team, changes in by_team.items():
            lines.append(f"*{team}* — {len(changes)} atualização(ões)")
            for c in changes:
                lines.append(
                    f"  • `{c['login']}` {month_names[c['month']]}: "
                    f"{c['old']}h → *{c['new']}h* (linha {c['row']})"
                )
            lines.append("")

        lines.append(f"_Total: {len(all_changes)} células atualizadas_")
        msg = "\n".join(lines)

    payload = {"text": msg}
    resp = requests.post(webhook_url, json=payload)
    resp.raise_for_status()


# ── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    run_date = date.today().strftime("%d/%m/%Y")
    print(f"[{run_date}] Starting finance hours reconciliation...")

    # Determine months to process: current month + remaining months of year
    today = date.today()
    months = list(range(today.month, 13))

    # Auth
    service      = get_sheets_service()
    github_token = get_github_token()
    webhook_url  = get_gchat_webhook()

    # Load biz days
    print("Loading business days from GitHub...")
    wb = download_biz_xlsx(github_token)
    biz_days, holidays = parse_biz_data(wb)

    # Load allocation
    print("Loading allocation sheet...")
    alloc_rows = read_allocation(service)

    all_changes = []

    for team in TEAMS:
        print(f"Processing {team['name']}...")
        leave_rows = read_sheet(service, team["leave_sheet"], team["leave_tab"])
        roster = build_roster_from_leave(leave_rows, alloc_rows, team["alloc_key"])

        if not roster:
            print(f"  No roster found for {team['name']}, skipping.")
            continue

        changes = reconcile_and_update(
            service, team, alloc_rows, leave_rows,
            biz_days, holidays, months, roster
        )
        print(f"  {len(changes)} updates applied.")
        all_changes.extend(changes)

    # Send GChat summary
    print("Sending GChat notification...")
    send_gchat(webhook_url, all_changes, run_date)
    print(f"Done. {len(all_changes)} total updates.")


if __name__ == "__main__":
    main()
