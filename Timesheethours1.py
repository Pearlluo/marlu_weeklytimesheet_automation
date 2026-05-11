import os
import json
import time
import random
import requests
import pandas as pd

from pathlib import Path
from datetime import datetime, timedelta
from base64 import b64encode
from typing import Optional, Dict, Any, List
from dotenv import load_dotenv


# =========================================================
# LOAD ENV
# =========================================================
load_dotenv()


# =========================================================
# OPMS CONFIG
# =========================================================
OPMS_TOKEN_URL = "https://auth.opms.com.au/api/authenticate/token"
OPMS_TS_URL = "https://api.opms.com.au/timesheets/entries"

OPMS_CLIENT_ID = os.getenv("OPMS_CLIENT_ID")
OPMS_CLIENT_SECRET = os.getenv("OPMS_CLIENT_SECRET")

DEFAULT_PAGE_SIZE = 25
DEFAULT_SLEEP_SECONDS = 0.3
DEFAULT_TIMEOUT_SECONDS = 60
DEFAULT_MAX_RETRIES = 8
DEFAULT_RETRY_STATUS = {502, 503, 504}

CHECKPOINT_FILE = Path("opms_timesheets_checkpoint.json")


# =========================================================
# SHAREPOINT CONFIG
# =========================================================
SHAREPOINT_TENANT_ID = os.getenv("SHAREPOINT_TENANT_ID")
SHAREPOINT_CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
SHAREPOINT_CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")
SHAREPOINT_HOST = os.getenv("SHAREPOINT_HOST")
SITE_NAME = os.getenv("SITE_NAME", "BMS")

ROSTER_LIST_NAME = os.getenv("LIST_NAME", "PPL-Rosters")

OPMS_FIELD = os.getenv("OPMS_FIELD", "OPMS")
ROSTER_DATE_FIELD = os.getenv("ROSTER_DATE_FIELD", "Date_x0020_From")
HOURS_FIELD = os.getenv("HOURS_FIELD", "Hours")


# =========================================================
# VALIDATION
# =========================================================
def validate_env():
    required = {
        "OPMS_CLIENT_ID": OPMS_CLIENT_ID,
        "OPMS_CLIENT_SECRET": OPMS_CLIENT_SECRET,
        "SHAREPOINT_TENANT_ID": SHAREPOINT_TENANT_ID,
        "SHAREPOINT_CLIENT_ID": SHAREPOINT_CLIENT_ID,
        "SHAREPOINT_CLIENT_SECRET": SHAREPOINT_CLIENT_SECRET,
        "SHAREPOINT_HOST": SHAREPOINT_HOST,
        "SITE_NAME": SITE_NAME,
    }

    missing = [k for k, v in required.items() if not v]

    if missing:
        raise ValueError(f"Missing environment variables: {', '.join(missing)}")


# =========================================================
# WEEK RANGE
# =========================================================
def get_last_full_week_range():
    today = datetime.now().date()
    this_monday = today - timedelta(days=today.weekday())
    last_monday = this_monday - timedelta(days=7)
    last_sunday = this_monday - timedelta(days=1)

    return last_monday, last_sunday


def get_modified_since_for_week(week_start):
    return f"{week_start.strftime('%Y-%m-%d')}T00:00:00Z"


# =========================================================
# CHECKPOINT
# =========================================================
def load_checkpoint(checkpoint_file: Path) -> Dict[str, Any]:
    if checkpoint_file.exists():
        try:
            return json.loads(checkpoint_file.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_checkpoint(
    checkpoint_file: Path,
    next_cursor: Optional[str],
    fetched_count: int,
    modified_since: str
):
    payload = {
        "next_cursor": next_cursor,
        "fetched_count": fetched_count,
        "modified_since": modified_since,
        "updated_at": datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
    }

    checkpoint_file.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# =========================================================
# OPMS AUTH
# =========================================================
def get_opms_token() -> str:
    auth = f"{OPMS_CLIENT_ID}:{OPMS_CLIENT_SECRET}"
    b64 = b64encode(auth.encode()).decode()

    headers = {
        "Authorization": f"Basic {b64}",
        "Content-Type": "application/x-www-form-urlencoded",
        "Accept": "application/json",
    }

    r = requests.post(
        OPMS_TOKEN_URL,
        headers=headers,
        data={"grant_type": "client_credentials"},
        timeout=DEFAULT_TIMEOUT_SECONDS,
    )
    r.raise_for_status()

    token = r.json().get("access_token")

    if not token:
        raise RuntimeError(f"OPMS token missing: {r.text[:500]}")

    return token


# =========================================================
# REQUEST RETRY
# =========================================================
def get_with_retry(
    url: str,
    headers: dict,
    params: dict,
    timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS,
    max_retries: int = DEFAULT_MAX_RETRIES,
    retry_status: Optional[set] = None,
) -> requests.Response:
    retry_status = retry_status or DEFAULT_RETRY_STATUS
    last_error = None

    for attempt in range(1, max_retries + 1):
        try:
            r = requests.get(
                url,
                headers=headers,
                params=params,
                timeout=timeout_seconds,
            )

            if r.status_code == 401:
                raise RuntimeError("401 Unauthorized")

            if r.status_code in retry_status:
                wait = min(60, 2 ** attempt) + random.uniform(0, 1.5)
                print(f"⚠️ HTTP {r.status_code}, retry {attempt}/{max_retries}, sleep {wait:.1f}s")
                time.sleep(wait)
                continue

            r.raise_for_status()
            return r

        except requests.exceptions.RequestException as e:
            last_error = e
            wait = min(60, 2 ** attempt) + random.uniform(0, 1.5)
            print(f"⚠️ Request error {attempt}/{max_retries}: {e}, sleep {wait:.1f}s")
            time.sleep(wait)

    raise RuntimeError(f"Failed after retries. Last error: {last_error}")


# =========================================================
# OPMS FETCH TIMESHEETS
# =========================================================
def fetch_all_opms_timesheets(
    modified_since: str,
    page_size: int = DEFAULT_PAGE_SIZE,
    checkpoint_file: Path = CHECKPOINT_FILE,
    use_checkpoint: bool = False,
) -> List[Dict[str, Any]]:
    token = get_opms_token()

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

    params = {
        "modified_since": modified_since,
        "page_size": page_size,
    }

    already = 0

    if use_checkpoint:
        checkpoint = load_checkpoint(checkpoint_file)

        resume_cursor = checkpoint.get("next_cursor")
        checkpoint_modified_since = checkpoint.get("modified_since")
        already = checkpoint.get("fetched_count", 0)

        if checkpoint_modified_since and checkpoint_modified_since != modified_since:
            print("🟡 Checkpoint modified_since different. Ignoring old checkpoint.")
        elif resume_cursor:
            params["after"] = resume_cursor
            print(f"🔁 Resume OPMS from cursor: {resume_cursor}")

    all_timesheets = []
    page = 0

    while True:
        try:
            r = get_with_retry(OPMS_TS_URL, headers, params)
        except RuntimeError as e:
            if "401 Unauthorized" in str(e):
                print("🔄 OPMS token expired. Refreshing...")
                token = get_opms_token()
                headers["Authorization"] = f"Bearer {token}"
                r = get_with_retry(OPMS_TS_URL, headers, params)
            else:
                raise

        data = r.json()
        batch = data.get("timesheets", []) or []

        all_timesheets.extend(batch)
        page += 1

        next_cursor = data.get("next_cursor")
        total = already + len(all_timesheets)

        print(f"✅ OPMS fetched page {page}, total so far: {total}")

        if use_checkpoint:
            save_checkpoint(
                checkpoint_file=checkpoint_file,
                next_cursor=next_cursor,
                fetched_count=total,
                modified_since=modified_since,
            )

        if not next_cursor:
            break

        params["after"] = next_cursor
        time.sleep(DEFAULT_SLEEP_SECONDS)

    return all_timesheets


# =========================================================
# OPMS TRANSFORM
# =========================================================
def flatten_opms_to_daily_hours(
    all_timesheets: List[Dict[str, Any]],
    week_start,
    week_end,
) -> pd.DataFrame:
    rows = []

    for ts in all_timesheets:
        ts_date = ts.get("date")

        if not ts_date:
            continue

        ts_date_obj = normalize_opms_date(ts_date)

        if ts_date_obj is None:
            continue

        if ts_date_obj < week_start or ts_date_obj > week_end:
            continue

        for entry in ts.get("entries", []) or []:
            emp = entry.get("employee", {}) or {}

            emp_id = emp.get("id")
            first_name = (emp.get("first_name") or "").strip()
            last_name = (emp.get("last_name") or "").strip()
            employee_name = f"{first_name} {last_name}".strip()

            hours = entry.get("value", 0) or 0

            rows.append({
                "date": ts_date,
                "employee_id": emp_id,
                "employee_name": employee_name,
                "hours": float(hours),
            })

    df = pd.DataFrame(rows)

    if df.empty:
        return pd.DataFrame(columns=["date", "employee_id", "employee_name", "hours"])

    df = df.groupby(
        ["date", "employee_id", "employee_name"],
        as_index=False,
    )["hours"].sum()

    return df


def get_opms_daily_hours(
    modified_since: str,
    week_start,
    week_end,
) -> pd.DataFrame:
    all_ts = fetch_all_opms_timesheets(
        modified_since=modified_since,
        page_size=DEFAULT_PAGE_SIZE,
        use_checkpoint=False,
    )

    print(f"✅ OPMS total timesheet objects fetched: {len(all_ts)}")

    df = flatten_opms_to_daily_hours(
        all_timesheets=all_ts,
        week_start=week_start,
        week_end=week_end,
    )

    print(f"✅ OPMS daily rows after aggregation for last full week: {len(df)}")

    return df


# =========================================================
# GRAPH AUTH
# =========================================================
def get_graph_token() -> str:
    url = f"https://login.microsoftonline.com/{SHAREPOINT_TENANT_ID}/oauth2/v2.0/token"

    data = {
        "client_id": SHAREPOINT_CLIENT_ID,
        "client_secret": SHAREPOINT_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }

    r = requests.post(url, data=data, timeout=60)
    r.raise_for_status()

    return r.json()["access_token"]


def graph_get(url: str, token: str, params: Optional[dict] = None) -> dict:
    headers = {"Authorization": f"Bearer {token}"}

    r = requests.get(url, headers=headers, params=params, timeout=60)
    r.raise_for_status()

    return r.json()


def graph_patch(url: str, token: str, payload: dict):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    r = requests.patch(url, headers=headers, json=payload, timeout=60)

    if r.status_code not in (200, 204):
        print("❌ PATCH failed")
        print("Status:", r.status_code)
        print("Body:", r.text)
        r.raise_for_status()

    return True


# =========================================================
# SHAREPOINT HELPERS
# =========================================================
def get_site_id(token: str) -> str:
    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOST}:/sites/{SITE_NAME}:"
    data = graph_get(url, token)
    return data["id"]


def get_list_id(token: str, site_id: str, list_name: str) -> str:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"

    data = graph_get(url, token)

    for item in data.get("value", []):
        if item.get("name") == list_name or item.get("displayName") == list_name:
            return item["id"]

    raise RuntimeError(f"SharePoint list not found: {list_name}")


def get_list_columns(token: str, site_id: str, list_id: str) -> Dict[str, str]:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/columns"
    data = graph_get(url, token)

    columns = {}

    print("\n=== SharePoint Columns ===")
    for col in data.get("value", []):
        display = col.get("displayName")
        internal = col.get("name")
        columns[display] = internal
        print(f"displayName={display} | internalName={internal}")

    return columns


def pick_field(fields: dict, possible_names: list):
    for name in possible_names:
        if name in fields:
            return fields.get(name)
    return None


def get_sharepoint_rosters(token: str, site_id: str, list_id: str) -> pd.DataFrame:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items"

    params = {
        "$expand": "fields",
        "$top": "5000",
    }

    rows = []
    page = 0

    while url:
        data = graph_get(url, token, params=params)
        page += 1

        for item in data.get("value", []):
            fields = item.get("fields", {}) or {}

            opms_value = pick_field(fields, [
                OPMS_FIELD,
                "OPMS",
                "OPMS0",
                "OPMSID",
                "OPMS_x0020_ID",
            ])

            roster_date_value = pick_field(fields, [
                ROSTER_DATE_FIELD,
                "Date_x0020_From",
                "RosterDate",
                "RosterDate0",
                "Roster_x0020_Date",
                "DateFrom",
                "Date_x0020_From0",
                "Date",
            ])

            hours_value = pick_field(fields, [
                HOURS_FIELD,
                "Hours",
                "Hours0",
            ])

            work_type_value = pick_field(fields, [
                "WorkType",
                "Work_x0020_Type",
                "Work_x0020_type",
            ])

            rows.append({
                "item_id": item.get("id"),
                "Title": fields.get("Title"),
                "FirstName": pick_field(fields, ["FirstName", "First_x0020_Name"]),
                "LastName": pick_field(fields, ["LastName", "Last_x0020_Name"]),
                "Position": fields.get("Position"),
                "OPMS": opms_value,
                "RosterDate": roster_date_value,
                "WorkType": work_type_value,
                "Project": fields.get("Project"),
                "Site": fields.get("Site"),
                "CurrentHours": hours_value,
                "_raw_fields": fields,
            })

        print(f"✅ Loaded SharePoint roster page {page}, rows so far: {len(rows)}")

        url = data.get("@odata.nextLink")
        params = None

    df = pd.DataFrame(rows)

    print(f"✅ SharePoint roster rows loaded: {len(df)}")

    if not df.empty:
        print("\n=== SharePoint Roster Sample ===")
        print(
            df[
                ["item_id", "Title", "OPMS", "RosterDate", "WorkType", "CurrentHours"]
            ].head(20).to_string(index=False)
        )

        print("\nRoster OPMS null:", df["OPMS"].isna().sum())
        print("RosterDate null:", df["RosterDate"].isna().sum())

        print("\n=== RAW FIELD SAMPLE FIRST ROW ===")
        raw = df.iloc[0]["_raw_fields"]
        for k, v in raw.items():
            if "Date" in k or "OPMS" in k or "Hours" in k:
                print(f"{k}: {v}")

    return df


def update_roster_hours(
    token: str,
    site_id: str,
    list_id: str,
    item_id: str,
    hours: float,
):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"

    payload = {
        HOURS_FIELD: float(hours)
    }

    return graph_patch(url, token, payload)


# =========================================================
# DATE + KEY NORMALISATION
# =========================================================
def normalize_roster_date(value):
    if value is None or pd.isna(value):
        return None

    text_value = str(value).strip()

    if not text_value or text_value.lower() == "nan":
        return None

    # SharePoint / Graph datetime, usually UTC.
    # Example: 2026-05-01T16:00:00Z -> Australia/Perth date 2026-05-02
    if "T" in text_value:
        dt = pd.to_datetime(text_value, utc=True, errors="coerce")

        if pd.isna(dt):
            print("❌ SharePoint date parse failed:", text_value)
            return None

        return dt.tz_convert("Australia/Perth").date()

    # Plain SharePoint date
    dt = pd.to_datetime(text_value, errors="coerce")

    if pd.isna(dt):
        print("❌ SharePoint plain date parse failed:", text_value)
        return None

    return dt.date()


def normalize_opms_date(value):
    if value is None or pd.isna(value):
        return None

    text_value = str(value).strip()

    if not text_value or text_value.lower() == "nan":
        return None

    # OPMS usually returns plain date, e.g. 2026-05-02.
    # If it ever returns datetime, convert it safely to Perth date.
    if "T" in text_value:
        dt = pd.to_datetime(text_value, utc=True, errors="coerce")

        if pd.isna(dt):
            print("❌ OPMS datetime parse failed:", text_value)
            return None

        return dt.tz_convert("Australia/Perth").date()

    dt = pd.to_datetime(text_value, errors="coerce")

    if pd.isna(dt):
        print("❌ OPMS date parse failed:", text_value)
        return None

    return dt.date()


def normalize_opms_key(value):
    if value is None or pd.isna(value):
        return None

    text = str(value).strip()

    if not text or text.lower() == "nan":
        return None

    try:
        return str(int(float(text)))
    except Exception:
        return text


# =========================================================
# MATCH
# =========================================================
def build_update_dataframe(
    roster_df: pd.DataFrame,
    opms_df: pd.DataFrame,
    week_start,
    week_end,
) -> pd.DataFrame:
    roster = roster_df.copy()
    opms = opms_df.copy()

    roster["OPMS_key"] = roster["OPMS"].apply(normalize_opms_key)
    opms["OPMS_key"] = opms["employee_id"].apply(normalize_opms_key)

    roster["Date_key"] = roster["RosterDate"].apply(normalize_roster_date)
    opms["Date_key"] = opms["date"].apply(normalize_opms_date)

    print("\n=== DEBUG BEFORE WEEK FILTER ===")
    print("Roster rows before filter:", len(roster))
    print("OPMS rows before filter:", len(opms))
    print(f"Target week: {week_start} to {week_end}")

    roster = roster[
        roster["OPMS_key"].notna()
        & roster["Date_key"].notna()
        & (roster["Date_key"] >= week_start)
        & (roster["Date_key"] <= week_end)
    ].copy()

    opms = opms[
        opms["OPMS_key"].notna()
        & opms["Date_key"].notna()
        & (opms["Date_key"] >= week_start)
        & (opms["Date_key"] <= week_end)
    ].copy()

    print("\n=== DEBUG AFTER WEEK FILTER ===")
    print("Roster rows in target week:", len(roster))
    print("OPMS rows in target week:", len(opms))

    if not roster.empty:
        print("\nRoster sample keys:")
        print(
            roster[
                ["item_id", "Title", "OPMS", "OPMS_key", "RosterDate", "Date_key", "CurrentHours"]
            ].head(30).to_string(index=False)
        )

    if not opms.empty:
        print("\nOPMS sample keys:")
        print(
            opms[
                ["date", "employee_id", "OPMS_key", "employee_name", "hours", "Date_key"]
            ].head(30).to_string(index=False)
        )

    if roster.empty:
        print("❌ No SharePoint roster rows found for last full week.")
        return pd.DataFrame()

    if opms.empty:
        print("⚠️ No OPMS timesheet rows found for last full week.")
        print("⚠️ Stop here to avoid replacing roster Hours with 0 incorrectly.")
        return pd.DataFrame()

    opms_sum = (
        opms.groupby(["OPMS_key", "Date_key"], as_index=False)["hours"]
        .sum()
        .rename(columns={"hours": "MatchedHours"})
    )

    print("\nOPMS grouped rows:", len(opms_sum))

    merged = roster.merge(
        opms_sum,
        on=["OPMS_key", "Date_key"],
        how="left",
    )

    # IMPORTANT:
    # Keep original match status BEFORE filling NaN.
    # If there is no OPMS match, do NOT update SharePoint Hours to 0.
    merged["HasOpmsMatch"] = merged["MatchedHours"].notna()

    merged["MatchedHours"] = pd.to_numeric(
        merged["MatchedHours"],
        errors="coerce"
    ).fillna(0)

    merged["CurrentHours_Number"] = pd.to_numeric(
        merged["CurrentHours"],
        errors="coerce"
    )

    def needs_update(row):
        if not row.get("HasOpmsMatch"):
            return False

        current = row.get("CurrentHours_Number")
        new = float(row.get("MatchedHours") or 0)

        if pd.isna(current):
            return True

        return abs(float(current) - new) > 0.001

    merged["NeedsUpdate"] = merged.apply(needs_update, axis=1)

    matched_count = merged["HasOpmsMatch"].sum()
    no_match_count = (~merged["HasOpmsMatch"]).sum()
    update_count = merged["NeedsUpdate"].sum()

    print(f"\n✅ Rows with OPMS match inside last week: {matched_count}")
    print(f"⚠️ Rows with no OPMS match inside last week, will NOT be updated: {no_match_count}")
    print(f"✍️ Rows needing Hours replacement: {update_count}")

    print("\n=== ROWS NEEDING UPDATE ===")
    update_preview_cols = [
        "item_id",
        "Title",
        "OPMS",
        "OPMS_key",
        "RosterDate",
        "Date_key",
        "WorkType",
        "CurrentHours",
        "MatchedHours",
        "HasOpmsMatch",
        "NeedsUpdate",
    ]

    existing_update_preview_cols = [
        c for c in update_preview_cols if c in merged.columns
    ]

    update_preview = merged[merged["NeedsUpdate"] == True]

    if update_preview.empty:
        print("No rows need update.")
    else:
        print(
            update_preview[existing_update_preview_cols]
            .to_string(index=False)
        )

    return merged


# =========================================================
# AUTO UPDATE PIPELINE
# =========================================================
def run_weekly_replace_roster_hours():
    validate_env()

    week_start, week_end = get_last_full_week_range()
    modified_since = get_modified_since_for_week(week_start)

    print("\n======================================")
    print("START WEEKLY ROSTER HOURS REPLACEMENT")
    print("======================================")

    print(f"Roster list: {ROSTER_LIST_NAME}")
    print(f"OPMS field: {OPMS_FIELD}")
    print(f"RosterDate field: {ROSTER_DATE_FIELD}")
    print(f"Hours field: {HOURS_FIELD}")
    print(f"Target week: {week_start} to {week_end}")
    print(f"OPMS modified_since: {modified_since}")
    print("Mode: replace SharePoint Hours with OPMS matched hours")
    print("Export: disabled")
    print("Dry run: disabled")

    print("\n📥 Step 1: Pull OPMS timesheet data")
    opms_df = get_opms_daily_hours(
        modified_since=modified_since,
        week_start=week_start,
        week_end=week_end,
    )

    if opms_df.empty:
        print("⚠️ No OPMS timesheet data found for last full week. Stop.")
        return

    print("\n🔐 Step 2: Connect SharePoint")
    graph_token = get_graph_token()
    site_id = get_site_id(graph_token)
    list_id = get_list_id(graph_token, site_id, ROSTER_LIST_NAME)

    print(f"✅ Site ID: {site_id}")
    print(f"✅ List ID: {list_id}")

    get_list_columns(graph_token, site_id, list_id)

    print("\n📥 Step 3: Pull PPL-Rosters")
    roster_df = get_sharepoint_rosters(graph_token, site_id, list_id)

    if roster_df.empty:
        print("⚠️ No roster data found. Stop.")
        return

    print("\n🔗 Step 4: Match by PPL-Rosters.OPMS + PPL-Rosters.RosterDate")
    merged = build_update_dataframe(
        roster_df=roster_df,
        opms_df=opms_df,
        week_start=week_start,
        week_end=week_end,
    )

    if merged.empty:
        print("\n⚠️ Nothing to update because merged result is empty.")
        return

    preview_cols = [
        "item_id",
        "Title",
        "OPMS",
        "OPMS_key",
        "RosterDate",
        "Date_key",
        "WorkType",
        "CurrentHours",
        "MatchedHours",
        "HasOpmsMatch",
        "NeedsUpdate",
    ]

    existing_preview_cols = [c for c in preview_cols if c in merged.columns]

    print("\n=== Preview first 50 rows - console only, no Excel export ===")
    print(merged[existing_preview_cols].head(50).to_string(index=False))

    update_df = merged[merged["NeedsUpdate"]].copy()

    print(f"\n🔎 Total roster rows in last full week: {len(merged)}")
    print(f"✍️ Rows to replace Hours: {len(update_df)}")

    if update_df.empty:
        print("✅ No rows need update. Done.")
        return

    print("\n✍️ Step 5: Replace SharePoint Hours")

    updated = 0
    failed = 0

    for _, row in update_df.iterrows():
        try:
            update_roster_hours(
                token=graph_token,
                site_id=site_id,
                list_id=list_id,
                item_id=row["item_id"],
                hours=float(row["MatchedHours"]),
            )

            updated += 1

            if updated % 50 == 0:
                print(f"✅ Updated {updated}/{len(update_df)}")

            time.sleep(0.05)

        except Exception as e:
            failed += 1
            print(
                f"❌ Failed item_id={row.get('item_id')} "
                f"Title={row.get('Title')} "
                f"OPMS={row.get('OPMS')} "
                f"RosterDate={row.get('RosterDate')} "
                f"CurrentHours={row.get('CurrentHours')} "
                f"MatchedHours={row.get('MatchedHours')} "
                f"Error={e}"
            )

    print("\n======================================")
    print("DONE WEEKLY ROSTER HOURS REPLACEMENT")
    print("======================================")
    print(f"Target week: {week_start} to {week_end}")
    print(f"✅ Updated: {updated}")
    print(f"❌ Failed: {failed}")


# =========================================================
# SCRIPT ENTRY
# =========================================================
if __name__ == "__main__":
    run_weekly_replace_roster_hours()