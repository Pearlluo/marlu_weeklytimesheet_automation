import os
import requests
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from dotenv import load_dotenv

load_dotenv()

# ======================================================
# ENV
# ======================================================

TENANT_ID = os.getenv("SHAREPOINT_TENANT_ID")
CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")

SHAREPOINT_HOST = os.getenv("SHAREPOINT_HOST")
SITE_NAME = os.getenv("SITE_NAME")

ROSTER_LIST_NAME = os.getenv("LIST_NAME")      # PPL-Rosters
PEOPLE_LIST_NAME = os.getenv("LIST_NAME1")     # PPL-People

PERTH_TZ = ZoneInfo("Australia/Perth")

# ======================================================
# AUTH
# ======================================================

def get_token():

    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }

    res = requests.post(url, data=data)
    res.raise_for_status()

    return res.json()["access_token"]


# ======================================================
# GRAPH HELPERS
# ======================================================

def graph_get(url, token):

    res = requests.get(
        url,
        headers={
            "Authorization": f"Bearer {token}"
        }
    )

    res.raise_for_status()

    return res.json()


def graph_patch(url, token, payload):

    res = requests.patch(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload
    )

    if not res.ok:

        print("PATCH FAILED")
        print(res.status_code)
        print(res.text)

    res.raise_for_status()


# ======================================================
# SITE / LIST
# ======================================================

def get_site_id(token):

    url = (
        f"https://graph.microsoft.com/v1.0/sites/"
        f"{SHAREPOINT_HOST}:/sites/{SITE_NAME}"
    )

    data = graph_get(url, token)

    return data["id"]


def get_list_id(token, site_id, list_name):

    url = (
        f"https://graph.microsoft.com/v1.0/sites/"
        f"{site_id}/lists"
    )

    data = graph_get(url, token)

    for lst in data.get("value", []):

        if lst.get("displayName") == list_name:
            return lst["id"]

    raise Exception(f"List not found: {list_name}")


# ======================================================
# FETCH ALL ITEMS
# ======================================================

def fetch_all_items(token, site_id, list_id):

    items = []

    url = (
        f"https://graph.microsoft.com/v1.0/sites/"
        f"{site_id}/lists/{list_id}/items"
        f"?expand=fields&$top=999"
    )

    while url:

        data = graph_get(url, token)

        items.extend(data.get("value", []))

        url = data.get("@odata.nextLink")

    return items


# ======================================================
# NORMALIZE OPMS
# ======================================================

def normalize_opms(value):

    if value is None:
        return ""

    try:

        value = str(value).strip()

        # SharePoint float fix
        # 6.0 -> 6
        if value.endswith(".0"):
            value = value[:-2]

        return value

    except:
        return ""


# ======================================================
# DATE
# ======================================================

def parse_date(value):

    if not value:
        return None

    try:
        return datetime.fromisoformat(
            value.replace("Z", "+00:00")
        )

    except:
        return None


def get_start_utc():

    # Perth 今天 -9 天
    perth_now = datetime.now(PERTH_TZ)

    perth_start = (
        perth_now.replace(
            hour=0,
            minute=0,
            second=0,
            microsecond=0
        )
        - timedelta(days=9)
    )

    return perth_start.astimezone(
        ZoneInfo("UTC")
    )


# ======================================================
# POSITION
# ======================================================

def is_z_position(position):

    return str(position or "").upper().startswith("Z.")


# ======================================================
# BUILD PEOPLE MAP
# ======================================================

def build_people_map(people_items):

    people_map = {}

    for item in people_items:

        fields = item.get("fields", {})

        opms = normalize_opms(
            fields.get("OPMS")
        )

        if not opms:
            continue

        supplier_id = (
            fields.get("SupplierLookupId")
            or fields.get("SupplierId")
        )

        if not supplier_id:
            continue

        try:
            supplier_id = int(supplier_id)

        except:
            continue

        # 避免 duplicate OPMS 覆盖
        if opms not in people_map:

            people_map[opms] = supplier_id

    return people_map


# ======================================================
# UPDATE SUPPLIER
# ======================================================

def update_supplier(
    token,
    site_id,
    roster_list_id,
    item_id,
    supplier_id
):

    url = (
        f"https://graph.microsoft.com/v1.0/sites/"
        f"{site_id}/lists/{roster_list_id}"
        f"/items/{item_id}/fields"
    )

    payload = {
        "SupplierLookupId": supplier_id
    }

    graph_patch(
        url,
        token,
        payload
    )


# ======================================================
# MAIN
# ======================================================

def main():

    print("\n==============================")
    print("START UPDATE SUPPLIER")
    print("==============================\n")

    token = get_token()

    print("✅ Token OK")

    site_id = get_site_id(token)

    print("✅ Site ID OK")

    roster_list_id = get_list_id(
        token,
        site_id,
        ROSTER_LIST_NAME
    )

    people_list_id = get_list_id(
        token,
        site_id,
        PEOPLE_LIST_NAME
    )

    print("✅ List IDs OK")

    # ==================================================
    # LOAD PEOPLE
    # ==================================================

    print("\n📥 Loading PPL-People...")

    people_items = fetch_all_items(
        token,
        site_id,
        people_list_id
    )

    print(f"✅ People rows: {len(people_items)}")

    # ==================================================
    # LOAD ROSTER
    # ==================================================

    print("\n📥 Loading PPL-Rosters...")

    roster_items = fetch_all_items(
        token,
        site_id,
        roster_list_id
    )

    print(f"✅ Roster rows: {len(roster_items)}")

    # ==================================================
    # BUILD MAP
    # ==================================================

    people_map = build_people_map(
        people_items
    )

    print(f"✅ Supplier map built: {len(people_map)}")

    # ==================================================
    # DATE FILTER
    # ==================================================

    start_utc = get_start_utc()

    print(f"\n📅 Perth last 9 days UTC start: {start_utc}")

    # ==================================================
    # UPDATE LOOP
    # ==================================================

    updated = 0
    skipped = 0
    failed = 0

    for item in roster_items:

        try:

            item_id = item.get("id")

            fields = item.get("fields", {})

            # already has supplier
            current_supplier = (
                fields.get("SupplierLookupId")
            )

            if current_supplier:
                skipped += 1
                continue

            # roster date
            roster_date = parse_date(
                fields.get("Date_x0020_From")
            )

            if not roster_date:
                skipped += 1
                continue

            # only last 9 days
            if roster_date < start_utc:
                skipped += 1
                continue

            # skip vehicles
            position = fields.get("Position", "")

            if is_z_position(position):
                skipped += 1
                continue

            # normalize opms
            opms = normalize_opms(
                fields.get("OPMS")
            )

            if not opms:
                skipped += 1
                continue

            supplier_id = people_map.get(opms)

            if not supplier_id:
                skipped += 1
                continue

            print(
                f"🔄 Updating "
                f"Item={item_id} "
                f"OPMS={opms} "
                f"Supplier={supplier_id}"
            )

            update_supplier(
                token,
                site_id,
                roster_list_id,
                item_id,
                supplier_id
            )

            updated += 1

        except Exception as e:

            failed += 1

            print(
                f"❌ FAILED "
                f"Item={item.get('id')} "
                f"Error={e}"
            )

    # ==================================================
    # DONE
    # ==================================================

    print("\n==============================")
    print("DONE")
    print("==============================")

    print(f"✅ Updated : {updated}")
    print(f"⏭️ Skipped : {skipped}")
    print(f"❌ Failed  : {failed}")


def run_update_supplier():

    main()


if __name__ == "__main__":
    main()