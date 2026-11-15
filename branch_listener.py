import os
import sys
import json
import time
import shutil
import textwrap
import py_compile
import urllib.request
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import win32print

# ================= OTA CONFIGURATION =================
CURRENT_VERSION = "1.0.1"
GITHUB_USER = "SieraCGT"
GITHUB_REPO = "branch-printers"
VERSION_URL = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/main/version.txt"
SCRIPT_URL = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/main/branch_listener.py"
OTA_CHECK_INTERVAL_SEC = 1800  # Check for updates every 30 minutes
# =====================================================

# ================= BRANCH CONFIGURATION ==============
BASE_DIR = r"C:\HardwarePrint"
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")
CONFIG_FILE = os.path.join(BASE_DIR, "branch_config.json")
SHEET_NAME = "Hardware_Store_Orders"
WORKSHEET_TAB = "Form Responses 1"
CHECK_INTERVAL_SECONDS = 10
# =====================================================

def load_local_branch():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("local_branch", "Rizza Branch").strip()
        except Exception as e:
            print(f"[WARNING] Error reading branch_config.json: {e}")
    return "Rizza Branch"

def check_for_ota_update():
    """Safely checks GitHub for a newer version, validates syntax, swaps file, and restarts."""
    try:
        req = urllib.request.Request(
            f"{VERSION_URL}?t={int(time.time())}",
            headers={'Cache-Control': 'no-cache', 'Pragma': 'no-cache'}
        )
        with urllib.request.urlopen(req, timeout=8) as res:
            remote_version = res.read().decode("utf-8").strip()

        if remote_version and remote_version != CURRENT_VERSION:
            print(f"[*] OTA Update found: v{CURRENT_VERSION} -> v{remote_version}. Downloading...")
            current_file = os.path.abspath(__file__)
            temp_file = current_file + ".tmp"
            backup_file = current_file + ".prev"

            # 1. Download new code to a temporary file
            script_req = urllib.request.Request(
                f"{SCRIPT_URL}?t={int(time.time())}",
                headers={'Cache-Control': 'no-cache', 'Pragma': 'no-cache'}
            )
            with urllib.request.urlopen(script_req, timeout=15) as s_res:
                new_code = s_res.read().decode("utf-8")

            with open(temp_file, "w", encoding="utf-8") as f:
                f.write(new_code)

            # 2. Verify Python syntax before replacing (prevents broken pushes from crashing PC)
            py_compile.compile(temp_file, doraise=True)

            # 3. Backup current working script and atomically swap in the new one
            shutil.copy2(current_file, backup_file)
            os.replace(temp_file, current_file)
            print(f"[*] Update v{remote_version} verified and applied! Restarting single process...")
            time.sleep(1)

            # 4. Replace current process in-place (never spawns a duplicate 2nd process)
            os.execv(sys.executable, [sys.executable] + sys.argv)

    except Exception as e:
        print(f"[!] OTA Check skipped/failed: {e}")
        temp_file = os.path.abspath(__file__) + ".tmp"
        if os.path.exists(temp_file):
            try:
                os.remove(temp_file)
            except Exception:
                pass

def get_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    return client.open(SHEET_NAME).worksheet(WORKSHEET_TAB)

def parse_timestamp(ts_str):
    for fmt in ("%m/%d/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S"):
        try:
            return datetime.strptime(ts_str.strip(), fmt)
        except ValueError:
            continue
    return datetime.min

def generate_slip(row, copy_type=""):
    order_type = str(row.get("Order Type", ""))

    if order_type == "Log Report":
        log_text = str(row.get("Items for Dispatch", ""))
        return log_text + "\n\nEnd of Report.\n\n\n\n\n"

    width = 80
    sep_thick = "=" * width
    sep_thin = "-" * width
    order_id = str(row.get("Order_ID", "N/A"))
    timestamp = str(row.get("Timestamp", "N/A"))
    doc_type = str(row.get("Document Type", "N/A"))
    doc_num = str(row.get("Document / Reference Number", "N/A"))
    origin = str(row.get("Origin Branch (Issuing Location)", "N/A"))
    fulfill = str(row.get("Fulfillment Branch (Stock Source / Dispatch Yard)", "N/A"))
    target_date = str(row.get("Target Delivery Date", "N/A"))
    slot = str(row.get("Dispatch Slot / Order Size", "N/A"))
    customer = str(row.get("Customer / Project Name", "N/A"))
    address = str(row.get("Delivery Address & Contact Number", "N/A"))
    notes = str(row.get("Special Instructions / Landmarks", "None"))
    items_raw = str(row.get("Items for Dispatch", ""))
    payment_term = str(row.get("Payment Term (Is this Cash On Delivery?)", "Prepaid / Charge"))
    cod_amount = str(row.get("COD Amount to Collect (PHP)", "0.00"))

    is_stock = (doc_type.upper() == "STOCK REQUEST" or str(doc_num).upper().startswith("STK-"))
    slip_title = "INTERNAL BRANCH STOCK TRANSFER" if is_stock else "WAREHOUSE DISPATCH SLIP"
    sig_label = "Receiving Branch Sig:" if is_stock else "Customer Signature:"

    lines = [sep_thick]

    if copy_type:
        lines.append(f"*** {copy_type.upper()} COPY ***".center(width))

    lines.extend([
        "SIERA MADRE CONSUMER GOODS TRADING".center(width),
        slip_title.center(width),
        sep_thick,
        f"Order ID      : {order_id:<28} Date/Time : {timestamp}",
        f"Doc Type      : {doc_type:<28} Ref No.   : {doc_num}",
        f"Issued At     : {origin:<28} Yard      : {fulfill}",
        f"Order Type    : {order_type:<28} Slot      : {target_date} ({slot})",
        sep_thin,
        f"Customer/Proj : {customer}",
        f"Address/Tel   : {address}",
        f"Landmarks     : {notes}",
        sep_thin,
        f"{'QTY / UNIT / ITEM DESCRIPTION':<60} | {'CHECK':^7} | {'LOADED':^7}",
        sep_thin
    ])

    for item_line in items_raw.splitlines():
        clean_item = item_line.strip()
        if not clean_item:
            continue
        wrapped = textwrap.wrap(clean_item, width=58) or [""]
        lines.append(f"{wrapped[0]:<60} |  [   ]  |  [   ] ")
        for cont_line in wrapped[1:]:
            lines.append(f"  {cont_line:<58} |         |        ")

    lines.append(sep_thin)
    if is_stock:
        lines.append("TRANSFER STATUS: INTERNAL STOCK TRANSFER (NO CASH TO COLLECT)")
    elif "C.O.D" in payment_term.upper() or "CASH ON DELIVERY" in payment_term.upper():
        lines.append("PAYMENT STATUS: CASH ON DELIVERY (C.O.D.)")
        lines.append(f">>> AMOUNT TO COLLECT FROM CUSTOMER: PHP {cod_amount} <<<")
    else:
        lines.append("PAYMENT STATUS: PREPAID / CHARGE (DO NOT COLLECT CASH)")

    lines.extend([
        sep_thin,
        f"Checked By (Warehouse): _________________  Received By (Name): _________________",
        f"Released By (Yard Lead): ________________  {sig_label:<20} _________________",
        f"Driver / Plate No.    : _________________  Date/Time Received: _________________",
        sep_thick,
    ])

    if copy_type:
        lines.append(f"*** {copy_type.upper()} COPY ***".center(width))

    lines.append("\n\n\n\n\n")
    return "\r\n".join(lines)

def send_to_printer(raw_text):
    printer_name = win32print.GetDefaultPrinter()
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        hJob = win32print.StartDocPrinter(hPrinter, 1, ("Warehouse Dispatch Slip", None, "RAW"))
        try:
            win32print.StartPagePrinter(hPrinter)
            payload = (raw_text + "\x0c").encode("ascii", "replace")
            win32print.WritePrinter(hPrinter, payload)
            win32print.EndPagePrinter(hPrinter)
        finally:
            win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)

def run():
    # 1. Check for OTA updates immediately on boot
    check_for_ota_update()
    last_ota_check = time.time()

    local_branch = load_local_branch()
    print(f"\n[{local_branch}] Listener v{CURRENT_VERSION} Started. Connecting to Google Sheets...")
    try:
        print(f"Target Printer: {win32print.GetDefaultPrinter()}")
    except Exception:
        print("[WARNING] No default Windows printer found.")

    try:
        sheet = get_sheet()
        print("Connected successfully! Waiting for PENDING orders...\n")
    except Exception as e:
        print(f"[ERROR] Failed to connect: {e}")
        return

    while True:
        # 2. Periodic OTA check while idle between print cycles
        if time.time() - last_ota_check > OTA_CHECK_INTERVAL_SEC:
            check_for_ota_update()
            last_ota_check = time.time()

        try:
            records = sheet.get_all_records()
            pending_queue = []
            for row_idx, row in enumerate(records, start=2):
                fulfill_branch = str(row.get("Fulfillment Branch (Stock Source / Dispatch Yard)", "")).strip()
                status = str(row.get("Print_Status", "")).strip().upper()

                if fulfill_branch == local_branch and status == "PENDING":
                    ts_raw = str(row.get("Timestamp", ""))
                    pending_queue.append((parse_timestamp(ts_raw), row_idx, row))

            pending_queue.sort(key=lambda x: (x[0], x[1]))

            for _, row_idx, row in pending_queue:
                doc_num = str(row.get("Document / Reference Number", "N/A"))
                doc_type = str(row.get("Document Type", "N/A"))
                order_type = str(row.get("Order Type", ""))
                is_stock = (doc_type.upper() == "STOCK REQUEST" or doc_num.upper().startswith("STK-"))

                if order_type == "Log Report":
                    print(f"[{datetime.now().strftime('%X')}] Printing EOD Log Report...")
                    send_to_printer(generate_slip(row))
                else:
                    print(f"[{datetime.now().strftime('%X')}] Found {doc_num}. Printing 2 copies...")
                    copy_1 = "DISPATCH YARD" if is_stock else "DELIVERY"
                    copy_2 = "RECEIVING BRANCH" if is_stock else "CUSTOMER"

                    send_to_printer(generate_slip(row, copy_type=copy_1))
                    time.sleep(1.5)
                    send_to_printer(generate_slip(row, copy_type=copy_2))

                sheet.update_cell(row_idx, 16, "PRINTED")
                print(f"[{datetime.now().strftime('%X')}] Success! Marked as PRINTED.\n")

        except Exception as e:
            print(f"[{datetime.now().strftime('%X')}] Loop Error: {e}. Retrying in 10s...")
            try:
                sheet = get_sheet()
            except Exception:
                pass

        time.sleep(CHECK_INTERVAL_SECONDS)

if __name__ == "__main__":
    run()
