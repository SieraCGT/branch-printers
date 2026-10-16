# ==============================================================================
# SCRIPT 1: RIZZA BRANCH
# Save as: branch_listener_rizza.py
# ==============================================================================
import time
import textwrap
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import win32print

# ================= CONFIGURATION =================
SHEET_NAME = "Hardware_Store_Orders"
WORKSHEET_TAB = "Form Responses 1"
CREDENTIALS_FILE = r"C:\HardwarePrint\credentials.json"
LOCAL_BRANCH = "Rizza Branch"
CHECK_INTERVAL_SECONDS = 10
# =================================================

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

def generate_slip(row):
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

    lines = [
        sep_thick,
        "SIERA MADRE CONSUMER GOODS TRADING".center(width),
        "WAREHOUSE DISPATCH SLIP".center(width),
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
    ]

    for item_line in items_raw.splitlines():
        clean_item = item_line.strip()
        if not clean_item: continue
        wrapped = textwrap.wrap(clean_item, width=58) or [""]
        lines.append(f"{wrapped[0]:<60} |  [   ]  |  [   ] ")
        for cont_line in wrapped[1:]:
            lines.append(f"  {cont_line:<58} |         |        ")

    lines.append(sep_thin)
    if "C.O.D" in payment_term.upper() or "CASH ON DELIVERY" in payment_term.upper():
        lines.append(f"PAYMENT STATUS: CASH ON DELIVERY (C.O.D.)")
        lines.append(f">>> AMOUNT TO COLLECT FROM CUSTOMER: PHP {cod_amount} <<<")
    else:
        lines.append("PAYMENT STATUS: PREPAID / CHARGE (DO NOT COLLECT CASH)")

    lines.extend([
        sep_thin,
        "Checked By (Warehouse): _________________  Received By (Name): _________________",
        "Released By (Yard Lead): ________________  Customer Signature: _________________",
        "Driver / Plate No.    : _________________  Date/Time Received: _________________",
        sep_thick,
        "\n\n\n\n\n"
    ])
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
    print(f"\n[{LOCAL_BRANCH}] Listener Started. Connecting to Google Sheets...")
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
        try:
            records = sheet.get_all_records()
            pending_queue = []
            for row_idx, row in enumerate(records, start=2):
                fulfill_branch = str(row.get("Fulfillment Branch (Stock Source / Dispatch Yard)", "")).strip()
                status = str(row.get("Print_Status", "")).strip().upper()
                
                if fulfill_branch == LOCAL_BRANCH and status == "PENDING":
                    ts_raw = str(row.get("Timestamp", ""))
                    pending_queue.append((parse_timestamp(ts_raw), row_idx, row))

            pending_queue.sort(key=lambda x: (x[0], x[1]))

            for _, row_idx, row in pending_queue:
                doc_num = row.get("Document / Reference Number", "N/A")
                print(f"[{datetime.now().strftime('%X')}] Found order {doc_num}. Printing...")
                
                slip_content = generate_slip(row)
                send_to_printer(slip_content)
                sheet.update_cell(row_idx, 16, "PRINTED")
                
                print(f"[{datetime.now().strftime('%X')}] Success! Marked as PRINTED.\n")
                
        except Exception as e:
            print(f"[{datetime.now().strftime('%X')}] Loop Error: {e}. Retrying in 10s...")
            try: sheet = get_sheet()
            except Exception: pass
            
        time.sleep(CHECK_INTERVAL_SECONDS)

if __name__ == "__main__":
    run()