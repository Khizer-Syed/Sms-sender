import openpyxl
import subprocess
import time
import os
from datetime import datetime

# === CONFIG ===
xlsx_file = "Contacts.xlsx"                   # XLSX file in the same folder
image_file = "image.png"                      # Relative image path
message_template = "Hello {name}, this is a hardcoded message with an image!"
log_file = "log.txt"                          # Log file in same folder
delay_seconds = 2                             # Delay between messages

# === Resolve relative paths ===
base_dir = os.path.dirname(os.path.abspath(__file__))
xlsx_path = os.path.join(base_dir, xlsx_file)
image_path = os.path.join(base_dir, image_file)
log_path = os.path.join(base_dir, log_file)

if not os.path.exists(xlsx_path):
    raise FileNotFoundError(f"XLSX file not found: {xlsx_path}")

if not os.path.exists(image_path):
    print(f"⚠️ Warning: Image not found at {image_path}. Will still send texts (SMS fallback will ignore images).")

# === Load Excel file ===
wb = openpyxl.load_workbook(xlsx_path)
sheet = wb.active

# === Open log file ===
with open(log_path, "w") as log:
    log.write("iMessage/SMS Sending Log\n")
    log.write(f"Started at: {datetime.now()}\n")
    log.write("========================\n\n")

    for row in sheet.iter_rows(min_row=2, values_only=True):
        phone_number, name = row
        if not phone_number:
            continue

        message_text = message_template.replace("{name}", name if name else "")

        # Step 1: Try iMessage (supports image)
        apple_script_imessage = f'''
        try
            tell application "Messages"
                set targetService to 1st service whose service type = iMessage
                set targetBuddy to buddy "{phone_number}" of targetService
                send "{message_text}" to targetBuddy
                {'send POSIX file "' + image_path + '" to targetBuddy' if os.path.exists(image_path) else ''}
            end tell
        on error errMsg
            error errMsg
        end try
        '''

        result = subprocess.run(
            ["osascript", "-e", apple_script_imessage],
            capture_output=True, text=True
        )

        if result.returncode == 0:
            log.write(f"[{datetime.now()}] ✅ SENT via iMessage: {phone_number}\n")
        else:
            log.write(f"[{datetime.now()}] ⚠️ iMessage failed for {phone_number}, trying SMS...\n")
        log.flush()  # ensure logs are written
        time.sleep(delay_seconds)

print(f"✅ All messages processed. Log saved to: {log_path}")
