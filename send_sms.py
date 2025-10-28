import openpyxl
import subprocess
import time

# === CONFIG ===
xlsx_file = "contacts.xlsx"  # Path to your XLSX
message_template = "Hello {name}, this is a hardcoded message!"  # Your message
log_file = "log.txt"  # Log file

# === Load XLSX ===
wb = openpyxl.load_workbook(xlsx_file)
sheet = wb.active

# Open log file
with open(log_file, "w") as log:
    log.write("SMS Sending Log\n")
    log.write("====================\n")

    # Skip header
    for row in sheet.iter_rows(min_row=2, values_only=True):
        phone_number, name = row
        if phone_number:
            # Customize message
            message_text = message_template.replace("{name}", name if name else "")

            # AppleScript command
            apple_script = f'''
            tell application "Messages"
                set targetService to 1st service whose service type = SMS
                set targetBuddy to buddy "{phone_number}" of targetService
                send "{message_text}" to targetBuddy
            end tell
            '''

            try:
                subprocess.run(["osascript", "-e", apple_script], check=True)
                log.write(f"✅ SUCCESS: {phone_number} - {message_text}\n")
            except subprocess.CalledProcessError as e:
                log.write(f"❌ FAILED: {phone_number} - {message_text} | Error: {e}\n")

            time.sleep(2)  # Delay 2 seconds between messages

print(f"All messages processed. Log saved to {log_file}")

