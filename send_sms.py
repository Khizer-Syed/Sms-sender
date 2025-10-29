import openpyxl
import subprocess
import time
import os

# === CONFIG ===
xlsx_file = "contacts.xlsx"
image_path = "/Users/khizer/Desktop/image.jpg"
message_template = "Hello {name}, this is a hardcoded message with an image!"
log_file = "log.txt"

if not os.path.exists(image_path):
    raise FileNotFoundError(f"Image not found: {image_path}")

wb = openpyxl.load_workbook(xlsx_file)
sheet = wb.active

with open(log_file, "w") as log:
    log.write("iMessage/SMS Sending Log\n")
    log.write("========================\n")

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
                send POSIX file "{image_path}" to targetBuddy
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
            log.write(f"✅ SENT via iMessage: {phone_number}\n")
        else:
            log.write(f"⚠️ iMessage failed for {phone_number}, trying SMS...\n")

            # Step 2: Try SMS (no image support)
            apple_script_sms = f'''
            tell application "Messages"
                set targetService to 1st service whose service type = SMS
                set targetBuddy to buddy "{phone_number}" of targetService
                send "{message_text}" to targetBuddy
            end tell
            '''
            sms_result = subprocess.run(
                ["osascript", "-e", apple_script_sms],
                capture_output=True, text=True
            )

            if sms_result.returncode == 0:
                log.write(f"✅ SENT via SMS: {phone_number}\n")
            else:
                log.write(f"❌ FAILED: {phone_number} - {sms_result.stderr}\n")

        time.sleep(2)

print(f"All messages processed. Log saved to {log_file}")
