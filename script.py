import pandas as pd
import pywhatkit as kit
import time
import numpy as np

# 🟢 Load configuration from input.txt
# this script will show to whome the message is sending...
variables = {}
try:
    with open('C:\\pg_data\\input.txt', 'r') as file:
        for line in file:
            if '=' in line:
                key, value = line.strip().split('=', 1)
                key = key.strip()
                value = value.strip()

                if value.lower() == 'true':
                    value = True
                elif value.lower() == 'false':
                    value = False
                elif value.startswith('"') and value.endswith('"'):
                    value = value[1:-1]
                variables[key] = value
except Exception as e:
    print(f"❌ Error reading input.txt: {e}")
    exit()

sheet_to_process = variables.get('sheet_to_process')
PATH = variables.get('PATH')
payment_upi_id = variables.get('payment_upi_id')
payment_phone_number = variables.get('payment_phone_number')
payment_date = variables.get('payment_date')
is_food = variables.get('is_food')

print(f"🚀 Starting process for: {sheet_to_process}")

# 🟢 Read Excel safely
# We remove specific dtypes for numeric columns here to prevent the crash
df = pd.read_excel(
    PATH,
    sheet_name=sheet_to_process,
    header=1,
    na_values=[' ', '', 'NA', 'N/A'],
    dtype={'PHONE NUMBER': str}  # Keep Phone as string to preserve leading zeros
)

# 🟢 Clean Numeric Columns
# This converts spaces/text to NaN, then fills with 0, then converts to whole numbers (int)
numeric_cols = ['RENT', 'FOOD', 'ELECTRICITY', 'PENDINGS', 'TOTAL']
for col in numeric_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
    else:
        # Create the column with 0 if it's missing to avoid KeyErrors later
        df[col] = 0

print(f"📄 Processing sheet: {sheet_to_process}")

for index, row in df.iterrows():
    phone_raw = row.get('PHONE NUMBER', '')

    if pd.isna(phone_raw) or str(phone_raw).strip() == "":
        continue

    # Clean phone number formatting
    phone_number = str(phone_raw).strip().split('.')[0] # Remove .0 if it exists
    if not phone_number.startswith('+'):
        phone_number = '+91' + phone_number

    if len(phone_number) < 13:
        print(f"⚠️ Skipping invalid number: {phone_number}")
        continue

    # Get data for message
    name = str(row.get('NAME', 'Guest')).strip()
    rent = row.get('RENT', 0)
    food = row.get('FOOD', 0)
    pending = row.get('PENDINGS', 0)
    electricity = row.get('ELECTRICITY', 0)
    total = row.get('TOTAL', 0)

    # 🟢 Construct Message
    if is_food:
        message = (
            f"🎉😄 Hello {name}! 😄🎉\n"
            f"Hope you’re doing well... ✨😊\n"
            f"This is your {sheet_to_process} month rent details:\n"
            f"* If paid, please share the screenshot 👍\n"
            f"* {payment_date}\n"
            f"* Kindly take food on time, otherwise I can't take responsibility.\n"
            f"Room rent: {rent}\n"
            f"Food amount: {food}\n"
            f"Electric bill: {electricity}\n"
            f"Pending: {pending}\n"
            f"Total: {total}\n"
            f"PhonePe number: {payment_phone_number}\n"
            f"UPI ID: {payment_upi_id}\n"
            f"Food timings:\n"
            f"* Weekdays: Breakfast & Lunch 8:00am to 9:00am, Dinner: 8:00pm to 9:00pm\n"
            f"* Weekend (+-10 mins): Breakfast 8:00am-9:00am, Lunch 1:10pm-2:30pm, Dinner 8:00pm-9:00pm\n"
            f"Thanks 😊🙏"
        )
    else:
        message = (
            f"Dear {name},\n"
            f"This is your {sheet_to_process} month rent details:\n"
            f"* If paid, please share the screenshot.\n"
            f"* {payment_date}.\n"
            f"Room rent: {rent}\n"
            f"Food amount: {food}\n"
            f"Electric bill: {electricity}\n"
            f"Pending: {pending}\n"
            f"Total: {total}\n"
            f"PhonePe number: {payment_phone_number}\n"
            f"UPI ID: {payment_upi_id}\n"
            f"Thank you."
        )

    try:
        print(f"📤 Sending to {name} ({phone_number})...")
        kit.sendwhatmsg_instantly(
            phone_no=phone_number,
            message=message,
            wait_time=13,
            tab_close=True
        )
        time.sleep(5) # Give the browser time to close and reset

    except Exception as e:
        print(f"❌ Failed to send to {phone_number}: {e}")
        with open("failed_numbers.txt", "a") as f:
            f.write(f"{phone_number} ({name}) - Error: {e}\n")

print(f"✅ All messages attempted for sheet: {sheet_to_process}")

