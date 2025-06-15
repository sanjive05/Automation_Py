import pandas as pd
import pywhatkit as kit
import time

# 🟢 Specify the sheet to process here
variables = {}

with open('C:\\pg_data\\input.txt', 'r') as file:
    for line in file:
        if '=' in line:
            key, value = line.strip().split('=', 1)
            key = key.strip()
            value = value.strip()

            # Convert string "True"/"False" to boolean, and remove quotes from strings
            if value.lower() == 'true':
                value = True
            elif value.lower() == 'false':
                value = False
            elif value.startswith('"') and value.endswith('"'):
                value = value[1:-1]

            variables[key] = value

# Now you can access the variables like this:
sheet_to_process = variables['sheet_to_process']
PATH = variables['PATH']
payment_upi_id = variables['payment_upi_id']
payment_phone_number = variables['payment_phone_number']
payment_date = variables['payment_date']
is_food = variables['is_food']

print(sheet_to_process, PATH, payment_upi_id, payment_phone_number, payment_date, is_food)

df = pd.read_excel(PATH,
    sheet_name=sheet_to_process,
    header=1,
    dtype={
        'PHONE NUMBER': str,
        'RENT': 'Int64',
        'FOOD': 'Int64',
        'ELECTRICITY': 'Int64',
        'PENDING': 'Int64'
    }
)
print(df.columns)



print(f"📄 Processing sheet: {sheet_to_process}")
print("Detected columns:", df.columns.tolist())

for index, row in df.iterrows():
    phone_raw = row.get('PHONE NUMBER', '')

    if pd.isna(phone_raw):
        continue

    phone_number = str(phone_raw).strip().replace(".0", "")

    if not phone_number.startswith('+'):
        phone_number = '+91' + phone_number

    if len(phone_number) < 13:
        print(f"Skipping invalid number: {phone_number}")
        continue

    name = str(row.get('NAME', '')).strip()
    rent = str(row.get('RENT', '')).strip()
    food = str(row.get('FOOD', '')).strip()
    pending = str(row.get('PENDINGS', '')).strip()
    electricity = str(row.get('ELECTRICITY', '')).strip()
    total = str(row.get('TOTAL', '')).strip()

    if (is_food):
        message = (
            f"Dear {name},\n"
            f"This is your {sheet_to_process} month rent details:\n"
            f"* If paid, please share the screenshot.\n"
            f"* {payment_date}.\n"
            f"* Kindly take food on time, otherwise I can't take responsibility.\n"
            f"Room rent: {rent}\n"
            f"Food amount: {food}\n"
            f"Electric bill: {electricity}\n"
            f"pending : {pending}\n"
            f"Total: {total}\n"
            f"PhonePe number: {payment_phone_number}\n"
            f"UPI ID: {payment_upi_id}\n"
            f"Food timings:\n"
            f"* Weekdays: Breakfast & Lunch 8:00am to 9:00am, Dinner: 8:00pm to 9:00pm\n"
            f"* Weekend (+-10 minutes): Breakfast 8:00am to 9:00am, Lunch 1:10pm to 2:30pm,\n"
            f"  Dinner 8:00pm to 9:00pm\n"
            f"Thank you."
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
            f"pending : {pending}\n"
            f"Total: {total}\n"
            f"PhonePe number: {payment_phone_number}\n"
            f"UPI ID: {payment_upi_id}\n"
            f"Thank you."
        )
    try:
        kit.sendwhatmsg_instantly(
            phone_no=phone_number,
            message=message,
            wait_time=13,
            tab_close=True
        )
        time.sleep(5)

    except Exception as e:
        print(f"❌ Failed to send to {phone_number}: {e}")
        with open("failed_numbers.txt", "a") as f:
            f.write(f"{phone_number} ({name}) - Error: {e}\n")

print("✅ All messages attempted for sheet:", sheet_to_process)
