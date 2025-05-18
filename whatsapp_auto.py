'''
import pandas as pd
import pywhatkit as kit
import time

# Read the Excel file
df = pd.read_excel("D:\\Automation_Py\\data\\men_data.xlsx")  # Replace with your actual file name

# Ensure the column names match exactly
# Example columns: 'PhoneNumber', 'Message'
for index, row in df.iterrows():
    phone_number = str(row['PhoneNumber'])
    message = str(row['Message'])  # Replace with your actual message column

    if not phone_number.startswith('+'):
        phone_number = '+91' + phone_number  # Adjust country code if needed

    print(f"Sending to {phone_number}...")

    # Schedule the message 1 minute ahead to allow for WhatsApp Web to load
    hour = time.localtime().tm_hour
    minute = (time.localtime().tm_min + 2 + index) % 100  # Adjust timing

    kit.sendwhatmsg(phone_number, message, hour, minute, wait_time=10, tab_close=True)

    time.sleep(10)  # Wait before sending the next message to avoid overlapping

print("All messages scheduled.")

'''
import pandas as pd
import pywhatkit as kit
import time

# Load Excel file, ensure PHONE NUMBER is read as string to avoid float issues
df = pd.read_excel(
    "D:\\Automation_Py\\data\\men_data1.xlsx",
    header=1,
    dtype={'PHONE NUMBER': str}
)

# Check column names for reference
print("Detected columns:", df.columns.tolist())

for index, row in df.iterrows():
    phone_raw = row.get('PHONE NUMBER', '')

    if pd.isna(phone_raw):
        continue

    phone_number = str(phone_raw).strip().replace(".0", "")  # Remove float decimals if any

    # Sanitize phone number
    if not phone_number.startswith('+'):
        phone_number = '+91' + phone_number

    if len(phone_number) < 13:  # '+91' + 10 digits = 13
        print(f"Skipping invalid number: {phone_number}")
        continue

    # Extract user info
    name = str(row.get('NAME', '')).strip()
    rent = str(row.get('RENT', '')).strip()
    food = str(row.get('FOOD', '')).strip()
    electricity = str(row.get('ELECTRICITY', '')).strip()
    total = str(row.get('TOTAL', '')).strip()

    # Compose the message
    message = (
        f"Dear {name},\n"
        f"This is your April month rent details:\n"
        f"* If paid, please share the screenshot.\n"
        f"* Pay before April 10th.\n"
        f"* Kindly take food on time, otherwise I can't take responsibility.\n"
        f"Room rent: {rent}\n"
        f"Food amount: {food}\n"
        f"Electric bill: {electricity}\n"
        f"Total: {total}\n"
        f"PhonePe number: 7094189502\n"
        f"UPI ID: sachinsubash57@ybl\n"
        f"Food timings:\n"
        f"* Weekdays: Breakfast & Lunch 8:00am to 9:00am, Dinner: 8:00pm to 9:00pm\n"
        f"* Weekend (+-10 minutes): Breakfast 8:00am to 9:00am, Lunch 1:10pm to 2:30pm,\n"
        f"  Dinner 8:00pm to 9:00pm\n"
        f"Thank you."
    )

    print(f"Sending to {phone_number}...")

    try:
        # Send message instantly

        kit.sendwhatmsg_instantly(
            phone_no=phone_number,
            message=message,
            wait_time=13,
            tab_close=True
        )
        time.sleep(5)  # Short delay between messages

    except Exception as e:
        print(f"❌ Failed to send to {phone_number}: {e}")
        with open("failed_numbers.txt", "a") as f:
            f.write(f"{phone_number} ({name}) - Error: {e}\n")

print("✅ All messages attempted.")


