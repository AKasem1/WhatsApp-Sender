import pywhatkit as kit
import pyautogui
import time
import pandas as pd

# Get inputs from user
textFile = input("Enter the path to your text file: ")
image = input("Enter the path to your image: ")
excel_file = input("Excel file: ")

# Prepare image path and load phone numbers from Excel
image = r"{}".format(image)
df = pd.read_excel(excel_file)
phone_numbers = [str(num[0]) for num in df.values.tolist()]

# Read the content of the text file
with open(textFile, "r", encoding="utf-8") as file:
    content = file.read()

i = 1
for num in phone_numbers:
    # Add the correct country code if necessary
    if str(num)[0] == "0":
        num = "+2" + str(num)
    elif str(num)[0] == "1":
        num = "+20" + str(num)
    elif str(num)[0] == "2" and str(num)[1] == "0":
        num = "+" + str(num)
    elif str(num)[0] == "+" and str(num)[1] == "2":
        num = str(num)

    print(f"Sending to {num}...")

    # Send WhatsApp image with caption
    kit.sendwhats_image(num, image, content)

    # Ensure sufficient delay before pressing enter
    time.sleep(8)  # Adjust this delay if needed (increase if needed)

    # Press enter to send the message
    pyautogui.press('enter')

    print(f"Message sent to {num}: {i}")

    # Wait for 20 seconds before sending the next message
    time.sleep(20)

    # Close the WhatsApp web tab
    pyautogui.hotkey('ctrl', 'w')
    
    # Wait briefly before moving to the next number
    time.sleep(2)
    
    i += 1

input("press enter to exit..")
