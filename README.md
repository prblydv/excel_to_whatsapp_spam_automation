# WhatsApp Message Automation using VBA

## Overview
This VBA script automates the process of sending WhatsApp messages using Internet Explorer. It retrieves phone numbers and corresponding messages from an Excel sheet and sends messages via the WhatsApp desktop application.

## Requirements
- Microsoft Excel (with VBA enabled)
- WhatsApp desktop application installed
- Internet Explorer (IE)
- Macro security settings adjusted to allow execution

## Installation and Setup
1. Open Microsoft Excel and enable macros.
2. Navigate to the VBA editor by pressing `ALT + F11`.
3. Insert a new module and copy-paste the `Sub WhatsAppMsg()` script.
4. Ensure that your data is structured as follows in the **Data** sheet:
   - Column A: Phone numbers (including country codes, e.g., `+1234567890`)
   - Column B: Messages to be sent

## How to Use
1. Run the macro `Sub WhatsAppMsg()`.
2. The script will:
   - Read each phone number and message from the **Data** sheet.
   - Construct the WhatsApp message URL.
   - Open Internet Explorer to navigate to WhatsApp.
   - Simulate pressing enter to send the message.
3. The script will wait for 5 seconds before processing the next message.

## Script Breakdown
- **Reads data** from an Excel sheet (phone number and message).
- **Formats the message URL** for WhatsApp Web.
- **Uses Internet Explorer (IE)** to open the WhatsApp chat.
- **Simulates pressing enter** to send the message.
- **Iterates over all entries** in the dataset.

## Notes
- The script relies on `InternetExplorer.Application`, which may not work on modern browsers. Consider using another automation approach such as Selenium.
- WhatsApp may block excessive automated messages due to spam detection.
- Ensure that your phone is connected to WhatsApp Web before running the script.
- Modify the `Application.Wait` timing if needed to account for slower or faster systems.

## Disclaimer
This script is provided as-is for educational purposes. Use it responsibly and ensure compliance with WhatsApp's terms of service.

