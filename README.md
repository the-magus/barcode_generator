# QR Code Generator for Purchase Orders

## Overview

This application generates QR codes for variants from purchase order PDFs and emails them to Kaps Finishing. It reads emails from a specified mailbox, extracts attachments, processes the PDFs to extract variant information, generates QR codes, and sends the QR codes via email.

## Features

- Reads emails from a specified mailbox.
- Extracts purchase order PDFs from the emails.
- Processes the PDFs to extract variant information.
- Generates QR codes for each variant.
- Sends the QR codes via email.
- Keeps track of processed purchase orders to avoid duplication.

## Requirements

- Python 3.6+
- Required Python packages (listed in `requirements.txt`):
  - pandas~=2.2.1
  - qrcode~=7.4.2
  - pillow~=10.3.0
  - PyPDF2
  - imaplib
  - smtplib
  - email
  - json
  - re
  - tkinter

## Installation

1. Clone the repository:
   ```sh
   git clone <https://github.com/the-magus/barcode_generator>
   ```

2. Install the required packages:
   ```sh
   pip install -r requirements.txt
   ```

## Configuration

- Update the email credentials in the `Mailer` class in `_main_.py`:
  ```python
  self.username = "email@domain"
  self.password = "password"
  ```

## Usage

1. Activate the virtual environment:
- Windows: 
  ```sh
  venv\Scripts\activate
  ```
- MacOS/Linux: 
  ```sh
  source venv/bin/activate
  ```
2. Run the script:
   ```sh
   python _main_.py
   ```
The script will continuously check for new purchase order emails every hour, process them, generate QR codes, and send the QR codes via email.

## File Structure

- `_main_.py`: Main script that runs the application.
- `finished_pos.json`: JSON file to keep track of processed purchase orders.
- `requirements.txt`: List of required Python packages.
- `README.md`: Project overview and instructions.
- `arial.ttf`: Font file for the QR code.
- `arialbd.ttf`: Bold font file for the QR code.
- `venv`: Virtual environment directory.