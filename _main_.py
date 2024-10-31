"""
This script generates QR codes for variants from purchase order PDFs and emails them to Kaps Finishing.
"""

import os
import shutil
import pandas as pd
import qrcode
from tkinter import filedialog
from PIL import Image, ImageDraw, ImageFont
import imaplib

import PyPDF2
import re
from pprint import pprint
from tempfile import gettempdir
import zipfile
import smtplib
import json
import time

import email
from email.header import decode_header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# set the base directory to the directory of the script file so that the script can be run from anywhere and still
# access the required files in the same directory as the script file (e.g. fonts) using relative paths instead of
# absolute paths (which would be different depending on where the script is run from) - this is useful for creating
# standalone executables using PyInstaller or similar tools that bundle the script and its dependencies into a single
# executable file that can be run without needing to install Python or any dependencies on the target machine.
basedir = os.path.dirname(__file__)


class PdfReader:
    """
    A class to extract variant information from a PDF file.
    """
    def __init__(self):
        pass

    @staticmethod
    def extract_variant_info(pdf):
        """
        Extracts variant information from a PDF file.
        :param pdf:
        :return:
        """

        with open(pdf, 'rb') as f:
            pdf = PyPDF2.PdfReader(f)
            # Get the text from the first page
            text = pdf.pages[0].extract_text()

        # Regex pattern for variant code and description
        pattern = r'(V\d{6}) (.*?)(?=\d{2}/\d{2}/\d{4}(?!\d))'

        # Find all matches of the pattern in the text
        matches = re.findall(pattern, text, re.DOTALL)

        # Create a list of tuples for each pair of variant code and description
        variant_info = [(match[0], match[1].rstrip()) for match in matches]

        return variant_info


class Mailer:
    """
    A class to read emails from a mailbox and extract attachments.
    """
    def __init__(self):

        self.username = "updgoodsinbot@gmail.com"
        self.password = "bnbnfcndseivxwus"

    def login(self):
        """
        Logs into the email account.
        :return: The IMAP4 object
        """

        # create an IMAP4 class with SSL
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        # authenticate
        try:
            mail.login(self.username, self.password)
            print(f"Login successful: {self.username}")
        except Exception as e:
            print(f"Login failed: {e}")
            return

        return mail

    def find_latest(self):

        """Finds the latest UPD PO email in the inbox, if not used already"""

        while True:
            # create an IMAP4 class with SSL
            mail = imaplib.IMAP4_SSL("imap.gmail.com")
            r, d = mail.login(self.username, self.password)
            assert r == 'OK', 'login failed'
            # authenticate
            try:

                mail.select('INBOX')

                # find all emails from a sender with an @upwood-distribution.co.uk domain
                sender = "upwood-distribution.co.uk"
                result, data = mail.uid('search', None, f'(FROM "{sender}")')

                # get all mail IDs
                mail_ids = data[0]
                id_list = mail_ids.split()

                # retrieve the email headers

                for mail_id in id_list:
                    # fetch the email body (RFC822) for the given ID
                    result, data = mail.uid('fetch', mail_id, '(RFC822)')
                    raw_email = data[0][1]

                    # converts byte literal to string removing b''
                    raw_email_string = raw_email.decode('utf-8')
                    email_message = email.message_from_string(raw_email_string)

                    # downloading attachments
                    for part in email_message.walk():
                        # this part comes from the snipped I don't understand yet...
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue

                        # save pdf file in temp folder
                        filename = part.get_filename()
                        if filename:
                            match = re.search(r'UPD-PO(\d+)', filename)
                            if match:
                                print(f"Found PO: {match.group(1)}")
                                po_number = match.group(1)

                                finished_pos = os.path.join(basedir, 'finished_pos.json')
                                with open(finished_pos, 'r') as f:
                                    data = json.load(f)
                                if f'UPD-PO{po_number}' in data:
                                    print(f"PO {po_number} has already been processed. Skipping...")
                                    continue

                                new_filename = f'UPD-PO{po_number}.pdf'
                                filepath = os.path.join(gettempdir(), new_filename)

                                # Save the file with the new filename
                                with open(filepath, 'wb') as f:
                                    f.write(part.get_payload(decode=True))

                                # Open the saved PDF with PdfReader.extract_variant_info
                                results = PdfReader.extract_variant_info(filepath)

                                return {'po': f'UPD-PO{po_number}', 'variants': results}

            except mail.abort as e:
                print(e)
                continue
            # close the connection
            mail.logout()
            break

    class Generator:

        """
        A class to generate QR codes for variant codes and descriptions.

        """

        def __init__(self):
            pass

        def create_barcodes(self):
            """
            Creates QR codes for variant codes and descriptions.
            :return:
            """

            po = Mailer().find_latest()
            if not po:
                print("No new POs found.")
                return

            else:
                output_directory = gettempdir()

                for variant in po['variants']:
                    variant_code = variant[0]
                    description = variant[1]
                    barcode_file_path = os.path.join(output_directory, variant_code)

                    # Create QR code
                    qr = qrcode.QRCode(
                        version=1,
                        error_correction=qrcode.constants.ERROR_CORRECT_L,
                        box_size=10,
                        border=4,
                    )
                    qr.add_data(variant_code)
                    qr.make(fit=True)
                    qr_image = qr.make_image(fill_color="black", back_color="white")
                    qr_image.resize((10, 10))

                    # Convert 700x30mm to pixels
                    dpi = 300
                    width = int(70 * 0.0393701 * dpi)  # convert to pixels
                    height = int(30 * 0.0393701 * dpi)  # convert to pixels

                    # Create a new image with space for text
                    new_image = Image.new('RGB', (width, height), 'white')

                    # Paste the QR code to the top right corner
                    new_image.paste(qr_image, ((width // 2) + 130, 40))

                    # smaller_font = ImageFont.truetype('arial.ttf', 32)

                    # Add text
                    draw = ImageDraw.Draw(new_image)
                    main_font = ImageFont.truetype('arial.ttf', 45)
                    bold_font = ImageFont.truetype('arialbd.ttf', 100)

                    draw.text((30, 30), variant_code, font=bold_font, fill=(0, 0, 0))

                    # Split description into words and wrap the text
                    description_words = description.split(" ")

                    max_line_width = (new_image.width // 2) + 90
                    wrapped_description = self.wrap_text(description_words, main_font, max_line_width)
                    draw.text((30, 160), wrapped_description, font=main_font, fill=(0, 0, 0))

                    # Save the image
                    new_image.save(barcode_file_path + ".png")

                with zipfile.ZipFile(os.path.join(output_directory, f'{po["po"]}_barcodes.zip'), 'w') as zip_f:
                    for variant in po['variants']:
                        zip_f.write(os.path.join(output_directory, f'{variant[0]}.png'), f'{variant[0]}.png')

                try:
                    Mailer().send_email(os.path.join(output_directory, f'{po["po"]}_barcodes.zip'), po['po'])
                    finished_pos = os.path.join(basedir, 'finished_pos.json')
                    with open(finished_pos, 'r') as f:
                        data = json.load(f)
                    data.append(po['po'])
                    with open(finished_pos, 'w') as f:
                        json.dump(data, f, indent=4)
                except Exception as e:
                    print(f"Failed to send email: {e}")

        @staticmethod
        def wrap_text(words: list, font: ImageFont, max_line_width_: int):
            """
            Wraps text to fit within a specified width.
            :param words: The text to wrap as a list of words
            :param font: The font to use for calculating the text size
            :param max_line_width_: The maximum width of the text line
            :return: The wrapped text as a string with newline characters
            """
            lines = []
            current_line = []
            for word in words:
                # If adding a new word doesn't exceed the max width, add the word to current line
                if font.getlength(' '.join(current_line + [word])) <= max_line_width_:
                    current_line.append(word)
                else:
                    # If it does, add current line to lines and start a new line
                    lines.append(' '.join(current_line))
                    current_line = [word]
            # Add the last line
            if current_line:
                lines.append(' '.join(current_line))

            return '\n'.join(lines)

    def send_email(self, attachment, po_number):
        """
        Sends an email with the QR codes as attachments.
        :return:
        """
        recipients = ['tiger@kapsfinishing.co.uk', 'purchasing@upwood-distribution.co.uk',
                      'goods-in@upwooddistribution.co.uk', 'hsatchell@upwood-distribution.co.uk']

        # Create a multipart message and set headers
        message = MIMEMultipart()
        message["From"] = self.username
        message["To"] = ", ".join(recipients)
        message["Subject"] = f"Barcodes for {po_number}"

        # Add body to email
        message.attach(MIMEText(
            f"Hello, \n\nPlease find the attached zip file containing the barcodes for {po_number}.")
        )

        # Open the zip file
        with open(attachment, 'rb') as f:
            # Add file as application/octet-stream
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())

        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            'Content-Disposition',
            f'attachment; filename={po_number}_barcodes.zip'
        )

        # Add attachment to message and convert message to string
        message.attach(part)
        text = message.as_string()

        # Send the email
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(self.username, self.password)
            server.sendmail(self.username, recipients, text)


if __name__ == "__main__":
    while True:
        Generator().create_barcodes()
        time.sleep(3600)
