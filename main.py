"""Generates QR codes from two columns of a user provided Excel file and saves them as PNG files,
with the first column providing the QR code data and the second providing the file name."""

import os
import pandas as pd
import qrcode
from tkinter import filedialog

from PIL import Image, ImageDraw, ImageFont

# Get file from user

file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])

if file_path:
    df = pd.read_excel(file_path)
    output_dir = filedialog.askdirectory(title="Select output directory")

    # Generate barcodes

    for index, row in df.iterrows():
        barcode = (row[0])
        barcode_fp = os.path.join(output_dir, row[1])

        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )

    # Resize the barcodes to 400x400 pixels

        qr.add_data(barcode)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        img = img.resize((400, 400))

        # Create a new image with space for text
        new_img = Image.new('RGB', (400, 450), 'white')  # Adjust the second parameter as needed
        new_img.paste(img, (0, 0))

        # Add text
        d = ImageDraw.Draw(new_img)
        fnt = ImageFont.truetype('arial.ttf', 45)  # You may need to adjust the font path
        d.text((10, 390), barcode, font=fnt, fill=(0, 0, 0))  # Adjust coordinates and color as needed

        new_img.save(barcode_fp + ".png")

    # open the output directory

    os.startfile(output_dir)
