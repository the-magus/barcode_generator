"""
This script generates QR codes for each row in an Excel file. The Excel file should have three columns:
1. Variant code (for the QR code data) - Column A
2. Tiger code (for the file name) - Column B
3. Description - Column C

The script will ask the user to select an Excel file and an output directory. It will then generate QR codes for each
row in the Excel file and save them as PNG files in the output directory. The QR code will contain the variant code,
and the file name will be the Tiger code. The PNG file will also contain the variant code, tiger code, and a wrapped
description text.
"""

import os
import pandas as pd
import qrcode
from tkinter import filedialog
from PIL import Image, ImageDraw, ImageFont

# set the base directory to the directory of the script file so that the script can be run from anywhere and still
# access the required files in the same directory as the script file (e.g. fonts) using relative paths instead of
# absolute paths (which would be different depending on where the script is run from) - this is useful for creating
# standalone executables using PyInstaller or similar tools that bundle the script and its dependencies into a single
# executable file that can be run without needing to install Python or any dependencies on the target machine.
basedir = os.path.dirname(__file__)

# Set the app icon for the script window (if running as a standalone executable)
try:
    # noinspection PyUnresolvedReferences
    # import ctypes only on Windows to avoid ImportError on other platforms (e.g. Linux, macOS)
    # ctypes is used to set the app icon for the script window
    from ctypes import windll  # Only exists on Windows.

    # myappid is a unique string that identifies the application to the OS (Windows) - this is used to associate the
    # script window with the app icon specified in the .ico file (e.g. app_icon.ico) - this is useful when the script
    # is run as a standalone executable, and you want the script window to have a custom icon in the taskbar, alt-tab
    # switcher, etc. instead of the default Python icon. The icon file should be in the same directory as the script
    # file and should be in .ico format (e.g. app_icon.ico).
    myappid = u'Eye of Magus.Sign Label Generator.1.0.0'  # Arbitrary string - can be any string (e.g. 'myappid')
    # but it should be unique to the application.
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

except ImportError:
    pass


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


# Ask user to select an Excel file
excel_file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])

if excel_file_path:
    data_frame = pd.read_excel(excel_file_path, header=None)

    # Ask user to select an output directory
    output_directory = filedialog.askdirectory(title="Select output directory")

    # Generate QR codes for each row in the data frame
    for index, row in data_frame.iterrows():
        variant_code = row[0]
        tiger_code = row[1]
        description = row[2]
        barcode_file_path = os.path.join(output_directory, tiger_code)

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
        new_image.paste(qr_image, ((width // 2) + 150, -20))

        smaller_font = ImageFont.truetype('arial.ttf', 32)

        # Add text
        draw = ImageDraw.Draw(new_image)
        main_font = ImageFont.truetype('arial.ttf', 45)
        draw.text(((width // 2) + 195, (height // 2) + 65), variant_code, font=main_font, fill=(0, 0, 0))
        draw.text(
            ((width // 2) + 195, (height // 2) + 115), tiger_code,
            font=main_font if len(tiger_code) < 11 else smaller_font, fill=(0, 0, 0)
        )

        # Split description into words and wrap the text
        description_words = description.split(" ")
        bold_font = ImageFont.truetype('arialbd.ttf', 50)

        # Customer branded signs have a Tiger code starting with '9'. There's typically a code at the beginning of their
        # description that is not part of the Tiger code. This code is bolded and placed at the top of the image.
        first_word = None
        if tiger_code.startswith('9'):
            first_word = description_words.pop(0)
            draw.text((10, 10), first_word, font=bold_font, fill=(0, 0, 0))

        max_line_width = (new_image.width // 2) + 90
        wrapped_description = wrap_text(description_words, main_font, max_line_width)
        draw.text((10, 60 if first_word else 10), wrapped_description, font=main_font, fill=(0, 0, 0))

        # Save the image
        new_image.save(barcode_file_path + ".png")

    # Open the output directory
    os.startfile(output_directory)
