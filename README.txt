Generates labels for each row in an Excel file saved as PNG images.

The Excel file should have three columns:
1. Variant code (for the QR code data) - Column A
2. Tiger code (for the file name) - Column B
3. Description - Column C

The program will ask the user to select an Excel file and an output directory. It will then generate QR codes for each
row in the Excel file and save them as PNG files in the output directory.

The QR code will contain the variant code,
and the file name will be the Tiger code. The PNG file will also contain the variant code, Tiger code, and a wrapped
description text.