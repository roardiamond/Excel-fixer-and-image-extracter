# Excel Image Extractor & Fixer

A pure Python tool that extracts all embedded images from Excel (`.xlsx`) files **and** converts modern "Place in Cell" pictures into classic floating pictures so they display correctly in Outlook, Gmail, older Excel, mobile apps, and email previews.

**No external dependencies** – only Python standard library.

---

## What This Tool Does

### 1. Image Extraction
- Pulls **every** embedded image (PNG, JPG, GIF, BMP, etc.) out of the workbook
- Names files by the row they are anchored to (e.g. image in cell `H23` → `23.png`)
- Handles multiple sheets and duplicate row numbers gracefully
- Optionally extract images from only one column

### 2. Email-Safe Workbook Creator
Excel 365 introduced **"Place in Cell"** pictures. These look great inside Excel 365, but:
- Show as `#VALUE!` in older Excel
- Break in Outlook / Gmail / phone previews
- Don't display properly in many email clients

This tool creates a **copy** of your workbook where every "Place in Cell" picture is converted into a normal floating picture that sits exactly over the same cell. The original file is **never modified**.

---

## Features

- Pure Python (no `pip install` needed)
- Works with both floating pictures and modern in-cell pictures
- Preserves original file completely (read-only)
- Smart naming of extracted images
- Creates email-compatible `.xlsx` copy
- Configurable via simple settings at the top of the script
- Cross-platform (Windows, macOS, Linux)

---

## How to Use

### Step 1 – Prepare your file
1. Place your Excel file next to `excel.py`
2. Rename it to `input.xlsx` (or change the setting inside the script)

### Step 2 – Run the script
```bash
python excel.py
```

### Step 3 – Check the results
- **Extracted images** → folder `extracted_images/`
- **Email-safe workbook** → `input_email_safe.xlsx`

That's it. Send the `*_email_safe.xlsx` file when you need the images to show up everywhere.

---

## Configuration (Top of `excel.py`)

```python
EXCEL_FILE = "input.xlsx"                 # Your source file
OUTPUT_FOLDER = "extracted_images"        # Where images are saved
EXTRACT_IMAGES = True                     # Turn extraction on/off
MAKE_EMAIL_SAFE_COPY = True               # Turn email-safe copy on/off
EMAIL_SAFE_FILE = "input_email_safe.xlsx" # Name of the safe copy
ONLY_COLUMN = None                        # e.g. "H" to extract only column H
```

Change any of these values and run the script again.

---

## Example Output

```
Extracted: Sheet1 H23 -> extracted_images/23.png
Extracted: Sheet1 H24 -> extracted_images/24.jpg
Converted: Sheet1 H23 (in-cell picture -> normal picture)
Converted: Sheet1 H24 (in-cell picture -> normal picture)

Total images extracted: 2
Email-safe copy created (2 picture(s) converted):
input_email_safe.xlsx
```

---

## Requirements

- Python 3.6 or higher
- No third-party packages required

---

## Important Notes

- Only `.xlsx` files are supported (old `.xls` is not)
- Original file is opened in read-only mode and is **never** changed
- The email-safe copy is a brand new file – safe to share

---

## License

MIT License – free to use, modify, and distribute.
