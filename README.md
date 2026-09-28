# Excel Image Extractor & Fixer

**Paid tool — ₹500 one-time**

A pure Python tool that extracts all embedded images from Excel (`.xlsx`) files **and** converts modern "Place in Cell" pictures into classic floating pictures so they display correctly in Outlook, Gmail, older Excel, mobile apps, and email previews.

---

## Pricing & How to Get It

| Item | Details |
|------|---------|
| **Price** | ₹500 (one-time) |
| **Buy / Get License** | [voidxhub.in](https://voidxhub.in) |
| **UPI** | `yashrawat9873-1@okicici` (Yash Rawat) |
| **After payment** | Message on voidxhub / share UTR → you receive the full working script |

This public repo only contains a stub. The complete working code is private and shared only with paying users.

---

## What the Full Tool Does

### 1. Image Extraction
- Pulls **every** embedded image (PNG, JPG, GIF, BMP, etc.) out of the workbook
- Names files by the row they are anchored to (e.g. image in cell `H23` → `23.png`)
- Handles multiple sheets and duplicate row numbers gracefully
- Optionally extract images from only one column

### 2. Email-Safe Workbook Creator
Excel 365 "Place in Cell" pictures show as `#VALUE!` in older Excel / Outlook / Gmail / phone apps.

The tool creates a **copy** where every in-cell picture becomes a normal floating picture that sits exactly over the same cell. Original file is never modified.

---

## Features (Full Version)

- Pure Python (standard library only)
- Floating + modern in-cell pictures
- Original file stays untouched
- Smart image naming
- Email-compatible `.xlsx` copy
- Easy settings at top of script
- Windows / macOS / Linux

---

## How to Use (After You Get the Full Script)

1. Place your Excel file next to `excel.py` and name it `input.xlsx`
2. Run:
   ```bash
   python excel.py
   ```
3. Results:
   - Images → `extracted_images/`
   - Email-safe file → `input_email_safe.xlsx`

---

## Configuration (in full script)

```python
EXCEL_FILE = "input.xlsx"
OUTPUT_FOLDER = "extracted_images"
EXTRACT_IMAGES = True
MAKE_EMAIL_SAFE_COPY = True
EMAIL_SAFE_FILE = "input_email_safe.xlsx"
ONLY_COLUMN = None          # e.g. "H" for only column H
```

---

## Requirements

- Python 3.6+
- No external packages needed

---

## Support / Purchase

Buy or get the full script here: **https://voidxhub.in**

Or pay ₹500 to UPI `yashrawat9873-1@okicici` and contact via the site with UTR.

---

Made by **yashxchi** · Powered by voidxhub
