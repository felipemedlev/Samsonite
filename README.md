[samsonite_readme.md](https://github.com/user-attachments/files/21746003/samsonite_readme.md)
# 🏷 Samsonite Product Tag & Label Generator

## 📌 Overview

This tool automates the creation of **Samsonite product collaterals** (labels, stickers, and tags) from a single **Excel Collaterals Matrix** file. It's designed to save time, reduce manual design work, and ensure **branding and export compliance** across multiple markets.

Given an Excel spreadsheet where **each row corresponds to one product**, the app generates:

- **Barcodes** (`EAN13`) in SVG format
- **PDF labels & stickers** with dynamic product details
- Organized output folders for easy exporting & printing

The program includes a **simple Tkinter-based graphical interface** so that no command-line knowledge is needed, and can be distributed as a **standalone executable** built with **PyInstaller** for easy installation on any computer.

## 🎯 Key Features

- **Excel to Labels in One Click** - Reads your **Collateral Matrix** and generates all required product assets.
- **Multiple Label Formats Supported** - Creates PDF & SVG collaterals for:
  - **Main Hangtag** (Front & Rear)
  - **Inner Label (2 versions)**
  - **LATAM Hangstickers** (All & Brazil-specific)
  - **EU Hangstickers** (All & Brazil-specific)
- **Barcode Generation** - Automatically generates **EAN13 barcodes** for each product.
- **Dynamic Product Data** - Inserts SKU, style name, product type, color, material composition, size, weight, liters, and origin country into the labels.
- **Custom Brand Fonts** - Uses Samsonite's specified fonts for brand consistency.
- **Automatic Folder Organization** - Outputs are stored in subfolders by product style name, ready for printing or export.
- **User-Friendly Tkinter Interface**
  - Browse for your **Excel file**
  - Choose your **output directory**
  - Click **Generate** and let the script do the work
- **PyInstaller Executable**
  - No Python installation required
  - Works on any Windows PC by running the `.exe` file

## 🖼 Output Folder Structure

After generation, your output will be organized like this:

```
output_folder/
├── STYLE_NAME_1/
│   ├── Main_Hangtag/
│   │   ├── Main_Hangtag_Front_SKU.pdf
│   │   ├── Main_Hangtag_Rear_SKU.pdf
│   ├── Inner_Label_2/
│   │   ├── Inner_Label_2_SKU.pdf
│   ├── Inner_Label/
│   │   ├── Inner_Label_Front_SKU.pdf
│   │   ├── Inner_Label_Rear_SKU.pdf
│   ├── HangSticker_Latam/
│   │   ├── Hangtag_Sticker_Latam_All_SKU.pdf
│   │   ├── Hangtag_Sticker_Latam_BR_SKU.pdf
│   ├── HangSticker_EU/
│   │   ├── Hangtag_Sticker_EU_All_SKU.pdf
│   │   ├── Hangtag_Sticker_EU_BR_SKU.pdf
│   ├── EAN13/
│   │   ├── EAN13_SKU.svg
│   │   ├── EAN13_SKU.png
├── STYLE_NAME_2/
│   └── ...
```

## 🚀 How It Works

1. **Open the App** - Launch the `.exe` file (or run the Python script if using source code).
2. **Select Excel File** - Use the GUI to browse for the **Collateral Matrix** Excel file.
3. **Choose Output Folder** - Pick where the generated collaterals should be saved.
4. **Click Generate** - The tool reads product data, creates barcodes, generates PDFs, and organizes them into folders.

## 📋 Requirements (if running from source)

- Python 3.8+
- Excel file with **exact column names** used in the Collateral Matrix
- Required fonts & template PDFs in the `src/` folder

**Dependencies:**
- pandas
- numpy
- python-barcode
- reportlab
- svglib
- pdfrw
- PyMuPDF (fitz)
- pillow
- tkinter (comes with Python)

## 💡 Use Cases

- **Manufacturing** → Preparing all product tags before shipping
- **Export Compliance** → Automatically includes all necessary market-specific labels
- **Brand Consistency** → Uses the same design for every label, avoiding manual errors

## 📦 Building the Executable

To create the standalone `.exe` file:

```bash
pyinstaller --onefile --noconsole main.py
```

The executable will be found inside the `dist/` folder and can be run on any Windows PC without installing Python.
