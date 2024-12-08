# NCLI: A CLI (Command Line Interface) Tool for Excel Column Removal

**NCLI** is a command-line application for removing specified columns from Excel `.xlsx` files, while preserving data types, formatting, and cell styles. Unlike solutions that infer data types (like Pandas), NCLI works directly with the Excel file structure via `openpyxl`, ensuring numeric cells remain numeric, date cells remain dates, and formatting stays intact.

---

## Features

- **Preserves Data Types & Formatting**:  
  Works at the workbook level using `openpyxl`, so no data type inference occurs. Your numeric cells stay numeric, dates stay dates, and formatting remains intact.
  
- **Command-Line Interface (CLI)**:  
  Easily integrate into your workflows. Run a single command to process entire directories of files.
  
- **Cross-Platform Compatibility**:  
  Works seamlessly on Linux, Windows, and macOS. As long as Python 3 and the required packages are installed, you’re good to go.
  
- **Logging & Progress Indication**:  
  Logs its actions to `process.log` and uses `tqdm` to show a progress bar for multiple files, enhancing user experience.

- **Extensible Architecture**:  
  Designed with subcommands, allowing for future extensions.

---

## Requirements

- **Python**: Version 3.7 or higher.
- **Python Libraries**:
  - `openpyxl`: For Excel manipulation.
  - `tqdm`: For progress bars.

Install dependencies:

```bash
pip install openpyxl tqdm
```

## Installation Guide

This section guides you through the installation process for **NCLI**, a CLI tool for manipulating Excel files.

---

### Installation Steps

1. **Clone or Download the Repository**:
   ```bash
   git clone https://github.com/dave07y/ncli.git
   cd ncli


## Setup Guide

Follow this section to configure and prepare your environment for using **NCLI**.

---

### Preparing the Input Directory

1. **Organize Excel Files**:
   - Place all `.xlsx` files you want to process in a single directory. For example:
     ```
     ./input
     ├── sample1.xlsx
     ├── sample2.xlsx
     └── sample3.xlsx
     ```

2. **Check Headers**:
   - Ensure the first row in your Excel files contains the column headers. These headers will be matched against the columns you specify for removal.

---

### Creating a Columns File

1. **What is a Columns File?**
   - A text file listing the column headers you want to remove, one per line.
   - Example (`columns_to_drop.txt`):
     ```
     Name
     DOB
     MR#
     etc...
     ```

2. **How to Create It**:
   - Use any text editor (e.g., Notepad, VSCode, nano).
   - Save the file in the same directory as `ncli.py` or specify its path when running the tool.

---

### Setting Up the Output Directory

1. **Default Directory**:
   - Processed files are saved in a directory named `processed` by default.

2. **Custom Output Directory**:
   - You can specify a custom output directory using the `--output-dir` option.

---

## Usage Guide

This section provides detailed instructions for using **NCLI** to remove columns from Excel files.

---

### Command Overview

Run the following command to use the `dropcolumns` subcommand:
```bash
python ncli.py dropcolumns --input-dir <input-dir> --output-dir <output-dir> --columns-file <columns-file>

