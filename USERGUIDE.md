# NCLI User Guide

NCLI is a command-line tool for removing specified columns from Excel `.xlsx` files while preserving data types and formatting. This guide explains how to use the prebuilt executable on your system.

---

## System Requirements

- **Operating System**: Windows (for the provided `.exe`).
- **No Python Installation Required**: The `.exe` is a standalone application.

---

## Setup

1. **Download the Executable**:
   - Download the `ncli.exe` file provided in the release.

2. **Prepare Input Files**:
   - Create a directory (e.g., `input`) containing the `.xlsx` files you want to process.
   - Ensure the first row in each file contains the column headers.

3. **Prepare a Columns File**:
   - Create a text file (e.g., `columns_to_drop.txt`) listing the headers of columns you want removed, one per line. 
   Example:

     ```
     Name
     DOB
     MR#
     ```

4. **Create an Output Directory**:
   - Decide where you want the processed files to be saved (e.g., `processed`). If not specified, the tool will use a default `processed` directory.

---

## Usage

1. **Open a Command Prompt**:
   - Press `Win + R`, type `cmd`, and press Enter.

2. **Navigate to the Executable**:
   - Use the `cd` command to go to the directory containing `ncli.exe`. Example:

     ```bash
     cd C:\path\to\exe
     ```

3. **Run the Command**:

   Use the following format:

   ```bash
   ncli.exe dropcolumns --input-dir "<path-to-input>" --output-dir "<path-to-output>" --columns-file "<path-to-columns-file>"
   ```

### Example

    ```bash
   ncli.exe dropcolumns --input-dir input --output-dir processed --columns-file columns_to_drop.txt
    ```
