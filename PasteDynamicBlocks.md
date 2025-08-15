============================================================
HOW TO SET UP EXCEL FOR "PasteDynamicBlocks" VBA MACRO
============================================================

1. **Open the VBA Editor**
   - Press ALT + F11 in Excel.
   - Go to Insert > Module.
   - Paste the entire "PasteDynamicBlocks" code provided into the new module.
   - Press CTRL + S to save the workbook as a **Macro-Enabled Workbook (.xlsm)**.

2. **Set up the Instruction Sheet**
   - Create a new sheet named "Instruction".
   - In column A, you can write labels for reference (optional).
   - In column B, put the required configuration values EXACTLY as below:

   ------------------------------------------------------------
   |   A (label)                      |    B (value)          |
   ------------------------------------------------------------
   | Client source sheet name         | Client                | <-- B2
   | Client source range (no row end) | A2:N                   | <-- B3
   | Client paste start column        | B                      | <-- B4
   | Goldman source sheet name        | Goldman                | <-- B5
   | Goldman source range (no row end)| A2:N                   | <-- B6
   | Goldman paste start column       | M                      | <-- B7
   | Transaction sheet name           | Transaction            | <-- B8
   | Timestamp sheet name             | Config                 | <-- B9
   | Timestamp cell                   | A2                     | <-- B10
   ------------------------------------------------------------

   **NOTES:**
   - "source range" means starting cell to last column, WITHOUT the last row number.
     Example: `A2:N` means the macro will copy from A2 to column N, down to the last row with data.
   - Paste start columns (B4 and B7) should be column letters (e.g., "B", "M").
   - Make sure sheet names match EXACTLY with your workbook.

3. **Set up the Source Sheets**
   - Create (or make sure you have) a sheet for the Client source data.
   - Create (or make sure you have) a sheet for the Goldman source data.
   - In each sheet, fill your data starting from the same row/columns you listed in the Instruction sheet (e.g., starting at A2).
   - Data will be copied from that starting cell to the last non-empty row in the first column of that range.

4. **Set up the Destination Sheet**
   - Create a sheet named exactly as specified in B8 of Instruction (e.g., "Transaction").
   - Column A of this sheet is reserved for timestamps.
   - The macro will paste:
     - Client data starting in the column given in Instruction!B4.
     - Goldman data starting in the column given in Instruction!B7.
   - The macro will also place the timestamp in Column A for every pasted row.

5. **Set up the Timestamp Sheet**
   - Create a sheet named exactly as specified in B9 of Instruction (e.g., "Config").
   - Put your timestamp in the cell specified in B10 (e.g., A2).
   - This value will be repeated in the Transaction sheet for all rows pasted in the run.

6. **Running the Macro**
   - Press ALT + F8 in Excel.
   - Select "PasteDynamicBlocks".
   - Click Run.
   - The macro will:
     1. Find the first empty row in the Transaction sheet (based on column A).
     2. Paste Client data and Goldman data starting at the same row, in their respective columns.
     3. Fill Column A with the timestamp from Config sheet.
     4. Draw a top border above the block and a bottom border below the block.

7. **Future Changes**
   - To change source sheets, paste columns, or ranges: ONLY edit the values in the Instruction sheet.
   - No need to edit VBA code.

============================================================
