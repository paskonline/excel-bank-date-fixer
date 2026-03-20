# excel-bank-date-fixer
An Excel Office Script to fix inconsistent date formats in bank statement exports using chronological logic and transaction reference codes.

# Excel Office Script: Smart Bank Statement Date Fixer

### The Problem
Many bank systems export CSVs with inconsistent date formats (e.g., mixing `DD/MM/YYYY` with `YYYY-DD-MM`). Standard Excel formatting often fails to recognize these or swaps months and days incorrectly.

### The Solution
This Office Script uses a two-step verification to fix dates:
1. **Transaction Reference Check:** It extracts dates from standard "E-codes" (YYMMDD) found in bank descriptions.
2. **Chronological Logic:** It compares the current row with the previous row to ensure the date sequence makes sense (e.g., preventing a jump from October to January if the intended month was November).

### How to use
1. Copy the code from `DateFixer.ts`.
2. In Excel (Web or Desktop), go to the **Automate** tab -> **New Script**.
3. Paste the code and click **Run**.
