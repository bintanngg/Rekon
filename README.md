# Rekon v1.1

**A Python utility for automated bank account reconciliation**

[Download the latest version](https://github.com/bintanngg/Rekon/releases/tag/v1.1)

Made By Bintanngg

---

## Overview

Bank Reconciliation Tool is a simple yet powerful application designed to help accountants and finance teams quickly identify discrepancies in bank transactions. It automatically compares debit and credit entries grouped by transaction reference, making it easy to spot posting errors and reconciliation issues.

### Why Use This?
- **Save Time**: Automate tedious manual reconciliation
- **Reduce Errors**: Catch discrepancies automatically
- **User-Friendly GUI**: Intuitive graphical interface
- **Excel Integration**: Works directly with your Excel files

## Features
 **Automatic Reconciliation** - Groups transactions by reference and compares totals  
 **Smart Detection** - Identifies transactions where total debit ≠ total credit  
 **Excel Reports** - Exports detailed results with multiple views:
  - Problem Transactions: Full details of all mismatched entries
  - Summary Report: Quick overview of discrepancies by transaction

 **User-Friendly GUI** - Intuitive graphical interface with dropdowns
 **Data Validation** - Automatic checks for file format and data types

## Quick Start

### Requirements
- Python 3.7+
- pip

## Technical Stack
- **Python** 3.7+ - Core language
- **Pandas** - Data processing
- **Tkinter** - GUI framework (built-in)
- **OpenPyXL** - Excel handling

## License
MIT License - feel free to use and modify

### Installation

```bash
# Clone the repository
git clone https://github.com/bintanngg/Reconciliation.git
cd Reconciliation

# Install dependencies
pip install -r requirements.txt
```

### Usage

**GUI Version:**

Open executable file or

```bash
python Rekon_v_1_1.py
```

1. Click **"Pilih File Excel"** to choose your data file
2. Select columns from dropdowns:
   - **Kolom Transaksi (Keterangan)**: ID or description identifying the transaction
   - **Kolom Debet**: Debit amounts
   - **Kolom Kredit**: Credit amounts
3. Click **"Proses Rekonsiliasi"**
4. View results and download Excel report

## Data Format

### Input Requirements
Your Excel file should contain:
- **Transaction Reference** (any column name): Identifies related transactions
- **Debit**: Numeric values for debit entries
- **Credit**: Numeric values for credit entries

Example:
| Reference | Debit | Credit | Date |
|-----------|-------|--------|------|
| INV001 | 1000 | 0 | 2025-01-15 |
| INV001 | 0 | 1000 | 2025-01-16 |
| INV002 | 500 | 0 | 2025-01-15 |
| INV002 | 0 | 600 | 2025-01-16 |

### Output
Generated `rekonsiliasi_output.xlsx` (or custom name) contains:
- **Sheet 1 - Normal Trx**: All rows with normal transactions
- **Sheet 2 - Problem Trx**: All rows with discrepancies
- **Sheet 3 - Difference Summary**: Aggregated debit, credit, and difference by reference

## Common Issues & Solutions

**Column not found?**
→ Check exact column names in your Excel file (case-sensitive)

**Non-numeric error?**
→ Ensure debit/credit columns contain only numbers

**Module not found?**
→ Run `pip install -r requirements.txt`

**GUI won't open?**
→ Test tkinter: `python -m tkinter`
