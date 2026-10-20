---
name: income-expense-processor
description: Use when processing, aggregating, splitting, and filtering income and expense Excel statement tables (收支流水).
---

# Income Expense Statement Processor (收支流水表格处理)

## Overview
Automates the complete workflow for processing e-commerce / logistics income & expense Excel statements (收支流水表格). Performs text-safe reading, amount parsing, main-fee filtering ("正向配送费"), order aggregation, ratio-based order splitting for large multi-item orders (>180 amount), fee item column renaming, and calculation breakdown tracking.

## When to Use
Use this skill whenever processing, cleaning, aggregating, splitting, or validating income/expense Excel statement tables (收支流水), specifically for requirements such as:
- Reading Excel files with string dtype to prevent 16+ digit Order ID precision loss.
- Cleaning currency symbols (like `CN￥`, `CNY`) to extract numeric amounts with signs.
- Filtering out invalid orders that lack the main fee ("正向配送费").
- Aggregating duplicate Order IDs into single rows.
- Splitting orders with multiple main fees (>180 total amount) into sub-orders (`_1`, `_2`, etc.) and distributing secondary fees proportionally based on main fee ratios.
- Renaming fee item column C (`费用项`) to `总金额(包含项目...)`.
- Appending an explicit calculation breakdown column (`金额计算明细`).
- Preserving original raw data on `Sheet1` and outputting processed results to `Sheet2`.

## Core Processing Rules

1. **Text-Safe Reading (`dtype=str`)**:
   Read Excel inputs with `dtype=str` to ensure Order IDs are never truncated, converted to floats, or formatted in scientific notation.

2. **Amount Cleaning (`变动金额`)**:
   Use regex `[^\d\.-]` to strip currency symbols and non-numeric characters while keeping signs and decimals.

3. **Main Fee Requirement ("正向配送费")**:
   Orders without "正向配送费" are invalid and MUST be removed.

4. **Order Splitting & Proportional Allocation**:
   - Trigger condition: Order contains `>1` "正向配送费" entries and `abs(total_amount) > 180`.
   - Sub-orders created: `_1`, `_2`, ..., `_N`.
   - Secondary fee distribution: Other fees total is split proportionally according to `main_fee_i / total_main_fee`.

5. **Fee Item Column Renaming (`费用项`)**:
   Column C is updated to `总金额(包含项目...)`.

6. **Calculation Breakdown Column (`金额计算明细`)**:
   Append a final column explaining the math behind the total.

## Execution

Run the Python script located in this skill package:

```bash
python3 skills/income-expense-processor/scripts/process_excel.py <input_excel_path> <output_excel_path>
```
