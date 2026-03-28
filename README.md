# Python Transaction Reconciliation Automation

This project automates a multi-step transaction reconciliation workflow using Python, Pandas, and OpenPyXL.

## Overview
The workflow prepares transaction sheets, loads opening balances, calculates weighted average cost, and validates closing balances.

## Files
- `01_prepare_transactions.py` – standardizes transaction data and loads it into the workbook
- `02_load_opening_balances.py` – fills opening balances by security code
- `03_calculate_weighted_average.py` – calculates weighted average cost and ending balances
- `04_validate_closing_balances.py` – validates closing balances and flags exceptions

## Tools
- Python
- Pandas
- NumPy
- OpenPyXL

## Project Structure
- `input/` – input Excel files such as opening, closing, and transaction data
- `output/` – generated transaction workbook and validation outputs

## Notes
This repository contains sample code only. Confidential business data and internal file paths have been removed.