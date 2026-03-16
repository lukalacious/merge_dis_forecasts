# Merge Distributor Forecasts

Consolidates seasonal assortment forecast Excel files from multiple ASICS EMEA distributors into a single standardised spreadsheet.

## What It Does

Each distributor (ALS, INX, KEY, MAP, MAR, MNT, etc.) submits a seasonal sell-in forecast as a separate Excel file. This tool:

1. Dynamically discovers all Excel files in a folder
2. Reads the "ASSORTMENT" sheet from each
3. Standardises headers and tags rows with the distributor name
4. Concatenates everything into one consolidated file
5. Exports a timestamped Excel output

## Files

| File | Description |
|------|-------------|
| `Merge DIS Forecasts.ipynb` | Interactive notebook with distributor row-count visualisations |
| `Merge DIS Forecasts - script.py` | Production CLI version with checkpoint logging and error handling |

## Tech Stack

Python, pandas, matplotlib, seaborn, openpyxl

## Usage

```bash
python "Merge DIS Forecasts - script.py"
```

Place distributor Excel files in the input folder. The script auto-discovers and processes all `.xlsx` files.
