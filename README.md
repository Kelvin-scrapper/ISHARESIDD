# ETF Effective Duration Data Pipeline

Automated system for extracting, mapping, and managing ETF Effective Duration data from iShares websites.

## 📁 Project Structure

```
ISHARESIDD/
├── orchestrator.py              # Main orchestrator - runs the complete pipeline
├── main.py                      # Step 1: Web scraping - extracts ETF data
├── mapping_code.py              # Step 2: Data mapping with dual date logic
├── requirements.txt             # Python dependencies
├── README.md                    # This file
├── ETF_Effective_Duration.xlsx  # Output from main.py (scraped data)
└── ISHARESIDD_DATA_.xlsx        # Final output with mapped data
```

## 🚀 Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Run Complete Pipeline
```bash
python orchestrator.py
```

This will:
1. Extract ETF data from iShares websites → `ETF_Effective_Duration.xlsx`
2. Map data to ISHARESIDD_DATA_.xlsx with dual date logic

### Run Individual Components

**Extract ETF Data Only:**
```bash
python main.py
```

**Map Data Only (if extraction already done):**
```bash
python mapping_code.py
```

## 📊 Pipeline Overview

### Step 1: Data Extraction (main.py)
- Scrapes 15 ETF websites from iShares
- Extracts: ETF Name, Effective Duration, As of Date
- Output: `ETF_Effective_Duration.xlsx`

### Step 2: Data Mapping (mapping_code.py)
Implements **DUAL DATE LOGIC**:

#### **STEP 1: Historical Data Correction**
- Reads "As of Date" from source file (e.g., 2025-10-15)
- Checks if that historical date exists in output
- If found: Compares and **corrects/updates** any changes
- Ensures historical data accuracy

#### **STEP 2: Yesterday's Date Mapping**
- Always maps to **YESTERDAY'S DATE** (today - 1 day)
- Creates new row if doesn't exist
- Updates existing row if present
- **Zero value handling**: Uses previous day's data for 0/invalid values

## 🎯 Key Features

### ✅ Robust Data Handling
- **Hardcoded headers** - No dependency on existing files
- **Automatic zero handling** - Invalid values replaced with previous data
- **Change detection** - Automatically detects and corrects data changes
- **Dual mapping** - Historical correction + current tracking

### ✅ Smart Date Logic
- Uses "As of Date" from source to correct historical data
- Maps to yesterday's date for current tracking
- Handles stale/outdated source data gracefully

### ✅ Data Validation
- Validates numeric values before writing
- Detects 0 or invalid duration values
- Fallback to previous day's data when needed

## 📝 Output Format

### ISHARESIDD_DATA_.xlsx Structure

**Row 1:** Technical identifiers (e.g., `ISHARESIDD.JPMEMBONDSETF.EFFECTDUR.B`)
**Row 2:** ETF friendly names (e.g., `JPM EM Bonds ETF`)
**Row 3+:** Date + data rows

| Date       | JPM EM Bonds | Iboxx High Yield | ... |
|------------|--------------|------------------|-----|
| 2025-10-15 | 6.88         | 2.86             | ... |
| 2025-10-16 | 6.88         | 2.86             | ... |

## 🔧 Configuration

### ETF List (in main.py)
15 ETFs currently configured:
- JPM EM Bonds ETF
- Iboxx High Yield Corp Bond ETF
- Iboxx IG Corp Bond ETF
- TIPS Bond ETF
- 20 Yr Treasury Bond ETF
- 7-10 Year Treasury Bond ETF
- iShares J.P. Morgan $ EM Bond UCITS ETF
- iShares $ Treasury Bond 1-3yr UCITS ETF
- iShares $ High Yield Corp Bond UCITS ETF
- iShares € Corp Bond 1-5yr UCITS ETF
- iShares € High Yield Corp Bond UCITS ETF
- iShares Core € Corp Bond UCITS ETF
- iShares MBS ETF
- iShares Short-Term Corporate Bond ETF
- iShares Intermediate-Term Corporate Bond ETF

### Headers (in mapping_code.py)
Hardcoded in `HEADER_ROW_1` and `HEADER_ROW_2` (lines 17-54)

## 📋 Dependencies

All dependencies are listed in `requirements.txt`:
- **openpyxl** - Excel file handling
- **selenium** - Web automation
- **undetected-chromedriver** - Chrome driver for web scraping
- **beautifulsoup4** - HTML parsing
- **pandas** - Data processing
- **lxml, html5lib** - HTML parsing support

Install all at once:
```bash
pip install -r requirements.txt
```

## 🔄 Workflow Example

**Day 1 (Oct 17):**
```
Source: As of Date = Oct 15
Output:
  Row 3: Oct 15 (historical, corrected if needed)
  Row 4: Oct 16 (yesterday's data)
```

**Day 2 (Oct 18):**
```
Source: As of Date = Oct 16
Output:
  Row 3: Oct 15 (unchanged)
  Row 4: Oct 16 (historical, corrected if needed)
  Row 5: Oct 17 (yesterday's new data)
```

## 🛠️ Error Handling

- **Missing source file**: Graceful error with clear message
- **Invalid duration values**: Replaced with previous day's data
- **Historical data mismatch**: Automatically corrected
- **Script failures**: Orchestrator reports detailed error information

## 📌 Notes

- The code always uses **yesterday's date** for current tracking
- Historical data is automatically reviewed and corrected on each run
- Zero or invalid values are handled by using previous day's data
- All headers are hardcoded for independence and reliability

## 🎓 Maintenance

To add new ETFs:
1. Add to `etf_data` list in `main.py`
2. Add headers to `HEADER_ROW_1` and `HEADER_ROW_2` in `mapping_code.py`

To modify date logic:
- Edit `create_mapping()` function in `mapping_code.py`

---

**Created:** October 2025
**Last Updated:** October 17, 2025
