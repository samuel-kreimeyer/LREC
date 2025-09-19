# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is the Little Rock Engineers Club (LREC) repository containing Python scripts for automating club administration tasks. The primary functions include generating PDF certificates of attendance and creating meeting notification messages from spreadsheet data.

## Commands

### Environment Setup
```bash
# Activate virtual environment (required before running any scripts)
source venv/bin/activate
```

### Certificate Generation
```bash
# Generate PDF certificates from Excel spreadsheet
python scripts/generate_coas.py
# Interactive prompt will ask for spreadsheet path (default: COA forms.xlsx)
```

### Notice Generation
```bash
# Generate meeting notices from event data
python scripts/generate_notices.py sample_events.csv --output notices.txt

# With optional speaker bio and lunch message
python scripts/generate_notices.py sample_events.csv --bio "Speaker bio text" --lunch-provided --output notices.txt
```

## Architecture

### Certificate Generation (`scripts/generate_coas.py`)
- **Input**: Excel spreadsheet with columns: Name, Speaker, Title, Date
- **Output**: Individual PDF certificates named `COA_{Name}_{Date}.pdf`
- **Dependencies**: pandas, openpyxl, reportlab
- Uses ReportLab to create landscape PDFs with club branding and skyline image
- Certificates include 1 PDH (Professional Development Hour) validation

### Notice Generation (`scripts/generate_notices.py`)
- **Input**: CSV/Excel with columns: date, topic, speaker, location, time
- **Template**: `scripts/notice_template` (Jinja2 format)
- **Output**: Text file with formatted meeting invitation
- Automatically selects next future meeting based on current date
- Supports optional speaker bio and customizable lunch messages

### Data Structure
- `PII/`: Contains membership rosters and sensitive club data
- `scripts/`: Python automation scripts and sample data files
- `venv/`: Python virtual environment with required dependencies
- Root level: Contains working Excel/PDF files and email templates

### Dependencies
The project uses a Python virtual environment with:
- pandas (Excel/CSV processing)
- openpyxl (Excel file support)
- reportlab (PDF generation)
- jinja2 (template rendering)

## File Formats

### COA Input Format (Excel/CSV)
```
Name, Speaker, Title, Date
Jon Doe, Jane Doe, Innovations in Engineering, September 9 2026
```

### Events Input Format (CSV/Excel)
```
date,topic,speaker,location,time
2025-10-15,Next Meeting Topic,Dr. Jane Smith,Engineering Building Room 101,6:00 PM
```