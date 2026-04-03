# OKZ Data Validation Framework

## Overview
This project is a Python-based data validation framework designed to validate large-scale data extracts between source and target systems.

It compares datasets at both:
- Row level (record comparison)
- Column level (data consistency)

## Key Features
- Automated data validation using Python (pandas)
- Handles header inconsistencies and data mismatches
- Generates:
  - Matched records report
  - Mismatch summary report
- Supports real-world ETL/data pipeline validation scenarios

## Tech Stack
- Python
- Pandas
- Excel (openpyxl)

## Use Case
Used to validate data between:
- Source system (Grant Tracking)
- Target system (OKZ Extract)

Ensures:
- Data completeness
- Data accuracy
- Transformation correctness

## How to Run
```bash
python src/validate_common_columns_ignore_known_header_defects.py