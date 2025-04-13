# automatedreporting
# Automated Financial Reporting System

## Overview
This Python script automates monthly financial reporting by processing 50 financial transactions, producing a multi-sheet Excel report with VBA macros, a CSV summary, and a Google Drive backup. It features multi-currency support, anomaly detection, cohort analysis, and PII masking, ensuring a robust, Power BI-compatible output. Descriptions are realistic, enhancing professional appeal for job applications.

## Features
- **Data Processing**: Handles 50 transactions (sales, expenses, returns) with USD, EUR, GBP currencies.
- **Realistic Descriptions**: Uses category-specific templates (e.g., "Sale of office equipment").
- **Data Validation**: Manually enforces schema (e.g., amount ≤ $100,000).
- **Anomaly Detection**: Isolation Forest identifies outliers (5% contamination).
- **Financial Analytics**: Computes revenue, expenses, profit, current ratio, and cohort trends.
- **Excel Report**: Includes raw data, summary, category aggregates, cohort analysis, and pivot table.
- **VBA Macros**: Provides `format_report.vba` for formatting and pie chart creation.
- **CSV Summary**: Outputs `financial_summary.csv` with key metrics.
- **Cloud Backup**: Saves Excel to Google Drive (`/Reports/financial_report.xlsx`).
- **Security**: Masks PII (emails, names) in descriptions.
- **Output**: Generates downloadable `financial_report.xlsx`, `financial_summary.csv`, and `format_report.vba`.
- **Observability**: Logs execution and anomalies.

## Prerequisites
- **Environment**: Google Colab.
- **Dependencies**: `pandas`, `openpyxl`, `xlsxwriter`, `faker`, `scikit-learn`.
- **Internet**: For installation, downloads, and Google Drive.
- **Local Excel**: For VBA macros (post-.xlsm conversion).
- **Google Account**: For Drive backup.

## Usage
1. Copy and execute the script in a Google Colab notebook.
2. The pipeline will:
   - Load 50 transactions for March 2025 with realistic descriptions.
   - Detect anomalies and mask PII.
   - Process data into metrics and cohort analysis.
   - Create `financial_summary.csv`, downloadable.
   - Create `financial_report.xlsx`, downloadable.
   - Produce `format_report.vba`, downloadable.
   - Back up Excel to Google Drive.
3. For VBA macros:
   - Open `financial_report.xlsx` in Excel.
   - Save As `financial_report.xlsm`.
   - Press Alt+F11, insert module, paste `format_report.vba`.
   - Run `FormatReport` to format tables and add pie chart.
4. Import `financial_report.xlsx` into Power BI for visuals.
5. Download `reporting_system.log` for details.

## Data Source
- **Transactions**: 50 records with:
  - `transaction_id`: E.g., TXN-1001.
  - `date`: March 2025.
  - `category`: Sales, Expenses, Returns.
  - `description`: Realistic note (e.g., "Payment for office rent").
  - `amount`: Original currency value.
  - `currency`: USD, EUR, GBP.
  - `base_amount`: USD equivalent.
  - `vendor`: Company name.

## Output
- **financial_report.xlsx**:
  - **Raw Data**: 50 transactions.
  - **Summary**: Revenue, expenses, profit, ratio.
  - **Category Summary**: Totals, counts, vendors.
  - **Cohort Analysis**: Weekly amounts.
  - **Pivot**: Category-vendor matrix.
- **financial_summary.csv**: Key metrics.
- **format_report.vba**: VBA code.
- **Google Drive**: Backup at `/Reports/financial_report.xlsx`.

## Technical Details
- **Data**: Loads transactions with realistic descriptions.
- **Validation**: Manual checks ensure integrity.
- **Anomaly Detection**: Isolation Forest flags outliers.
- **Analytics**: Pandas computes ratios and cohorts.
- **Excel**: OpenPyXL creates sheets with formatting.
- **VBA**: Formats tables and adds charts locally.
- **CSV**: Summarizes metrics.
- **Cloud**: Google Drive backup.
- **Security**: Regex masks PII.
- **Logging**: Tracks execution and errors.

## Customization
- **Real Data**: Integrate with QuickBooks API.
- **Metrics**: Add debt ratio or forecasts.
- **Reports**: Include charts in Excel.
- **Cloud**: Use AWS S3.
- **Security**: Add encryption.

## Limitations
- **Colab**: Outputs are temporary unless backed up.
- **VBA**: Requires local Excel for macros.
- **Data**: Simulated; needs APIs for real data.

## Future Enhancements
- Add accounting API integration.
- Support Power BI API publishing.
- Enhance Excel with embedded charts.
- Deploy to cloud for automation.

## Notes
This system demonstrates financial automation, data validation, and analytics, ideal for BI roles. Realistic descriptions enhance its professional appeal for job applications.

For further information, please contact me on linkedin at https://www.linkedin.com/in/edward-antwi-8a01a1196/
"""
