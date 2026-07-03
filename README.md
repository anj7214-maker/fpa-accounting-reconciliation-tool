# FPA Monthly Accounting Reconciliation Tool

This is a focused monthly reconciliation app for filling only two Excel columns from one or more Tally ledger PDFs:

- `Tally Date`
- `Tally Voucher`

The goal is to explain every Tally PDF voucher safely. It does not try to force every Excel row to match.

It reads the student code, student name, payment `Date`, and `Fees Received with GST` from the workbook, extracts Tally PDF transactions, and matches in this order by default:

1. Student code + Fees Received with GST amount + payment date
2. Student code + Fees Received with GST amount
3. Unique student code

The extractor recognises FPA course code families such as `IBOC`, `ACCA`, `CBE`, `CMA`, `CET`, `WORK`, `O`, and short ACCA legacy codes such as `A998`.

Review mode can also try:

4. Student code + exact student name
5. Exact student name + amount
6. Fuzzy student name above 95% + amount

If there are multiple possible PDF transactions for the same student, the workbook row is left unchanged and the case is sent to the duplicate review report.

## Run

### App

```powershell
python -m streamlit run streamlit_app.py
```

The app has three tabs:

- `PDF Extract`: upload one or more Tally PDFs and download one combined Excel file with `Tally Date`, `Tally Voucher`, and source PDF details.
- `Reconciliation`: upload one or more Tally PDFs and the Excel workbook to update only the workbook's `Tally Date` and `Tally Voucher` columns.
- `Post to Income Sheet`: upload a verified reconciliation sheet and the final income sheet, then copy only `Tally Date` and `Tally Voucher` into `IS25-26` columns `AK` and `AL`.

The posting tab is intentionally conservative:

- It previews posting metrics before downloads are shown.
- It does not overwrite existing `AK`/`AL` values unless enabled.
- It highlights review rows when requested.
- It generates posting, review/discrepancy, and summary reports.

### Command line

```powershell
python fpa_monthly_reconcile.py --pdf "C:\path\to\tally-ledger-1.pdf" "C:\path\to\tally-ledger-2.pdf" --excel "C:\path\to\student-data.xlsx" --output-dir outputs\monthly_reconcile
```

To enable review mode name fallback from the command line:

```powershell
python fpa_monthly_reconcile.py --pdf "C:\path\to\tally-ledger-1.pdf" "C:\path\to\tally-ledger-2.pdf" --excel "C:\path\to\student-data.xlsx" --output-dir outputs\monthly_reconcile --match-mode review
```

## Outputs

The output folder will contain:

- updated Excel workbook
- `matched_report.xlsx`
- `excel_unmatched_report.xlsx`
- `pdf_unmatched_report.xlsx`
- `duplicate_manual_review_report.xlsx`
- `summary_report.xlsx`

The script preserves the original workbook by saving a new updated file. It only writes to the detected `Tally Date` and `Tally Voucher` columns.

The summary separates Excel-side and PDF-side status:

- PDF transactions extracted
- PDF files uploaded
- PDF transactions matched
- PDF transactions unmatched
- PDF transactions needing manual review
- Excel rows scanned
- Excel rows updated
- Excel rows unmatched
- Duplicate/manual review count

## Notes

The script can run without `rapidfuzz`; it falls back to Python's built-in name similarity. Install the packages in `requirements.txt` for the preferred setup.
