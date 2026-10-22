# Verification performed

54 core regression checks passed using the three supplied marketplace sample files plus synthetic boundary cases.

- Shopee: 5 product rows, including 1 blank SKU replaced with `0` in both SKU columns. First combined ID matched `42611267358~265946333448`.
- Lazada: 6 product rows; the three instruction rows were skipped.
- Zalora: 12 Timberland product rows.
- Text leading zeros were retained. Missing IDs, invalid columns, duplicate required headers, brand mismatches, formula IDs, Excel errors, and unsafe numeric identifiers blocked conversion.
- Duplicate mappings were retained with a count. Entirely empty rows were skipped.
- Export boundaries of 1, 5,198, 5,199, and 12,000 records were verified. The 12,000-record export split into 5,198 / 5,198 / 1,604 records.
- All exported Sheet1 rows were parsed independently with Python's XML parser. The first and last records and row counts matched each batch.
- Both template header rows were retained exactly. All other workbook ZIP entries, including Sheet2 and styles, were byte-identical to the source template.
- Formula-like strings were exported as escaped text, never as formulas.
- JavaScript syntax checks passed.

## Outstanding validation

Browser automation was attempted but could not run because the test browser was unavailable and its download failed in the execution environment. The included `npm run test:browser` suite covers UI selections, stale-result clearing, file/ZIP downloads, mobile overflow, and browser errors. Run it in an environment with Playwright Chromium installed.

No visual browser inspection or live eRetail import has been completed. Verify a small export through the normal eRetail import workflow before using bulk outputs.

Original marketplace sample files are intentionally excluded from the source project.
