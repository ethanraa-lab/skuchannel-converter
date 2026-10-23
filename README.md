# SkuChannel Converter

A standalone browser app that converts Shopee, Lazada, and Zalora exports into the supplied eRetail SkuChannel template. Files are read and processed in browser memory. No product data is sent to a server, saved in local storage, or included in this repository. All runtime assets are bundled locally; there are no CDN dependencies or analytics.

## Run

Serve this directory with any static HTTP server. With Python installed:

```sh
python -m http.server 8080 --bind 127.0.0.1
```

Open http://localhost:8080. Opening index.html directly as a file will not work because browsers block loading the template over file://.

Choose marketplace, country, and brand, select the export workbook and worksheet, click **Convert & preview**, then download individual files or a ZIP. A marketplace/brand selection applies to the entire input file. For Zalora, a populated Brand column is checked against the selected brand.

## Mapping

| Marketplace | eRetail SKU and Channel SKU | Channel Product Id |
|---|---|---|
| Shopee | SKU | Product ID + `~` + Variation ID |
| Lazada | SellerSKU | `sku.skuId` (including hidden columns) |
| Zalora | SellerSku | CatalogProductId |

| Marketplace | Columbia | Vans | Timberland |
|---|---|---|---|
| Shopee | M01 | M02 | M03 |
| Lazada | M04 | M06 | M05 |
| Zalora | — | M19 | M16 |

MY and SG share these codes. Country is used in filenames only. Blank or whitespace-only SKU values become the text `0` in both SKU columns. Status is always `Active`, Mode is always `Create`, and other output data columns remain blank, including all stock and price fields. This tool creates SKU-channel mappings; it does not update inventory quantities.

Headers are matched by trimmed, case-insensitive names, independently of column positions. Lazada's three metadata rows are skipped when its Optional/Mandatory row is detected. Entirely empty rows are ignored. Duplicate mappings are retained and counted in a warning. Records are never silently deduplicated or removed based on source status.

## Template preservation and format

`template.xls` is the supplied blank template. Despite its .xls filename, its contents are an Office Open XML (ZIP/.xlsx) workbook. Export deliberately preserves the same internal format and .xls extension used by the supplied template. Excel may show an extension/format mismatch notice, as it can for the original. This is not a BIFF .xls conversion.

The exporter replaces only Sheet1's product rows and its used-range dimension. Both original header rows, their styles, column widths, validations, Sheet2, and all other workbook package entries are preserved. Each file contains at most **5,198 data rows + 2 header rows = 5,200 Sheet1 rows**. Sheet2 is a template support sheet and is excluded from the Sheet1 row limit.

Filename: `SkuChannel_Shopee_MY_Vans_Part01.xls`.

Input supports modern XML/ZIP Excel workbooks (.xlsx and equivalent files named .xls), up to 30 MB compressed and 160 MB expanded. Legacy binary .xls, encrypted Excel files, and CSV files must first be saved as .xlsx in Excel. Identifier formulas, Excel error cells, missing IDs, scientific-notation numeric identifiers, and numeric identifiers longer than 15 digits block conversion. Text identifiers and leading zeros are preserved. Precision already lost in a source file cannot be recovered. Missing Shopee Variation IDs must be resolved in the source; the tool never invents one.

## Verification

```sh
npm install
npx playwright install chromium
npm test
npm run test:browser
```

Core tests use Node.js and Python 3 for independent XML verification. The browser test runner starts its own local static server and uses synthetic workbooks. To also verify the original marketplace samples without adding them to the repository:

```sh
SAMPLE_DIR=/absolute/path/to/samples npm test
```

Sample filenames: ShopeeLayoutExample.xlsx, LazadaLayoutExample.xlsx, ZaloraLayoutExample.xlsx. Tests cover mapping, missing SKUs, leading zeros, unsafe IDs, missing columns/IDs, Lazada instructions, XML/formula-like strings, stale results, file and ZIP downloads, 12,000-row splitting, and byte preservation of unchanged template components.

The files have not been imported into a live eRetail instance. Confirm acceptance with a small import through your normal company process before bulk use.

## GitHub

Use a separate repository named `skuchannel-converter`. Do not upload marketplace exports or credentials. This project can run locally, on an approved internal static host, or on static hosting if authorised. No deployment is configured automatically.

## Third-party dependency

JSZip 3.10.1 is vendored under `vendor/`; its MIT license is included. Playwright is a development-only test dependency.
