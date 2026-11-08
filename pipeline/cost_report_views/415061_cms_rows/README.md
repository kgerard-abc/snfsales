# Brentwood's rows from CMS's original cost report files

These files hold every row CMS publishes for Brentwood Nursing Home, now Brentwood Health Center (CMS number 415061), in its Medicare cost reports for 2015 to 2025. The rows and values are copied unchanged from CMS's Healthcare Cost Report Information System (HCRIS) files. The only thing added is the header row at the top of each file, using CMS's own column names, because CMS's files have no header row.

The 2025 report is on CMS's newer form, CMS-2540-24 (file SNF24FY2025.ZIP). Its worksheet codes are the same, but some columns are numbered differently: on Worksheet A-8-1, the amount the home recorded is in column 6 instead of column 5.

## Where the originals are

CMS posts one zip file per fiscal year on its [Cost Reports by Fiscal Year](https://www.cms.gov/data-research/statistics-trends-and-reports/cost-reports/cost-reports-fiscal-year) page, for example `https://downloads.cms.gov/FILES/HCRIS/SNF10FY2024.ZIP`. Each zip holds three CSV files covering every skilled nursing facility in the country. Sizes below are for the fiscal year 2022 zip:

| File in the zip | What it holds | Size, unzipped |
| --- | --- | --- |
| SNF10_2022_rpt.csv | One row per cost report: record number, CMS number, period, status | about 1.6 MB |
| SNF10_2022_nmrc.csv | Every number on every worksheet | about 497 MB |
| SNF10_2022_alpha.csv | Every piece of text on every worksheet (names, cost centers, expense items) | about 158 MB |

The number and text files are too large for Excel, which is why these smaller copies exist. To check any row here against CMS's original, open the zip and search the file for the report's record number (for 2024, `1435504`).

## How to read a row

Each row in a `_nmrc` or `_alpha` file is one cell on one worksheet of the cost report:

`1435504,A810001,00100,00500,1964709`

| Field | Value | Meaning |
| --- | --- | --- |
| RPT_REC_NUM | 1435504 | The report's CMS record number (Brentwood, 2024) |
| WKSHT_CD | A810001 | Worksheet A-8-1, Part A |
| LINE_NUM | 00100 | Line 1 (the last two digits are a sub-line: 00102 is line 1.02) |
| CLMN_NUM | 00500 | Column 5 |
| ITM_VAL_NUM | 1964709 | The amount: $1,964,709 |

The `_alpha` file for the same report has `1435504,A810001,00100,00200,CAP REL COSTS - BLDGS & FIXTURES` and `1435504,A810001,00100,00300,INTEREST EXPENSE`: line 1, columns 2 and 3 of the same worksheet, so line 1 is a building cost of $1,964,709 recorded as paid to a related organization.

Worksheet codes used in the readable version (`../415061.md`): A810001 is Worksheet A-8-1 Part A, A810002 is A-8-1 Part B, G300000 is Worksheet G-3, S200001 is Worksheet S-2 Part I.

## The files

One set of three files per report, named by the last day of the report period and the record number:

| Period | Owner during the period | Record number | CMS file |
| --- | --- | --- | --- |
| January 1 to December 31, 2015 | Brentwood Health Center, Ltd. | 1137882 | SNF10FY2015 |
| January 1 to December 31, 2016 | Brentwood Health Center, Ltd. | 1173999 | SNF10FY2016 |
| January 1 to December 31, 2017 | Brentwood Health Center, Ltd. | 1207159 | SNF10FY2017 |
| January 1 to December 31, 2018 | Brentwood Health Center, Ltd. | 1237707 | SNF10FY2018 |
| January 1 to December 31, 2019 | Brentwood Health Center, Ltd. | 1276235 | SNF10FY2019 |
| January 1 to December 31, 2020 | Brentwood Health Center, Ltd. | 1301305 | SNF10FY2020 |
| January 1 to December 31, 2021 | Brentwood Health Center, Ltd. | 1331348 | SNF10FY2021 |
| January 1 to June 6, 2022 | Brentwood Health Center, Ltd. | 1363748 | SNF10FY2022 |
| June 7 to December 31, 2022 | 4000 Post Road Eden Operations LLC | 1363749 | SNF10FY2022 |
| January 1 to December 31, 2023 | 4000 Post Road Eden Operations LLC | 1391251 | SNF10FY2023 |
| January 1 to December 31, 2024 | 4000 Post Road Eden Operations LLC | 1435504 | SNF10FY2024 |
| January 1 to December 31, 2025 | 4000 Post Road Eden Operations LLC | 16786 | SNF24FY2025 |

Made with `pipeline/scripts/cost_report_extract.py 415061 SNF10 SNF24`.
