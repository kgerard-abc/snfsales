# Lincolnwood's rows from CMS's original cost report files

These files hold every row CMS publishes for Lincolnwood Rehabilitation and Healthcare Center (CMS number 415035, formerly Hopkins Manor, Ltd.) in its Medicare cost reports for 2015 to 2024. The rows and values are copied unchanged from CMS's Healthcare Cost Report Information System (HCRIS) files. The only thing added is the header row at the top of each file, using CMS's own column names, because CMS's files have no header row.

## Where the originals are

CMS posts one zip file per fiscal year on its [Cost Reports by Fiscal Year](https://www.cms.gov/data-research/statistics-trends-and-reports/cost-reports/cost-reports-fiscal-year) page, for example `https://downloads.cms.gov/FILES/HCRIS/SNF10FY2024.ZIP`. Each zip holds three CSV files covering every skilled nursing facility in the country. Sizes below are for the fiscal year 2022 zip:

| File in the zip | What it holds | Size, unzipped |
| --- | --- | --- |
| SNF10_2022_rpt.csv | One row per cost report: record number, CMS number, period, status | about 1.6 MB |
| SNF10_2022_nmrc.csv | Every number on every worksheet | about 497 MB |
| SNF10_2022_alpha.csv | Every piece of text on every worksheet (names, cost centers, expense items) | about 158 MB |

The number and text files are too large for Excel, which is why these smaller copies exist. To check any row here against CMS's original, open the zip and search the file for the report's record number (for 2024, `1435496`).

## How to read a row

Each row in a `_nmrc` or `_alpha` file is one cell on one worksheet of the cost report:

`1435496,A810001,00100,00500,4397809`

| Field | Value | Meaning |
| --- | --- | --- |
| RPT_REC_NUM | 1435496 | The report's CMS record number (Lincolnwood, 2024) |
| WKSHT_CD | A810001 | Worksheet A-8-1, Part A |
| LINE_NUM | 00100 | Line 1 (the last two digits are a sub-line: 00102 is line 1.02) |
| CLMN_NUM | 00500 | Column 5 |
| ITM_VAL_NUM | 4397809 | The amount: $4,397,809 |

The `_alpha` file for the same report has `1435496,A810001,00100,00300,RENT EXPENSE`: line 1, column 3 of the same worksheet, so line 1 is rent of $4,397,809.

Worksheet codes used in the readable version (`../415035.md`): A810001 is Worksheet A-8-1 Part A, A810002 is A-8-1 Part B, G300000 is Worksheet G-3, S200001 is Worksheet S-2 Part I.

## The files

One set of three files per report, named by the last day of the report period and the record number:

| Period | Owner filing | Record number |
| --- | --- | --- |
| January 1 to December 31, 2015 | Hopkins Manor, Ltd. | 1125470 |
| January 1 to December 31, 2016 | Hopkins Manor, Ltd. | 1157517 |
| January 1 to December 31, 2017 | Hopkins Manor, Ltd. | 1207384 |
| January 1 to December 31, 2018 | Hopkins Manor, Ltd. | 1237308 |
| January 1 to December 31, 2019 | Hopkins Manor, Ltd. | 1262086 |
| January 1 to December 31, 2020 | Hopkins Manor, Ltd. | 1284720 |
| January 1 to March 7, 2021 | Hopkins Manor, Ltd. | 1305406 |
| March 8 to December 31, 2021 | Lincolnwood Operator LLC | 1331581 |
| January 1 to December 31, 2022 | Lincolnwood Operator LLC | 1375728 |
| January 1 to December 31, 2023 | Lincolnwood Operator LLC | 1420853 |
| January 1 to December 31, 2024 | Lincolnwood Operator LLC | 1435496 |

Made with `pipeline/scripts/cost_report_extract.py 415035`.
