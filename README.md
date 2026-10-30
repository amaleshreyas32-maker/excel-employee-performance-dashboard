# Excel Employee Performance Dashboard

An Excel MIS dashboard with VBA macros that turns raw employee call-center data into a single summary report with interactive charts.

## Overview
The workbook consolidates data for **149 employees** from three raw data sheets into one Summary sheet. It then presents working hours and customer-active time in Pivot Tables and charts on a Dashboard sheet.

## Workbook Structure

| Sheet | Purpose |
|---|---|
| **Break** | Break details per employee: lunch, normal break, feedback, training, meeting, unwanted break, total break taken |
| **Efficiency** | Total working hours, work efficiency, attended calls, active time with customer, not-working %, hold details |
| **OverAll Performance** | Attended, transferred, conferenced and redirected calls, hold calls, average call duration |
| **Summary** | Combines all three sheets into one table using 1,000+ formulas |
| **Dashboard** | Pivot Tables, charts and macro buttons for reporting |

## Key Features
- **Data consolidation:** VLOOKUP with named ranges pulls data from three sheets into one Summary sheet.
- **Data cleaning:** MID and FIND extract clean employee names from IDs like `Mum_Name (ID)`.
- **Time calculations:** TIMEVALUE converts text time values into numbers that can be summed and charted.
- **Pivot Tables and charts:** Report 1 shows Working Hours and a second report shows Active Time with Customer.
- **VBA automation:** Spinner controls filter the Top-N employees, and macro buttons show or hide charts.

## Skills Used
MS Excel, VLOOKUP, Pivot Tables, Pivot Charts, Named Ranges, TIMEVALUE, MID/FIND, VBA Macros, Form Controls, Data Validation, MIS Reporting, Dashboard Design

## How to Use
1. Download `Dashboard.xlsm`.
2. Open it in Microsoft Excel.
3. Click **Enable Content** when asked, so the macros can run.
4. Open the **Dashboard** sheet and use the spinner buttons to change the Top-N filter.

## Screenshots
Add screenshots of the Dashboard sheet here.

## Note
All employee names and IDs in this file are sample data.

## Author
**Shreyas Surendra Amale**
M.Sc. Computer Science | [GitHub](https://github.com/amaleshreyas32-maker) | [LinkedIn](https://linkedin.com/in/shreyasamale32)
