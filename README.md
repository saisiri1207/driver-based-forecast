# Driver-based forecast

Driver-based P&L for a fictional CPG company (**Northline Consumer Products**): volume × price by product line, COGS from unit-cost drivers, OpEx by cost center, plan vs reforecast.

**Deliverable:** [`Northline_Driver_Based_Forecast.xlsx`](Northline_Driver_Based_Forecast.xlsx)

`build.py` only regenerates that workbook. The file a hiring manager should open is the `.xlsx`.

## Business question

If volume or unit cost moves, what happens to the P&L versus plan — and which walk explains it?

## How to review this file (8 minutes)

1. Open `01_Assumptions`. Yellow / blue cells are inputs (scenario, prices, unit costs, volumes, OpEx).
2. Flip Base / Upside / Downside. Confirm only the forecast scales, not the plan.
3. Read revenue on `02_Revenue_Build`, margin on `03_COGS_Build`.
4. Close on `06_Bridge` (volume / mix / price / cost / OpEx). Black font is formulas. Amounts in $000s.

## Screen-share test

Cut volume ~10% on `01_Assumptions` and leave price and unit cost fixed. Revenue and COGS volume walks should move; price and rate walks should stay near zero. Then raise material cost per case and confirm the COGS cost walk absorbs it.

## Tabs

| Tab | Role |
| --- | --- |
| `00_Cover` | Purpose and how to use |
| `01_Assumptions` | Volume, price, cost/unit, OpEx rates, scenario |
| `02_Revenue_Build` | Volume × price, mix, plan vs forecast |
| `03_COGS_Build` | Volume × CPU, gross margin |
| `04_OpEx` | Selling, marketing, distribution, G&A, R&D |
| `05_P&L_Forecast` | 12-month P&L, FY plan vs forecast |
| `06_Bridge` | Plan vs forecast walks |
| `07_Data_Dictionary` | Field definitions |

Excel formulas only. No VBA, no live ERP or demand-planning feed, no employer data. This is the driver-based P&L / reforecast slice, not a three-statement model.

[Profile](https://github.com/saisiri-bandaru) · [Portfolio](https://saisiri-bandaru.github.io) · [LinkedIn](https://www.linkedin.com/in/bandarusaisiri) · [bandarusaisiri1207@gmail.com](mailto:bandarusaisiri1207@gmail.com)
