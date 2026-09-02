# Driver-based forecast

Driver-based P&L for a fictional CPG company (**Northline Consumer Products**): volume × price by product line, COGS from unit-cost drivers, OpEx by cost center, and a plan vs reforecast bridge.

**File to open:** `Northline_Driver_Based_Forecast.xlsx`

## What you will see

- Product-line drivers: volume (000 cases), net $/case, material / labor / overhead per case
- Revenue = volume × price; mix % and ASP; plan vs latest-view forecast
- COGS = volume × cost per unit, with a cost-inflation input and a scenario unit-cost factor
- OpEx by cost center: selling and marketing (% of sales + monthly fixed), distribution ($/case + fixed), G&A and R&D (fixed)
- A 12-month P&L (forecast) with FY plan, variance, and % 
- Plan vs forecast walks for revenue (volume / mix / price), COGS, OpEx, and EBITDA
- A Base / Upside / Downside toggle that scales the forecast only

Yellow cells with blue font are inputs. Black font is formulas.

## How to use

1. Open `01_Assumptions` and change the yellow cells (scenario, prices, unit costs, volumes, OpEx rates, monthly fixed).
2. Read volume × price, mix, and revenue variance on `02_Revenue_Build`.
3. Read unit-cost COGS and gross margin on `03_COGS_Build`.
4. Read cost-center OpEx on `04_OpEx`.
5. Read the 12-month P&L on `05_P&L_Forecast`. Walk plan → forecast on `06_Bridge`.

## Tabs

| Tab | Role |
| --- | --- |
| `00_Cover` | Purpose and how to use |
| `01_Assumptions` | Volume, price, cost/unit, OpEx rates, scenario |
| `02_Revenue_Build` | Volume × price, mix, plan vs forecast |
| `03_COGS_Build` | Volume × CPU, gross margin |
| `04_OpEx` | Selling, marketing, distribution, G&A, R&D |
| `05_P&L_Forecast` | 12-month P&L, FY plan vs forecast |
| `06_Bridge` | Plan vs forecast walks (volume / mix / price / cost / OpEx) |
| `07_Data_Dictionary` | Field definitions |

## Stack

Excel (formulas only — no VBA). Built so another analyst can inherit the file from the data dictionary. Amounts in $000s.

## Not included on purpose

- Live ERP or demand-planning feeds
- Confidential employer data
- A full three-statement model (this is the *driver-based P&L / reforecast* slice)

All sample numbers are fictional.

## Profile

Sai Siri Bandaru — Financial Analyst | FP&A | forecasting, variance analysis, Excel
