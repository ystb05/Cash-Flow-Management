# Cash-Flow-Management
# Project Cash Flow Management: 13-Week Cash Flow Forecast for a Landscaping Business

Built a **13-week rolling cash flow forecast** in Excel for Greenway Landscaping LLC to answer the CFO's key question: *"Over the next 13 weeks (Dec 1, 2026 – Feb 28, 2027), will we have enough cash, and which week is the most critical?"*

<img width="1300" height="633" alt="image" src="https://github.com/user-attachments/assets/372674d1-849f-4eaa-877b-3772b53b837b" />


## Dataset
| Sheet | Description |
|:---|:---|
| `Cash Position` | Opening cash balance **85,000** as of 11/30/2026 (First National Bank – Operating Checking) |
| `AR Collections` | 24 open customer invoices (total **210,760**) with invoice date, due date and aging bucket |
| `Committed Expenses` | 11 fixed expense items with frequency (weekly, bi-weekly, monthly, annual), amount and next due date |

<img width="733" height="610" alt="image" src="https://github.com/user-attachments/assets/a93d2b8a-14c2-46fb-9cd2-ea623471fb75" />
<img width="1147" height="396" alt="image" src="https://github.com/user-attachments/assets/7ccf8ba5-9d93-449e-ad34-6267fc351455" />


## Preview
The company **never runs out of cash**, but it **burns 66,180** over 13 weeks. The balance falls from 85,000 to **9,500 in Week 7**, the most critical week.

<img width="1248" height="510" alt="image" src="https://github.com/user-attachments/assets/c2ca0d87-805c-4254-8ec7-d9562aaa023f" />

## How to implement
**Core logic:** `Ending Balance = Beginning Balance + Cash In − Cash Out`, and *next week's Beginning = this week's Ending*.

### Step 1: Build weekly Cash Inflow (`Cash Inflow` sheet)
- Generate 13 weekly buckets: Start date `=DATE(2026,12,1)`, End date `= Start + 6`.
- Assume customers pay **on the due date**. Collections per week = cumulative AR due up to the week's end date, minus what was already collected in previous weeks:
```excel
Week 01: =SUMIFS('AR Collections'!$E$5:$E$28,'AR Collections'!$D$5:$D$28,"<="&C3)
Week 02+: =SUMIFS('AR Collections'!$E$5:$E$28,'AR Collections'!$D$5:$D$28,"<="&C4)-SUM($D$3:D3)
```
→ The cumulative approach makes sure overdue invoices (e.g., the 31–60 day City of Pinecrest invoice) are picked up in Week 1 rather than dropped.

<img width="457" height="431" alt="image" src="https://github.com/user-attachments/assets/19ba98f8-8f1e-4580-8d75-991358024afe" />

### Step 2: Allocate expenses to weeks (`Cash Outflow` sheet)
Each expense is placed in the week that contains its due date, based on frequency:
| Frequency | Rule | Example |
|:---|:---|:---|
| Weekly | Every week | Payroll 12,000 (Fridays), Fuel 2,800 |
| Bi-weekly | Every other week from 12/05 | Landscaping materials 4,500 → Weeks 1, 3, 5, 7, 9, 11, 13 |
| Monthly (1st) | Week containing the 1st | Rent 8,500, Insurance 1,800 + 950, Software 310 |
| Monthly (10th / 15th / 20th) | Week containing that day | Utilities 420, Equipment lease 3,200, SBA loan 2,100 |
| Annual | Week containing Jan 15 | Vehicle registration 1,200 → Week 7 |

Weekly totals are then transposed (`=TRANSPOSE(B16:N16)`) into a vertical list that feeds the forecast.

<img width="1410" height="270" alt="image" src="https://github.com/user-attachments/assets/3e10bd22-60da-4a3c-b203-68ca9ab72888" />


### Step 3: Link everything in the 13-week forecast (`Cash Flow Forecast` sheet)
| Column | Formula |
|:---|:---|
| Beginning | Week 1 = `85,000`; following weeks = previous week's Ending |
| Cash In | `='Cash Inflow'!D3` |
| Cash Out | `='Cash Outflow '!B19` |
| Net Cash Flow | `= Cash In − Cash Out` |
| Ending | `= Beginning + Net Cash Flow` |

### Step 4: Pivot Tables & Dashboard
4 Pivot Tables summarize Cash In vs Cash Out, Beginning vs Ending balance, Net Cash Flow by week and Outflow by expense category. They feed a dashboard with **Week** and **Expense/Week** slicers:
- **Net Cash Flow** (line chart)
- **Beginning and Ending Balance by Week** (clustered column)
- **Total Outflow by Category** (clustered column)

## Key Findings
| Metric | Value |
|:---|---:|
| Opening cash (12/01/2026) | 85,000 |
| Total Cash In (13 weeks) | 210,760 |
| Total Cash Out (13 weeks) | 276,940 |
| **Net cash burn** | **(66,180)** |
| Ending cash (Week 13) | 18,820 |
| **Lowest balance** | **9,500 – Week 7 (Jan 12–18, 2027)** |
| Weeks with negative net cash flow | 7 of 13 |

1. **Lowest balance:** Cash falls steadily from Week 1 to **Week 7 (9,500)**. Early-week collections are thin because receivables take 1–2 months to come in, while payroll and fuel go out every week. Week 7 also bunches together the equipment lease, materials and the annual vehicle registration.
2. **No shortfall, but a razor-thin cushion:** the balance never turns negative, yet 9,500 is **less than one week of payroll (12,000)**. A few late customer payments would push the company into overdraft.
3. **Burning cash:** outflows exceed inflows by 66,180. **Payroll alone is 156,000 (~56% of total outflows).**

## Recommendations to the CFO
- **Speed up collections:** offer early-payment discounts, follow up on invoices before they are due, and monitor customers' financial health.
- **Secure a revolving credit line** before Week 5 as a buffer for the Week 6–8 low point.
- **Manage labor costs:** as the largest line item, defer bonuses or adjust crew scheduling if the balance trends lower.

> 💡 **Key takeaway:** *Profit ≠ Cash.* The timing mismatch between collections (AR cycle) and disbursements (expense cycle) drives the real cash position.

## Assumptions & Limitations
- Customers pay exactly on the due date; no bad debts or late payments are modelled.
- Only committed fixed expenses are included; no new sales or variable costs after Nov 30, 2026.
- Week 13 is labelled *Feb 23 – Mar 01*, but expenses due on Mar 1 (rent, insurance, software ≈ 11,560) are excluded because the forecast period ends Feb 28.

# Author
Huynh Truong Giang
Email: *truonggiang281005@gmail.com*
