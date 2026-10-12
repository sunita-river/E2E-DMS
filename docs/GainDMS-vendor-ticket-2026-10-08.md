# GainDMS vendor ticket - MIS report defects and removed reports (8 Oct 2026)

**Environment:** rivermobility.gaindms.com (River Mobility), MIS menu
**Raised by:** River DMS QA automation (draft - not yet sent)
**Evidence:** automated MIS health check runs of 29 Sep 2026 13:23 (report dates 01/09/2026-28/09/2026, 268 reports) and 8 Oct 2026 11:53 (report dates 01/10/2026-07/10/2026, 262 reports). Source files: `MIS Reports/MIS_Report_Check_2026-10-08_1153.csv`, `MIS_Report_Snapshot_2026-10-08_1153.json`, `MIS_Report_Summary_2026-10-08_1153.txt` (and the 2026-09-29_1323 equivalents).

## Part A - Reports missing from the MIS menu

**Priority: High.** `MIS > Accounts Reports > GSTR Reports > ProductWise Sales Details` is no longer in the menu. It worked on 29 Sep (result OK, 44 columns). Our GST reconciliation depends on it. We need its `ReferenceNo`, `InvoiceNo`, `InvoiceDate`, `ReportOn`, `TaxableAmount`, `TotalAmount`, `DiscountAmount` and `Chassis No` columns.

We checked all 262 reports in the 8 Oct menu and found no report with this name, or a similar one, under any MIS menu. The nearest match is `Accounts Reports > GSTR Reports > Sales Tax Details`. It was already there on 29 Sep and is not a replacement: it has 31 of the 44 columns but lacks ReferenceNo, Chassis No, ItemName, ItemCode, InvoiceValue, UnitPrice, TCS, Subsidy Value/Offer, CustomerMobileNo, PlaceOfSupply, CustomerPanCardNo and AccountName. `Vehicle Reports > Sales > Product-wise sales details` is a different report: it shares only 3 of the 44 columns and has no GST breakdown.

Nine Accounts reports were removed between 29 Sep and 8 Oct, plus one MFR Workshop report:

| # | Menu path (as of 29 Sep) | Result on 29 Sep | Possible replacement seen on 8 Oct |
|---|---|---|---|
| 1 | MIS > Accounts Reports > GSTR Reports > ProductWise Sales Details | OK | None found |
| 2 | MIS > Accounts Reports > GSTR Reports > Advanced Received Details | Blank / no data | None found |
| 3 | MIS > Accounts Reports > GSTR Reports > Advanced Adjusted Details | Blank / no data | None found |
| 4 | MIS > Accounts Reports > GSTR Reports > Inward Supplies Details | OK | None found |
| 5 | MIS > Accounts Reports > All Reports > Receipt List For Cheque | OK | Unclear - new `MFR Account Reports > All Reports > Receipt List` (manufacturer-level, all dealers) |
| 6 | MIS > Accounts Reports > All Reports > Cash Counter Summary for the day | Blank / no data | None found |
| 7 | MIS > Accounts Reports > All Reports > Overall Transaction Report | OK | None found |
| 8 | MIS > Accounts Reports > All Reports > Overall OutStanding | OK | None found |
| 9 | MIS > Accounts Reports > All Reports > First And Last Invoice No Details | OK | None found |
| 10 | MIS > MFR Workshop Reports > Detail > TFR Details | OK | Likely `MFR Workshop Reports > Warranty And Free Service > TFR Report` (new; same purpose) |

Reports added to the menu in the same period: `MFR Workshop Reports > Detail > RecallvDetail Report` (does not open), `MFR Workshop Reports > Warranty And Free Service > TFR Report`, `MFR Account Reports > All Reports > Receipt List` and `MFR Account Reports > All Reports > Payment List`.

**Questions for the vendor:**
1. Was each report above **removed on purpose, renamed, or moved**? If it moved, what is the new menu path? If it was renamed, what is the new name?
2. If ProductWise Sales Details was removed by mistake, please restore it under Accounts Reports > GSTR Reports with the same columns. If it was replaced, which report now gives the same per-invoice GST data, including ReferenceNo and Chassis No?
3. Was this an intended release or a role/permission change? Please share the release notes or change log for this period.
4. `RecallvDetail Report` looks like a typo for "Recall Detail Report", and it does not open.

## Part B - Defects found by the 8 Oct MIS check

Of 262 reports, 157 were OK, 18 gave server errors, 7 were slow or timed out, 76 were blank and 4 never opened.

| Failure type | Count |
|---|---|
| Blank - viewer window empty | 70 |
| Server error - Crystal viewer crv.js HTTP 404 (blank page) | 11 |
| Slow - over 30s | 5 |
| Never opened - clicking the report opened nothing | 4 |
| Blank - Crystal viewer shows blank page | 3 |
| Server error - Report Viewer exception (HTTP 200 page with script exception) | 2 |
| Server error - report page HTTP 500 | 2 |
| Blank - Opens in new tab, table empty (script error: Cannot read properties of null (reading 'getElementsByTagName')) | 2 |
| Slow - new tab did not load within 60s | 2 |
| Server error - ASP.NET runtime error (Server Error in '/' Application) | 2 |
| Server error - Report Viewer 'Service Unavailable' | 1 |
| Blank - Opens in new tab, table empty (script error: $ is not defined) | 1 |

The "29 Sep" column gives each report's result in the previous run. "OK (regression)" means the report worked on 29 Sep.

### Server errors (18)

| Report name | Menu path | Failure type | Seconds | 29 Sep |
|---|---|---|---|---|
| Cases Closed Report & CSN Pending | MIS > MFR CRM Reports > All Reports | Server error - Report Viewer 'Service Unavailable' | 10 | OK (regression) |
| VOR Details | MIS > Workshop Reports > Job based | Server error - report page HTTP 500 | 33 | OK (regression) |
| VOR Details | MIS > MFR Workshop Reports > Job based | Server error - report page HTTP 500 | 33 | OK (regression) |
| Customer Complaints History | MIS > MFR CRM Reports > All Reports | Server error - Crystal viewer crv.js HTTP 404 (blank page) | 31 | same |
| Vin Movement | MIS > Vehicle Reports > Stock | Server error - Crystal viewer crv.js HTTP 404 (blank page) | 7 | same |
| Product-wise Sale count | MIS > Vehicle Reports > Summary | Server error - Crystal viewer crv.js HTTP 404 (blank page) | 11 | same |
| Vehicle History | MIS > Workshop Reports > Job based | Server error - Report Viewer exception (HTTP 200 page with script exception) | 4 | same |
| Performance Summary Report | MIS > Workshop Reports > Summary | Server error - Crystal viewer crv.js HTTP 404 (blank page) | 10 | same |
| Service Inflow/OutFlow Details | MIS > Workshop Reports > Summary | Server error - Crystal viewer crv.js HTTP 404 (blank page) | 11 | same |
| Vehicle History | MIS > MFR Workshop Reports > Detail | Server error - Report Viewer exception (HTTP 200 page with script exception) | 9 | same |
| Standard Part Current Stock | MIS > MFR Workshop Reports > Spares | Server error - Crystal viewer crv.js HTTP 404 (blank page) | 14 | same |
| Dealer Wise Spares Aging | MIS > Spares Reports > Stock | Server error - ASP.NET runtime error (Server Error in '/' Application) | 1 | same |
| Dealer Wise Spares Aging | MIS > Accessories Reports > Stock | Server error - ASP.NET runtime error (Server Error in '/' Application) | 3 | same |
| Enquiry analysis | MIS > Enquiry Reports > Leads Summary | Server error - Crystal viewer crv.js HTTP 404 (blank page) | 7 | same |
| Product Sales Planning | MIS > Enquiry Reports > Others | Server error - Crystal viewer crv.js HTTP 404 (blank page) | 7 | same |
| Visited Enquiry Details | MIS > Enquiry Reports > Others | Server error - Crystal viewer crv.js HTTP 404 (blank page) | 10 | same |
| Testride Summary Product Wise | MIS > Enquiry Reports > Testride | Server error - Crystal viewer crv.js HTTP 404 (blank page) | 6 | same |
| Testride Summary Executive Wise | MIS > Enquiry Reports > Testride | Server error - Crystal viewer crv.js HTTP 404 (blank page) | 6 | same |

### Slow (>30s) / timed out (7)

| Report name | Menu path | Failure type | Seconds | 29 Sep |
|---|---|---|---|---|
| Order Status Report | MIS > MFR Spares Reports > Detail | Slow - over 30s | 33 | OK (regression) |
| Warrant Claim Detail EV | MIS > Workshop Reports > Warranty And Free Service | Slow - over 30s | 34 | OK (regression) |
| PMP and PML detail | MIS > MFR Workshop Reports > Summary | Slow - new tab did not load within 60s | 60 | OK (regression) |
| City wise Enquiry List | MIS > Enquiry Reports > Detail | Slow - over 30s | 31 | OK (regression) |
| Executive-wise enquiry summary analysis | MIS > Enquiry Reports > Enquiry Summary | Slow - over 30s | 112 | OK (regression) |
| PMP PML overall detail | MIS > Workshop Reports > Summary | Slow - new tab did not load within 60s | 60 | BLANK / NO DATA |
| Sales tax Report | MIS > MFR Account Reports > All Reports | Slow - over 30s | 92 | same |

### Never opened (4)

| Report name | Menu path | Failure type | Seconds | 29 Sep |
|---|---|---|---|---|
| Vehicle Stock Aging As On Date | MIS > Vehicle Reports > Stock | Never opened - clicking the report opened nothing | 19 | same |
| WIP on a date (actionable) | MIS > Workshop Reports > Job based | Never opened - clicking the report opened nothing | 1 | same |
| WIP on a date (actionable) | MIS > MFR Workshop Reports > Detail | Never opened - clicking the report opened nothing | 1 | same |
| RecallvDetail Report | MIS > MFR Workshop Reports > Detail | Never opened - clicking the report opened nothing | 5 | new in menu |

### Blank / no data (76)

Some blank reports may just have no data for 01-07 Oct. Please start with the rows marked "OK (regression)" and the ones showing a script error.

| Report name | Menu path | Failure type | Seconds | 29 Sep |
|---|---|---|---|---|
| Spares Purchase Order Status Summary | MIS > MFR Spares Reports > Summary | Blank - viewer window empty | 5 | OK (regression) |
| Scheme Based Sale | MIS > Vehicle Reports > Sales | Blank - viewer window empty | 1 | OK (regression) |
| Receipt List | MIS > Accounts Reports > All Reports | Blank - viewer window empty | 4 | OK (regression) |
| Form GSTR-1 Table-7B | MIS > Accounts Reports > GSTR Reports | Blank - viewer window empty | 2 | OK (regression) |
| Customer Satisfaction Index Rating(V) | MIS > MFR CRM Reports > All Reports | Blank - viewer window empty | 4 | same |
| Sale FeedBack | MIS > MFR CRM Reports > Feedback | Blank - viewer window empty | 20 | TIMEOUT |
| Spares FSN Wise Details | MIS > MFR Spares Reports > Detail | Blank - viewer window empty | 3 | same |
| Dead Stock Analysis | MIS > MFR Spares Reports > Detail | Blank - viewer window empty | 3 | TIMEOUT |
| Spares Purchase Order Detail Status Report | MIS > MFR Spares Reports > Purchase | Blank - viewer window empty | 2 | TIMEOUT |
| All Dealer Dead Stock | MIS > MFR Spares Reports > Summary | Blank - viewer window empty | 1 | same |
| Spares FSN Analysis | MIS > MFR Spares Reports > Summary | Blank - viewer window empty | 20 | same |
| Customer List With DOB & Anniversary | MIS > Vehicle Reports > CRM | Blank - viewer window empty | 6 | same |
| Product Sale Invoice Cancel Details | MIS > Vehicle Reports > Detail | Blank - viewer window empty | 31 | same |
| Supplierwise Purchases Details | MIS > Vehicle Reports > Purchase | Blank - viewer window empty | 1 | TIMEOUT |
| Product Sales Additional Services | MIS > Vehicle Reports > Sales | Blank - viewer window empty | 3 | same |
| Product Stock Complete Details | MIS > Vehicle Reports > Stock | Blank - viewer window empty | 10 | same |
| PDI Status Detail | MIS > Vehicle Reports > Stock | Blank - viewer window empty | 3 | same |
| Web Enquiry Tracker | MIS > MFR Vehicles Reports > Detail | Blank - viewer window empty | 2 | same |
| Visited Enquiry Details | MIS > MFR Vehicles Reports > Pre-sales | Blank - Crystal viewer shows blank page | 32 | TIMEOUT |
| Loaner Vehicle Allocation | MIS > MFR Vehicles Reports > Sales | Blank - viewer window empty | 2 | same |
| Product Stock Complete Details | MIS > MFR Vehicles Reports > Stock | Blank - viewer window empty | 7 | same |
| Workshop Service Register (Billed) | MIS > Workshop Reports > Billing | Blank - viewer window empty | 12 | same |
| Workshop Spares bill register | MIS > Workshop Reports > Billing | Blank - viewer window empty | 6 | same |
| Workshop Labour bill register | MIS > Workshop Reports > Billing | Blank - viewer window empty | 2 | same |
| Service Job OutStanding Payments | MIS > Workshop Reports > Billing | Blank - viewer window empty | 4 | same |
| Post Service Feedback List | MIS > Workshop Reports > CRM | Blank - viewer window empty | 2 | same |
| TFR Details | MIS > Workshop Reports > Detail | Blank - viewer window empty | 7 | same |
| Service Denied Details | MIS > Workshop Reports > Detail | Blank - viewer window empty | 7 | same |
| Workshop Service Register | MIS > Workshop Reports > Job based | Blank - viewer window empty | 3 | same |
| Service type Wise Jobs | MIS > Workshop Reports > Job based | Blank - viewer window empty | 3 | same |
| Repeat Job | MIS > Workshop Reports > Job based | Blank - viewer window empty | 5 | same |
| Job Card Concern | MIS > Workshop Reports > Job based | Blank - viewer window empty | 11 | same |
| Repeat Visit | MIS > Workshop Reports > Job based | Blank - viewer window empty | 4 | same |
| VOR Pending Reason History Details | MIS > Workshop Reports > Job based | Blank - viewer window empty | 2 | same |
| Mechanic Wise Activity | MIS > Workshop Reports > Labour | Blank - viewer window empty | 1 | same |
| Labours Issued Report: Job Card | MIS > Workshop Reports > Labour | Blank - viewer window empty | 2 | same |
| Spares Required Today | MIS > Workshop Reports > Planning | Blank - viewer window empty | 2 | same |
| Leakage And Redemption Summary | MIS > Workshop Reports > Redemption And Leakage | Blank - viewer window empty | 3 | same |
| PMP Top Spares/Model Details | MIS > Workshop Reports > Spares | Blank - Opens in new tab, table empty (script error: Cannot read properties of null (reading 'getElementsByTagName')) | 3 | same |
| Technical Ticket Detail | MIS > Workshop Reports > Summary | Blank - viewer window empty | 3 | same |
| Labours Issued Report: Job Card | MIS > MFR Workshop Reports > Detail | Blank - viewer window empty | 1 | same |
| Repeat Visit | MIS > MFR Workshop Reports > Job based | Blank - viewer window empty | 3 | same |
| Leakage And Redemption Summary | MIS > MFR Workshop Reports > Redemption And Leakage | Blank - viewer window empty | 6 | same |
| PMP top Spares/Model Details | MIS > MFR Workshop Reports > Spares | Blank - Opens in new tab, table empty (script error: Cannot read properties of null (reading 'getElementsByTagName')) | 3 | same |
| ReCall Summary | MIS > MFR Workshop Reports > Summary | Blank - Opens in new tab, table empty (script error: $ is not defined) | 33 | same |
| Movement Analysis | MIS > Spares Reports > General | Blank - Crystal viewer shows blank page | 17 | SERVER ERROR |
| ABC Analysis | MIS > Spares Reports > General | Blank - viewer window empty | 1 | same |
| Spares issued with value | MIS > Spares Reports > Job based | Blank - viewer window empty | 6 | same |
| Spares Purchase Register | MIS > Spares Reports > Purchase | Blank - viewer window empty | 19 | same |
| Spares Purchase Detail | MIS > Spares Reports > Purchase | Blank - viewer window empty | 6 | same |
| Spares Sale Register | MIS > Spares Reports > Sales | Blank - viewer window empty | 2 | same |
| SparesReturnDetail | MIS > Spares Reports > Sales | Blank - viewer window empty | 1 | same |
| Sale Tax | MIS > Spares Reports > Spares | Blank - viewer window empty | 1 | same |
| Movement Analysis | MIS > Accessories Reports > General | Blank - Crystal viewer shows blank page | 7 | SERVER ERROR |
| Spares issued with value | MIS > Accessories Reports > Job based | Blank - viewer window empty | 9 | same |
| SparesReturnDetail | MIS > Accessories Reports > Sales | Blank - viewer window empty | 1 | same |
| Form GSTR-1 Table-5 | MIS > Accounts Reports > GSTR Reports | Blank - viewer window empty | 5 | same |
| Form GSTR-1 Table-11A-1 | MIS > Accounts Reports > GSTR Reports | Blank - viewer window empty | 2 | TIMEOUT |
| Form GSTR-1 Table-11A-2 | MIS > Accounts Reports > GSTR Reports | Blank - viewer window empty | 2 | same |
| Form GSTR-1 Table-11B-1 | MIS > Accounts Reports > GSTR Reports | Blank - viewer window empty | 8 | same |
| Form GSTR-1 Table-11B-2 | MIS > Accounts Reports > GSTR Reports | Blank - viewer window empty | 2 | same |
| Cases Opened Report | MIS > CRM Reports > All Reports | Blank - viewer window empty | 1 | same |
| Cases Closed Report & CSN Uploaded | MIS > CRM Reports > All Reports | Blank - viewer window empty | 1 | same |
| Cases Concern Report | MIS > CRM Reports > All Reports | Blank - viewer window empty | 1 | same |
| Cases Closed Report & CSN Pending | MIS > CRM Reports > All Reports | Blank - viewer window empty | 1 | same |
| Customer Anniversary | MIS > CRM Reports > Anniversary | Blank - viewer window empty | 6 | same |
| Enquiry FeedBack | MIS > CRM Reports > Feedback | Blank - viewer window empty | 1 | same |
| Sale FeedBack | MIS > CRM Reports > Feedback | Blank - viewer window empty | 6 | same |
| Testride FeedBack | MIS > CRM Reports > Feedback | Blank - viewer window empty | 1 | same |
| Service FeedBack | MIS > CRM Reports > Feedback | Blank - viewer window empty | 3 | same |
| Sales Analysis | MIS > Enquiry Reports > Enquiry Summary | Blank - viewer window empty | 2 | same |
| Mobile App Utilisation Summary | MIS > Enquiry Reports > Mobile App Utilisation | Blank - viewer window empty | 1 | same |
| Mobile App Utilisation Details | MIS > Enquiry Reports > Mobile App Utilisation | Blank - viewer window empty | 1 | same |
| Suspect Failure Details | MIS > Enquiry Reports > Suspect | Blank - viewer window empty | 1 | same |
| Suspect Details | MIS > Enquiry Reports > Suspect | Blank - viewer window empty | 1 | same |
| Leads list with testride details | MIS > Enquiry Reports > Testride | Blank - viewer window empty | 1 | same |

**Asks for Part B:**
- **Crystal viewer crv.js 404:** the Crystal Reports viewer script `crv.js` returns HTTP 404, so these reports show a blank page. This looks like one missing or misconfigured static resource on the server, and fixing it should clear most of the server errors.
- **ASP.NET runtime errors and HTTP 500:** please send the server-side exception for Dealer Wise Spares Aging (Spares and Accessories) and VOR Details (Workshop and MFR Workshop).
- **Report Viewer exceptions and Service Unavailable:** Vehicle History (Workshop and MFR Workshop) and Cases Closed Report & CSN Pending (MFR CRM).
- **Slow reports:** we time out at 30s (60s for reports that open in a new tab). Executive-wise enquiry summary analysis took 112s and MFR Sales tax Report took 92s in our run. These reports took longer than 30s on the 01-07 Oct range, so their columns could only be read from a 1-day run.
- **Never opened:** clicking these reports opens no viewer or tab.
