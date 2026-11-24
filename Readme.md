# E-commerce Operations: Pricing, Inventory Sync and Account Health Analysis

End to end Excel Project

**Author:** Krishnendu Das

---

## Disclaimer

This is a portfolio project and a simulation. The company, suppliers, products and marketplace performance figures are fictional, and both datasets were generated with a seeded Python script, not taken from a real business. The findings show the method, not real-world results.

---

## Project Overview

This project simulates the work of an E-commerce Operations Analyst managing a third-party reseller presence on Amazon, Walmart and eBay. It joins a 25,000-line supplier inventory feed to a 25,000-listing marketplace catalog, then answers 13 business questions about margin, repricing, account health and catalog quality, all in Excel.

I built it to practise the day-to-day work of the role: syncing stock and cost into listings, supporting B2B repricing with advanced formulas, and monitoring account health and catalog quality at scale.

---

## My Thinking Process

I started where a real project starts: with a Business Requirements Document. I wrote down what the business needs (a synced catalog, safe repricing, account-health monitoring, catalog QA) and how we would know each need is met, before opening Excel.

Then I needed data with real operational problems in it. Clean data teaches very little, so I generated two datasets with imperfections I would have to find and handle: duplicate SKUs, missing UPCs, zero and low stock, over-length titles and high late-shipment rates.

Finally, I treated the workbook as something a manager would read. The first sheet holds only the questions and answers, and everything behind it (calculations, checks, raw data) is on separate, clearly labelled sheets.

---

## Business Problem

A reseller sells thousands of products it does not make. Supplier cost, stock, competitor prices and seller performance live in separate files, so margin leaks, out-of-stock items stay live and account-health problems go unnoticed. The business needs one joined view that shows where it loses money, where it cannot fulfil orders, and which listings put its seller accounts at risk.

---

## Business Questions Answered

- Can every listing be linked to a supplier cost and stock level, and how reliable is the supplier feed?
- How many listings are priced below break-even, and what does each sale cost?
- Which marketplaces and categories have the weakest margins?
- What does a rules-based B2B repricing plan change, and what is its impact on margin and profit?
- Where can we not compete on price without breaking our margin floor?
- Which listings put our seller accounts at risk, and how many cannot be fulfilled or fail catalog QA?

The workbook answers 13 questions in total, and every answer is a live formula.

---

## Key Findings

- **7.8%** of listings (1,959) are priced below break-even, losing **$1.62** per sale on average.
- Repricing changes **85.5%** of listings (14,866 increases, 6,511 decreases) and lifts average net margin from **13.6% to 17.6%**.
- Loss-making listings in scope fall from **1,671 to 0** after repricing.
- **56%** of repriced listings face a competitor price below our 15% margin floor, so they cannot be matched without losing margin.
- **500 listings (2.0%)** exceed the 4% late-shipment threshold, with another **108** on watch.
- **1,082** active listings have zero supplier stock and **1,735** more hold fewer than 15 units.
- The supplier feed has **500 duplicate SKUs**, **250 missing UPCs** and **375 titles** over 200 characters in the catalog.

---

## Data Source

Both datasets are synthetic, generated for this project with a seeded Python script so the results can be reproduced.

- **Supplier_Inventory_Feed.xlsx:** 25,000 lines with SKU, UPC, supplier, wholesale cost, stock, warehouse zone and lead time. It includes about 2% duplicate SKUs, 1% missing UPCs, and a mix of zero-stock and low-stock lines.
- **Marketplace_Master_Catalog.xlsx:** 25,000 listings across five categories (Electronics, Home & Kitchen, Office Supplies, Automotive, Fitness) and three marketplaces, with price, competitor price, fee, shipping, late-shipment rate and status. Every SKU matches the feed, and about 1.5% of titles exceed 200 characters.

---

## Tools and Tech Stack

| Tool | Purpose |
|------|---------|
| Microsoft Excel | Formulas, named ranges, conditional formatting, charts and the full analysis |
| Business Requirements Document | Requirements, scope, acceptance criteria and business rules |
| Python (openpyxl) | Generating the synthetic datasets only; no analysis is done in Python |
| GitHub | Version control and portfolio hosting |

---

## Core Concepts Demonstrated

- **Multi-file lookups:** SUMIFS and VLOOKUP with IFERROR to sync supplier stock and cost into listings
- **Nested logic:** IF with AND and OR for repricing, account health and catalog QA rules
- **Aggregation:** COUNTIFS, AVERAGEIFS and SUMPRODUCT for marketplace, category and supplier analysis
- **Pricing mechanics:** break-even price, net margin and a margin-floor repricing rule
- **Data cleansing:** helper columns to flag duplicate SKUs, missing UPCs and stock status
- **Controlled assumptions:** named ranges so thresholds are changed in one place
- **Validation:** reconciliation checks that confirm totals tie out and no formula errors exist

---

## Project Workflow

**Step 1: Requirements.** I wrote a short BRD covering objectives, scope, nine business requirements with acceptance criteria, and the business rules (target margin, undercut, thresholds).

**Step 2: Data preparation.** I generated the supplier feed and marketplace catalog with realistic values and deliberate imperfections, then confirmed the issue counts matched the plan.

**Step 3: Analysis in Excel.** I completed the calculated columns (synced stock and cost, break-even price, repricing, net margin, account health, QA check) and built the analysis sheets on top of them.

**Step 4: Validation and reporting.** I added reconciliation checks, cross-checked the key figures against an independent calculation, and wrote the answers sheet so every number traces back to the detail data.

---

## Workbook Structure

**Business_Questions.** The answer sheet, in a simple two-column Questions and Answers format. Each answer is a formula that reads from the analysis sheets.

**Analysis sheets.** Pricing_Analysis (break-even, margins, repricing actions and impact), Account_Health (late-shipment risk), Catalog_QA (status, QA results and stock exposure) and Feed_Analysis (feed quality and supplier scorecard), plus an Assumptions sheet that holds every editable input.

**Data and controls.** Active_Listings and Vendor_Feed hold the raw data with the calculated columns, and Data_Checks confirms the workbook reconciles.

---

## Sample Excel: Syncing Supplier Stock and Cost

```
Synced_Stock_Qty (sums duplicate feed lines for the same SKU)
=SUMIFS(Vendor_Feed!$E$2:$E$25001, Vendor_Feed!$A$2:$A$25001, B2)

Synced_Wholesale_Cost (exact match, unmatched SKUs are labelled)
=IFERROR(VLOOKUP(B2, Vendor_Feed!$A$2:$D$25001, 4, FALSE), "NOT IN FEED")

Break_Even_Price
=ROUNDUP((M2+I2)/(1-H2),2)
```

---

## Sample Excel: B2B Repricing Rule

```
Minimum price at target margin
=ROUNDUP((M2+I2)/(1-H2-Target_Margin),2)

New_B2B_Reprice_USD
(hold inactive or out-of-stock listings, otherwise undercut the competitor
 but never go below the minimum price)
=IF(OR(K2<>"Active", L2=0), F2,
   IF(ROUND(G2-Undercut_USD,2)>=T2, ROUND(G2-Undercut_USD,2), T2))

Account_Health_Flag
=IF(J2>Late_Critical, "CRITICAL",
   IF(AND(J2>=Late_Watch, J2<=Late_Critical), "WATCH", "HEALTHY"))
```
