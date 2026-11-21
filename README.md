# Final Project: Sales Dashboard for Farila Muslim Store (Group A5)

My Excel Bootcamp final project, a group dashboard built from 1,000 online orders for a Muslim fashion brand, documenting how my team turned raw transaction records into a Pivot Table dashboard and a business recommendation.

## Project Background

This project was completed as the final group assignment of an Excel Bootcamp I participated in, together with my teammates in Group A5 (Kelompok A5), La Tansa Mahabbata and Risdayanti, under the guidance of our mentor, Naufal Ramadhan Adibperdana.

The assignment involved working with 1,000 orders from Farila Muslim Store, sold across four marketplaces (Shopee, Tokopedia, BliBli, and Bukalapak) during April 2023. The objective was to prepare the raw data with formulas and lookups, analyse sales with Pivot Tables, assemble a dashboard with at least four charts and one slicer, and present what a Brand Owner should do based on the insights.

This was a team effort, and I want to be clear about who did what. La Tansa Mahabbata, Risdayanti, and I built the data preparation, KPI strip, marketplace, rating and region charts, slicers, and dashboard layout together. My individual part was three of the analyses (weekly trend vs profit, discount code vs profit, and product profit vs margin) and the final business conclusion. The project presentation labels each page as "Mine" or "Team" so the split is easy to see.

Compared with my earlier individual Excel projects, this was my first time working on one shared workbook with a team, which meant agreeing on a common data structure before anyone could build their own analysis.

This project represents another step in developing practical data analysis skills using Microsoft Excel.

## Project Files

- [Excel Workbook (.xlsx)](./FARILA_FINAL_PROJECT_KELOMPOK_A5.xlsx)
- [Project Presentation](./Group%20Final%20Project.pdf)

The workbook contains eight sheets: Guide, Main Data, Clean Data, Pivot Table, Dashboard, Rating, Kode Diskon, and Katalog.

## Learning Objectives

Through this assignment, I aimed to:

- Build a unique OrderID in the format ORD-(marketplace code)-(number)-(dd/mm/yy).
- Split a text field that holds several items into separate item, quantity, and rating columns.
- Pull product names, categories, costs, prices, and discount rates from lookup tables.
- Calculate total quantity, transaction value, discount, grand total, profit, and margin.
- Summarize the data with Pivot Tables and Pivot Charts.
- Build a dashboard with at least four charts and one slicer.
- Turn the findings into a clear recommendation for a Brand Owner.

## Tasks & Analysis Process

The assignment consisted of five main steps:

**1. Prepare the Data (Team)**

Built the OrderID, then extracted the item code, quantity, and rating of up to three items per order from a single text column. Looked up each item's name, category, production cost, and selling price from the Katalog sheet, and each discount rate from the Kode Diskon sheet.

Functions used: CONCATENATE, TEXT, FIND, MID, VALUE, VLOOKUP, IFNA.

**2. Calculate Transaction Values (Team)**

Calculated total quantity, total transaction value, discount, grand total, total production cost, total profit, and margin %. Added a week-number column so sales can be grouped by week.

Functions used: IF, IFERROR, INT, DAY.

**3. Build the Pivot Table Sheet (Team and Individual)**

Built six Pivot Tables that feed six charts. Three of them are mine: weekly revenue and order count, total profit per discount code, and margin % and profit per product.

**4. Assemble the Dashboard and Slicers (Team)**

Combined a KPI strip, six charts, and two slicers (Marketplace and Jenis promo). The slicers re-cut every chart at once, so the same dashboard can be filtered by marketplace or promo type.

**5. Analyse and Recommend (Individual)**

Wrote three findings from my pivots and turned them into three recommendations for the Brand Owner.

## Key Findings

Based on the completed workbook:

1. The dataset holds 1,000 orders and 4,080 items, with total revenue of Rp800,057,500, a grand total (after discounts) of Rp706,736,170, and total profit of Rp392,311,170.
2. Week 1 (1-7 April) alone accounted for 510 of the 1,000 orders, or 51% of all orders and about 51% of profit. Filtering the Promo 4.4 slicer leaves 492 orders, all in week 1, so the spike comes from the promo and not from normal demand.
3. A deeper discount did not bring more orders. DISC20 attracted 123 orders, about the same as DISC10 (116) and DISC15 (130).
4. DISC20 gave away the most discount (Rp18.56M) and earned the lowest profit of all discount codes (Rp37.7M), while full-price orders earned the most profit (Rp68.4M).
5. Jubah Gamis earned the most profit (Rp79.3M) at a 56% margin, followed by Dress (Rp48.6M). The highest margins (60%) belong to Khimar Bandana and Pashmina Viscose, but they earn much less profit in total.
6. Shopee generated the largest share of revenue (43%), and Jabodetabek was the top region by purchasing power.

**Recommendations to the Brand Owner:**

1. Promote in short, scheduled bursts instead of steady discounting.
2. Cap discounts at 10%, since more discount did not bring more orders.
3. Sell bundles that pair best sellers (Jubah Gamis, Dress) with high-margin accessories (Khimar, Pashmina, Sajadah), then test the results.

## Tools & Techniques

**Tool:** Microsoft Excel

**Skills Practised:**
- Text Functions (CONCATENATE, TEXT, FIND, MID)
- Lookup Functions (VLOOKUP, IFNA)
- Logical and Error-Handling Functions (IF, IFERROR)
- Date Calculations
- Pivot Tables and Pivot Charts
- Slicers
- Dashboard Design
- Data Storytelling and Business Recommendations
- Teamwork on a Shared Workbook

## Challenges & Learning Reflection

The biggest challenge in this project was data preparation. Each order stored its items, quantities, and ratings in one text field, so splitting it into clean columns with FIND and MID took careful work. Every analysis after that depends on those columns, so a small error there would have affected every chart.

Another important takeaway was working in a team, and I am grateful to La Tansa Mahabbata and Risdayanti for the collaboration and to our mentor, Naufal Ramadhan Adibperdana, for the guidance. Splitting the work only works when everyone agrees on the same prepared data, column names, and definitions of profit and margin. Labelling which parts were mine and which were my teammates' also made me more careful about presenting my own contribution honestly.

I also learned that the most obvious number is not always the most useful one. The product with the highest revenue was not the one with the highest margin, and the biggest discount earned the least profit. Looking at profit and margin together gave a much better basis for a recommendation than looking at sales alone.

## What's Next?

This project builds on the cleaning, lookup, and dashboard techniques I practised in my earlier Excel assignments.

I plan to continue strengthening my Excel skills and gradually expand my toolkit with SQL, Python, and data visualization tools as I work toward a career in data analytics.

I will continue documenting each project, including the techniques I practise, the challenges I encounter, and how my analytical skills develop over time.

---

**Disclaimer**

This project was completed as part of an Excel Bootcamp for educational and portfolio purposes only. The analysis is based on a dataset of 1,000 orders (April 2023) supplied for the assignment. The dataset contains no customer names or contact details, only order, product, marketplace, and region information. The project is not intended for commercial use or to represent actual business results. Parts of the workbook and dashboard were built collaboratively by the members of Group A5, and the project presentation indicates which analyses are my individual work.
