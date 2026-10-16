# Regional Sales & Target Tracker (Power BI)

**Regional Sales & Target Tracker** is an end-to-end Power BI project that tracks sales against targets by region, category, product and city. Messy raw tables are cleaned and merged in Power Query, shaped into a star-schema data model, protected with region-based row-level security (RLS), and presented as an interactive one-page dashboard.


---

## 1. Project at a glance

| Area | What was done |
|---|---|
| Data cleaning | Raw staging queries cleaned in Power Query (filter, remove, rename, unpivot, split, trim, proper case, surrogate keys) |
| Merge queries | 20 left-outer merges to build dimension keys and enrich fact tables |
| Data modeling | Star schema: 6 fact tables, 5 dimensions, a date table and a measures table |
| DAX | 14 reusable measures (sales, target, profit, orders, marketing, inventory) |
| Security | Dynamic RLS: each user sees only their region, driven by a `security` table and `USERPRINCIPALNAME()` |
| Visuals | One-page dashboard: 3 slicers, 5 KPI cards, 4 bar charts (including Top 10 filters) |

Tools: Power BI Desktop, Power Query (M), DAX, Excel (source workbook).

### Questions the tracker answers

- Are we meeting our sales target, and by how much (Achievement % and Variance)?
- Which regions, categories, products and customer cities drive the most sales?
- How do sales and targets compare year over year?
- What is the order volume and the average order value?
- Does each regional user see only their own region's data? (RLS)

---

## 2. Data sources

The raw data is an Excel workbook (`raw_tables.xlsx`) with one sheet per source table. Each sheet is loaded as a **staging query** and then transformed into the final model tables.

| Staging query | Used to build |
|---|---|
| `CUST_MASTER`, `customer_contacts`, `Address`, `cities`, `user_details` | `dim_customer` |
| `products`, `subcategories` | `dim_products` |
| `orders`, `channels` | `dim_order_flags` |
| `cities` | `dim_geo` |
| `order_line_items`, `orders` | `fact_sales` |
| `inventory` | `fact_inventory` |
| `CAMPAIGN_LOG` | `dim_campaign`, `fact_campaign_spend` |
| `campaign_skus` | `fact_promotion_coverage` |
| `orders`, `shipments`, `INVOICES`, `payments` | `fact_order_process` |
| `sales_targets` | `fact_Sales_targets` |
| `security` sheet | `security` |



**Tip:** in the Excel connector, set the file path as a Power Query **parameter** so the project runs on any machine without editing each query.

---

## 3. Step 1: Clean the messy data (Power Query)

Open **Home → Transform data**. For every table, the same cleaning routine was applied.

| Cleaning step | Where it was used | How |
|---|---|---|
| Remove dummy/test rows | `dim_customer` | Filter `CustomerID <> 9999` |
| Promote headers + set data types | `security`, `fact_promotion_coverage` | Transform → Use first row as headers, then Detect data type |
| Remove unused columns | all tables | Right-click column → Remove (for example `hash_key`, `source_id`, `ProductDescription`) |
| Remove duplicates | `dim_order_flags`, `dim_campaign` | Select columns → Remove rows → Remove duplicates |
| Create surrogate keys | `dim_products`, `dim_geo`, `dim_order_flags`, `dim_campaign` | Add column → **Index column** from 1, renamed to `*_key` |
| Standardise text | `dim_products` | Transform → Format → **Capitalize Each Word** on `category` |
| Trim spaces | `fact_promotion_coverage` | Transform → Format → **Trim** on `PromotedSKUs` |
| One item per row | `fact_promotion_coverage` | Split column by delimiter `,` **into rows** |
| Wide to long | `fact_inventory` | Select `ProductName` → **Unpivot Other Columns**; month columns become rows |
| Strip hidden time from dates | `fact_sales` | Add column → Custom column `order_date_clean = Date.From([order_date])`, type Date; raw `order_date` hidden |
| Consistent naming | all tables | Rename to `snake_case` (`CustomerID` becomes `customer_id`) |
| Correct data types | all tables | Dates as Date, amounts as Decimal/Whole number, keys as Whole number |

Rule followed throughout: **facts keep only keys, dates and numbers; descriptive text lives in dimensions.**

---

## 4. Step 2: Merge queries (how to do it)

Merging replaces long text keys (customer name, product name, city) with compact numeric keys, so facts can relate to dimensions.

### How to merge in Power Query

1. Open **Transform data** and select the query you want to enrich (the "left" table, for example `fact_sales`).
2. Click **Home → Merge Queries** (use *Merge Queries as New* only if you want a separate result table).
3. Pick the second query (the "right" table, for example `dim_products`).
4. Click the matching column in each table (`ProductName` and `product_name`). Select several columns with `Ctrl` for composite keys; keep the same column order in both tables.
5. Choose **Join kind: Left Outer** (keep every row of the left table).
6. Click **OK**, then use the expand icon on the new column and tick only the key you need (for example `products_key`). Untick *Use original column name as prefix* if you want clean names.
7. Remove the original text column and rename the key column.
8. Check for `null` values after expanding. A `null` means a row did not match (typo, extra space, or missing dimension row).

### Merges used in this project

| Left table | Right table | Join on | Brought in |
|---|---|---|---|
| `dim_customer` | `customer_contacts` | `CustomerID` | contact name, email |
| `dim_customer` | `Address` | `AddressID` | street, city name |
| `dim_customer` | `cities` | `CityName` | region |
| `dim_customer` | `user_details` | `CustomerID` = `UserID` | credit limit, phone |
| `dim_products` | `subcategories` | `SubcategoryName` = `subcategory` | category |
| `dim_order_flags` | `channels` | `OrderChannel` = `channel_code` | channel name |
| `fact_sales` | `orders` | `OrderID` | order date, ship/bill city, channel, status, priority |
| `fact_sales` | `dim_customer` | customer name | `customer_id` |
| `fact_sales` | `dim_products` | product name | `products_key` |
| `fact_sales` | `dim_order_flags` | channel + status + priority | `flag_key` |
| `fact_sales` | `dim_geo` (twice) | ship city and bill city | `ship_to_city_key`, `bill_to_city_key` |
| `fact_inventory` | `dim_products` | product name | `product_key` |
| `fact_campaign_spend` | `dim_campaign` | campaign name + channel | `campaign_key` |
| `fact_promotion_coverage` | `dim_campaign` | campaign name | `campaign_key` |
| `fact_promotion_coverage` | `dim_products` | SKU = `product_code` | `products_key` |
| `fact_order_process` | `dim_customer` | customer name | `customer_id` |
| `fact_order_process` | `shipments` | `OrderID` | ship date, delivery date |
| `fact_order_process` | `INVOICES` | `OrderID` | invoice date |
| `fact_order_process` | `payments` | `InvoiceID` | payment date |

**Good practice:** merging on names works for this dataset, but production models should merge on IDs. Names break when two customers or products share a name or are spelled differently.

---

## 5. Step 3: Data model (star schema)

### Tables

| Type | Tables |
|---|---|
| Fact | `fact_sales`, `fact_order_process`, `fact_inventory`, `fact_campaign_spend`, `fact_promotion_coverage` (bridge), `fact_Sales_targets` |
| Dimension | `dim_customer`, `dim_products`, `dim_geo`, `dim_order_flags`, `dim_campaign`, `dim_date` |
| Helper | `_measure` (holds all measures), `security` (RLS lookup) |

### Relationships

All relationships are **many-to-one** with **single** filter direction (dimension filters fact).

| From (many) | To (one) | Status |
|---|---|---|
| `fact_sales[customer_id]` | `dim_customer[customer_id]` | Active |
| `fact_sales[product_key]` | `dim_products[products_key]` | Active |
| `fact_sales[flag_key]` | `dim_order_flags[flag_key]` | Active |
| `fact_sales[ship_to_city_key]` | `dim_geo[geo_key]` | Active |
| `fact_sales[bill_to_city_key]` | `dim_geo[geo_key]` | Inactive (use `USERELATIONSHIP`) |
| `fact_inventory[product_key]` | `dim_products[products_key]` | Active |
| `fact_promotion_coverage[campaign_key]` | `dim_campaign[campaign_key]` | Active |
| `fact_promotion_coverage[product_key]` | `dim_products[products_key]` | Active |
| `fact_campaign_spend[campaign_key]` | `dim_campaign[campaign_key]` | Active |
| `fact_order_process[customer_id]` | `dim_customer[customer_id]` | Active |
| `dim_customer[region]` | `security[region]` | Active |
| `fact_sales[order_date_clean]` | `dim_date[Date]` | Active |
| `fact_order_process[order_date]` | `dim_date[Date]` | Active (invoice, ship, deliver, payment dates stay inactive) |
| `fact_inventory[month]` | `dim_date[Date]` | Active |
| `fact_campaign_spend[date]` | `dim_date[Date]` | Active |
| `fact_Sales_targets[date]` | `dim_date[Date]` | Active |

Design decisions:

- **Role-playing dimension:** `dim_geo` serves both ship-to and bill-to city. One relationship is active, the other inactive and activated inside a measure.
- **Bridge table:** `fact_promotion_coverage` resolves the many-to-many between campaigns and products.
- **One shared date table** (`dim_date`) so every fact can be analysed by the same Year/Month.
- **Surrogate keys** (`*_key`) created in Power Query instead of joining on text.

### Date table

```DAX
dim_date =
ADDCOLUMNS(
    CALENDARAUTO(),
    "year", YEAR([Date]),
    "month", MONTH([Date]),
    "Month Year", FORMAT([Date], "MMM YYYY"),
    "YearMonth", YEAR([Date]) * 100 + MONTH([Date])
)
```

Then: set `Month Year` **Sort by column** to `YearMonth`, and use **Mark as date table** on `Date`.

---

## 6. Step 4: DAX measures

All measures live in the `_measure` table.

```DAX
Total Sales        = SUM(fact_sales[line_total])
Total Orders       = DISTINCTCOUNT(fact_sales[order_id])
Avg Order Value    = DIVIDE([Total Sales], [Total Orders])
Total Profit       = SUMX(fact_sales, fact_sales[line_total] - fact_sales[quantity] * fact_sales[unit_cost])
Profit Margin %    = DIVIDE([Total Profit], [Total Sales])

Total Target       = SUM(fact_Sales_targets[target_revenue])
Achievement %      = DIVIDE([Total Sales], [Total Target])
Variance           = [Total Sales] - [Total Target]

Total Spend        = SUM(fact_campaign_spend[spend])
Total Clicks       = SUM(fact_campaign_spend[clicks])
Total Impressions  = SUM(fact_campaign_spend[impressions])
Cost Per Click     = DIVIDE([Total Spend], [Total Clicks])
CTR %              = DIVIDE([Total Clicks], [Total Impressions])

Total Units        = SUM(fact_inventory[unit])
```

`Total Orders` counts from `fact_sales` (not `fact_order_process`) so it responds to the category and region slicers.

Always use **measures** (not calculated columns) for totals: a calculated column cannot respond to slicers and charts the way a measure does.

Using the inactive bill-to relationship:

```DAX
Sales by Bill City =
CALCULATE([Total Sales], USERELATIONSHIP(fact_sales[bill_to_city_key], dim_geo[geo_key]))
```

---

## 7. Step 5: Row-level security (RLS)

**Goal:** a user should only see the data of their own region.

### Security table

`security` has two columns: `user_email` and `region`. It is loaded from the Excel workbook and renamed in Power Query.

| user_email | region |
|---|---|
| user1@company.com | Europe |
| user2@company.com | North America |

### Build the role

1. Make sure `dim_customer[region]` is related to `security[region]` (many-to-one).
2. In Power BI Desktop, open **Modeling → Manage roles → Create**.
3. Name the role `region_access`.
4. Select the table `dim_customer` and enter this DAX filter:

```DAX
[region] = LOOKUPVALUE(security[region], security[user_email], USERPRINCIPALNAME())
```

5. Click **Save**.

Because `dim_customer` filters `fact_sales` and `fact_order_process`, the restriction flows to the sales numbers automatically.

### Test the role

1. **Modeling → View as**.
2. Tick `region_access`, and also tick **Other user** and type an email from the `security` table.
3. Click **OK**. Cards and charts now show only that user's region.
4. Stop viewing with **Stop viewing** at the top banner.

### Publish and assign users

1. **Home → Publish** to a Power BI workspace.
2. In the Power BI service: dataset → **... → Security**.
3. Add users or groups to `region_access`.

### Limitations (read before relying on it)

- `LOOKUPVALUE` returns one region, so each user should map to **one** region. For several regions per user, use `CONTAINS`/`IN` with `CALCULATETABLE(VALUES(security[region]), security[user_email] = USERPRINCIPALNAME())`.
- The role filters customers. Tables not linked to `dim_customer` (`fact_Sales_targets`, `fact_campaign_spend`, `fact_inventory`) are **not** restricted.
- The **Sales by Region** chart uses `dim_geo[region]` (ship-to city). A restricted user can still see several regions there because the filter is on the **customer's** region.

---

## 8. Step 6: Visuals

The tracker is a single 1280 x 720 page titled **Regional Sales & Target Tracker**.

| Section | Visual | Fields |
|---|---|---|
| Filters | 3 slicers | `dim_date[year]`, `dim_geo[region]`, `dim_products[category]` |
| KPIs | 5 cards | Total Sales, Total Target, Achievement %, Avg Order Value, Total Orders |
| Chart | Sales by Category (bar) | `dim_products[category]`, `[Total Sales]` |
| Chart | Sales by Region (bar) | `dim_geo[region]`, `[Total Sales]` |
| Chart | Top 10 Products (bar) | `dim_products[product_name]`, `[Total Sales]`, Top N = 10 |
| Chart | Top 10 Customer Cities (bar) | `dim_customer[city]`, `[Total Sales]`, Top N = 10 |

Design choices: one primary colour (`#1F4E79`), sorted descending bars, data labels on, axis titles off, consistent card layout.

### How to build the Top 10 charts

1. Add a clustered bar chart with the category field (product name or city) and `[Total Sales]`.
2. In the **Filters** pane, open the category field and set **Filter type = Top N**.
3. Set **Show items = Top 10**, **By value = `[Total Sales]`**, click **Apply filter**.
4. Sort the chart by `Total Sales` descending.

---

## 9. Key results

| Year | Sales | Target | Achievement |
|---|---|---|---|
| 2025 | 2,66,271 | 2,80,000 | 95.1% |
| 2026 | 2,60,373 | 2,72,000 | 95.7% |
| Total | 5,26,644 | 5,52,000 | 95.4% |

- 80 orders in total, average order value around 6,583.
- Sales data runs until October 2026, while the 2026 target covers the full year, so 2026 achievement is slightly understated.

## 10. Challenges and how I solved them

| Problem | Cause | Fix |
|---|---|---|
| Every month showed the same total | Visuals mixed a calculated column and fields from tables not related to each other | Rebuilt totals as DAX **measures** and used `dim_date` fields only |
| All sales landed in one `(Blank)` group | Hidden time part in `order_date` did not match the date table | Created `order_date_clean` with `Date.From` and rebuilt the relationship |
| Targets and spend gave odd results | Relationships were created as one-to-one with both-direction filtering | Changed to many-to-one, single direction |
| Order count blank for a year | Orders were counted from a table with unmatched dates | Counted distinct `order_id` from `fact_sales` |
| Months sorted alphabetically | `Month Year` had no sort column | Added `YearMonth` and used Sort by column |

## 11. How to run the project

1. Clone the repository.
2. Open `sales_report_v2.pbix` in Power BI Desktop.
3. **Home → Transform data → Data source settings** and point the Excel source to your local copy of the workbook, then **Refresh**.
4. Use **Modeling → View as** to test the `region_access` role.

---

## 12. Repository structure

```
regional-sales-target-tracker/
├── README.md
├── sales_report_v2.pbix
├── power_query_scripts.md
├── .gitignore
├── data/                 
└── images/
    ├── dashboard.png
    ├── model_view.png
    ├── rls_view_as.png
    └── power_query_merge.png
```

---

## Author

**<Krupa Patel>**: Data Analyst | Power BI, SQL

- LinkedIn: <https://www.linkedin.com/in/krupa-patel02/>
- GitHub: <https://github.com/krupa-patel24/>
