# Logistics Operation: Power BI Build Log

## 1. Project goal and status

Pull insights out of a logistics database (14 tables, described in `Key Guide.docx`) and present them in four dashboards:

1. Executive Overview
2. Customer Route and Delivery Service
3. Driver Performances and Safety
4. Fleet Fuel and Maintenance

---

## 2. The data source (14 tables)
TS Academy Capstone Project raw csv files

| Table | Type | Notes |
|---|---|---|
| drivers | dimension | about 150 drivers |
| trucks | dimension | about 120 trucks in use |
| trailers | dimension | about 180 trailers in use |
| customers | dimension | 200 customers |
| facilities | dimension | terminals, warehouses, hubs, cross-docks |
| routes | dimension | origin and destination city/state, rates |
| loads | fact | revenue, customer_id, route_id (about 85.4K loads) |
| trips | fact | miles, fuel gallons, duration, one trip per load |
| fuel_purchases | fact | 196,442 rows |
| maintenance_records | fact | 2,920 rows |
| delivery_events | fact | 170,820 rows |
| safety_incidents | fact | 170 rows |
| driver_monthly_metrics | aggregated | 4,464 rows (driver x month) |
| truck_utilization_metrics | aggregated | 3,312 rows (truck x month) |

**Rule used throughout:** trips and loads are the main fact tables. The two aggregated tables are used only for monthly trend charts, and are never mixed with the same metric calculated from raw trips in one visual (it double counts or disagrees).

---

## 3. Data cleaning (Power Query)

- Checked and set data types on every table: IDs as text, dates as date, miles/gallons/costs as numbers.
- `fuel_purchases[truck_id]` and `[driver_id]`: replaced empty strings with `null` (about 2% blank). Rows were **not** deleted, because each is a real fuel purchase.
- Tried merging fuel_purchases with trips to fill the blanks. trips has the same blanks (about 2% of trips have no truck), so the merge was discarded. Blanks were left as null.
- `fuel_purchases[purchase_date]`: changed from date/time to **Date** so it can match the Date table.
- `delivery_events[on_time_flag]` holds true/false. `detention_minutes` is numeric.
- Applied with **Close & Apply**.

---

## 4. Data model

### Relationships (many-to-one unless noted)

| From | To |
|---|---|
| loads[customer_id] | customers |
| loads[route_id] | routes |
| trips[load_id] | loads (one-to-one, both directions) |
| trips[driver_id] | drivers |
| trips[truck_id] | trucks |
| trips[trailer_id] | trailers |
| fuel_purchases[trip_id] | trips |
| maintenance_records[truck_id] | trucks |
| delivery_events[load_id] | loads |
| delivery_events[facility_id] | facilities |
| safety_incidents[trip_id] | trips (one-to-one) |
| driver_monthly_metrics[driver_id] | drivers |
| truck_utilization_metrics[truck_id] | trucks |

### Changes made in Model view

- **Deleted** `fuel_purchases[driver_id] -> drivers` and `fuel_purchases[truck_id] -> trucks`. They gave Power BI two routes to the same table and made the trips links inactive.
- **Activated** `trips[driver_id] -> drivers` and `trips[truck_id] -> trucks`.
- Left **inactive** (fine): `safety_incidents[driver_id] -> drivers`, `safety_incidents[truck_id] -> trucks`, `delivery_events[trip_id] -> trips`.

### Date table

`Date` (calculated table, Jan 2022 to Dec 2025), **marked as date table**. Linked to:

- `trips[dispatch_date]` (main date link)
- `maintenance_records[maintenance_date]`
- `fuel_purchases[purchase_date]` (Power BI may keep this one inactive, because fuel already reaches Date through trips)
- `driver_monthly_metrics[month]`
- `truck_utilization_metrics[month]`

The Date table runs to Dec 2025 because fuel purchases go up to 2 Jan 2025. Every Year slicer excludes 2025, so only 2022 to 2024 show.

---

## 5. DAX

### Calculated tables

```dax
Date =
VAR MinD = MIN ( MIN ( trips[dispatch_date] ), MIN ( fuel_purchases[purchase_date] ) )
VAR MaxD = MAX ( MAX ( trips[dispatch_date] ), MAX ( fuel_purchases[purchase_date] ) )
RETURN
CALENDAR ( DATE ( YEAR ( MinD ), 1, 1 ), DATE ( YEAR ( MaxD ), 12, 31 ) )
```

`DateAxis` is a disconnected table with one row per year and one row per year-month. A measure reads which row it is on. Set `DateAxis[Label]` to **Sort by column -> Sort**.

```dax
DateAxis =
VAR Ys = SUMMARIZE ( trips, 'Date'[Year] )
VAR Ms = SUMMARIZE ( trips, 'Date'[Year], 'Date'[MonthNo], 'Date'[Month] )
RETURN
UNION (
    SELECTCOLUMNS ( Ys,
        "Label", FORMAT ( 'Date'[Year], "0" ),
        "Level", "Year",
        "Yr", 'Date'[Year],
        "Mon", 0,
        "Sort", 'Date'[Year] * 100
    ),
    SELECTCOLUMNS ( Ms,
        "Label", 'Date'[Year] & " " & 'Date'[Month],
        "Level", "Month",
        "Yr", 'Date'[Year],
        "Mon", 'Date'[MonthNo],
        "Sort", 'Date'[Year] * 100 + 'Date'[MonthNo]
    )
)
```

### Calculated columns

These must be created with **New column** (not New measure).

```dax
// Date table
Year = YEAR ( 'Date'[Date] )
MonthNo = MONTH ( 'Date'[Date] )
Month = FORMAT ( 'Date'[Date], "MMM" )          // Sort by column -> MonthNo
YearMonth = FORMAT ( 'Date'[Date], "yyyy-MM" )  // text label (no longer used on a chart)
MonthStart = DATE ( YEAR ( 'Date'[Date] ), MONTH ( 'Date'[Date] ), 1 )  // real date, used on the fuel price axis

// routes table
Route = routes[origin_city] & " → " & routes[destination_city]
```

### Measures: revenue, cost and volume

```dax
Total Revenue = SUM ( loads[revenue] )
Fuel Cost = SUM ( fuel_purchases[total_cost] )
Maintenance Cost = SUM ( maintenance_records[total_cost] )
Profit = [Total Revenue] - [Fuel Cost] - [Maintenance Cost]   // before driver pay and overheads
Margin % = DIVIDE ( [Profit], [Total Revenue] )
Revenue after Fuel = [Total Revenue] - [Fuel Cost]
Total Loads = DISTINCTCOUNT ( loads[load_id] )
Total Miles = SUM ( trips[actual_distance_miles] )   // trips is the trusted source for miles
Revenue per Customer = DIVIDE ( [Total Revenue], DISTINCTCOUNT ( loads[customer_id] ) )
```

### Measures: service and safety

```dax
Incident Count = COUNTROWS ( safety_incidents )

// If on_time_flag is text, use = "true" instead of TRUE ()
On-Time % =
DIVIDE (
    CALCULATE ( COUNTROWS ( delivery_events ), delivery_events[on_time_flag] = TRUE () ),
    COUNTROWS ( delivery_events )
)

Avg Detention (min) = AVERAGE ( delivery_events[detention_minutes] )

// "& """ turns the flag into text, so it works for True/False or text columns
Preventable Accidents =
CALCULATE (
    [Incident Count],
    FILTER ( safety_incidents, safety_incidents[preventable_flag] & "" = "True" )
)
```

`Active Customers %` (lives in the customers table, now formatted as a percentage and showing 84%): the share of customers that are active:

```dax
Active Customers % = 
VAR TotalCustomers = DISTINCTCOUNT(customers[customer_id])
VAR ActiveCustomers = 
    CALCULATE(
        DISTINCTCOUNT(customers[customer_id]),
        customers[account_status] = "Active"
    )
RETURN
    DIVIDE(ActiveCustomers, TotalCustomers, 0)
```

### Measures: drivers

```dax
Avg MPG = DIVIDE ( SUM ( trips[actual_distance_miles] ), SUM ( trips[fuel_gallons_used] ) )
Revenue per Mile = DIVIDE ( [Total Revenue], SUM ( trips[actual_distance_miles] ) )
Incidents per Driver = DIVIDE ( [Incident Count], DISTINCTCOUNT ( trips[driver_id] ) )
Monthly MPG = DIVIDE ( SUM ( driver_monthly_metrics[total_miles] ), SUM ( driver_monthly_metrics[total_fuel_gallons] ) )
```

### Measures: fleet, fuel and maintenance

```dax
Miles per Truck = DIVIDE ( [Total Miles], DISTINCTCOUNT ( trips[truck_id] ) )
Revenue per Truck = DIVIDE ( [Total Revenue], DISTINCTCOUNT ( trips[truck_id] ) )
Fuel Cost per Mile = DIVIDE ( [Fuel Cost], [Total Miles] )
Maintenance Cost per Mile = DIVIDE ( [Maintenance Cost], [Total Miles] )
Total Cost per Mile = [Fuel Cost per Mile] + [Maintenance Cost per Mile]
Downtime Days = DIVIDE ( SUM ( maintenance_records[downtime_hours] ), 24 )
Avg Fuel Price = DIVIDE ( SUM ( fuel_purchases[total_cost] ), SUM ( fuel_purchases[gallons] ) )
```

### Measures: dynamic year to month axis

By default the chart shows years. When one year is picked in the Year slicer it switches to that year's months.

```dax
Dynamic Revenue by Date =
VAR _Level = SELECTEDVALUE ( DateAxis[Level] )
VAR _Yr = SELECTEDVALUE ( DateAxis[Yr] )
VAR _Mon = SELECTEDVALUE ( DateAxis[Mon] )
VAR _OneYear = HASONEVALUE ( 'Date'[Year] )
VAR _PickedYear = SELECTEDVALUE ( 'Date'[Year] )
RETURN
SWITCH (
    TRUE (),
    NOT _OneYear && _Level = "Year",
        CALCULATE ( [Total Revenue], KEEPFILTERS ( 'Date'[Year] = _Yr ) ),
    _OneYear && _Level = "Month" && _Yr = _PickedYear,
        CALCULATE ( [Total Revenue], KEEPFILTERS ( 'Date'[MonthNo] = _Mon ) ),
    BLANK ()
)
```

`Dynamic Loads by Date` is the same formula with `[Total Loads]` in place of `[Total Revenue]`.

### Which table each measure lives in (from the finished report)

| Table | Measures |
|---|---|
| loads | Fuel Cost, Maintenance Cost, Incident Count, Total Loads, On-Time %, Avg Detention (min), Revenue per Customer, Revenue after Fuel, Dynamic Revenue by Date |
| trips | Avg MPG, Revenue per Mile, Incidents per Driver, Preventable Accidents, Total Cost per Mile, Avg Fuel Price |
| customers | Active Customers % |
| driver_monthly_metrics | Monthly MPG |
| truck_utilization_metrics | Fuel Cost per Mile, Maintenance Cost per Mile, Miles per Truck, Revenue per Truck |
| maintenance_records | Downtime Days |

A measure works the same wherever it lives. The home table only decides where you find it in the Data pane.

---

## 6. Report design
**Status: finished.** All four pages are built, branded (LogiTrack), linked with page navigation buttons, and all planned corrections are done. Section 12 lists only optional extras.
- **Branding:** LogiTrack logo (top left) and a truck image (top right), both added as images.
- **Navigation:** a page navigator (four buttons) in the same spot under the title on every page. The active page's button is dark.
- **Left panel:** for the slicers
- **KPI cards:** five across the top of each page.
- **Layout on every page:** title and navigation, KPI row, charts, then tables.

---

## 7. The four final pages

### Page 1: Executive Overview (how is the business doing?)

| Visual | Type | Fields |
|---|---|---|
| Year | tile slicer | Date[Year] (2025 excluded) |
| Revenue by customer type | tile slicer | customers[customer_type], labelled with revenue |
| Total Revenue 262.53M | card | Sum of loads[revenue] |
| Total load 85.41K | card | Count of loads[load_id] |
| Fuel Cost 95.59M | card | measure Fuel Cost |
| Total distance 122M | card | Sum of trips[actual_distance_miles] |
| Incident Count 170 | card | measure Incident Count |
| Revenue by credit_terms_days | column chart | customers[credit_terms_days] x Sum of loads[revenue] |
| Total Loads by Month | line chart | Date[Month] x Total Loads (seasonal pattern) |
| Revenue by Cargo Category | bar chart | customers[primary_freight_type] x Sum of loads[revenue] |
| Dynamic Revenue by Date | line chart | DateAxis[Label] x Dynamic Revenue by Date |
| Revenue by city | map | routes[origin_city], size = Sum of loads[revenue] |

### Page 2: Customer Route and Delivery Service (which lanes and clients are profitable, and are we delivering well?)

| Visual | Type | Fields |
|---|---|---|
| Year, Facility type | tile slicers | Date[Year], facilities[facility_type] |
| On-Time % 55.67% | card | measure On-Time % |
| Avg Detention (min) 92 | card | measure Avg Detention (min) |
| Revenue per Customer 1.31M | card | measure Revenue per Customer |
| Active Customers % 84% | card | measure Active Customers % (formatted as a percentage) |
| Unique Customer 200 | card | Distinct count of loads[customer_id] |
| On-Time % and Revenue by Customer | scatter | customers[customer_name], X = On-Time %, Y = Sum of loads[revenue] |
| Top routes by revenue after fuel | table | routes[Route], Sum of revenue, Fuel Cost, Revenue after Fuel. Top 10 by Revenue after Fuel |
| Routes losing money after fuel | table | same fields. Top 10, lowest Revenue after Fuel first |
| Top 10 Customer by Revenue | table | customers[customer_name], Sum of loads[revenue]. Top 10 |
| Revenue by Customer type | donut | customers[customer_type] x Sum of loads[revenue] |

### Page 3: Driver Performances and Safety (who performs well, who is a risk?)

| Visual | Type | Fields |
|---|---|---|
| Year, Employment status, Preventable Incident | tile slicers | Date[Year], drivers[employment_status], safety_incidents[preventable_flag] (labelled with Incident Count) |
| On-Time %, Preventable Accidents, Avg MPG, Incidents per Driver, Revenue per Mile | 5 cards | the five measures of the same names |
| Monthly Mile Per Gallon | line chart | driver_monthly_metrics[month] x Monthly MPG |
| Top drivers with incident counts | table | drivers[driver_id], Avg MPG, On-Time %, Revenue per Mile, Incident Count. Top 10 by Incident Count, plus a deliberate second filter, **On-Time % below 56.2%**, that trims tied drivers (see the note under this table) |
| Avg MPG and On-Time % by driver_id | scatter | drivers[driver_id], X = Avg MPG, Y = On-Time % |
| Count of incident types | column chart | safety_incidents[incident_type] x Incident Count |
| Delivery: on time vs late | donut | delivery_events[on_time_flag] x Count of delivery_events[load_id] (counts delivery events) |

**Note: why the leaderboard has two filters (deliberate)**

- A Top 10 filter in Power BI keeps **ties**. After DRV00058 (7 incidents) and DRV00003 (6), many drivers are tied at 3 incidents, so Top 10 alone showed 12 to 15 rows.
- A second visual filter, **On-Time % is less than 0.562**, removes enough of the tied drivers to leave exactly 10 rows.
- Both are set in the **Filters** pane under **Filters on this visual**: `driver_id` (Top 10 by Incident Count) and `On-Time %` (less than 0.562). Do not clear either one.
- Caveat for presenting: the 10-row result is exact for the all-years view. If a Year is picked, the ties change, so the table can show a few more or fewer rows.
- The empty filter cards for Avg MPG, Revenue per Mile and Incident Count in the same pane are just placeholders and do nothing.

### Page 4: Fleet Fuel and Maintenance (are trucks earning money or costing it?)

| Visual | Type | Fields |
|---|---|---|
| Year | tile slicer | Date[Year] (2025 excluded) |
| Fuel Cost per Mile, Maintenance Cost per Mile, Miles per Truck, Revenue per Truck, Downtime Days | 5 cards | the five measures of the same names |
| Downtime Days and Maintenance Cost by truck_id | scatter | trucks[truck_id], X = Downtime Days, Y = Maintenance Cost |
| Average of utilization_rate by month | line chart | truck_utilization_metrics[month] x **Average** of utilization_rate. Kept as a plain average on a 0.80 to 0.86 scale (no percentage format), by choice |
| Maintenance type table | table | maintenance_records[maintenance_type], Count of truck_id (renamed "No of truck") |
| Avg Fuel Price by MonthStart | line chart | Date[MonthStart] x Avg Fuel Price |
| Truck ranking | table | trucks[truck_id], Total distance in Miles, Revenue per Mile, Fuel Cost per Mile, Maintenance Cost per Mile, Total Cost per Mile. Top 10 by Total Cost per Mile, sorted by Total Cost per Mile (highest first) |

---

## 8. Key findings
<img width="2000" height="1129" alt="Untitled Design" src="https://github.com/user-attachments/assets/f5910da4-66e7-4039-8353-5263f8b75551" />

**Business**
- Revenue 262.53M; fuel cost 95.59M; maintenance about 5.7M; profit before driver pay and overheads about 161M.
- Revenue is almost identical each year (about 87.5M), and loads are about 28K a year.
- Customer type: Contract 98.8M (37.6%), Spot 82.85M (31.6%), Dedicated 80.88M (30.8%).
- By credit terms: 60 days 73M, 30 days 73M, 15 days 64M, 45 days 53M.
- By cargo: Automotive 55M, Food/Beverage 49M, Retail 46M, Electronics 43M, Consumer Goods 39M, General 31M.
- 200 customers, 84% active. Revenue per customer averages 1.31M. The largest customer is First Group (about 9.1M).

**Service**
- On-time rate is 55.67% (95,095 of 170,820 delivery events), and 75.73K events were late.
- On-time % sits between roughly 50% and 58% for almost every customer, so service is about equal for all.
- Average detention is 92 minutes, and every facility sits between 91 and 94.

**Drivers and safety**
- 170 incidents, of which 64 are preventable. That is 1.36 incidents per driver.
- Incident types: DOT Violation 39, Accident 35, Equipment Damage 35, Customer Complaint 34, Moving Violation 27.
- DRV00058 has 7 incidents and DRV00003 has 6. Most drivers with incidents have 3 or fewer.
- MPG (6.41 to 6.49 per driver), on-time % and revenue per mile are almost identical across drivers. Incidents are what separates them.

**Fleet and fuel**
- Fuel cost per mile 0.78, maintenance cost per mile 0.05, total 0.83. Revenue per mile 2.15.
- Costliest truck is TRK00014 at 0.91 per mile.
- Trucks with more shop days also cost more to repair. Downtime totals about 3,010 days (about 32 per truck over three years).
- Average truck utilization sits around 80% to 86%.
- Fuel price falls in steps at the start of each year: about 4.2 (2022), 3.85 (2023), 3.65 (2024).
- Fuel price does not depend on city (every city is about 3.90).
- Maintenance jobs by type: Inspection 432, Repair 424, Preventive 422, Engine 415, Tire 414, Transmission 411, Brake 402 (2,920 in total).

---

## 10. Business insights and strategic recommendations

This section turns the dashboard numbers into business conclusions and actions. Figures cover 2022 to 2024 and come from the four dashboards.

### 10.1 Business insights

| Area | Insight | Evidence |
|---|---|---|
| Revenue | Steady and balanced. No single customer or cargo type carries the business. | 262.53M in total, about 87.5M a year. Contract 37.6%, Spot 31.6%, Dedicated 30.8%. The biggest customer (First Group, 9.14M) is only about 3.5% of revenue, and the top 10 customers together are about 20%. |
| Cost | Fuel is the big cost, and maintenance is small. | Fuel 95.59M is about 36 cents of every revenue dollar. Maintenance is about 5.7M, about 2 cents. |
| Margin | A healthy margin before driver pay and overheads. | Cost per mile is 0.83 (0.78 fuel + 0.05 maintenance) against 2.15 revenue per mile. That leaves about 61% of revenue (161.20M). |
| Service | Late delivery is a system-wide problem, not a few bad customers. | On-time is 55.67% (95.1K of 170.8K deliveries), with 75.7K late. Almost every customer sits between 50% and 58%, every facility has detention between 91 and 94 minutes, and drivers all sit around 54% to 56%. |
| Safety | Risk is concentrated in a few drivers and in preventable events. | 170 incidents, 64 of them preventable (38%). DRV00058 has 7 and DRV00003 has 6, while most others have 3 or fewer. DOT violations are the most common type (39). MPG and on-time barely differ between drivers, so incidents are the real risk signal. |
| Fleet | Repair cost rises with time in the shop. | Trucks with more downtime days also cost more to repair. Total downtime is about 3,010 days (about 32 per truck over three years). The costliest truck, TRK00014, is 0.91 per mile. |
| Routes | A small group of routes loses money after fuel. | 7 routes lose a combined 5.1M after fuel, led by New York to Philadelphia (-1.42M) and Philadelphia to New York (-1.39M). The best route, Philadelphia to Seattle, earns 8.46M after fuel. |
| Fuel price | Fuel is getting cheaper, in yearly steps. | About 4.2 (2022), 3.85 (2023) and 3.65 (2024) per gallon, a drop of about 13%. Read from the dashboard chart. |

### 10.2 Strategic recommendations

| Priority | Recommendation | Why | How to measure it |
|---|---|---|---|
| 1 | Treat on-time delivery as a process problem. Review dispatch planning, appointment scheduling and detention handling, and set a target. | 55.67% is low and equally low everywhere, so the cause is shared. Detention averages 92 minutes at every facility. | On-Time % each month |
| 2 | Review the 7 loss-making routes. Check rates and real route mileage before repricing or dropping them. | They lose a combined 5.1M after fuel. | Revenue after Fuel per route |
| 3 | Coach the high-incident drivers and target preventable events, starting with DOT violations. | 64 of 170 incidents are preventable, and two drivers have 6 to 7 each. | Incidents per Driver, Preventable Accidents |
| 4 | Plan inspections or replacement for the highest-downtime trucks, and shift work toward preventive maintenance. | Downtime and repair cost rise together. Some trucks lose 40 to 47 days. | Downtime Days, Maintenance Cost per Mile |
| 5 | Make fuel efficiency the main cost lever. Track MPG, review fuel contracts and fuel surcharges now that prices have fallen. | Fuel is about 36% of revenue. As an illustration, a 5% MPG gain would cut fuel use by about 4.8% and save about 4.5M over three years, at the same miles and price. | Avg MPG, Fuel Cost per Mile |
| 6 | Add driver pay and overheads to the data so profit is real. | Today's margin is before those costs, so it is overstated. | Profit and Margin % |

## 11. Data quirks

- About 2% of trips and fuel purchases have no truck or driver recorded. They show as a blank row.
- Total Miles from trips (about 122M) is about 2% higher than the sum in the monthly summary table (about 120M), because trips with no truck or driver are missing from the summaries. Use trips for KPIs.
- Fuel cost is about 1.6M on every route, so "revenue after fuel" by route is mostly driven by revenue.
- Fuel price steps down once a year, which suggests one price level per year in this dataset, not real month-to-month market swings.
- Driver pay is not in the data, so profit and margin are before driver pay and overheads.

---

## 12. Mistakes made and lessons learned

1. **New measure vs New column.** `Route` and `YearMonth` must be **columns**. Created as measures they give "a single value cannot be determined".
2. **"Sum of ..." fields are wrong for ratios.** MPG, utilization and prices must be measures or Average. Adding them up gives nonsense (utilization read 74 to 80 until it was set to Average).
3. **Top N tables show a wrong Total.** The total only covers the rows shown. Turn off Totals or title the table "Top 10".
4. **Two active paths to one table.** Power BI switches one off. Fixed by deleting the fuel_purchases links to drivers and trucks.
5. **Check a flag's values before writing DAX.** `on_time_flag` is true/false, not 1/0.
6. **Use the right truck_id.** A table that mixes trips and maintenance must use `trucks[truck_id]`. Using `trips[truck_id]` divides total maintenance by one truck's miles.
7. **Axis choices.** Month on the axis averages the years together, so use Year (or a month-start date) to see trends. Bar charts should start at 0. Line charts do not need to.
8. **Use a real date for line chart axes.** A text axis (YearMonth) shows every label tilted. A date axis (MonthStart, Continuous) spaces labels cleanly.
9. **Chart labels are cut off with "...".** Y-axis labels cannot wrap in Power BI. Widen **Y-axis -> Values -> Maximum width**, or shorten the label.
10. **Top N keeps ties.** A Top 10 filter shows every item tied at the 10th value. Here many drivers tied at 3 incidents, so a second filter (On-Time % below 56.2%) trims the table to exactly 10.

---
