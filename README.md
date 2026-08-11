E-COMMERCE CONVERSION ANALYTICS — POWER BI DASHBOARD
=================================

Purpose
-------
This document provides a complete step-by-step guide for creating the
E-Commerce Conversion Analytics dashboard shown in the supplied dashboard
snapshot. It covers Excel import, Power Query cleaning, DAX measures,
dashboard visuals, slicers, formatting, and the final dashboard layout.

DASHBOARD SNAPSHOT
------------------
Image: Images/00_complete_dashboard.png

The dashboard contains:
• Header: E-COMMERCE CONVERSION ANALYTICS
• Slicers: Date, Traffic Source, Device, Region, Category
• KPI Cards: Total Visitors, Total Purchases, Conversion Rate,
  Total Revenue, Average Order Value, Total Profit
• Charts: Monthly Conversion Rate, Traffic Source Conversion,
  Customer Conversion Funnel, Revenue by Product Category,
  Revenue by Traffic Source, Conversion Rate by Device
• Detail Matrix: Category & Device Performance

----
PART 1 — IMPORT THE EXCEL DATASET
----

Open Power BI Desktop.

STEP 1 — Import Excel
---------------------
1. Go to Home → Get Data → Excel Workbook.
2. Select:

   ecommerce_conversion_rate_10000_rows.xlsx

3. In Navigator, select Sheet1.
4. Click Transform Data, not Load, because the data will be cleaned first.


----
PART 2 — POWER QUERY CLEANING
----
We will perform 15 cleaning and preparation steps.

STEP 1 — Check Column Names
---------------------------
Verify that the columns are:

Session_ID
Date
Customer_ID
Customer_Type
Traffic_Source
Device
Region
Product_Category
Product_Name
Visitors
Product_Views
Add_To_Cart
Checkout
Purchases
Order_ID
Quantity
Unit_Price
Discount_Pct
Revenue
Cost
Profit
Month
Month_Name

If Power Query shows Column1, Column2, etc.:
Home → Use First Row as Headers.


STEP 2 — Remove Duplicate Sessions
-----------------------------------
1. Select Session_ID.
2. Go to Home → Remove Rows → Remove Duplicates.
3. Every session should have a unique Session_ID.


STEP 3 — Trim and Clean Text Columns
-------------------------------------
Select:
• Customer_Type
• Traffic_Source
• Device
• Region
• Product_Category
• Product_Name

Then:
Transform → Format → Trim
and
Transform → Format → Clean

This removes unnecessary spaces and invisible characters.


STEP 4 — Set Date Type
----------------------
1. Select Date.
2. Transform → Data Type → Date.
3. Confirm that the column uses the calendar/date data type.


STEP 5 — Set Text Data Types
----------------------------
Set these columns to Text:

Session_ID
Customer_ID
Customer_Type
Traffic_Source
Device
Region
Product_Category
Product_Name
Order_ID
Month_Name


STEP 6 — Set Whole Number Columns
----------------------------------
Set these columns to Whole Number:

Visitors
Product_Views
Add_To_Cart
Checkout
Purchases
Quantity
Month


STEP 7 — Set Decimal Number Columns
-----------------------------------
Set these columns to Decimal Number:

Unit_Price
Discount_Pct
Revenue
Cost
Profit


STEP 8 — Check Blank Values
----------------------------
Check these important columns for (null) or blank values:

Date
Traffic_Source
Device
Product_Category
Visitors
Purchases
Revenue

Remove genuinely invalid blank records if they occur.


STEP 9 — Check Numeric Errors
-----------------------------
Select the numeric columns.

To inspect errors:
Home → Keep Rows → Keep Errors

After checking the errors, remove invalid records if necessary:
Home → Remove Rows → Remove Errors


STEP 10 — Check Negative Values
--------------------------------
Check these columns:

Visitors
Product_Views
Add_To_Cart
Checkout
Purchases
Quantity
Unit_Price
Revenue
Cost
Profit

Filter out impossible negative values. For example, Revenue < 0 should not
normally be present in this conversion dashboard.


STEP 11 — Validate the Conversion Funnel
----------------------------------------
The customer journey should logically follow:

Visitors
   ↓
Product Views
   ↓
Add To Cart
   ↓
Checkout
   ↓
Purchases

Check that:

Purchases <= Checkout
Checkout <= Add_To_Cart

These relationships make the funnel meaningful.


STEP 12 — Create Month Number
-----------------------------
If Month is not already available:

Add Column → Date → Month → Month

Name the resulting column:

Month_Number

Use the month number for correct chronological sorting.


STEP 13 — Create Month Name
---------------------------
If required:

Add Column → Date → Month → Name of Month

Name it:

Month_Name_New

You can use this column instead of the existing Month_Name.


STEP 14 — Sort Month Name Correctly
-----------------------------------
After loading the data into Power BI:

1. Select Month_Name.
2. Go to Column tools → Sort by column → Month
   (or Month_Number if that is the column you created).

This prevents alphabetical month ordering and gives:

January
February
March
April
May
June
July
August
September
October
November
December


STEP 15 — Rename the Query
---------------------------
In the Queries panel, rename:

Sheet1

to:

Ecommerce_Data

Then click:

Home → Close & Apply.


----
PART 3 — DAX MEASURES
----

Go to:
Report View → Modeling → New Measure

KPI MEASURE 1 — TOTAL VISITORS
------------------------------
Total Visitors =
SUM(Ecommerce_Data[Visitors])


KPI MEASURE 2 — TOTAL PRODUCT VIEWS
------------------------------------
Total Product Views =
SUM(Ecommerce_Data[Product_Views])


KPI MEASURE 3 — TOTAL ADD TO CART
----------------------------------
Total Add To Cart =
SUM(Ecommerce_Data[Add_To_Cart])


KPI MEASURE 4 — TOTAL CHECKOUTS
--------------------------------
Total Checkouts =
SUM(Ecommerce_Data[Checkout])


KPI MEASURE 5 — TOTAL PURCHASES
-------------------------------
Total Purchases =
SUM(Ecommerce_Data[Purchases])


KPI MEASURE 6 — TOTAL REVENUE
-----------------------------
Total Revenue =
SUM(Ecommerce_Data[Revenue])

Format this measure as Currency.


----
PART 4 — CONVERSION MEASURES
----

1. OVERALL CONVERSION RATE
--------------------------
Conversion Rate =
DIVIDE(
    [Total Purchases],
    [Total Visitors],
    0
)

Format as Percentage with 2 decimal places.


2. PRODUCT VIEW RATE
--------------------
Product View Rate =
DIVIDE(
    [Total Product Views],
    [Total Visitors],
    0
)


3. ADD TO CART RATE
-------------------
Add To Cart Rate =
DIVIDE(
    [Total Add To Cart],
    [Total Visitors],
    0
)


4. CHECKOUT RATE
----------------
Checkout Rate =
DIVIDE(
    [Total Checkouts],
    [Total Add To Cart],
    0
)


5. PURCHASE COMPLETION RATE
---------------------------
Purchase Completion Rate =
DIVIDE(
    [Total Purchases],
    [Total Checkouts],
    0
)


6. CART ABANDONMENT RATE
------------------------
Cart Abandonment Rate =
1 -
DIVIDE(
    [Total Purchases],
    [Total Add To Cart],
    0
)

Format as Percentage.

----
PART 5 — REVENUE AND PROFIT MEASURES
----

AVERAGE ORDER VALUE
-------------------
Average Order Value =
DIVIDE(
    [Total Revenue],
    [Total Purchases],
    0
)


TOTAL COST
----------
Total Cost =
SUM(Ecommerce_Data[Cost])


TOTAL PROFIT
------------
Total Profit =
SUM(Ecommerce_Data[Profit])


PROFIT MARGIN
-------------
Profit Margin =
DIVIDE(
    [Total Profit],
    [Total Revenue],
    0
)


REVENUE PER VISITOR
-------------------
Revenue Per Visitor =
DIVIDE(
    [Total Revenue],
    [Total Visitors],
    0
)


AVERAGE QUANTITY PER ORDER
--------------------------
Average Quantity Per Order =
DIVIDE(
    SUM(Ecommerce_Data[Quantity]),
    [Total Purchases],
    0
)

----
PART 6 — CUSTOMER MEASURES
----

NEW CUSTOMERS
-------------
New Customers =
CALCULATE(
    DISTINCTCOUNT(Ecommerce_Data[Customer_ID]),
    Ecommerce_Data[Customer_Type] = "New"
)


RETURNING CUSTOMERS
-------------------
Returning Customers =
CALCULATE(
    DISTINCTCOUNT(Ecommerce_Data[Customer_ID]),
    Ecommerce_Data[Customer_Type] = "Returning"
)


RETURNING CUSTOMER RATE
-----------------------
Returning Customer Rate =
DIVIDE(
    [Returning Customers],
    DISTINCTCOUNT(Ecommerce_Data[Customer_ID]),
    0
)

--------------------------------------------------
Create a rectangle across the top.

Title:
E-COMMERCE CONVERSION ANALYTICS

Optional subtitle:
Customer Journey • Conversion • Revenue • Performance

The supplied dashboard uses a strong magenta title and a clean white header.


----
SECOND SECTION — SLICERS
----

Place five slicers underneath the title:

Date | Traffic Source | Device | Region | Category

SLICER 1
--------
Field: Date
Style: Between


SLICER 2
--------
Field: Traffic_Source


SLICER 3
--------
Field: Device


SLICER 4
--------
Field: Region


SLICER 5
--------
Field: Product_Category


----
THIRD SECTION — KPI CARDS
----

Use six cards arranged horizontally:

1. TOTAL VISITORS
   Measure: [Total Visitors]

2. TOTAL PURCHASES
   Measure: [Total Purchases]

3. CONVERSION RATE
   Measure: [Conversion Rate]

4. TOTAL REVENUE
   Measure: [Total Revenue]

5. AVERAGE ORDER VALUE
   Measure: [Average Order Value]

6. TOTAL PROFIT
   Measure: [Total Profit]

Recommended card style:
• White background
• Rounded/subtle border
• Soft shadow
• Large value
• Smaller bold label
• Consistent spacing


----
FOURTH SECTION — CONVERSION TREND AND CHANNEL ANALYSIS
----

VISUAL 1 — MONTHLY CONVERSION RATE
-----------------------------------
Chart:
Line Chart

X-axis:
Month_Name

Y-axis:
[Conversion Rate]

Title:
Monthly Conversion Rate

Important:
Sort Month_Name by Month or Month_Number.

CHART SNAPSHOT:
Images/01_monthly_conversion_rate.png


VISUAL 2 — TRAFFIC SOURCE CONVERSION
-------------------------------------
Chart:
Clustered Bar Chart

Y-axis:
Traffic_Source

X-axis:
[Conversion Rate]

Title:
Traffic Source Conversion

Turn on data labels.

CHART SNAPSHOT:
Images/02_traffic_source_conversion.png

----
FIFTH SECTION — CUSTOMER CONVERSION FUNNEL
----

Create a Funnel Chart.

Power BI funnel visuals need a category/value structure. Create a small
disconnected funnel table.

Go to:
Modeling → New Table

Funnel_Stages =
DATATABLE(
    "Stage", STRING,
    "SortOrder", INTEGER,
    {
        {"Visitors", 1},
        {"Product Views", 2},
        {"Add To Cart", 3},
        {"Checkout", 4},
        {"Purchases", 5}
    }
)


Create the funnel measure:

Funnel Value =
SWITCH(
    SELECTEDVALUE(Funnel_Stages[Stage]),
    "Visitors", [Total Visitors],
    "Product Views", [Total Product Views],
    "Add To Cart", [Total Add To Cart],
    "Checkout", [Total Checkouts],
    "Purchases", [Total Purchases]
)


Sort Stage by:
SortOrder

Funnel chart fields:
Category → Stage
Values → Funnel Value

Title:
Customer Conversion Funnel

CHART SNAPSHOT:
Images/03_customer_conversion_funnel.png


----
SIXTH SECTION — CATEGORY AND DEVICE ANALYSIS
----

VISUAL 4 — REVENUE BY PRODUCT CATEGORY
---------------------------------------
Chart:
Clustered Column Chart

X-axis:
Product_Category

Y-axis:
[Total Revenue]

Title:
Revenue by Product Category

Turn on data labels.

CHART SNAPSHOT:
Images/04_revenue_by_product_category.png


VISUAL 5 — CONVERSION RATE BY DEVICE
-------------------------------------
Chart:
Bar Chart

Axis:
Device

Values:
[Conversion Rate]

Title:
Conversion Rate by Device

A bar chart is recommended because conversion-rate comparisons are easier
to read than a donut chart.

CHART SNAPSHOT:
Images/06_conversion_rate_by_device.png


----
SEVENTH SECTION — TRAFFIC SOURCE REVENUE
----

VISUAL 6 — REVENUE BY TRAFFIC SOURCE
-------------------------------------
Chart:
Bar Chart

Y-axis:
Traffic_Source

X-axis:
[Total Revenue]

Title:
Revenue by Traffic Source

Turn on Data labels.

CHART SNAPSHOT:
Images/05_revenue_by_traffic_source.png


----
EIGHTH SECTION — CATEGORY AND DEVICE PERFORMANCE MATRIX
----

Add a Matrix visual.

Rows:
Product_Category

Columns:
Device

Values:
[Total Visitors]
[Total Purchases]
[Conversion Rate]
[Total Revenue]
[Average Order Value]
[Profit Margin]

Title:
Category & Device Performance

CHART SNAPSHOT:
Images/07_category_device_performance_matrix.png

================================================================
END OF E-COMMERCE CONVERSION ANALYTICS DASHBOARD GUIDE
================================================================
