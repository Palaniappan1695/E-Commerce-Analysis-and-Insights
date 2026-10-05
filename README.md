🛒 E-Commerce Sales Analysis using Excel
📌 Project Overview

This project analyzes an E-Commerce Sales Dataset using Microsoft Excel to identify sales patterns, product performance, customer behavior, and factors influencing revenue.

The project follows the four fundamental types of data analysis:

Descriptive Analysis – What happened?
Diagnostic Analysis – Why did it happen?
Predictive Analysis – What could happen next?
Prescriptive Analysis – What should the business do?

The objective is to transform raw e-commerce transaction data into meaningful and actionable business insights that can support better sales and business decisions.

🎯 Project Objective

The primary objective of this project is to analyze e-commerce sales data and identify meaningful patterns and trends in business performance.

The analysis focuses on:

Sales revenue
Product and category performance
Quantity sold
Pricing
Discounts
Customer behavior
Payment methods
Regional performance
Sales trends over time

The project also explores why certain sales patterns occur, what trends may continue in the future, and what actions could help improve sales and profitability.

📊 Dataset Description

The dataset contains transactional e-commerce sales information covering customers, products, stores, sales quantities, unit prices, discounts, payment types, and revenue.

Key fields include:

Field	Description
Month	Month of the order
Year	Year of the order
Customer ID	Unique customer identifier
Gender	Customer gender
Loyalty Level	Customer loyalty category
Product ID	Unique product identifier
Category	Product category
Sub Category	Product sub-category
Store ID	Unique store identifier
Region	Store region
City	Store city
Store Type	Type of store
Quantity	Quantity of products sold
Unit Price	Unit price of the product
Discount	Discount provided to the customer
Payment Type	Payment method used
Revenue	Revenue generated from the sale




🧹 Data Cleaning & Preparation

The raw dataset was prepared before performing the analysis.

Data preparation activities included:
Converted raw data into an Excel Table
Cleaned text using TRIM, CLEAN, and PROPER
Handled missing loyalty-level values
Checked and handled duplicate values
Identified missing values
Filled missing unit prices using product-level average prices
Filled missing quantity values using calculated averages
Recalculated total sales amounts
Split date information into Month and Year
Combined information from different dimension tables

Example formula used for cleaning names:

=PROPER(TRIM(CLEAN([@Name])))




🧮 Data Transformation & Calculations

Several Excel formulas and techniques were used to prepare the dataset for analysis.

Total Amount
=Quantity*Unit_Price*(1-Discount)
Unit Price

Unit price was derived using:

Unit Price = Total Amount / (Quantity × (1 − Discount))

Product-level average prices were also calculated and used to fill missing unit-price values.

Discount
Discount % = (Gross Amount − Total Amount) / Gross Amount × 100
Other Excel Techniques
IF
OR
ISBLANK
AVERAGE
XLOOKUP
VLOOKUP
TEXT
PivotTables
Excel Data Analysis ToolPak




📈 Data Analysis
1. Descriptive Analysis

The descriptive analysis focuses on understanding the overall sales performance.

Key questions:

What is the total sales revenue?
Which product category generated the highest sales?
Which region performed best?
Which year had the highest sales?
2. Diagnostic Analysis

The diagnostic analysis explores the reasons behind the observed sales patterns.

Key questions:

Why did the top-performing category generate higher sales?
Why is profit higher or lower for certain products?
What factors may be influencing category performance?
3. Predictive Analysis

The predictive analysis uses historical sales trends to understand possible future performance.

Key questions:

Based on the sales trend, what could next month's sales be?
Which category is likely to perform well in the future?
4. Prescriptive Analysis

The prescriptive analysis converts the findings into possible business actions.

Key questions:

Which category should the company focus on?
Should the company increase or reduce discounts?
What three actions could improve profitability?




📊 Dashboard

An Excel dashboard was created to present the analysis in a simple and visual format.

The dashboard helps communicate:

Overall sales performance
Category performance
Regional performance
Sales trends
Customer and payment behavior
Future sales expectations

The dashboard was designed to convert the analytical results into an easy-to-understand business story.

💡 Key Insights

The analysis identified several important findings:

The dataset contains 2,000 orders and 4,960 items sold.
Total sales revenue was approximately $2.17 million.
Average order value was approximately $1,087.
Sports was the highest-performing product category, followed by Home and Electronics.
Sports and Home showed balanced purchasing patterns across male and female customers.
Credit Cards and PayPal were the major payment methods.
The forecast suggested future monthly sales could remain within an approximate $10,000–$15,000 range based on the trend analysis.
Electronics and Sports were identified as categories with potential for continued strong performance.
The West region represented approximately 12% of sales, indicating an opportunity for improvement.




💼 Business Recommendations

Based on the analysis, the following actions were recommended:

Increase marketing and inventory focus on high-performing categories such as Sports and Home.
Consider bundle offers instead of excessive percentage discounts to encourage purchases while protecting profitability.
Focus on improving the West region's performance through targeted promotions and marketing campaigns.
Continue monitoring sales trends and category performance to identify changes in customer demand.
🛠️ Tools & Skills Used
Tools
Microsoft Excel
Excel PivotTables
Excel Charts & Dashboard
Excel Data Analysis ToolPak
Excel Skills
Data Cleaning
Data Transformation
IF
OR
ISBLANK
AVERAGE
XLOOKUP
VLOOKUP
TEXT
PivotTables
Descriptive Statistics
Forecasting
Data Visualization
Business Insights

