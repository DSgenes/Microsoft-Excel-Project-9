# Microsoft-Excel-Project-9

# Case study

Jamie, at Adventure Works, is attending a management meeting. She has been asked to prepare an Excel worksheet that presents sales figures for the first quarter of the year and compares these figures to the results for the same period in the previous year. 

This worksheet is called Summary and is in the workbook Quarter One Report.xlsx. In this worksheet, you’ll need to complete the following actions:

⦁ Create formulas that show the total quarter-one sales for both 2022 and 2023.

⦁ Create formulas that show the percentage increase in sales in 2023. 

⦁ And break down these totals by month with the use of further calculations.
______________________________________________________________________________________________________________________________________________________________________________________________________________________________________________

# Executive Data Summary: Q1 Sales Performance Analysis

# Project Overview

This project involved analyzing the month-by-month profit margin performance for Quarter 1 (Q1) of a business for 2023 and comparing it to the same period in 2022. The goal was to present key data in a clear and impactful executive summary. Various Excel techniques, including data formatting, calculations, and functions, were applied to transform raw sales data into actionable insights.

# Key Tasks and Techniques

# 1. Data Organization and Formatting

      ⦁ Adjusted column widths and merged cells to make the report visually appealing.
      ⦁ Added headings and applied formatting to ensure clarity and focus.
      ⦁ Inserted new columns for Product ID and formatted text using the PROPER function.

# 2. Data Customization

      ⦁ Sorted the data by date from oldest to newest.
      ⦁ Used the Freeze Panes option to keep headers static while scrolling.
      ⦁ Removed irrelevant data by hiding unnecessary columns.

# 3. Formulas and Calculations

      ⦁ MONTH and YEAR functions: Extracted month and year from date values for use in calculations.
      ⦁ SUMIF function: Calculated quarterly sales totals for 2022 and 2023 by month and year.
      ⦁ IF function: Applied tax calculation logic based on order value.
      ⦁ Percentage Difference: Calculated the percentage increase in sales for Q1 2023 compared to 2022.

# 4. Data Insights

      ⦁ Summarized sales for each month of Q1 (January, February, March) for both years (2022 and 2023).
      ⦁ Displayed profit margin changes using percentage difference calculations.

# Technical Details

      ⦁ Tools Used: Microsoft Excel
      ⦁ Functions Implemented:
          ⦁ =PROPER(G2)
          ⦁ =MONTH(J2)
          ⦁ =YEAR(J2)
          ⦁ =SUMIF(...)
          ⦁ =IF(P2>2000, P2*5%, 0)
          ⦁ Percentage Difference Formula: =(C6-B6)/B6
          
# Sample Formulas

      ⦁ SUMIF for Q1 Sales 2022:
      ⦁ =SUMIF(L2:L246,2022,R2:R246)
      ⦁ Result: $330,500

# SUMIF for Q1 Sales 2023:

      ⦁ =SUMIF(L2:L246,2023,R2:R246)
      ⦁ Result: $453,830

# Percentage Difference:

      ⦁ =(C6-B6)/B6
      ⦁ Result: 37.32%

# Conclusion

This exercise transformed raw sales data into a professional and actionable executive summary using advanced Excel features. The skills learned included data formatting, sorting, filtering, and applying advanced Excel functions to calculate key performance indicators such as total sales, tax, and percentage differences in sales growth.

