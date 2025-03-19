# Learning Microsoft Excel

## Table of Contents

- [Learning Microsoft Excel](#learning-microsoft-excel)
- [Pivot Tables](#pivot-tables)
- [Formulas](#formulas)
  - [MAX & MIN](#max--min)
  - [IF & IFS](#if--ifs)
  - [LEN](#len)
  - [LEFT & RIGHT](#left--right)
  - [Converting Date to Text](#converting-date-to-text)
  - [TRIM](#trim)
  - [SUBSTITUTE](#substitute)
  - [SUM & SUMIF](#sum--sumif)
  - [COUNT & COUNTIF](#count--countif)
  - [CONCATENATE](#concatenate)
  - [DAYS & NETWORKDAYS](#days--networkdays)
- [XLOOKUP & VLOOKUP](#xlookup--vlookup)
  - [VLOOKUP](#vlookup)
  - [XLOOKUP (Basic)](#xlookup-basic)
  - [XLOOKUP (Multiple Rows)](#xlookup-multiple-rows)
  - [XLOOKUP (Exact Match)](#xlookup-exact-match)
  - [XLOOKUP (Search Order)](#xlookup-search-order)
  - [XLOOKUP (Horizontal)](#xlookup-horizontal)
  - [XLOOKUP with Sum](#xlookup-with-sum)
- [Conditional Formatting](#conditional-formatting)
  - [Example 1 - Monthly Stationary Sales](#example-1---monthly-stationary-sales)
  - [Example 2 - Highlighting Specific Thresholds](#example-2---highlighting-specific-thresholds)
- [Charts](#charts)
  - [Example 1 - Monthly Breakdown](#example-1---monthly-breakdown)
  - [Example 2 - Trend Analysis](#example-2---trend-analysis)
  - [Example 3 - Year-End Distribution](#example-3---year-end-distribution)
- [Cleaning Data in Excel](#cleaning-data-in-excel)
  - [Common Data Cleaning Steps](#common-data-cleaning-steps)
    - [1. Open and Inspect the Data](#1-open-and-inspect-the-data)
    - [2. Standardize Column Names](#2-standardize-column-names)
    - [3. Remove Duplicate Rows](#3-remove-duplicate-rows)
    - [4. Trim Extra Spaces in Data Cells](#4-trim-extra-spaces-in-data-cells)
    - [5. Handle Missing Values](#5-handle-missing-values)
    - [6. Convert Data Types](#6-convert-data-types)
    - [7. Optional: Outlier Detection and Conditional Formatting](#7-optional-outlier-detection-and-conditional-formatting)


## Pivot Tables

- We will be using this dataset to learn and practice: https://www.kaggle.com/code/sadiqshah/bike-store-sales-in-europe/data
- **Pivot tables** let you quickly summarize and analyze large datasets in Excel by automatically grouping or calculating totals.
- They’re crucial in data analytics because they reduce manual work, help uncover insights, and allow for easy reorganization of data.
- We will create 2 pivot tables to:
    - Show a financial overview - revenue, cost, and profit - broken down by region.
        - **Rows:** Country/Region (with states/provinces under each country)
        - **Values:** Sum of Revenue, Sum of Cost, and Sum of Profit
    - Reveal how revenue changes over different years across regions.
        - **Rows:** Country/Region
        - **Columns:** Year
        - **Values:** Sum of Revenue
    
    ![image.png](https://github.com/ndomah1/Learning-Microsoft-Excel/blob/main/images/01_pivot_tables.png)
    

## Formulas

- Excel functions are essential for quickly processing and analyzing data.
- We will use this data set to practice the most important formulas in Excel for data analytics:
    
    ![image.png](https://github.com/ndomah1/Learning-Microsoft-Excel/blob/main/images/02_formulas.png)
    
- We can use our formula data set to demonstrate several key functions with real values:

### **MAX & MIN:**

- `=MAX(SalaryRange)` returns the highest salary ($65,000).
- `=MIN(SalaryRange)` returns the lowest salary ($36,000).

### **IF & IFS:**

- `=IF(Age > 35, "Senior", "Junior")` labels employees as Senior or Junior.
- `=IFS(Gender = "Male", "M", Gender = "Female", "F", TRUE, "Other")` handles labeling for multiple conditions.

### **LEN:**

- `LEN(FirstName)` measures the length of a first name, e.g., `=LEN("Michael")` = 7.

### **LEFT & RIGHT:**

- `=LEFT(LastName, 3)` extracts the first three letters, e.g., `=LEFT("Halpert", 3)` = “Hal”.
- `=RIGHT(LastName, 3)` extracts the last three letters, e.g., `=RIGHT(”Halpert”, 3)` = “ert”.

### **Converting Date to Text**

- `=TEXT(StartDate, "mm/dd/yyyy")` converts a date (e.g., 11/3/2001) into a text string “11/03/2001”.

### **TRIM**

- `=TRIM(FirstName)` removes any extra spaces, ensuring clean text data.

### **SUBSTITUTE**

- `=SUBSTITUTE(JobTitle, "Salesman", "Sales Rep")` replaces “Salesman” with “Sales Rep”.

### **SUM & SUMIF**

- `=SUM(SalaryRange)` calculates total salaries (e.g., $405,000).
- `=SUMIF(GenderRange, "Male", SalaryRange)` sums salaries for males only.

### **COUNT & COUNTIF**

- `=COUNT(SalaryRange)` counts the number of salary entries (e.g., 9).
- `=COUNTIF(GenderRange, "Female")` returns how many are female (e.g., 3).

### **CONCATENATE**

- `=CONCATENATE(FirstName, " ", LastName)` combines names into one string (e.g., “Jim Halpert”).

### **DAYS & NETWORKDAYS**

- `=DAYS(EndDate, StartDate)` calculates total days between two dates (e.g., 5078).
- `=NETWORKDAYS(StartDate, EndDate)` returns business days excluding weekends (e.g., 3620).

## `XLOOKUP` & `VLOOKUP`

- `XLOOKUP` and `VLOOKUP` are powerful lookup functions in Excel that help you retrieve information from large datasets.
- `XLOOKUP` searches for a lookup value in a specified range (or array) and returns a corresponding value from another range.
    - Unlike `VLOOKUP`, `XLOOKUP` can search from left to right or right to left, making it more flexible.
- `VLOOKUP` searches for a lookup value in the leftmost column of a table array and returns a value from a specified column in the same row.
    - It’s older than `XLOOKUP` and requires your lookup column to be on the left side of the data you want to return.
- We will use the following table for our examples:
    
    ![image.png](https://github.com/ndomah1/Learning-Microsoft-Excel/blob/main/images/03_lookups.png)
    

### `VLOOKUP`

- Get the *JobTitle* (8th column) for the EmployeeID in cell A2:
    - `=VLOOKUP(A2, $A:$L, 8, False)`

### `XLOOKUP` (Basic)

- Retrieve the *Salary* for “Jim Halpert” by matching FullName (column D) to Salary (column I):
    - `=XLOOKUP("Jim Halpert", $D:$D, $I:$I)`

### `XLOOKUP` (Multiple Rows)

- Pull *Salary* for several names in cells D2:D5:
    - `=XLOOKUP(D2:D5, $D:$D, $I:$I)`

### `XLOOKUP` (Exact Match)

- Find Salary by EmployeeID in A2 or show “Not Found”:
    - `=XLOOKUP(A2, $A:$A, $I:$I, "Not Found", 0)`

### `XLOOKUP` (Search Order)

- Search for the last match of “Halpert” in LastName (column C) and return Salary (column I):
    - `=XLOOKUP("Halpert", $C:$C, $I:$I, , 0, -1)`

### `XLOOKUP` (Horizontal)

- If row 2 holds data horizontally (A2:L2), find “JobTitle” in the header row (A1:L1):
    - `=XLOOKUP("JobTitle", $A$1:$L$1, A2:L2)`

### `XLOOKUP` with Sum

- Sum the salaries of “Jim Halpert” and “Pam Beasly”:
    
    ```
    =SUM(
    	XLOOKUP({"Jim Halpert", "Pam Beasly"}, $D:$D, $I:$I)
    )
    ```
    

## Conditional Formatting

- Conditional formatting helps data analysts quickly spot trends, outliers, or key thresholds in large datasets by applying color highlights, icons, or other visual cues.
- It streamlines decision-making because critical data points stand out immediately.
- They are used for:
    - **Fast Insights:** Identify high or low values at a glance.
    - **Trend Detection:** Highlight month-over-month changes or performance indicators.
    - **Outlier Analysis:** Pinpoint anomalies (e.g., exceptionally high or low sales, salaries).

### Example 1 - Monthly Stationary Sales

![image.png](https://github.com/ndomah1/Learning-Microsoft-Excel/blob/main/images/04_cf_before.png)

1. **Select Data Range:** Highlight all the monthly sales cells (January-December for Paper, Printer, and Manila Folder).
2. **Navigate to Conditional Formatting:** On the Home tab in Excel, click **Conditional Formatting → Color Scales.**
3. **Choose a Color Scale:** Select a preset (e.g., “Green-Yellow-Red”). This automatically assigns green to higher numbers, yellow to mid-range, and red to lower values.

**Result:** Large sales figures (like 750 for Paper in April) show in green, and smaller ones (like 40 for Printer in January) appear in red:

![image.png](https://github.com/ndomah1/Learning-Microsoft-Excel/blob/main/images/05_cf_after.png)

### Example 2 - Highlighting Specific Thresholds

![image.png](https://github.com/ndomah1/Learning-Microsoft-Excel/blob/main/images/06_cf2_before.png)

1. **Select the Salary Column:** Click and drag over all salary cells.
2. **Set High-Salary Rule:**
    1. Go to **Home → Conditional Formatting → Highlight Cell Rules → Greater Than…**
    2. Enter **60000** (or another threshold) and pick a green fill or custom format.
3. **Set Middle-Salary Rule:**
    1. Repeat step 2. a. but choose **Between**
    2. Enter **40000** as lower limit and **60000** as upper limit.
4. **Set Low-Salary Rule:**
    1. Repeat the above process but choose **Less Than…**
    2. Enter **40000** and pick a red fill.

**Result:** Salaries above $60,000 (like Dwight’s $63,000) appear in green, while lower salaries (like Pam’s $36,000) are flagged in red:

![image.png](https://github.com/ndomah1/Learning-Microsoft-Excel/blob/main/images/07_cf2_after.png)

## Charts

- Data analysts rely on charts to transform raw numbers into visuals, making it easier to spot trends, compare categories, and communicate insights.
- Instead of sifting through rows and columns of data, charts quickly highlight outliers, patterns, and relationships.
- By selecting the right chart type for each question - comparison, trend, or distribution, you can quickly uncover insights in your data.
- We will use the following data set for our examples:
    
    ![image.png](https://github.com/ndomah1/Learning-Microsoft-Excel/blob/main/images/08_charts.png)
    

### Example 1 - Monthly Breakdown

- We want to compare multiple items side-by-side across months.
- We will create a visual comparison of how each item’s sales stack up each month with a **Clustered Column** chart**:**
    1. **Highlight Data:** Select the item names (e.g., A2:A8) and their monthly columns (B2:M8).
    2. **Insert Chart:** Go to **Insert → Column →** choose **Clustered Column.**
    3. **Customize:**
        1. Use **Chart Title** to label the chart “Monthly Sales by Item”.
        2. Add a **Legend** if Excel doesn’t do so automatically.
    
    ![image.png](https://github.com/ndomah1/Learning-Microsoft-Excel/blob/main/images/09_monthly_breakdown.png)
    

### Example 2 - Trend Analysis

- We want to view overall sales trends over time.
- We will create a **Line c**hart that clearly shows sales trends, peaks, and dips throughout the year:
    1. **Highlight Data:** Select the months (B1:M1) and the “Total Items Per Month” row (B9:M9).
    2. **Insert Chart:** Go to **Insert → Line →** choose **Line with Markers** (or another style you prefer).
    3. **Customize:**
        1. Rename the data series to “Monthly Totals” under **Select Data → Series → Edit.**
        2. Add **Axis Titles** (e.g., “Months” on the x-axis, “Total Sales” on the y-axis).
        3. Add a **Trendline**
    
    ![image.png](https://github.com/ndomah1/Learning-Microsoft-Excel/blob/main/images/10_trend_analysis.png)
    

### Example 3 - Year-End Distribution

- We want to show how each item contributes to total annual sales, which can be done with a **Pie Chart:**
    1. **Highlight Data:** Select item names (A2:A8) and their “Year End Total” column (N2:N8).
    2. **Insert Chart:** Go to **Insert → Pie → 2-D Pie.**
    3. **Customize:**
        1. Click **Add Data Labels** to show each slice’s percentage or value.
        2. Right-click slices to **Format Data Series** (e.g., explode a slice, adjust colors).
    
    ![image.png](https://github.com/ndomah1/Learning-Microsoft-Excel/blob/main/images/11_year_end%20distribution.png)
    

## Cleaning Data in Excel

- Data cleaning ensures accuracy, consistency, and usability.
- Even a small dataset can have inconsistencies - like missing values or inconsistent formats - that can skew analysis.
- By cleaning your data first, you avoid errors and build a solid foundation for deeper analysis.
- We will use the following data set for our cleaning example:
    
    ![image.png](https://github.com/ndomah1/Learning-Microsoft-Excel/blob/main/images/12_cleaning_data.png)
    

### Common Data Cleaning Steps

### 1. Open and Inspect the Data

- **Open the File:** Open your Excel file (e.g., *Cleaned Data.xlsx*).
- **Initial Check:**
    - Look at the number of rows and columns.
    - Review the header row and a sample of data to spot obvious issues (inconsistent column names, extra spaces, missing values, etc.).

### 2. Standardize Column Names

- **Why:** Consistent headers (all lowercase, no extra spaces) make further analysis easier.
- **How:**
    - **Manually:** Click on each header cell and edit it to remove extra spaces, convert to lowercase, and replace spaces with underscores (e.g., change “Sales Amount” to “sales_amount”).
    - **Using Formulas (Optional):**
        1. In a helper row (or new sheet), use a formula such as:`=LOWER(TRIM(A1))`
        to standardize each header.
        2. Copy the results and paste them back as values over your original headers.

### 3. Remove Duplicate Rows

- **Why:** Duplicates can bias your analysis.
- **How:**
    1. Select your data range (or click any cell within the dataset).
    2. Go to the **Data** tab.
    3. Click on **Remove Duplicates**.
    4. In the dialog, select the columns you want to check (or all columns) and click **OK**.

### 4. Trim Extra Spaces in Data Cells

- **Why:** Extra spaces in text entries can lead to mismatches or errors in analysis.
- **How:**
    1. For a given text column (e.g., column B), insert a new helper column.
    2. In the helper column, use the formula:`=TRIM(B2)`
    and drag it down to cover all rows.
    3. Once completed, copy the helper column and paste it as values over the original column.
    4. Remove the helper column if desired.

### 5. Handle Missing Values

- **Identify Missing Data:**
    - Use **Filter** (click the filter icon in the header row) to quickly see blank cells in each column.
- **For Numeric Columns:**
    - **Option A (Imputation with Median):**
        1. In an empty cell, calculate the median using the formula:`=MEDIAN(C2:C100)` (adjust the range as needed).
        2. Manually fill blank cells with the median value or use Find & Replace after filtering blanks.
- **For Categorical/Text Columns:**
    - **Option A (Fill with Mode or 'Unknown'):**
        1. You can create a **PivotTable** to identify the most frequent value (mode).
        2. Replace blank cells with this value or simply type in “Unknown” where appropriate.

### 6. Convert Data Types

- **Why:** Ensuring that numbers, dates, and text are formatted correctly is crucial for analysis.
- **How:**
    1. Select a column (e.g., a date column).
    2. Right-click and choose **Format Cells**.
    3. Choose the appropriate format (e.g., Date, Number, or Text).
- **Tip:** For converting text dates into Excel dates, you might also use the **Text to Columns** feature (found in the **Data** tab).

### 7. Optional: Outlier Detection and Conditional Formatting

- **How:**
    1. Select a numeric column.
    2. Go to **Home > Conditional Formatting > Highlight Cells Rules**.
    3. Set rules to highlight cells that fall outside a chosen range. This helps in spotting potential data errors.
