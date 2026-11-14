/*
BlinkIT Grocery Sales Dashboard - SQL Server script
Source: BlinkIT Grocery Data.xlsx (sheet: BlinkIT Grocery Data)
Rows: 8,523

SETUP
1. In SSMS, create/select your database.
2. Import the Excel sheet using Tasks > Import Data (or the Import/Export Wizard).
3. Name the imported table: dbo.BlinkIT_Grocery_Data
4. Run this script after the import.

The source has 1,463 NULL Item Weight values. They are intentionally retained as NULL.
Fat-content labels are standardized in the dashboard view.
*/

-- 1) Optional: create the raw table before importing, if you prefer to map columns manually.
-- If you use the Import Wizard, skip this CREATE TABLE if it already created the table.
IF OBJECT_ID('dbo.BlinkIT_Grocery_Data', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.BlinkIT_Grocery_Data
    (
        [Item Fat Content]          NVARCHAR(30)  NULL,
        [Item Identifier]           NVARCHAR(20)  NOT NULL,
        [Item Type]                 NVARCHAR(100) NULL,
        [Outlet Establishment Year] INT           NULL,
        [Outlet Identifier]         NVARCHAR(20)  NOT NULL,
        [Outlet Location Type]      NVARCHAR(30)  NULL,
        [Outlet Size]               NVARCHAR(30)  NULL,
        [Outlet Type]               NVARCHAR(50)  NULL,
        [Item Visibility]           DECIMAL(12,9) NULL,
        [Item Weight]               DECIMAL(10,3) NULL,
        [Sales]                     DECIMAL(12,4) NULL,
        [Rating]                    DECIMAL(4,2)  NULL
    );
END;
GO

-- 2) Cleaned view used by the Power BI dashboard.
CREATE OR ALTER VIEW dbo.vw_BlinkIT_Grocery_Clean AS
SELECT
    UPPER(LTRIM(RTRIM([Item Identifier]))) AS Item_Identifier,
    CASE
        WHEN LOWER(LTRIM(RTRIM([Item Fat Content]))) IN ('lf', 'low fat') THEN 'Low Fat'
        WHEN LOWER(LTRIM(RTRIM([Item Fat Content]))) IN ('reg', 'regular') THEN 'Regular'
        ELSE LTRIM(RTRIM([Item Fat Content]))
    END AS Item_Fat_Content,
    LTRIM(RTRIM([Item Type])) AS Item_Type,
    [Outlet Establishment Year] AS Outlet_Establishment_Year,
    UPPER(LTRIM(RTRIM([Outlet Identifier]))) AS Outlet_Identifier,
    LTRIM(RTRIM([Outlet Location Type])) AS Outlet_Location_Type,
    LTRIM(RTRIM([Outlet Size])) AS Outlet_Size,
    LTRIM(RTRIM([Outlet Type])) AS Outlet_Type,
    CAST([Item Visibility] AS DECIMAL(12,6)) AS Item_Visibility,
    CAST([Item Weight] AS DECIMAL(10,3)) AS Item_Weight,
    CAST([Sales] AS DECIMAL(12,2)) AS Sales,
    CAST([Rating] AS DECIMAL(4,2)) AS Rating
FROM dbo.BlinkIT_Grocery_Data;
GO

-- 3) KPI cards
SELECT
    SUM(Sales) AS Total_Sales,
    AVG(Sales) AS Average_Sales_Per_Item,
    COUNT(*) AS Number_of_Items,
    AVG(Rating) AS Average_Rating
FROM dbo.vw_BlinkIT_Grocery_Clean;
GO

-- 4) Sales by outlet establishment year (line chart)
SELECT
    Outlet_Establishment_Year,
    SUM(Sales) AS Total_Sales
FROM dbo.vw_BlinkIT_Grocery_Clean
GROUP BY Outlet_Establishment_Year
ORDER BY Outlet_Establishment_Year;
GO

-- 5) Sales by item fat content (donut chart)
SELECT
    Item_Fat_Content,
    SUM(Sales) AS Total_Sales
FROM dbo.vw_BlinkIT_Grocery_Clean
GROUP BY Item_Fat_Content
ORDER BY Item_Fat_Content;
GO

-- 6) Sales by item type (bar chart)
SELECT
    Item_Type,
    SUM(Sales) AS Total_Sales
FROM dbo.vw_BlinkIT_Grocery_Clean
GROUP BY Item_Type
ORDER BY Total_Sales DESC;
GO

-- 7) Sales by outlet size (donut chart)
SELECT
    Outlet_Size,
    SUM(Sales) AS Total_Sales
FROM dbo.vw_BlinkIT_Grocery_Clean
GROUP BY Outlet_Size
ORDER BY Total_Sales DESC;
GO

-- 8) Sales by outlet location tier (bar chart)
SELECT
    Outlet_Location_Type,
    SUM(Sales) AS Total_Sales
FROM dbo.vw_BlinkIT_Grocery_Clean
GROUP BY Outlet_Location_Type
ORDER BY Outlet_Location_Type;
GO

-- 9) Fat content sales by outlet location tier (clustered/stacked bar chart)
SELECT
    Outlet_Location_Type,
    Item_Fat_Content,
    SUM(Sales) AS Total_Sales
FROM dbo.vw_BlinkIT_Grocery_Clean
GROUP BY Outlet_Location_Type, Item_Fat_Content
ORDER BY Outlet_Location_Type, Item_Fat_Content;
GO

-- 10) Outlet type summary table
SELECT
    Outlet_Type,
    SUM(Sales) AS Total_Sales,
    COUNT(*) AS Number_of_Items,
    AVG(Sales) AS Average_Sales,
    AVG(Rating) AS Average_Rating,
    AVG(Item_Visibility) AS Average_Item_Visibility
FROM dbo.vw_BlinkIT_Grocery_Clean
GROUP BY Outlet_Type
ORDER BY Total_Sales DESC;
GO

-- 11) Optional slicer/helper summaries
SELECT DISTINCT Outlet_Location_Type
FROM dbo.vw_BlinkIT_Grocery_Clean
ORDER BY Outlet_Location_Type;

SELECT DISTINCT Outlet_Size
FROM dbo.vw_BlinkIT_Grocery_Clean
ORDER BY Outlet_Size;

SELECT DISTINCT Item_Type
FROM dbo.vw_BlinkIT_Grocery_Clean
ORDER BY Item_Type;
GO

/*
Power BI connection:
- Connect to SQL Server and select the database.
- Load dbo.vw_BlinkIT_Grocery_Clean.
- Use the KPI query's measures or recreate them as DAX measures.
- Use the result sets above for the matching visuals.
*/
