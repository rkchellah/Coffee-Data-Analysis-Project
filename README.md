# Coffee-Data-Analysis-Project

- My inspiration for this project came from a YouTube tutorial I watched on [Mo Chen's]( https://youtu.be/m13o5aqeCbM?si=zzyVH2i8y0BedkLL) channel. He showcased a fascinating Excel project that caught my attention. The project began with data gathering, involving three sheets: the Order table, the Customer table, and the Product table. The Customer and Product tables from the Coffee Dataset that held data that needed to be incorporated into the Order table. To achieve this, Mo used the XLOOKUP and INDEX MATCH formulas to pull data from the Customer and Product tables. 
He then utilized the IF function to create two new columns, correcting the values for coffee type and roast type names. This step replaced the abbreviated values in the coffee type and roast type columns with more readable names. Additionally, Mo created a new "Sales" column by multiplying the Unit Price and Quantity columns, formatting the sales values into curreny, and changing the number format of the "Size" column to display values in kilograms (Kg). Before analyzing the dataset using pivot tables, he ensured there were no duplicate values in the Customer ID column. The final result was a well-designed dashboard.
Inspired by Mo’s approach, I decided to recreate the project with a different method. My goal was to gather data from other tables without relying on XLOOKUP and INDEX MATCH formulas, opting instead for Power Pivot and DAX functions. Although Power Pivot could handle most of the tasks, I still used XLOOKUP to ensure accurate Sales column values.

Here's my step-by-step guide to recreating this project

### Data Preparation
- I used XLOOKUP to retrieve unit price values, created a new column named "Sales," and calculated the total sales for each product by multiplying the Unit Price and Quantity. I also applied the IF function to correct coffee type and roast type names, but only in the Product table for clarity, unlike Mo, who adjusted them in the Orders table.

### Data Formatting
- I formatted the dataset, starting with the Size column, using a custom number format to display values in kilograms (Kg). I also formatted the Sales column to show values in currency.
   
### Checking for Duplicates
- I checked for duplicate Customer IDs in the Orders table before proceeding.
  
### Converting to Tables
- I converted the ranges from each sheet into tables, naming them appropriately—Orders, Customers, and Products—since Power Pivot requires tables. To add Power Pivot to the Excel ribbon, I enabled it as a COM-Add-in through the Excel Options menu.
  
### Adding Data to Power Pivot
- I clicked on Power Pivot and began adding tables to the data model. I started with the Orders table by selecting any cell within it and choosing "Add to Data Model." I repeated this process for the other tables, adding each one to the Power Pivot data model. Once all three tables were added, I switched to the Diagram View to establish relationships between them. I connected the Customer ID from the Orders table to the Customer ID in the Customers table, creating a one-to-many relationship. Similarly, I linked the Product ID from the Orders table to the Product ID in the Products table, forming another one-to-many relationship.

### Data Integration Using DAX
- I exited the Diagram View and began connecting the tables using the RELATED DAX function. I added the necessary columns from the Customer and Product tables, replicating the columns Mo Chen gathered using XLOOKUP and INDEX MATCH.

### Data Analysis
- For my data analysis, I used Pivot Tables with Power Pivot. After clicking on Pivot Tables, a new sheet was generated where I saw three tables in the Pivot Table Fields: Orders, Customers, and Products. I focused on the Orders table since it contained all the relevant information I needed. Following a similar process to Mo, I utilized Pivot Tables and Pivot Charts, ultimately creating a dashboard in the same style as Mo’s, but with a different theme.
  
### Conclusion
- This project demonstrated the effective use of Pivot Tables and Power Pivot to analyze and visualize data. By focusing on the Orders table, I was able to extract the key information needed for comprehensive analysis. Leveraging Pivot Charts, I was able to present the data clearly and intuitively, culminating in a well-structured dashboard. While the overall style mirrored Mo's approach, I customized the theme to reflect a unique visual aesthetic, enhancing the presentation and usability of the final dashboard. This project highlights the versatility and power of Excel tools for data-driven decision-making. I’m open to any tips or suggestions if there's an easier way to achieve the same results.
