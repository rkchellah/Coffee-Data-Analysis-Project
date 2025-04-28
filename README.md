# Coffee-Data-Analysis-Project

## Summary
- Goal
  
Inspired by a [YouTube]( https://youtu.be/m13o5aqeCbM?si=zzyVH2i8y0BedkLL) tutorial, I used Power Pivot and DAX functions to recreate an Excel project with a unique approach. Instead of relying solely on XLOOKUP and INDEX MATCH, I    utilized Power Pivot for data integration across tables while retaining XLOOKUP for precise Sales column calculations. This approach enhanced my understanding of data analysi for business decision-making.

- Processs

A critical step in this project is understanding the data and accurately calculating coffee beans quantities. Individual pivot tables are first created and then combined into a dashboard, enabling the presentation of a compelling story through Excel.

- Insight

The insights for this project may include:  

1. **Sales Performance**:  
   - Identifying the best-selling coffee types and roast types.  
   - Determining the total sales revenue and highest-performing products.  

2. **Customer Behavior**:  
   - Analyzing purchase patterns by customer demographics or regions.  
   - Highlighting loyal customers based on order frequency.  

3. **Product Analysis**:  
   - Comparing sales by coffee size (e.g., kilograms).  
   - Evaluating the impact of pricing on product demand.  

4. **Operational Efficiency**:  
   - Identifying potential errors or inconsistencies in data (e.g., duplicate Customer IDs).  
   - Streamlining data integration and reporting processes using Power Pivot and DAX.  

5. **Dashboard Insights**:  
   - Summarizing key performance indicators (KPIs) for quick decision-making.  
   - Providing actionable insights through visualizations to inform business strategies.  

The project not only enhances data-driven decision-making but also demonstrates how tools like Excel, Power Pivot, and DAX can simplify complex data workflows.

Here's my step-by-step guide to recreating this project

### Data Preparation
- I used XLOOKUP to retrieve unit price values. Then, I created a new column named “Sales.” I calculated the total sales for each product by multiplying the Unit Price and Quantity. I applied the IF function to correct coffee type and roast type names. It was done only in the Product table for clarity. Mo adjusted them in the Orders table.

### Data Formatting
- I formatted the dataset. I started with the Size column and used a custom number format to display values in kilograms (Kg). I also formatted the Sales column to show values in currency.
   
### Checking for Duplicates
- I checked for duplicate Customer IDs in the Orders table before proceeding.
  
### Converting to Tables
- I converted the ranges from each sheet into tables. I named them appropriately—Orders, Customers, and Products. This is necessary since Power Pivot requires tables. To add Power Pivot to the Excel ribbon, I enabled it as a COM-Add-in through the Excel Options menu.
  
![powerpivot](https://github.com/user-attachments/assets/3e134948-cdcc-4b47-a3d4-ba1a07c9a748)


### Adding Data to Power Pivot
- I clicked on Power Pivot and began adding tables to the data model. I started with the Orders table by selecting any cell within it and choosing “Add to Data Model.” I repeated this process for the other tables, adding each one to the Power Pivot data model. Once all three tables were added, I switched to the Diagram View to establish relationships between them. I connected the Customer ID from the Orders table to the Customer ID in the Customers table. This connection created a one-to-many relationship. Similarly, I connected the Product ID from the Orders table to its counterpart in the Products table. This link formed another one-to-many relationship.
  
![relationship](https://github.com/user-attachments/assets/2cff8e0a-994a-49bc-bd65-c04e7cb6f6bb)


### Data Integration Using DAX
- I exited the Diagram View and began connecting the tables using the RELATED DAX function. I added the necessary columns from the Customer table. I also added the necessary columns from the Product table. This replicated the columns Mo Chen gathered using XLOOKUP and INDEX MATCH.
  
![related](https://github.com/user-attachments/assets/81690353-8c3a-44c9-af8e-59075cf4dd6c)


In the screenshot below, I initially forgot to include the Size column but later added it, along with the Loyalty Card information. These were necessary for creating slicers in the final dashboard

![table](https://github.com/user-attachments/assets/d22debec-4328-4d7c-96d0-6dbea6170ebe)


### Data Analysis
- For my data analysis, I used Pivot Tables with Power Pivot. After clicking on Pivot Tables, a new sheet was generated. On this sheet, I saw three tables in the Pivot Table Fields: Orders, Customers, and Products. I focused on the Orders table since it contained all the relevant information I needed. I followed a similar process to Mo. I utilized Pivot Tables and Pivot Charts. I created a dashboard in the same style as Mo’s, but with a different theme.
  

![Coffee Dashboard](https://github.com/user-attachments/assets/fde50a44-2e92-4830-b44a-0550b553784f)

  
### Conclusion
- This project demonstrated the effective use of Pivot Tables and Power Pivot to analyze and visualize data. By focusing on the Orders table, I was able to extract the key information needed for comprehensive analysis. Leveraging Pivot Charts, I was able to present the data clearly and intuitively, culminating in a well-structured dashboard. I customized the theme to create a unique visual aesthetic. This enhanced both presentation and usability of the final dashboard. This project highlights the versatility and power of Excel tools for data-driven decision-making. I’m open to any tips or suggestions if there’s an easier way to achieve the same results.
