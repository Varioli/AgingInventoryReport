# Tire Aging Report Generator

Created by Josh Clark, this GUI-based Python application helps tire warehouse and inventory managers identify aging stock using FIFO principles to reduce waste and maximize supplier credits.

## What It Does

This tool cross-analyzes two CSV reports:  
- **Inventory Report**: Current stock, quantity, and description  
- **Sales Report**: Past sales data with completion dates  

The application:  
- Matches SKUs by cleaned descriptions  
- Calculates how many days it has been since each item was last sold  
- Classifies items as:  
  - ACTIVE  
  - WARNING – 60+ DAYS  
  - EXTREME – 90+ DAYS  
  - NO SALES HISTORY  
- Outputs a structured Excel report for review

## Features

- GUI interface built with Tkinter  
- Custom color scheme for ease of use  
- Upload prompts for inventory and sales files  
- Automatically generates a formatted `.xlsx` aging report  
- Calculates days idle and assigns status based on last recorded sale

## File Requirements

- Inventory CSV must contain: `SkuCode`, `Description`, `QuantityOnHand`  
- Sales CSV must contain: `Description`, `DateCompleted`  
- Both files must skip the first 3 rows (headers start at row 4)

## Why Use This

- Helps identify slow-moving or dead inventory  
- Supports FIFO enforcement and supplier credit tracking  
- Reduces manual report building time through automation  
- Enhances warehouse visibility and operational decision-making

## How to Run

1. Download or clone the repository from GitHub:  
   [https://github.com/Varioli/AgingInventoryReport](https://github.com/Varioli/AgingInventoryReport)

2. Ensure both your Inventory and Sales reports are saved as `.csv` files.

3. Run the program using:

    pip install -r requirements.txt
    python tire_aging_reporter.py
    
4. In the application:
   - Upload the Inventory file first
   - Upload the Sales file second
   - Click "Run Report" and choose where to save the generated Excel file

## Output

The final Excel report includes:
- SkuCode  
- Description  
- QuantityOnHand  
- LastSoldDate  
- DaysIdle  
- Status

## Use Cases

- Inventory audits  
- Weekly warehouse aging reviews  
- Supplier credit claim support  
- Identifying reorder candidates or dead stock