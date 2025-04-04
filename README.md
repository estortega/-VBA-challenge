# VBA Stock Market Analysis

# Project Overview

This project demonstrates the use of VBA scripting to analyze stock market data. The script loops through each quarter’s data to calculate key metrics for each stock, including quarterly change, percentage change, and total volume. Additionally, the script identifies the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

By automating data analysis tasks with VBA, this project reduces the tediousness of manual calculations, providing accurate results with just a click of a button.

# Analysis and Features

Key Calculations
The VBA script performs the following calculations for each stock on every worksheet (quarter):

- Ticker Symbol: Identifies the unique stock ticker.

- Quarterly Change: Calculates the difference between the opening price at the beginning and the closing price at the end of the quarter.

- Percentage Change: Calculates the percentage change from the opening to the closing price.
 
- Total Stock Volume: Sums up the volume traded throughout the quarter.

- Advanced Features

Greatest Values:

Greatest Percentage Increase

Greatest Percentage Decrease

Greatest Total Volume

Conditional Formatting:

Positive changes are highlighted in green.

Negative changes are highlighted in red.

Multi-Sheet Support:

The script automatically processes each worksheet (representing a quarter) without manual intervention.

# How to Run the Project

Prerequisites
Microsoft Excel (Windows version with VBA support)
Running the VBA Script
Open Excel: Open the alphabetical_testing.xlsx file from the data/ folder.
Access VBA Editor: Press ALT + F11 to open the VBA editor.
Import the Script: Go to File -> Import File and select the stock_analysis.vba script from the scripts/ folder.
Run the Script:
Press F5 or click Run to execute the script.
The script will automatically process each worksheet and generate the analysis.
View Results: The results will be displayed on each sheet, with conditional formatting applied to indicate positive and negative changes.

# Example Output

The script produces the following key outputs:

Quarterly Analysis: Displays ticker symbol, quarterly change, percentage change, and total volume.
Summary Report: Identifies the stock with the greatest percentage increase, decrease, and total volume.
Conditional Formatting: Clearly highlights gains and losses with green and red formatting, respectively.

# Performance Optimization

Testing on Smaller Dataset: Use the alphabetical_testing.xlsx file for faster testing and debugging.
Efficient Looping: The script is optimized to minimize processing time, even when handling large datasets.

# Challenges and Solutions

Challenge: Multi-Sheet Processing
Solution: The script dynamically detects all worksheets and processes each one individually.

Challenge: Accurate Percentage Calculations
Solution: Implemented error handling for division by zero and ensured formatting consistency for all calculated values.

Challenge: Performance on Large Datasets
Solution: Optimized the loop and reduced unnecessary operations to improve script execution speed.

# Troubleshooting

Issue: Script Not Running
Ensure that macros are enabled in Excel (File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> Enable all macros).
Issue: Unexpected Formatting
Check the conditional formatting rules and ensure they are applied correctly for positive and negative changes.
