# Stock_Analysis_VBA Script

Stock_Analysis_VBA Script performs stock analysis in the user-selected Excel workbook. The script iterates through the all exisitng worksheets that contain stock data within a workbook. Stock_Analysis_VBA calculates variables such as yearly change, percentage change and total volume for each of the stock teakcers, and creates additional column with this varabiles in the respective worksheet. Additionaly, it provides data with another summary tables with greatest percentage increase, greatest percentage decrease, and greatest total volume among the stock tickers.

## Restrictions

The script is position sensative. Before runing the script, user should insure that the stock data are within columns A to G, with the respective data in each column:
  1. Column A: "Stock Ticker"
  2. Column B: "Date"
  3. Column C: "Open Value of the Stock on Respective Date"
  4. Column D: "Highest Value of the Stock on Respective Date"
  5. Column E: "Lowest Value of the Stock on Respective Date"
  4. Column F: "Close Value of the Stock on Respective Date"
  6. Column G: "Stock's Volume on Respective Date"
Note: The names of the columns can be changed.

## Files

 - "stock_analysis_vba.bas": the script with the subrutines for the analysis.
 - Three .png files containing the screenshots of with the results of the analysis.

## Usage 

In order to run the script you can follow this process:
  1. Open Excel file that contains data for the aanalysis. Please use macros-enabled file format (.xlsm).
  2. Go to developer tab and open Visual Basic (aleternaively press Alt + F11).
  3. Import "stock_analysis_vba.bas" file to the editor.
  4. Open the module with the script and press run button on your toolbar (alternatively press F5).
  5. In a new pop-up window select "final analysis" and press run.
  6. Wait until the notification "Analysis Complete" will pop up.

## Description

The script contains four subrutines that preform following functions:
  - "stock(ws)": Analysis input data and creates a summary table containing each ticker with respective computations of its early change, percentage change, and total volume.
  - "format(ws)": Applies conditional formatting and appropriate data formats to the summary table.
  - "functionality(ws)": Creates additional table with greatest percent increase, greates decrease, and  greatest total volume among all the tickers.
  - "final_analysis()": Iterates through all the worksheets in the workbook and runs abovementioned subrutines. Displays notification regarding the end of analysis.
    
## Results 
The Stock_Analysis_VBA Script analyzes the stock data and organaizes results into two summary tables. The first summary table refelcts yearly change, percentage change and total volume of each stock. **Yearly change** variable shows the change of the price from the beginning to the end of the year. **Percentage change** further enhaces this understanding providing relative context as calculate relatively to the open price. **Total Stock Volume** agregates the volume of the each stock showing position/presence of respective stock on the market. This analysis is further enhanced by appling conditional formating to the data and appling proper formats to properly reflect the data. In the **Yearly change** column green cells indicate positive growth, red cells point to a decrease, and yellow cells reflect no change. 

In addition to the summary table Stock_Analysis_VBA Script will create another table with highlits of the most prominant variables and its changes. The table will show tickers with greatest percentage increase and decrease. It can provide user with the quick insights into the top and bottom performers. It will also identify the stock with the highest trading volume. 

When analysis complete the pop up message will notify user about its completeion. The final analysis is aimed to be clear and user-friendly which will provide quick and understandable results for further analysis and/or application.

