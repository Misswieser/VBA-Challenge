# Stock Market Data Analysis with VBA

This project involves analyzing stock market data using VBA scripting in Excel. The VBA script loops through multiple worksheets, calculates various metrics, and identifies stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

## Dataset

The analysis is performed on the following:

* Alphabetical testing 
Worksheets: A, B, C, D, E, and F.

* Multiple year stock data
Worksheets: 2018, 2019, and 2020.

For each worksheet contains stock data for a specific year, with columns for the ticker symbol, volume of stock, open price, and close price.

## Analysis Results

The VBA script provides the following analysis results:

### Alphabetical Testing

Worksheet A:
  * Greatest % Increase: Ticker - AAR, Percentage Increase - 22.95%
  * Greatest % Decrease: Ticker - AYO, Percentage Decrease - -52.67%
  * Greatest Total Volume: Ticker - ATS, Total Volume - 46673637787

Worksheet B:
  * Greatest % Increase: Ticker - BLC, Percentage Increase - 42.41%
  * Greatest % Decrease: Ticker - AYO, Percentage Decrease - -52.67%
  * Greatest Total Volume: Ticker - BDNG, Total Volume - 128132364932

Worksheet C:
  * Greatest % Increase: Ticker - BLC, Percentage Increase - 42.41%
  * Greatest % Decrease: Ticker - AYO, Percentage Decrease - -52.67%
  * Greatest Total Volume: Ticker - CKB, Total Volume - 261608911298

Worksheet D:
  * Greatest % Increase: Ticker - DLUW, Percentage Increase - 69.58%
  * Greatest % Decrease: Ticker - AYO, Percentage Decrease - -52.67%
  * Greatest Total Volume: Ticker - DJF, Total Volume - 1886657023522

Worksheet E:
  * Greatest % Increase: Ticker - DLUW, Percentage Increase - 69.58%
  * Greatest % Decrease: Ticker - EEA, Percentage Decrease - -63.17%
  * Greatest Total Volume: Ticker - DJF, Total Volume - 1886657023522

Worksheet F:
  * Greatest % Increase: Ticker - FGH, Percentage Increase - 77.77%
  * Greatest % Decrease: Ticker - EEA, Percentage Decrease - -63.17%
  * Greatest Total Volume: Ticker - DJF, Total Volume - 1886657023522

### Multiple year stock data

Year 2018
  * Greatest % Increase: Ticker - THB, Percentage Increase - 141.42%
  * Greatest % Decrease: Ticker - RKS, Percentage Decrease - -90.02%
  * Greatest Total Volume: Ticker - QKI, Total Volume - 1689179168176

Year 2019
  * Greatest % Increase: Ticker - RYU, Percentage Increase - 190.03%
  * Greatest % Decrease: Ticker - RKS, Percentage Decrease - -91.60%
  * Greatest Total Volume: Ticker - ZQ, Total Volume - 4372719580069

Year 2020
  * Greatest % Increase: Ticker - RYU, Percentage Increase - 190.03%
  * Greatest % Decrease: Ticker - RKS, Percentage Decrease - -91.60%
  * Greatest Total Volume: Ticker - ZQ, Total Volume - 4372719580069

## Usage

1. Open the Excel file provided, "alphabetical_testing.xlsx".
2. Run the VBA script by following these steps:
  * Press Alt + F11 to open the VBA editor in Excel.
  * In the VBA editor, locate the module containing the VBA code for stock data analysis.
  * Run the macro named "AnalyzeStockData" in each worksheet to perform the analysis.
3. The analyzed results will be displayed in the respective worksheets, along with conditional formatting highlighting positive and negative changes.


# Requirements

## Retrieval of Data 
* The script loops through one year of stock data and reads/ stores all of the following values from each row:
  * ticker symbol 
  * volume of stock 
  * open price 
  * close price 

## Column Creation 
* On the same worksheet as the raw data, or on a new worksheet all columns were correctly created for:
  * ticker symbol
  * total stock volume 
  * yearly change ($) 
  * percent change 

## Conditional Formatting 
* Conditional formatting is applied correctly and appropriately to the yearly change column 
* Conditional formatting is applied correctly and appropriately to the percent change column

## Calculated Values 
* All three of the following values are calculated correctly and displayed in the output:
  * Greatest % Increase 
  * Greatest % Decrease 
  * Greatest Total Volume
 
## Looping Across Worksheet 
* The VBA script can run on all sheets successfully.

## GitHub/GitLab Submission 
* All three of the following are uploaded to GitHub/GitLab:
  * Screenshots of the results 
  * Separate VBA script files 
  * README file 
