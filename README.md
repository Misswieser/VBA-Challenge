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
  * Greatest % Increase: Ticker - AIC, Percentage Increase - 4.268886044
  * Greatest % Decrease: Ticker - AEY, Percentage Decrease - -0.936406317
  * Greatest Total Volume: Ticker - ATS, Total Volume - 46673637787

Worksheet B:
  * Greatest % Increase: Ticker - BXK, Percentage Increase - 16.49637217
  * Greatest % Decrease: Ticker - AEY, Percentage Decrease - -0.936406317
  * Greatest Total Volume: Ticker - BDNG, Total Volume - 1.28132364932

Worksheet C:
  * Greatest % Increase: Ticker - CHKI, Percentage Increase - 21.2761417
  * Greatest % Decrease: Ticker - CUO, Percentage Decrease - -0.950917627
  * Greatest Total Volume: Ticker - CKB, Total Volume - 2.61608911298

Worksheet D:
  * Greatest % Increase: Ticker - CHKI, Percentage Increase - 21.2761417
  * Greatest % Decrease: Ticker - CUO, Percentage Decrease - -0.950917627
  * Greatest Total Volume: Ticker - DJF, Total Volume - 1.886657023522

Worksheet E:
  * Greatest % Increase: Ticker - CHKI, Percentage Increase - 21.2761417
  * Greatest % Decrease: Ticker - CUO, Percentage Decrease - -0.950917627
  * Greatest Total Volume: Ticker - DJF, Total Volume - 1.886657023522

Worksheet F:
  * Greatest % Increase: Ticker - CHKI, Percentage Increase - 21.2761417
  * Greatest % Decrease: Ticker - CUO, Percentage Decrease - -0.950917627
  * Greatest Total Volume: Ticker - DJF, Total Volume - 1.88666E+12

### Multiple year stock data

Year 2018
  * Greatest % Increase: Ticker - LVR, Percentage Increase - 57.26472995
  * Greatest % Decrease: Ticker - WSF, Percentage Decrease - -0.973498363
  * Greatest Total Volume: Ticker - QKI, Total Volume - 1.689179168176

Year 2019
  *Greatest % Increase: Ticker - LVR, Percentage Increase - 57.26472995
  *Greatest % Decrease: Ticker - LKJ, Percentage Decrease - -0.975200491
  *Greatest Total Volume: Ticker - ZQ, Total Volume - 4.372719580069

Year 2020
  * Greatest % Increase: Ticker - LVR, Percentage Increase - 59.1804419
  * Greatest % Decrease: Ticker - RKS, Percentage Decrease - -0.997184943
  * Greatest Total Volume: Ticker - ZQ, Total Volume - 4.372719580069

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
