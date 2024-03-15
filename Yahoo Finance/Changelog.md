## [1.2] - 2024-03-15

### Major changes

- Added support for multiple tickers
    - Each ticker now has its own output file
- Added `Historical Data -Last Year` and `Historical Data - All` sheets, displaying the historical data of each ticker since its creation

### Changes

- Updated the name of the file to `FDR - YahooFinance`
- Creation a `FDR - YahooFinance -test` file for internal testing
- Added the `Consensus EPS Estimate` as a replacement for the Analysts’ data
- The name of the `General Information` sheet has now been changed to `*TICKER* - Information`

### Fixed issues

- Fixed: The data from the `Historical Data` pages was seemingly false, especially the adjusted close
- Fixed: The date in the first page was displaying *today*’s date, and not the last date registered for the ticker
- Fixed: The Analysts’ data was not being correctly retrieved. This section has now been deleted.

## [1.1] - 2024-02-29

- Updated the name of the file to `Company Financial Information Retriever`
- Updated the entire code with new features

### Major changes

- Added a `General Information` sheet to the file output. This page displays relevant data about the company, such as:
    - Company name
    - Ticker
    - Currency
    - Exchange Market
    - Date
    - Previous close
    - And much more
- The data and columns in the `Income Statement`, `Balance Sheet`, and `Cash Flow`sheets has been reversed to better reflect the presentation on Yahoo Finance
- The output file will now display the name of the ticker of the company
- The output file is now placed in the user’s Downloads folder

### Fixed issues

- The output file would either not be downloaded or deleted after running the code

## [1.0] - 2024-02-28

Obtained the original code of the `Financial Data Retriever`
