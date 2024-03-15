## [0.1] - 2024-03-12

- The code is now perfectly functioning, and allows users to successfully export/download financial data from MorningStar based on the selected ticker.
- The code uses selenium for web-scraping, and automates the job of selecting the needed documents (Income Statement, Balance Sheet, and Cash Flow Statement), and downloads them in separated excel sheets using the Export Data located on the MS website

### Major changes

- The code now retrieves automatically the financial documents and downloads them
- Added the `download_morningstar_excel` function
    - Based on the selected ticker, downloads the IS, BS, and CF from MorningStar automatically
- The output in the terminal now correctly displays when the data has been successfully downloaded
- The downloaded documents are now being automatically converted to XLSX and merged together in one single file
- Once the files have been merged, the downloaded documents are removed from the user’s Downloads location
- The code now supports multiple tickers
- Each ticker will have, as an output, one file merging the `Income Statement`, `Balance Sheet`, and `Cash Flow` from MorningStar

### Changes

- New `download_morningstar_excel` function
- the `main` function has been entirely rewritten

### Fixed issues

- The `print` functions output is mainly irrelevant, as it displays text whatever happened before

## [alpha.0.0.0] - 2024-03-03

- Created the `MS - Financial Information Retriever` file
- *This version is in alpha stage. Many issues are still being encountered, and most of the code is, at this stage, not working. This is a work in progress.*