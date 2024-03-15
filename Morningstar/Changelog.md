## [0.2] - 2024-03-15

- The code can now download the financial information of multiple tickers without a problem
- The ETA of the code has been slightly increased to allow for slower connections to fully run the code

### Major changes

- Added the possibility to download information from multiple tickers

### Changes

- Modified the `main` and the `download_morningstar_excel` functions to allow for the download of information from multiple tickers
- The `time.pause()` has been increased from `3` to `5`, to allow downloads for slower connections
- Moreover, this increased length gives time for the code to rename, move, and merge successfully each downloaded file
- The code now deletes the folders and files it created successfully

### Fixed issues

- Fixed: The code was creating unnecessary files on the Desktop.

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
- The `main` function has been entirely rewritten

### Fixed issues

- The `print` functions output is mainly irrelevant, as it displays text whatever happened before

## [alpha.0.0.0] - 2024-03-03

- Created the `MS - Financial Information Retriever` file
- *This version is in alpha stage. Many issues are still being encountered, and most of the code is, at this stage, not working. This is a work in progress.*
