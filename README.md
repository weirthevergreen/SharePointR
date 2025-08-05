# SharePointR

SharePointR provides credentialed access to Microsoft SharePoint from R. It allows you to read and write various file types directly to/from SharePoint, eliminating the need to sync drives locally or use user-specific file paths in shared scripts.

## Installation

You can install SharePointR from GitHub with:

``` r
# Install from GitHub
remotes::install_github("weirthevergreen/SharePointR")

# Or using devtools
devtools::install_github("weirthevergreen/SharePointR")
```

## Important: Authentication Requirements

**Browser Authentication Required:** Authentication from R to SharePoint requires that you can successfully log into and access the target SharePoint site through your default web browser.

**Troubleshooting:** If you encounter HTTP 401 (Unauthorized) errors during connection, ensure that you can open and navigate the desired SharePoint site in your default browser before attempting to connect via SharePointR. This browser-based authentication prerequisite is necessary for the underlying Microsoft 365 authentication flow.

## Key Features

-   **Simple Authentication:** Connect to SharePoint sites with Microsoft 365 credentials
-   **CSV Operations:** Read and write CSV files to/from SharePoint with automatic data type detection
-   **Excel Support:** Full Excel file support including multiple sheets and workbook operations
-   **Fast Processing:** Optimized reading/writing with data.table's fread/fwrite for large files
-   **R Data Files:** Save and load R data objects (.RData, .rda) directly to/from SharePoint
-   **Session Management:** Automatic token management and reuse during R sessions

## Quick Start

``` r
library(SharePointR)

# Step 1: Connect to your SharePoint site
site <- connect_sharepoint("https://yourcompany.sharepoint.com/sites/yoursite")

# Step 2: Get a drive object (usually "Documents")
drive <- get_sp_drive(site, "Documents")

# Step 3: Read a CSV file from SharePoint
data <- read_sp_csv("path/to/file.csv", drive)

# Step 4: Write data back to SharePoint
write_sp_csv(data, "path/to/newfile.csv", drive)
```

## Main Functions

### Connection Functions

-   **connect_sharepoint():** Authenticate and connect to a SharePoint site
-   **get_sp_drive():** Retrieve a SharePoint drive object for file operations

### CSV Operations

-   **read_sp_csv():** Download and read CSV files with automatic type detection
-   **write_sp_csv():** Write data frames to CSV files on SharePoint
-   **fread_sp():** Fast CSV reading for large files using data.table
-   **fwrite_sp():** Fast CSV writing for large files using data.table

### Excel Operations

-   **read_sp_excel():** Read Excel files with sheet selection and range options
-   **write_sp_excel():** Write data to Excel files with formatting options

### R Data Operations

-   **sp_load():** Load R data objects (.RData, .rda) from SharePoint
-   **sp_save():** Save R objects directly to SharePoint as .RData files

## Detailed Examples

### Working with CSV Files

``` r
library(SharePointR)

# Connect to SharePoint
site <- connect_sharepoint("https://yourcompany.sharepoint.com/sites/yoursite")
drive <- get_sp_drive(site, "Documents")

# Read a CSV file
sales_data <- read_sp_csv("reports/sales_2024.csv", drive)

# Modify the data
sales_data$profit <- sales_data$revenue - sales_data$costs

# Write back to SharePoint
write_sp_csv(sales_data, "reports/sales_2024_updated.csv", drive)
```

### Working with Excel Files

``` r
# Read specific sheet from Excel file
budget <- read_sp_excel("finance/budget.xlsx", drive, sheet = "Q4_Budget")

# Write to Excel with multiple sheets
financial_data <- list(
  revenue = revenue_df,
  expenses = expenses_df,
  summary = summary_df
)
write_sp_excel(financial_data, "finance/financial_report.xlsx", drive)
```

### Fast Read/Write Operations

``` r
# For large CSV files, use the fast data.table versions
large_dataset <- fread_sp("data/large_file.csv", drive)

# Process the data
processed_data <- large_dataset[revenue > 1000, .(total_revenue = sum(revenue)), by = region]

# Write back efficiently
fwrite_sp(processed_data, "results/processed_large_data.csv", drive)
```

### Saving R Objects

``` r
# Save multiple R objects to SharePoint
model_results <- list(
  model = trained_model,
  predictions = predictions,
  metrics = performance_metrics
)

sp_save(model_results, "models/latest_model.RData", drive)

# Load them back later
loaded_results <- sp_load("models/latest_model.RData", drive)
```

## Requirements

-   R (\>= 3.6.0)
-   Microsoft365R package
-   Valid Microsoft 365 credentials with SharePoint access
-   Internet connection
-   Open browser access to target SharePoint sites

## Getting Help

For detailed documentation on any function:

``` r
# Package overview
?SharePointR

# Specific functions
?connect_sharepoint
?read_sp_csv
?write_sp_excel
?sp_save
```

## License

MIT License - see LICENSE file for details.

## Author

Alex Weirth, Evergreen Economics

------------------------------------------------------------------------

*SharePointR connects directly to SharePoint from your R session, eliminating local file syncing and user-specific paths making scripts collaborative across teams.*
