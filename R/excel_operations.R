#' Read an Excel File from SharePoint
#'
#' @description
#' Downloads and reads an Excel file (.xlsx, .xls) from SharePoint into an R data frame.
#'
#' @param file_path Character string. The path to the Excel file on the SharePoint drive.
#' @param drive A SharePoint drive object, as returned by \code{\link{get_sp_drive}}.
#' @param sheet Either a string naming a sheet or an integer specifying the sheet's position.
#'              Default is 1 (the first sheet).
#' @param startRow Integer. The row to start reading from. Default is 1.
#' @param colNames Logical. Whether the first row contains column names. Default is TRUE.
#' @param ... Additional arguments passed to \code{\link[openxlsx]{read.xlsx}}.
#'
#' @return A data frame containing the Excel data.
#'
#' @details
#' This function downloads the Excel file to a temporary location and then reads it into R
#' using the \code{\link[openxlsx]{read.xlsx}} function. The temporary file is
#' automatically deleted after reading.
#'
#' The function supports both .xlsx and .xls file formats. For complex Excel files with
#' formatting, formulas, or multiple sheets, you may need to use additional parameters
#' via the ... argument.
#'
#' @examples
#' \dontrun{
#' # Connect to SharePoint and get drive
#' site <- connect_sharepoint("https://example.sharepoint.com/sites/mysite")
#' drive <- get_sp_drive(site, "Documents")
#'
#' # Read the first sheet of an Excel file
#' data <- read_sp_excel("reports/quarterly_data.xlsx", drive)
#'
#' # Read a specific named sheet
#' sales_data <- read_sp_excel("reports/quarterly_data.xlsx", drive, sheet = "Sales")
#'
#' # Read with specific options
#' data <- read_sp_excel("reports/quarterly_data.xlsx", drive,
#'                       sheet = 2, startRow = 3, colNames = FALSE)
#' }
#'
#' @export
read_sp_excel <- function(file_path, drive, sheet = 1, startRow = 1, colNames = TRUE, ...) {
  # Validate inputs (single character string)
  if (!is.character(file_path) || length(file_path) != 1) {
    stop("file_path must be a single character string")
  }

  # Make sure the drive object is in the users environment (run get_sp_drive)
  if (is.null(drive)) {
    stop("SharePoint drive object not found")
  }

  # Check if openxlsx is installed
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("Package 'openxlsx' needed for Excel files. Please install it with install.packages('openxlsx')")
  }

  # Create temporary file to store the download
  temp_path <- tempfile(fileext = ".xlsx")

  # Download the file
  message(paste0("Downloading Excel from SharePoint: ", file_path))
  tryCatch({
    drive$download_file(file_path, dest = temp_path)
  }, error = function(e) {
    stop(paste0("Error downloading file: ", e$message))
  })

  # Read the Excel file
  excel_data <- tryCatch({
    openxlsx::read.xlsx(temp_path, sheet = sheet, startRow = startRow, colNames = colNames, ...)
  }, error = function(e) {
    # Make sure to clean up the temp file even if reading fails
    unlink(temp_path)
    stop(paste0("Error reading Excel: ", e$message)) # print error message
  })

  # Clean up the temporary file
  unlink(temp_path)
  message("Excel file successfully loaded")

  # Return the data that was read into excel data
  return(excel_data)
}

#' Write a Data Frame to an Excel File on SharePoint
#'
#' @description
#' Writes an R data frame, or multiple data frames, to an Excel file on a SharePoint drive.
#'
#' @param data A data frame or a named list of data frames to write to Excel.
#' @param file_path Character string. The path where the file should be written on the SharePoint drive.
#' @param drive A SharePoint drive object, as returned by \code{\link{get_sp_drive}}.
#' @param sheet Character string. The name of the sheet to write to (when data is a single data frame).
#'              Default is "Sheet1".
#' @param overwrite Logical. If TRUE, overwrites the file if it already exists. Default is FALSE.
#' @param ... Additional arguments passed to \code{\link[openxlsx]{write.xlsx}} or \code{\link[openxlsx]{writeData}}.
#'
#' @return Returns console message indicating status of export.
#'
#' @details
#' This function writes a data frame or multiple data frames to an Excel file on SharePoint.
#' It first creates a temporary file and then uploads it to the specified location.
#' The temporary file is automatically deleted after uploading.
#'
#' Folders in the file_path must already exist. The function will not create new folders.
#' If you want to write to a nested folder structure, ensure all folders exist
#' on the SharePoint drive first.
#'
#' To write multiple data frames to different sheets, provide a named list of data frames
#' where the names will be used as sheet names.
#'
#' @examples
#' \dontrun{
#' # Connect to SharePoint and get drive
#' site <- connect_sharepoint("https://example.sharepoint.com/sites/mysite")
#' drive <- get_sp_drive(site, "Documents")
#'
#' # Write a single data frame to Excel
#' df <- data.frame(x = 1:10, y = letters[1:10])
#' write_sp_excel(df, "reports/data.xlsx", drive)
#'
#' # Write to a specific sheet name
#' write_sp_excel(df, "reports/data.xlsx", drive, sheet = "MyData", overwrite = TRUE)
#'
#' # Write multiple data frames to different sheets
#' sales <- data.frame(month = month.name, amount = runif(12, 1000, 5000))
#' expenses <- data.frame(month = month.name, amount = runif(12, 500, 3000))
#' profit <- data.frame(month = month.name, amount = sales$amount - expenses$amount)
#'
#' # Create a named list of data frames
#' report_data <- list(
#'   "Sales" = sales,
#'   "Expenses" = expenses,
#'   "Profit" = profit
#' )
#'
#' # Write all data frames to a single Excel file with multiple sheets
#' write_sp_excel(report_data, "reports/financial_report.xlsx", drive)
#' }
#'
#' @export
write_sp_excel <- function(data, file_path, drive, sheet = "Sheet1", overwrite = FALSE, ...) {
  # Validate inputs - ensure single string and character type
  if (!is.character(file_path) || length(file_path) != 1) {
    stop("file_path must be a single character string")
  }

  # Ensure the drive object exists in the users environment (run get_sp_drive)
  if (is.null(drive)) {
    stop("SharePoint drive object not found")
  }

  # Check if openxlsx is installed
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("Package 'openxlsx' needed for Excel files. Please install it with install.packages('openxlsx')")
  }

  # Check if file already exists
  file_exists <- tryCatch({
    drive$get_item(file_path)
    TRUE
  }, error = function(e) {
    FALSE
  })

  # If file exists and overwrite is FALSE, stop with an error
  if (file_exists && !overwrite) {
    stop(paste0("File already exists: ", file_path, ". Set overwrite=TRUE to replace it."))
  }

  # Create temporary file
  temp_path <- tempfile(fileext = ".xlsx")

  # Write data to temporary file
  tryCatch({
    # Check if the data is a list of data frames for multiple sheets
    if (is.list(data) && !is.data.frame(data)) {
      # Create a new workbook
      wb <- openxlsx::createWorkbook()

      # Add each data frame as a sheet
      for (sheet_name in names(data)) {
        openxlsx::addWorksheet(wb, sheet_name)
        openxlsx::writeData(wb, sheet_name, data[[sheet_name]], ...)
      }

      # Save the workbook - AND overwrite since we are writing to the temp file
      openxlsx::saveWorkbook(wb, temp_path, overwrite = TRUE)
    } else {
      # Single data frame case
      openxlsx::write.xlsx(data, temp_path, sheetName = sheet, ...)
    }
  }, error = function(e) {
    unlink(temp_path)  # If error - still clean up the temp file
    stop(paste0("Error writing to Excel: ", e$message)) # print error message
  })

  # Split path to get folder and file name - similar to previous methodologies used in this script
  if (grepl("/", file_path)) {
    path_parts <- strsplit(file_path, "/")[[1]]
    file_name <- path_parts[length(path_parts)]
    folder_path <- paste(path_parts[-length(path_parts)], collapse = "/")
  } else {
    # File is meant to be written to root directory in the drive
    file_name <- file_path
    folder_path <- ""
  }

  # Check if destination folder exists
  folder_exists <- TRUE
  if (folder_path != "") {
    folder_exists <- tryCatch({
      drive$get_item(folder_path)
      TRUE
    }, error = function(e) {
      FALSE
    })
  }

  # If above evaluates as false then trow an error, (have to manually create the folder structure)
  if (!folder_exists) {
    unlink(temp_path)  # Clean up
    stop(paste0("Destination folder doesn't exist: ", folder_path,
                ". Please create the folder structure first."))
  }

  # Upload the file
  # We already made sure the file doesn't exist or if it does - overwrite is TRUE at this point
  message(paste0("Writing Excel to SharePoint: ", file_path))
  upload_result <- tryCatch({
    if (folder_path == "") {
      # Upload to root
      drive$upload(src = temp_path, dest = file_name) # drive$upload is from Microsoft365R package
    } else {
      # Upload to specific provided folder
      folder <- drive$get_item(folder_path)
      folder$upload(src = temp_path, dest = file_name)
    }
    TRUE
  }, error = function(e) {
    message(paste0("Error uploading file: ", e$message))
    FALSE
  })

  # Clean up the temporary file
  unlink(temp_path)

  # Give the user some confirmation
  if (upload_result) {
    if (file_exists && overwrite) {
      message("Existing Excel file was overwritten successfully")
    } else {
      message("Excel file was written successfully")
    }
    return(invisible(TRUE))
  } else {
    stop("Failed to write Excel file to SharePoint")
  }
}
