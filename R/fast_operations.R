#' Fast Read of a CSV File from SharePoint
#'
#' @description
#' Quickly reads a CSV file from SharePoint into an R data.table/data.frame using
#' data.table's fread function, which is optimized for large files.
#'
#' @param file_path Character string. The path to the CSV file on the SharePoint drive.
#' @param drive A SharePoint drive object, as returned by \code{\link{get_sp_drive}}.
#' @param ... Additional arguments passed to \code{\link[data.table]{fread}}.
#'
#' @return A data.table (which is also a data.frame) containing the CSV data.
#'
#' @details
#' This function downloads the CSV file to a temporary location and then reads it using
#' \code{\link[data.table]{fread}}, which is much faster than standard read.csv, especially
#' for large files. The temporary file is automatically deleted after reading.
#'
#' The function leverages data.table's fast and flexible CSV parsing capabilities, including:
#' - Automatic detection of column types
#' - Faster processing of large files
#' - More robust handling of various CSV formats
#'
#' @examples
#' \dontrun{
#' # Connect to SharePoint and get drive
#' site <- connect_sharepoint("https://example.sharepoint.com/sites/mysite")
#' drive <- get_sp_drive(site, "Documents")
#'
#' # Fast read of a large CSV file
#' large_data <- fread_sp("large_data/a_gazillion_rows.csv", drive)
#'
#' # With additional fread parameters
#' data <- fread_sp("data.csv", drive, sep = ";", dec = ",", select = c(1, 3, 5))
#'
#' # Read only specific columns by name
#' data <- fread_sp("data.csv", drive, select = c("Date", "Value", "Category"))
#' }
#'
#' @seealso \code{\link{read_sp_csv}} for standard CSV reading, \code{\link{fwrite_sp}} for fast CSV writing.
#'
#' @export
fread_sp <- function(file_path, drive, ...) {
  # Validate inputs - ensure single string and character type
  if (!is.character(file_path) || length(file_path) != 1) {
    stop("file_path must be a single character string")
  }

  # Make sure the drive object is in the users environment (run get_sp_drive)
  if (is.null(drive)) {
    stop("SharePoint drive object not found")
  }

  # Need to make sure the data table package is installed and loaded
  if (!requireNamespace("data.table", quietly = TRUE)) {
    stop("Package 'data.table' needed for fread. Please install it with install.packages('data.table')")
  }

  # Create temporary file to store the download
  temp_path <- tempfile(fileext = ".csv")

  # Download the file to temp path
  message(paste0("Downloading CSV from SharePoint: ", file_path))
  tryCatch({
    drive$download_file(file_path, dest = temp_path) # download_file is from Microsoft365R package
  }, error = function(e) { # e is the error object that is created from the tryCatch function
    stop(paste0("Error downloading file: ", e$message)) # If there is an error, function ends and error message from e is printed
  })

  # Read the CSV file with fread
  csv_data <- tryCatch({
    data.table::fread(temp_path, ...) # fread the data into csv_data
  }, error = function(e) {
    # Make sure to clean up the temp file even if reading fails
    unlink(temp_path)
    stop(paste0("Error reading CSV with fread: ", e$message))
  })

  # Clean up the temporary file
  unlink(temp_path)
  message("CSV file successfully loaded with fread")

  # Return the data that was f-read into csv_data
  return(csv_data)
}

#' Fast Write of a Data Frame to a CSV File on SharePoint
#'
#' @description
#' Quickly writes an R data frame to a CSV file on a SharePoint drive using
#' data.table's fwrite function, which is optimized for large files.
#'
#' @param data A data frame or data.table to write to CSV.
#' @param file_path Character string. The path where the file should be written on the SharePoint drive.
#' @param drive A SharePoint drive object, as returned by \code{\link{get_sp_drive}}.
#' @param overwrite Logical. If TRUE, overwrites the file if it already exists. Default is FALSE.
#' @param ... Additional arguments passed to \code{\link[data.table]{fwrite}}.
#'
#' @return Returns console message indicating status of export.
#'
#' @details
#' This function writes the data to a temporary CSV file using data.table's fast fwrite function
#' and then uploads it to SharePoint. The temporary file is automatically deleted after uploading.
#'
#' The function leverages data.table's fast CSV writing capabilities, which can be significantly
#' faster than standard write.csv, especially for large datasets.
#'
#' Folders in the file_path must already exist. The function will not create new folders.
#' If you want to write to a nested folder structure, ensure all folders exist
#' on the SharePoint drive first.
#'
#' @examples
#' \dontrun{
#' # Connect to SharePoint and get drive
#' site <- connect_sharepoint("https://example.sharepoint.com/sites/mysite")
#' drive <- get_sp_drive(site, "Documents")
#'
#' # Create a large data frame
#' large_df <- data.frame(
#'   id = 1:1000000,
#'   value = rnorm(1000000),
#'   group = sample(letters, 1000000, replace = TRUE)
#' )
#'
#' # Fast write to SharePoint
#' fwrite_sp(large_df, "large_data/a_gazillion_rows.csv", drive)
#'
#' # With additional fwrite parameters
#' fwrite_sp(large_df, "data.csv", drive, sep = ";", dec = ",", quote = FALSE)
#' }
#'
#' @seealso \code{\link{write_sp_csv}} for standard CSV writing, \code{\link{fread_sp}} for fast CSV reading.
#'
#' @export
fwrite_sp <- function(data, file_path, drive, overwrite = FALSE, ...) {
  # Validate inputs - ensure single string and character type
  if (!is.character(file_path) || length(file_path) != 1) {
    stop("file_path must be a single character string")
  }

  # Ensure the drive object exists in the users environment (run get_sp_drive)
  if (is.null(drive)) {
    stop("SharePoint drive object not found")
  }

  # Need to make sure the data table package is installed and loaded
  if (!requireNamespace("data.table", quietly = TRUE)) {
    stop("Package 'data.table' needed for fwrite. Please install it with install.packages('data.table')")
  }

  # For overwriting purposes, check if the file already exists
  file_exists <- tryCatch({
    drive$get_item(file_path)
    TRUE
  }, error = function(e) {
    FALSE
  })

  # If above evaluates as true and overwrite (provided by user) is false, stop with an error
  if (file_exists && !overwrite) {
    stop(paste0("File already exists: ", file_path, ". Set overwrite=TRUE to replace it."))
  }

  # Create temporary file
  temp_path <- tempfile(fileext = ".csv")

  # Write data to temporary file using fwrite
  tryCatch({
    data.table::fwrite(data, temp_path, ...) # fwrite to temp file
  }, error = function(e) {
    unlink(temp_path)  # if the function errors, remove the temp file
    stop(paste0("Error writing to CSV with fwrite: ", e$message)) # and return an error message
  })

  # Differentiate between file name and folder path
  if (grepl("/", file_path)) {
    path_parts <- strsplit(file_path, "/")[[1]] # returns a list with one element [[1]] grabs the character vector (chr [1:3] "folder1" "folder2" "file.csv")
    file_name <- path_parts[length(path_parts)] # from the vector, grab the last part (filename)
    folder_path <- paste(path_parts[-length(path_parts)], collapse = "/") # from the vector, grab everything but the last part (folder path)
  } else {
    # File is meant to be written to root path
    file_name <- file_path
    folder_path <- ""
  }

  # Check if destination folder exists (return TRUE or FALSE)
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
  message(paste0("Writing CSV to SharePoint with fwrite: ", file_path))
  upload_result <- tryCatch({
    if (folder_path == "") {
      # Upload to root if no folder path
      drive$upload(src = temp_path, dest = file_name)
    } else {
      # Upload to specific folder if the path exists
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
      message("Existing file was overwritten successfully")
    } else {
      message("CSV file was written successfully")
    }
    return(invisible(TRUE))
  } else {
    stop("Failed to write CSV file to SharePoint")
  }
}
