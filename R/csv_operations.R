#' Read in a CSV file from a SharePoint site
#'
#' @description
#' Using an established connection to a SharePoint drive object, the function reads
#' a .csv data file into the users environment.
#'
#' @param file_path Character string. The file path of the .csv file on the SharePoint site.)
#' @param drive A SharePoint drive object, as returned by \code{\link{get_sp_drive}}.
#' @param ... Additional arguments passed to \code{\link[utils]{read.csv}}.
#'
#' @return A data frame object loaded into the users R environment.
#'
#' @details
#' This function downloads the CSV file to a temporary location and then reads it into R
#' using the standard \code{\link[utils]{read.csv}} function. The temporary file is
#' automatically deleted after reading.
#'
#' Ensure that your drive object is active and loaded in your environment.
#'
#' @examples
#' \dontrun{
#' # Connect to SharePoint and get drive
#' site <- connect_sharepoint("https://example.sharepoint.com/sites/mysite")
#' drive <- get_sp_drive(site, "Documents")
#'
#' # Read CSV file
#' df <- read_sp_csv("folder_in_drive_object/data/data.csv", drive)
#'
#' # With additional read.csv parameters
#' df2 <- read_sp_csv("data.csv", drive, sep = ";", header = FALSE)
#' }
#'
#' @export
read_sp_csv <- function(file_path, drive, ...) { # ... allows extra arguments to be passed that are standard for read.csv
  # Make sure the file path is a single string 'example_folder/file.csv'
  if (!is.character(file_path) || length(file_path) != 1) {
    stop("file_path must be a single character string")
  }

  # Make sure the drive object is in the users environment and passed to the function (run get_sp_drive)
  if (is.null(drive)) {
    stop("SharePoint drive object not found")
  }

  # Create temporary file to store the download - will be cleaned up later
  temp_path <- tempfile(fileext = ".csv")

  # Download the file
  message(paste0("Downloading CSV from SharePoint: ", file_path))
  tryCatch({
    drive$download_file(file_path, dest = temp_path) # download_file is from Microsoft365R package
  }, error = function(e) { # e is the error object that is created from the tryCatch function
    stop(paste0("Error downloading file: ", e$message)) # If there is an error, function ends and error message from e is printed
  })

  # Read the CSV file
  csv_data <- tryCatch({
    utils::read.csv(temp_path, stringsAsFactors = FALSE, ...) # attempt to read into temp file - if successful, assigned to csv_data
  }, error = function(e) { # If there is an error:
    unlink(temp_path) # Delete the temp file
    stop(paste0("Error reading CSV: ", e$message)) # ... and print the error message
  })

  # Result if successful - data is in 'csv_data' and temp file is removed
  unlink(temp_path)
  message("CSV file successfully loaded")

  # Return the data
  return(csv_data)
}

#' Write a Data Frame to a CSV File on SharePoint
#'
#' @description
#' Writes an R data frame to a CSV file on a SharePoint drive.
#'
#' @param data A data frame to write to CSV.
#' @param file_path Character string. The path where the file should be written within the SharePoint drive.
#' @param drive A SharePoint drive object, as returned by \code{\link{get_sp_drive}}.
#' @param overwrite Logical. If TRUE, overwrites the file if it already exists. Default is FALSE.
#' @param ... Additional arguments passed to \code{\link[utils]{write.csv}}.
#'
#' @return Returns console message indicating status of export.
#'
#' @details
#' This function first writes the data to a temporary file and then uploads it to SharePoint.
#' The temporary file is automatically deleted after uploading.
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
#' # Create a simple data frame
#' df <- data.frame(x = 1:10, y = letters[1:10])
#'
#' # Write to a file in the root of the drive
#' write_sp_csv(df, "mydata.csv", drive)
#'
#' # Write to a subfolder, overwriting if exists
#' write_sp_csv(df, "reports/monthly/data.csv", drive, overwrite = TRUE)
#'
#' # Use additional arguments for write.csv
#' write_sp_csv(df, "data.csv", drive, row.names = TRUE, na = "NA")
#' }
#'
#' @export
write_sp_csv <- function(data, file_path, drive, overwrite = FALSE, ...) {
  # Make sure the file path is a single string 'example_folder/file.csv'
  if (!is.character(file_path) || length(file_path) != 1) {
    stop("file_path must be a single character string")
  }

  # Ensure the drive object exists in the users environment (run get_sp_drive)
  if (is.null(drive)) {
    stop("SharePoint drive object not found")
  }

  # Check if file already exists
  file_exists <- tryCatch({
    drive$get_item(file_path) # get_item is from Microsoft365R package
    TRUE
  }, error = function(e) { # if true, file exists and return TRUE, else: FALSE
    FALSE
  })

  # If file exists and overwrite is FALSE, stop with an error
  if (file_exists && !overwrite) {
    stop(paste0("File already exists: ", file_path, ". Set overwrite=TRUE to replace it."))
  }

  # Create temporary file
  temp_path <- tempfile(fileext = ".csv")

  # Write data to temporary file
  tryCatch({
    utils::write.csv(data, temp_path, row.names = FALSE, ...) # Attempt to write
  }, error = function(e) {
    unlink(temp_path)  # Clean up
    stop(paste0("Error writing to CSV: ", e$message)) # If errors, then stop and print error message
  })

  # Need to determine what is path vs. file name

  # Split by "/" - if / exists then it is a nested path
  if (grepl("/", file_path)) {
    path_parts <- strsplit(file_path, "/")[[1]] # returns a list with one element [[1]] grabs the character vector (chr [1:3] "folder1" "folder2" "file.csv")
    file_name <- path_parts[length(path_parts)] # from the vector, grab the last part (filename)
    folder_path <- paste(path_parts[-length(path_parts)], collapse = "/") # from the vector, grab everything but the last part (folder path)
  } else {
    # No nested path:
    file_name <- file_path
    folder_path <- ""
  }

  # Check if destination folder exists (return true or false)
  folder_exists <- TRUE
  if (folder_path != "") {
    folder_exists <- tryCatch({
      drive$get_item(folder_path)
      TRUE
    }, error = function(e) {
      FALSE
    })
  }

  # If above returns FALSE, then throw an error (assuming folders will be created manually on sharepoint for now)
  if (!folder_exists) {
    unlink(temp_path)  # Clean up
    stop(paste0("Destination folder doesn't exist: ", folder_path,
                ". Please create the folder structure first."))
  }

  # Write the file
  message(paste0("Writing CSV to SharePoint: ", file_path))
  upload_result <- tryCatch({
    if (folder_path == "") {
      # Upload to the root path (drive)
      drive$upload(src = temp_path, dest = file_name) # drive$upload is from Microsoft365R package
    } else {
      # Upload to specific folder
      folder <- drive$get_item(folder_path)
      folder$upload(src = temp_path, dest = file_name)
    }
    TRUE
  }, error = function(e) {
    message(paste0("Error uploading file: ", e$message)) # print error if happens
    FALSE
  })

  # Clean up the temporary file
  unlink(temp_path)

  # Some messages to communicate what happened
  if (upload_result) {
    if (file_exists && overwrite) {
      message("Existing file was overwritten successfully")
    } else {
      message("CSV file was written successfully")
    }
    return(invisible(TRUE)) # dont want this true to print in the console
  } else {
    stop("Failed to write CSV file to SharePoint")
  }
}
