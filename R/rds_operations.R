#' Load RDS Files from SharePoint
#'
#' @description
#' Loads an RDS data file (.rds) from SharePoint into the global environment.
#'
#' @param file_path Character string. The path to the R data file on the SharePoint drive.
#' @param drive A SharePoint drive object, as returned by \code{\link{get_sp_drive}}.
#'
#' @return The R objects in the data file are loaded into the global environment.
#'
#' @details
#' This function downloads the RDS file to a temporary location and then loads it into
#' the global environment using the standard \code{\link[base]{load}} function.
#' The temporary file is automatically deleted after loading.
#'
#'
#' @examples
#' \dontrun{
#' # Connect to SharePoint and get drive
#' site <- connect_sharepoint("https://example.sharepoint.com/sites/mysite")
#' drive <- get_sp_drive(site, "Documents")
#'
#' # Load R data file
#' sp_read_rds("models/regression_model.rds", drive)
#'
#' # After running this, the objects saved in regression_model.RData
#' # will be available in your global environment
#' }
#'
#' @export
sp_read_rds <- function(file_path, drive) {
# Make sure the file path is a single string 'example_folder/file.rds'
if (!is.character(file_path) || length(file_path) != 1) {
  stop("file_path must be a single character string")
}

# Make sure the drive object is in the user's environment and passed to the function
if (is.null(drive)) {
  stop("SharePoint drive object not found")
}

# Determine file extension - make sure user input was an rds file
file_ext <- tolower(tools::file_ext(file_path))
if (file_ext != "rds") {
  stop("File must have .rds extension.")
}

# Create temporary file to store the download
temp_path <- tempfile(fileext = ".rds")

# Download the file
message(paste0("Downloading RDS file from SharePoint: ", file_path))
tryCatch({
  drive$download_file(file_path, dest = temp_path)
}, error = function(e) {
  stop(paste0("Error downloading file: ", e$message))
})

# Read the RDS file
rds_object <- tryCatch({
  readRDS(temp_path)
}, error = function(e) {
  unlink(temp_path) # Delete the temp file
  stop(paste0("Error reading RDS file: ", e$message))
})

# Clean up the temporary file
unlink(temp_path)

message("RDS file successfully loaded")
return(rds_object)
}


#' Save an R Object to an RDS File on SharePoint
#'
#' @description
#' Saves a single R object to an .rds file and uploads it to SharePoint.
#'
#' @param object The R object to save (not quoted).
#' @param file_path Character string. The path where the file should be written on the SharePoint drive.
#' @param drive A SharePoint drive object, as returned by \code{\link{get_sp_drive}}.
#' @param overwrite Logical. If TRUE, overwrites the file if it already exists. Default is FALSE.
#'
#' @return Returns console message indicating status of export.
#'
#' @details
#' This function saves a single R object to a temporary .rds file using \code{saveRDS()}
#' and then uploads it to SharePoint. The temporary file is automatically deleted after uploading.
#'
#' Folders in the file_path must already exist. The function will not create new folders.
#' If you want to save to a nested folder structure, ensure all folders exist
#' on the SharePoint drive first.
#'
#' Unlike \code{sp_save()} which can save multiple objects, this function saves
#' only one object at a time but preserves the exact object structure.
#'
#' @examples
#' \dontrun{
#' # Connect to SharePoint and get drive
#' site <- connect_sharepoint("https://example.sharepoint.com/sites/mysite")
#' drive <- get_sp_drive(site, "Documents")
#'
#' # Create a model to save
#' model <- lm(mpg ~ wt + hp, data = mtcars)
#'
#' # Save the model to SharePoint
#' sp_write_rds(model, "models/regression_model.rds", drive)
#'
#' # Overwrite an existing file
#' sp_write_rds(model, "models/regression_model.rds", drive, overwrite = TRUE)
#'
#' # Save a data frame
#' sp_write_rds(mtcars, "data/mtcars_dataset.rds", drive)
#' }
#'
#' @export
sp_write_rds <- function(object, file_path, drive, overwrite = FALSE) {
  # Make sure the file path is a single string 'example_folder/file.rds'
  if (!is.character(file_path) || length(file_path) != 1) {
    stop("file_path must be a single character string")
  }

  # Ensure the drive object exists
  if (is.null(drive)) {
    stop("SharePoint drive object not found")
  }

  # Determine file extension - make sure user specified an rds file
  file_ext <- tolower(tools::file_ext(file_path))
  if (file_ext != "rds") {
    stop("File must have .rds extension.")
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
  temp_path <- tempfile(fileext = ".rds")

  # Save object to temporary file
  save_result <- tryCatch({
    saveRDS(object, temp_path)
    TRUE
  }, error = function(e) {
    unlink(temp_path)  # Clean up
    stop(paste0("Error saving RDS file: ", e$message))
  })

  # Determine path vs. file name
  if (grepl("/", file_path)) {
    path_parts <- strsplit(file_path, "/")[[1]]
    file_name <- path_parts[length(path_parts)]
    folder_path <- paste(path_parts[-length(path_parts)], collapse = "/")
  } else {
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

  # If folder doesn't exist, throw an error
  if (!folder_exists) {
    unlink(temp_path)  # Clean up
    stop(paste0("Destination folder doesn't exist: ", folder_path,
                ". Please create the folder structure first."))
  }

  # Upload the file
  message(paste0("Writing RDS file to SharePoint: ", file_path))
  upload_result <- tryCatch({
    if (folder_path == "") {
      # Upload to the root path
      drive$upload(src = temp_path, dest = file_name)
    } else {
      # Upload to specific folder
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

  # Return status messages
  if (upload_result) {
    if (file_exists && overwrite) {
      message("Existing file was overwritten successfully")
    } else {
      message("RDS file was saved successfully")
    }
    return(invisible(TRUE))
  } else {
    stop("Failed to save RDS file to SharePoint")
  }
}
