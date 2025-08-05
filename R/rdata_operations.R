#' Load R Data Files from SharePoint
#'
#' @description
#' Loads an R data file (.RData or .rda) from SharePoint into the global environment.
#'
#' @param file_path Character string. The path to the R data file on the SharePoint drive.
#' @param drive A SharePoint drive object, as returned by \code{\link{get_sp_drive}}.
#'
#' @return The R objects in the data file are loaded into the global environment.
#'
#' @details
#' This function downloads the R data file to a temporary location and then loads it into
#' the global environment using the standard \code{\link[base]{load}} function.
#' The temporary file is automatically deleted after loading.
#'
#' The function supports both .RData and .rda file extensions, which are standard
#' formats for saving R objects.
#'
#' @examples
#' \dontrun{
#' # Connect to SharePoint and get drive
#' site <- connect_sharepoint("https://example.sharepoint.com/sites/mysite")
#' drive <- get_sp_drive(site, "Documents")
#'
#' # Load R data file
#' sp_load("models/regression_model.RData", drive)
#'
#' # After running this, the objects saved in regression_model.RData
#' # will be available in your global environment
#' }
#'
#' @export
sp_load <- function(file_path, drive) {
  # Make sure the file path is a single string 'example_folder/file.RData'
  if (!is.character(file_path) || length(file_path) != 1) {
    stop("file_path must be a single character string")
  }

  # Make sure the drive object is in the user's environment and passed to the function
  if (is.null(drive)) {
    stop("SharePoint drive object not found")
  }

  # Determine file extension - make sure user input was an r data file
  file_ext <- tolower(tools::file_ext(file_path))
  if (!(file_ext %in% c("rdata", "rda"))) {
    stop("File must have .RData or .rda extension.")
  }

  # Create temporary file to store the download
  temp_path <- tempfile(fileext = paste0(".", file_ext))

  # Download the file
  message(paste0("Downloading R data from SharePoint: ", file_path))
  tryCatch({
    drive$download_file(file_path, dest = temp_path)
  }, error = function(e) {
    stop(paste0("Error downloading file: ", e$message))
  })

  # Load the R data file into the global environment
  load_result <- tryCatch({
    # load() loads objects into the environment where it's called
    # envir = .GlobalEnv ensures objects are loaded into the global environment
    load(temp_path, envir = .GlobalEnv)
    TRUE
  }, error = function(e) {
    unlink(temp_path) # Delete the temp file
    stop(paste0("Error loading R data: ", e$message))
  })

  # Clean up the temporary file
  unlink(temp_path)

  if (load_result) {
    message("R data file successfully loaded into global environment")
    return(invisible(TRUE))
  }
}

#' Save R Objects to an R Data File on SharePoint
#'
#' @description
#' Saves specified R objects to an .RData file and uploads it to SharePoint.
#'
#' @param objects Character vector. Names of objects in the global environment to save.
#' @param file_path Character string. The path where the file should be written on the SharePoint drive.
#' @param drive A SharePoint drive object, as returned by \code{\link{get_sp_drive}}.
#' @param overwrite Logical. If TRUE, overwrites the file if it already exists. Default is FALSE.
#'
#' @return Returns console message indicating status of export.
#'
#' @details
#' This function saves specified R objects to a temporary .RData file and then uploads it to
#' SharePoint. The temporary file is automatically deleted after uploading.
#'
#' Folders in the file_path must already exist. The function will not create new folders.
#' If you want to save to a nested folder structure, ensure all folders exist
#' on the SharePoint drive first.
#'
#' Only objects that exist in the global environment can be saved. If any object
#' in the list doesn't exist, an error will be returned.
#'
#' @examples
#' \dontrun{
#' # Connect to SharePoint and get drive
#' site <- connect_sharepoint("https://example.sharepoint.com/sites/mysite")
#' drive <- get_sp_drive(site, "Documents")
#'
#' # Create some objects to save
#' model <- lm(mpg ~ wt + hp, data = mtcars)
#' summary_stats <- summary(model)
#'
#' # Save the objects to SharePoint
#' sp_save(c("model", "summary_stats"), "models/regression_results.RData", drive)
#'
#' # Overwrite an existing file
#' sp_save(c("model", "summary_stats"), "models/regression_results.RData", drive, overwrite = TRUE)
#' }
#'
#' @export
sp_save <- function(objects, file_path, drive, overwrite = FALSE) {
  # Make sure the file path is a single string 'example_folder/file.RData'
  if (!is.character(file_path) || length(file_path) != 1) {
    stop("file_path must be a single character string")
  }

  # Ensure the drive object exists
  if (is.null(drive)) {
    stop("SharePoint drive object not found")
  }

  # Ensure objects parameter is a character vector with object names
  if (!is.character(objects)) {
    stop("objects must be a character vector of object names")
  }

  # Check if all objects exist in the environment
  missing_objects <- objects[!sapply(objects, exists)]
  if (length(missing_objects) > 0) {
    stop(paste0("The following objects don't exist in your environment: ",
                paste(missing_objects, collapse = ", ")))
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
  temp_path <- tempfile(fileext = ".RData")

  # Save objects to temporary file
  save_result <- tryCatch({
    # Get the objects from their names
    object_list <- mget(objects, envir = .GlobalEnv)
    # Save with the original names
    save(list = objects, file = temp_path, envir = .GlobalEnv)
    TRUE
  }, error = function(e) {
    unlink(temp_path)  # Clean up
    stop(paste0("Error saving R data: ", e$message))
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
  message(paste0("Writing R data to SharePoint: ", file_path))
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
      message("R data file was saved successfully")
    }
    return(invisible(TRUE))
  } else {
    stop("Failed to save R data file to SharePoint")
  }
}
