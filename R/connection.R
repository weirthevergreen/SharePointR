#' Connect to SharePoint Site
#'
#' @description
#' Establishes a connection to a Microsoft SharePoint site using the Microsoft365R package.
#' Authentication is handled automatically, using your Microsoft 365 credentials.
#'
#' @param site_url Character string. The URL of the SharePoint site to connect to.
#' @param tenant Character string. Tenant name of the desired SharePoint site.
#' @param app Character string. The client ID of the application to use for authentication.
#'
#' @return A SharePoint site object that can be used with other SharePointR functions.
#'
#' @details
#' For the initial verification, you will be redirected to your default browser
#' to log in, or to request access from the sites administrator
#'
#' If you encounter HTTP 401 (unauthorized) errors, you need to ensure you are able
#' to access the SharePoint site in your default browser.
#'
#' @examples
#' \dontrun{
#' # Connect to a SharePoint site
#' site <- connect_sharepoint("https://example.sharepoint.com/sites/mysite")
#'
#' # Connect with explicit tenant
#' site <- connect_sharepoint("https://example.sharepoint.com/sites/mysite", "example")
#' }
#'
#' @export
connect_sharepoint <- function(site_url, tenant, app, ...) {
  message("Connecting to SharePoint site...")

  if (is.null(site_url) || length(site_url) == 0 || site_url == "") {
    stop("Error: site_url is required and cannot be empty")
  }
  if (is.null(tenant) || length(tenant) == 0 || tenant == "") {
    stop("Error: tenant is required and cannot be empty")
  }
  if (is.null(app) || length(app) == 0 || app == "") {
    stop("Error: app is required and cannot be empty")
  }

  # Check if user tried to pass scopes
  dots <- list(...)
  if ("scopes" %in% names(dots)) {
    stop("Error: Custom scopes are not allowed for security compliance.")
  }

  # Build safe argument list
  args <- list(
    site_url = site_url,
    tenant = tenant,
    app = app,
    scopes = "Sites.ReadWrite.All"
  )

  # Add other safe parameters (exclude blocked ones)
  safe_params <- dots[!names(dots) %in% c("scopes", "site_url", "tenant", "app")]
  args <- c(args, safe_params)

  site <- do.call(Microsoft365R::get_sharepoint_site, args)

  # List granted scopes for this app
  if (!is.null(site$token$credentials$scope)) {
    granted_scopes <- strsplit(site$token$credentials$scope, " ")[[1]]
    message("Granted scopes: ", paste(granted_scopes, collapse = ", "))
  }
  return(site)
}

#' Get a Drive from a SharePoint Site
#'
#' @description
#' Retrieves a specific document library (drive) from a connected SharePoint site.
#'
#' @param site A SharePoint site object, as returned by \code{\link{connect_sharepoint}}.
#' @param drive_name Character string. The name of the document library/drive to access.
#'
#' @return A SharePoint drive object that can be used with read/write functions.
#'
#' @examples
#' \dontrun{
#' # Connect to a SharePoint site
#' site <- connect_sharepoint("https://example.sharepoint.com/sites/mysite")
#'
#' # Get the "Documents" drive
#' drive <- get_sp_drive(site, "Documents")
#'
#' # Get a different drive called "Shared Files"
#' drive <- get_sp_drive(site, "Shared Files")
#' }
#'
#' @export
get_sp_drive <- function(site, drive_name) {
  message(paste0("Accessing drive: ", drive_name))
  drive <- site$get_drive(drive_name)
  return(drive)
}
