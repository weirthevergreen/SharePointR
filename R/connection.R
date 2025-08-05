#' Connect to SharePoint Site
#'
#' @description
#' Establishes a connection to a Microsoft SharePoint site using the Microsoft365R package.
#' Authentication is handled automatically, using your Microsoft 365 credentials.
#'
#' @param site_url Character string. The URL of the SharePoint site to connect to.
#' @param tenant Character string. Tenant name of the desired SharePoint site.
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
connect_sharepoint <- function(site_url, tenant = NULL) {
  message("Connecting to SharePoint site...")
  # Provide flexibility if there is no tenant name
  if (is.null(tenant)) {
    site <- Microsoft365R::get_sharepoint_site(site_url = site_url)
  } else {
    site <- Microsoft365R::get_sharepoint_site(site_url = site_url, tenant = tenant)
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
