#' function to make sure that extrafont package is installed properly and fonts are imported and loaded to be used in the pacakge
#' @details To use custom fonts, please install the `extrafont` package and run `extrafont::font_import()` and `extrafont::loadfonts()`.
#'
#' @title Function to make sure necessary fonts are imported from system
#' @description
#' function to make sure that extrafont package is installed properly and fonts are imported and loaded to be used in the pacakge
#' @name check_extrafont
check_extrafont <- function() {
  if (!requireNamespace("extrafont", quietly = TRUE)) {
    stop("The 'extrafont' package is required. Please install it with install.packages('extrafont').")
  }
  if (length(extrafont::fonts()) == 0) {
    stop("Fonts are not imported yet. Please run 'extrafont::font_import()' and 'extrafont::loadfonts()' (requires administrator privileges).")
  }

  # Optionally, load the fonts for the current session
  extrafont::loadfonts(device = "pdf") # or "win" for Windows, etc.
}
