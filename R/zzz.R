.pkg_quotes <- c(
  "Statistics is the art of never having to say you are certain.",
  "Statistical thinking will one day be necessary for citizenship.",
  "The plural of anecdote is not data.",
  "All models are wrong, but some are useful.",
  "Without data, you are just another person with an opinion.",
  "Numbers do not lie, but liars use numbers.",
  "Data scientists are like modern day explorers.",
  "The world is one big data problem.",
  "Random is not random.",
  "Probability is expectation founded upon partial knowledge.",
  "I am not failing, I am collecting data points on what does not work.",
  "If at first you do not succeed, try two more times for significance.",
  "My failures are normally distributed - all over the place.",
  "Success is just statistically significant failure reduction.",
  "A statistician can have his head in an oven and his feet in ice and say he feels fine.",
  "Statisticians do not fail - they just achieve a high p-value.",
  "I have a joke about statistics, but it might not be significant.",
  "The best thing about being a statistician is that you get to play with numbers all day.",
  "Statistics: The only science that enables different experts using the same figures to draw different conclusions.",
  "If you torture the data long enough, it will confess to anything."
)
.onAttach <- function(libname, pkgname) {
  check_and_import_fonts <- function() {
    # Check if a common font is available to determine if fonts are loaded
    if (!"Arial" %in% extrafont::fonts()) {
      packageStartupMessage("Fonts not found. Importing fonts...")
      extrafont::font_import(prompt = FALSE)

      # Load fonts based on OS
      if (Sys.info()["sysname"] == "Windows") {
        extrafont::loadfonts(device = "win")
      } else {
        extrafont::loadfonts(device = "pdf")
      }
    }
  }

  # Run the check and import fonts if needed
  check_and_import_fonts()
  packageStartupMessage("Fonts loaded successfully for ", Sys.info()["sysname"])
  random_quote <- sample(.pkg_quotes, 1)
  msg <- paste0(
    "\n",
    "\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\n",
    "Quote of the day: ", random_quote, "\n",
    "\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\n"
  )

  packageStartupMessage(msg)
  packageStartupMessage("Thank you for installing 'visvaR'!")
  packageStartupMessage("For more information visit: https://rameshram96.github.io/visvaR/")
}

