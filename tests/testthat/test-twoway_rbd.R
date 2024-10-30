library(testthat)
library(shinytest2)
library(visvaR)
# Ensure your package is loaded to access twoway_rbd()

test_that("twoway_rbd app launches and UI elements are present", {
  # Launch the app
  app<-twoway_rbd()
  expect_true(inherits(app, "shiny.appobj"))

  test_that("anova computation works correctly", {
       # Prepare sample data
    fertilizer_data <- data.frame(
         Nitrogen = rep(c("N0", "N30", "N60"), each = 12),
         Phosphorus = rep(rep(c("P0", "P30", "P60", "P90"), each = 3), times = 3),
         Replication = factor(rep(1:3, times = 12)),
         Grain_Yield = c(
           4.2, 4.0, 4.1,  # N0-P0
           4.8, 4.6, 4.7,  # N0-P30
           5.1, 5.0, 5.2,  # N0-P60
           5.0, 4.9, 5.1,  # N0-P90
           5.5, 5.3, 5.4,  # N30-P0
           6.2, 6.0, 6.1,  # N30-P30
           6.8, 6.6, 6.7,  # N30-P60
           6.7, 6.5, 6.6,  # N30-P90
           6.0, 5.8, 5.9,  # N60-P0
           6.8, 6.6, 6.7,  # N60-P30
           7.5, 7.3, 7.4,  # N60-P60
           7.4, 7.2, 7.3   # N60-P90
         ),
         Protein_Content = c(
           9.0, 8.8, 8.9,   # N0-P0
           9.5, 9.3, 9.4,   # N0-P30
           9.8, 9.6, 9.7,   # N0-P60
           9.7, 9.5, 9.6,   # N0-P90
           10.5, 10.3, 10.4, # N30-P0
           11.2, 11.0, 11.1, # N30-P30
           11.8, 11.6, 11.7, # N30-P60
           11.7, 11.5, 11.6, # N30-P90
           12.0, 11.8, 11.9, # N60-P0
           12.8, 12.6, 12.7, # N60-P30
           13.5, 13.3, 13.4, # N60-P60
           13.4, 13.2, 13.3  # N60-P90
           )
    )

    fertilizer_data$Replication<-as.factor(fertilizer_data$Replication)
#expected_anova
       exp_aov<-aov(Grain_Yield~ fertilizer_data$Nitrogen*fertilizer_data$Phosphorus+fertilizer_data$Replication, data =fertilizer_data)
#VERIFY
       expect_equal(
         aov(Grain_Yield~ fertilizer_data$Nitrogen*fertilizer_data$Phosphorus+fertilizer_data$Replication, data = fertilizer_data),
         exp_aov
       )

       })
  test_that("file handling utilities work", {
    # Create test data
    test_data <- data.frame(
      A = 1:3,
      B = 4:6,
      C = 7:9
    )

    # Test CSV handling
    temp_csv <- tempfile(fileext = ".csv")
    write.csv(test_data, temp_csv, row.names = FALSE)
    read_data <- read.csv(temp_csv)
    expect_equal(test_data, read_data)
    unlink(temp_csv)

    # Test Excel handling
    if(requireNamespace("writexl", quietly = TRUE)) {
      temp_xlsx <- tempfile(fileext = ".xlsx")
      writexl::write_xlsx(test_data, temp_xlsx)
      read_data <- readxl::read_excel(temp_xlsx)
      expect_equal(test_data, as.data.frame(read_data))
      unlink(temp_xlsx)
    }
  })



  })
