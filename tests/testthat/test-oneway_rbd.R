library(testthat)
library(shinytest2)
library(visvaR)
# Ensure your package is loaded to access oneway_rbd()

test_that("oneway_rbd app launches and UI elements are present", {
  # Launch the app
  app<-oneway_rbd()
  expect_true(inherits(app, "shiny.appobj"))

  test_that("anova computation works correctly", {
    # Prepare sample data
    plant_data <- data.frame(
      Treatment    = rep(c("Control", "Low", "Medium", "High"), each = 3),
      Replication  = factor(rep(1:3, times = 4)),
      Plant_Height = c(25.3, 24.8, 25.1,
                       27.6, 28.1, 27.9,
                       30.2, 29.8, 30.5,
                       26.8, 27.2, 26.5),
      Leaf_Count = c(8, 7, 8,
                     10, 11, 10,
                     12, 13, 12,
                     9, 8, 9)
    )

    #expected_anova
    exp_aov<-aov(plant_data$Plant_Height ~ plant_data$Treatment+plant_data$Replication, data = plant_data)
    #VERIFY
    expect_equal(
      aov(plant_data$Plant_Height ~ plant_data$Treatment+plant_data$Replication , data = plant_data),
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
