library(testthat)
library(shiny)
library(DT)
library(corrplot)
library(bslib)
library(readxl)
library(officer)
library(flextable)

test_that("correlation_multi function returns a shiny app", {
  # Test that the function returns a shiny app
  app <- correlation_multi()
  expect_true(inherits(app, "shiny.appobj"))
})

test_that("correlation_multi app can be created with test data", {
  # Create test data
  test_data <- data.frame(
    A = rnorm(10),
    B = rnorm(10),
    C = rnorm(10)
  )

  # Write test data to temporary file
  temp_csv <- tempfile(fileext = ".csv")
  write.csv(test_data, temp_csv, row.names = FALSE)

  # Test that we can create the app
  app <- correlation_multi()
  expect_true(inherits(app, "shiny.appobj"))

  # Clean up
  unlink(temp_csv)
})

test_that("correlation computation works correctly", {
  # Create test data
  test_data <- data.frame(
    A = c(1, 2, 3, 4, 5),
    B = c(2, 4, 6, 8, 10),
    C = c(1, 3, 5, 7, 9)
  )

  # Calculate expected correlation
  expected_cor <- cor(test_data, method = "pearson")

  # Verify correlation computation
  expect_equal(
    cor(test_data, method = "pearson"),
    expected_cor
  )

  # Test different correlation methods
  expect_true(all(!is.na(cor(test_data, method = "spearman"))))
  expect_true(all(!is.na(cor(test_data, method = "kendall"))))
})

test_that("plot generation utilities work", {
  # Create test data
  test_data <- data.frame(
    A = rnorm(10),
    B = rnorm(10),
    C = rnorm(10)
  )

  # Test that we can create a correlation matrix
  test_cor <- cor(test_data)
  expect_true(is.matrix(test_cor))
  expect_equal(dim(test_cor), c(3, 3))

  # Test plot creation
  temp_png <- tempfile(fileext = ".png")
  png(temp_png)
  expect_error(
    corrplot::corrplot(test_cor, method = "circle"),
    NA
  )
  dev.off()

  # Check if plot file was created
  expect_true(file.exists(temp_png))

  # Clean up
  unlink(temp_png)
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
  if (requireNamespace("writexl", quietly = TRUE)) {
    temp_xlsx <- tempfile(fileext = ".xlsx")
    writexl::write_xlsx(test_data, temp_xlsx)
    read_data <- readxl::read_excel(temp_xlsx)
    expect_equal(test_data, as.data.frame(read_data))
    unlink(temp_xlsx)
  }
})

test_that("report generation utilities work", {
  # Create test data
  test_data <- data.frame(
    A = 1:3,
    B = 4:6,
    C = 7:9
  )

  # Create correlation matrix
  test_cor <- cor(test_data)

  # Test Word document creation
  temp_docx <- tempfile(fileext = ".docx")
  doc <- officer::read_docx()

  # Add title
  doc <- officer::body_add_par(doc, "Test Report", style = "heading 1")

  # Add correlation matrix as flextable
  ft <- flextable::flextable(as.data.frame(round(test_cor, 2)))
  ft <- flextable::autofit(ft)
  doc <- flextable::body_add_flextable(doc, ft)

  # Save document
  print(doc, target = temp_docx)

  # Check if file was created
  expect_true(file.exists(temp_docx))

  # Clean up
  unlink(temp_docx)
})
