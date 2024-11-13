#' Run the Shiny App for Correlation Analysis of multiple variables
#'
#' App allows user to import data through excel/.csv files or through clipboard
#' and select the correlation method and download the results with few customization options
#'
#' @details
#' This Shiny app is part of the `visvaR` package and is designed for
#' correlation analysis and user can download the report in word format also has
#' option to download the correlation plot as .png file. To use custom fonts, please install the `extrafont` package and run `extrafont::font_import()` and `extrafont::loadfonts()`.
#'
#' @return
#' This function runs a local instance of the Shiny app in your default web
#' browser. The app interface allows users to upload data, select analysis
#' methods, and download outputs.
#' @usage correlation_multi
#' @aliases cor_multiple
#' @aliases correlation_multi
#' @name correlation_multi
#' @examples
#' # Example 1: Basic usage
#' if(interactive()) {
#'   correlation_multi()
#' }
#'
#' # Example 2: Sample workflow with iris dataset
#' if(interactive()) {
#'   # Prepare sample data
#'   write.csv(iris[, 1:4], "sample_data.csv", row.names = FALSE)
#'
#'   # Launch the app
#'   correlation_multi()
#'
#'   # Instructions for users:
#'   # 1. Click "Choose a file" and select sample_data.csv
#'   # 2. Select correlation method (e.g., "pearson")
#'   # 3. Click "Analyze"
#'   # 4. View results in different tabs
#'   # 5. Download plot or Word report as needed
#'
#'   # Clean up
#'   unlink("sample_data.csv")
#' }
#'
#' # Example 3: Using clipboard data
#' if(interactive()) {
#'   # Copy this to clipboard:
#'   # Var1,Var2,Var3
#'   # 1,2,3
#'   # 4,5,6
#'   # 7,8,9
#'
#'   correlation_multi()
#'   # Click "Paste from Clipboard" after copying data
#' }
#'
#' @author Ramesh Ramasamy
#' @author Mathiyarsai Kulandaivadivel
#' @author Tamilselvan Arumugam
#' @references Wei, T., & Simko, V. (2021). R package 'corrplot': Visualization of a Correlation Matrix (Version 0.92). Available from https://CRAN.R-project.org/package=corrplot.
#'             Friendly, M. (2002). Corrgrams: Exploratory displays for correlation matrices. The American Statistician, 56(4), 316-324.
#' @import DT
#' @import corrplot
#' @import bslib
#' @import corrplot
#' @importFrom readxl read_excel
#' @import officer
#' @import flextable
#' @importFrom tools file_ext
#' @importFrom htmltools HTML
#' @importFrom utils read.delim read.csv
#' @importFrom stats cor
#' @importFrom graphics par
#' @importFrom grDevices png dev.off
NULL
#' @export
correlation_multi<-function(){
  default_data <- data.frame(
    Var1 = stats::rnorm(10),
    Var2 = stats::rnorm(10),
    Var3 = stats::rnorm(10),
    Var4 = stats::rnorm(10)
  )

  correlation_multi_ui <- shiny::fluidPage(
    theme = bslib::bs_theme(
      version = 5,
      bootswatch = "simplex",
      primary = "#007bff",
      "font-size-base" = "0.95rem"
    ),
    shiny::tags$head(
      shiny::tags$style(HTML("
      h1 {
        text-align: center;
      }
      .btn-info {
        margin: 5px;
      }
    "))
    ),
    shiny::h1("Correlation Analysis (Multiple variables)",
              style = "font-family:Times New Roman; font-weight: bold; font-size: 36px; color: #890304;"),

    shiny::sidebarLayout(
      shiny::sidebarPanel(
        position = "right",
        fluid = TRUE,
        shiny::actionButton("clipboard_input", "Paste from Clipboard",
                            class = "btn-primary",
                            style = "margin-bottom: 10px;"),
        shiny::fileInput("file", "Or choose a file (CSV or Excel)",
                         accept = c(".csv", ".xlsx")),
        shiny::selectInput("method", "Correlation Method:",
                           choices = c("pearson", "kendall", "spearman"),
                           selected = "pearson"),  # Default method
        shiny::selectInput("plot_method", "Plot Method:",
                           choices = c('circle', 'square', 'ellipse', 'number', 'shade', 'color', 'pie'),
                           selected = "circle"),  # Default method
        shiny::hr(),
        shiny::selectInput("font_family", "Choose Font Style:",
                           choices = c("Times New Roman", "Arial", "Helvetica", "serif", "sans", "mono"),
                           selected = "Times New Roman"),  # Default font
        shiny::numericInput("font_size", "Font Size:",
                            value = 12, min = 8, max = 20),
        shiny::selectInput("color_palette", "Choose Color Palette:",
                           choices = c("RdYlBu", "RdBu", "PuOr", "PRGn", "BrBG", "PiYG"),
                           selected = "RdYlBu"),  # Default palette
        shiny::numericInput("plot_width", "Plot Width (inches):",
                            value = 10, min = 1, max = 20),
        shiny::numericInput("plot_height", "Plot Height (inches):",
                            value = 10, min = 1, max = 20),
        shiny::numericInput("plot_dpi", "Plot Resolution(dpi):",
                            value = 300, min = 72, max = 600),
        shiny::actionButton("analyze", "Analyze",
                            style = "color: green; background-color: #4CAF50; font-size: 20px; margin: 10px 0;"),
        shiny::downloadButton("download_plot", "Download Plot",
                              class = "btn-info"),
        shiny::downloadButton("download_word", "Download Word Report",
                              class = "btn-info"),
        shiny::fluidPage(
          shiny::fluidRow(
            shiny::column(9, offset = 1,
                          bslib::card(
                            uiOutput("logo"),
                            height = "auto",
                            bslib::card_header(
                              shiny::div(
                                class = "text-center",
                                "visvaR- VISualize VARiance"
                              ),
                              class = "bg-primary text-white"
                            ),
                            shiny::HTML("
                  <p style='text-align: center;'>
                    <strong>Developed by</strong><br>
                    <span style='font-size: 24px;'>Ramesh R</span><br>
                    <span>PhD Scholar</span><br>
                    <span>Division of Plant Physiology </span><br> ICAR-IARI, New Delhi</span><br>
                    <span>ramesh.rahu96@gmail.com</span>
                  </p>
                ")
                          )
            )
          )
        )
      ),

      shiny::mainPanel(
        shiny::tabsetPanel(
          shiny::tabPanel("Data Preview",
                          shiny::div(
                            style = "margin-top: 20px;",
                            DT::DTOutput("data_preview")
                          )
          ),
          shiny::tabPanel("Correlation Plot",
                          shiny::div(
                            style = "margin-top: 20px;",
                            shiny::plotOutput("correlation_plot", height = "600px")
                          )
          ),
          shiny::tabPanel("Correlation Matrix",
                          shiny::div(
                            style = "margin-top: 20px;",
                            DT::DTOutput("correlation_matrix")
                          )
          )
        )
      )
    )
  )

  correlation_multi_server <- function(input, output, session) {
    # Initialize reactive values with default data
    data_reactive <- shiny::reactiveVal(default_data)
    correlation_matrix <- shiny::reactiveVal(cor(default_data))
    original_names <- shiny::reactiveVal(colnames(default_data))
    session$onSessionEnded(function() {
    shiny::stopApp()
      })

    # Observe event for clipboard input
    shiny::observeEvent(input$clipboard_input, {
      tryCatch({
        df <- read.delim("clipboard", header = TRUE, check.names = FALSE)
        data_reactive(df)
        original_names(colnames(df))
        shiny::showNotification("Data successfully read from clipboard", type = "message")
      }, error = function(e) {
        shiny::showNotification("Error reading from clipboard. Please try again.", type = "error")
      })
    })

    # Observe event for file input
    shiny::observeEvent(input$file, {
      shiny::req(input$file)
      ext <- tools::file_ext(input$file$name)

      tryCatch({
        if (ext == "csv") {
          df <- read.csv(input$file$datapath, header = TRUE, check.names = FALSE)
        } else if (ext == "xlsx") {
          df <- readxl::read_excel(input$file$datapath, col_names = TRUE)
        } else {
          stop("Unsupported file format")
        }
        data_reactive(df)
        original_names(colnames(df))
        shiny::showNotification("File successfully uploaded", type = "message")
      }, error = function(e) {
        shiny::showNotification(paste("Error reading file:", e$message), type = "error")
      })
    })

    # Render data table with improved styling
    output$data_preview <- DT::renderDT({
      shiny::req(data_reactive())
      DT::datatable(
        data_reactive(),
        options = list(
          scrollX = TRUE,
          scrollY = "400px",
          pageLength = 15,
          dom = 'Bfrtip',
          buttons = c('copy', 'csv', 'excel')
        ),
        class = 'cell-border stripe'
      )
    })

    # Analyze data for correlation matrix
    shiny::observeEvent(input$analyze, {
      shiny::req(data_reactive())
      M <- cor(data_reactive(), method = input$method)
      correlation_matrix(M)
    })

    # Create the correlation plot with improved styling
    create_correlation_plot <- function() {
      shiny::req(correlation_matrix())

      M <- correlation_matrix()
      res1 <- corrplot::cor.mtest(data_reactive(), conf.level = 0.95)
      oldpar <- par(no.readonly = TRUE)
      on.exit(par(oldpar))
      par(mfrow = c(1, 1),
          mar = c(5, 4, 4, 2) + 0.1,
          family = input$font_family)

      corrplot::corrplot(
        M,
        p.mat = res1$p,
        order = "AOE",
        method = input$plot_method,
        type = "upper",
        tl.pos = "lt",
        tl.col = "black",
        insig = "label_sig",
        sig.level = c(0.001, 0.01, 0.05),
        pch.cex = 2,
        pch.col = "black",
        tl.srt = 45,
        col = corrplot::COL2(input$color_palette, 10),
        tl.offset = 0.5
      )

      corrplot::corrplot(
        M,
        add = TRUE,
        type = "lower",
        method = "number",
        col = "black",
        number.cex = input$font_size / 12,
        order = "AOE",
        diag = FALSE,
        tl.pos = "n",
        cl.pos = "n"
      )
    }

    # Render the correlation plot
    output$correlation_plot <- shiny::renderPlot({
      create_correlation_plot()
    })

    # Render the correlation matrix with improved styling
    output$correlation_matrix <- DT::renderDT({
      shiny::req(correlation_matrix())
      matrix_with_names <- correlation_matrix()
      colnames(matrix_with_names) <- original_names()
      rownames(matrix_with_names) <- original_names()
      DT::datatable(
        round(matrix_with_names, 2),
        options = list(
          scrollX = TRUE,
          scrollY = "400px",
          dom = 'Bfrtip',
          buttons = c('copy', 'csv', 'excel')
        ),
        class = 'cell-border stripe'
      )
    })

    # Download handlers remain the same
    output$download_plot <- shiny::downloadHandler(
      filename = function() {
        paste("correlation_plot_", Sys.Date(), ".png", sep = "")
      },
      content = function(file) {
        png(file, width = input$plot_width, height = input$plot_height,
            units = "in", res = input$plot_dpi)
        create_correlation_plot()
        dev.off()
      }
    )

    output$download_word <- shiny::downloadHandler(
      filename = function() {
        paste("correlation_analysis_report_", Sys.Date(), ".docx", sep = "")
      },
      content = function(file) {
        doc <- officer::read_docx()
        fp_bold <- officer::fp_text_lite(bold = TRUE)
        fp_refnote <- officer::fp_text_lite(vertical.align = "baseline")
        bl <- officer::block_list(
          officer::fpar(officer::ftext("Analysis done using visvaR R package ", fp_bold)),
          officer::fpar(officer::ftext("Developed by Ramesh R PhD scholar, Division of Plant Physiology, ICAR-IARI, New Delhi", fp_bold))
        )
        a_par <- officer::fpar(
          "visvaR- Visualize Variance                                           ",
          officer::run_footnote(x = bl, prop = fp_refnote),
          ""
        )

        doc <- officer::body_add_fpar(doc, value = a_par, style = "Normal")

        flextable::set_flextable_defaults(
          font.size = 12,
          font.family = "Times New Roman",
          font.color = "#333333",
          table.layout = "autofit",
          border.color = "black"
        )

        doc <- doc %>%
          officer::body_add_par("Correlation Analysis Report", style = "heading 1")

        temp_plot <- tempfile(fileext = ".png")
        png(temp_plot, width = input$plot_width, height = input$plot_height,
            units = "in", res = input$plot_dpi)
        create_correlation_plot()
        dev.off()

        doc <- doc %>%
          officer::body_add_par("Correlation Plot", style = "heading 2") %>%
          officer::body_add_img(src = temp_plot, width = 6, height = 6)

        doc <- doc %>%
          officer::body_add_par("Correlation Matrix", style = "heading 2")

        matrix_with_names <- correlation_matrix()
        colnames(matrix_with_names) <- original_names()
        rownames(matrix_with_names) <- original_names()
        matrix_table <- flextable::flextable(round(as.data.frame(matrix_with_names), 2))
        matrix_table <- flextable::autofit(matrix_table)
        matrix_table <- flextable::add_footer_lines(matrix_table,
                                                    paste("Method used:", input$method))
        doc <- doc %>%
          body_add_flextable(matrix_table)

        print(doc, target = file)
      }
    )
  }

  shiny::shinyApp(ui = correlation_multi_ui, server = correlation_multi_server)
}
