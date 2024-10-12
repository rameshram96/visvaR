#' Run the Shiny App for Correlation Analysis of multiple variables
#'
#' App allows user to import data through excel/.csv files or through clipboard
#' and select the correlation method and download the results with few customization options
#'
#' @details
#' This Shiny app is part of the `visvaR` package and is designed for
#' correlation analysis and user can download the report in word format also has
#' option to download the correlation plot as .png file
#'
#' @return
#' This function runs a local instance of the Shiny app in your default web
#' browser. The app interface allows users to upload data, select analysis
#' methods, and download outputs.
#'
#' @usage correlation_multi()
#'
#' @name correlation_multi()
#' @author Ramesh Ramasamy
#' @author Mathiyarsai Kulandaivadivel
#' @author Tamilselvan Arumugam
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
correlation_multi<-function(){
  ui <- shiny::fluidPage(
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
    "))
    ),
    shiny::h1("Correlation Analysis (Multiple variables)",
              style = "font-family:Times New Roman; font-weight: bold; font-size: 36px; color: #4CAF50;"),

    shiny::sidebarLayout(
      shiny::sidebarPanel(
        position="right",
        fluid=TRUE,
        shiny::actionButton("clipboard_input", "Paste from Clipboard"),
        shiny::fileInput("file", "Or choose a file (CSV or Excel)", accept = c(".csv", ".xlsx")),
        shiny::selectInput("method", "Correlation Method:",
                           choices = c("pearson", "kendall", "spearman")),
        shiny::actionButton("analyze", "Analyze",style = "color: royalblue; font-size: 20px;"),
        shiny::hr(),
        shiny::selectInput("font_family", "Choose Font Style:",
                           choices = c("serif", "sans", "mono", "Arial", "Helvetica", "Times New Roman")),
        shiny::numericInput("font_size", "Font Size:", value = 12, min = 8, max = 20),
        shiny::selectInput("color_palette", "Choose Color Palette:",
                           choices = c("RdYlBu", "RdBu", "PuOr", "PRGn", "BrBG", "PiYG")),
        shiny::numericInput("plot_width", "Plot Width (inches):", value = 10, min = 1, max = 20),
        shiny::numericInput("plot_height", "Plot Height (inches):", value = 8, min = 1, max = 20),
        shiny::numericInput("plot_dpi", "Plot Resolution(dpi):", value = 300, min = 72, max = 600),
        shiny::downloadButton("download_plot", "Download Plot",class = "btn-info"),
        shiny::downloadButton("download_word", "Download Word Report",class = "btn-info"),
        shiny::fluidPage(
          shiny::fluidRow(
            shiny::column(9, offset = 1,
                          bslib::card(uiOutput("logo"),
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
          ),
        ),
      ),

      shiny::mainPanel(
        shiny::tabsetPanel(
          shiny::tabPanel("Data Preview", DT::DTOutput("data_preview")),
          shiny::tabPanel("Correlation Plot", shiny::plotOutput("correlation_plot", height = "600px")),
          shiny::tabPanel("Correlation Matrix", DT::DTOutput("correlation_matrix"))
        )
      )
    )
  )

  # Define server logic
  server <- function(input, output, session) {

    output$logo <- renderUI({
      img_path <- system.file("~/visvaR/inst/www", "hex_visvaR.png", package = "visvaR")
      tags$img(src = img_path, alt = "logo", style = "max-width:50%; height: auto;")
    })

    # Define reactive values to hold data
    data_reactive <- shiny::reactiveVal(NULL)  # Holds the dataset locally in reactive value
    correlation_matrix <- shiny::reactiveVal(NULL)
    original_names <- shiny::reactiveVal(NULL)

    # Observe event for clipboard input
    shiny::observeEvent(input$clipboard_input, {
      tryCatch({
        df <- read.delim("clipboard", header = TRUE, check.names = FALSE)
        data_reactive(df)  # Store the dataframe in the reactive value
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
        data_reactive(df)  # Store the dataframe in the reactive value
        original_names(colnames(df))
        shiny::showNotification("File successfully uploaded", type = "message")
      }, error = function(e) {
        shiny::showNotification(paste("Error reading file:", e$message), type = "error")
      })
    })

    # Render data table
    output$data_preview <- DT::renderDT({
      shiny::req(data_reactive())  # Access data from reactive value
      DT::datatable(data_reactive(), options = list(scrollX = TRUE, scrollY = "400px", pageLength = 15))
    })

    # Analyze data for correlation matrix
    shiny::observeEvent(input$analyze, {
      shiny::req(data_reactive())
      M <- cor(data_reactive(), method = input$method)  # Perform correlation on reactive data
      correlation_matrix(M)  # Store correlation matrix in reactive value
    })

    # Create the correlation plot
    create_correlation_plot <- function() {
      shiny::req(correlation_matrix())

      M <- correlation_matrix()
      res1 <- corrplot::cor.mtest(data_reactive(), conf.level = 0.95)

      par(mfrow = c(1, 1), mar = c(5, 4, 4, 2) + 0.1, family = input$font_family)

      corrplot::corrplot(M, p.mat = res1$p, order = "AOE", type = "upper", tl.pos = "lt", tl.col = "black",
                         insig = "label_sig", sig.level = c(0.001, 0.01, 0.05), pch.cex = 2, pch.col = "black",
                         tl.cex = input$font_size / 12, cl.cex = input$font_size / 12, tl.srt = 45,
                         col = corrplot::COL2(input$color_palette, 10),
                         tl.offset = 0.5)

      corrplot::corrplot(M, add = TRUE, type = "lower", method = "number", col = "black",
                         number.cex = input$font_size / 12, order = "AOE", diag = FALSE,
                         tl.pos = "n", cl.pos = "n", tl.cex = input$font_size / 12, tl.srt = 45)
    }

    # Render the correlation plot
    output$correlation_plot <- shiny::renderPlot({
      create_correlation_plot()
    })

    # Render the correlation matrix
    output$correlation_matrix <- DT::renderDT({
      shiny::req(correlation_matrix())
      matrix_with_names <- correlation_matrix()
      colnames(matrix_with_names) <- original_names()
      rownames(matrix_with_names) <- original_names()
      DT::datatable(round(matrix_with_names, 2), options = list(scrollX = TRUE, scrollY = "400px"))
    })

    # Download the correlation plot
    output$download_plot <- shiny::downloadHandler(
      filename = function() {
        paste("correlation_plot_", Sys.Date(), ".png", sep = "")
      },
      content = function(file) {
        png(file, width = input$plot_width, height = input$plot_height, units = "in", res = input$plot_dpi)
        create_correlation_plot()
        dev.off()
      }
    )

    # Download the Word report
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
          font.size = 12, font.family = "Times New Roman",
          font.color = "#333333",
          table.layout = "autofit",
          border.color = "black")

        # Add title
        doc <- doc %>%
          officer::body_add_par("Correlation Analysis Report", style = "heading 1")

        # Add correlation plot
        temp_plot <- tempfile(fileext = ".png")
        png(temp_plot, width = input$plot_width, height = input$plot_height, units = "in", res = input$plot_dpi)
        create_correlation_plot()
        dev.off()

        doc <- doc %>%
          officer::body_add_par("Correlation Plot", style = "heading 2") %>%
          officer::body_add_img(src = temp_plot, width = 6, height = 6)

        # Add correlation matrix
        doc <- doc %>%
          officer::body_add_par("Correlation Matrix", style = "heading 2")

        matrix_with_names <- correlation_matrix()
        colnames(matrix_with_names) <- original_names()
        rownames(matrix_with_names) <- original_names()
        matrix_table <- flextable::flextable(round(as.data.frame(matrix_with_names), 2))
        matrix_table <- flextable::autofit(matrix_table)
        matrix_table <- flextable::add_footer_lines(matrix_table, paste("Method used:", input$method))
        doc <- doc %>%
          body_add_flextable(matrix_table)

        print(doc, target = file)
      }
    )
  }

  # Run the Shiny app
  shiny::shinyApp(ui = ui, server = server)}
