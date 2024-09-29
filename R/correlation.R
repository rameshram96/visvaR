# Load required libraries
library(shiny)
library(corrplot)
library(stats)
library(ggcorrplot)
library(readxl)
library(openxlsx)
library(DT)
library(officer)
library(flextable)
library(bslib)

# Define UI
ui <- fluidPage(
  theme = bs_theme(
    version = 5,
    bootswatch = "simplex",
    primary = "#007bff",
    "font-size-base" = "0.95rem"
  ),
  tags$head(
    tags$style(HTML("
      h1 {
        text-align: center;
      }
    "))
  ),
  h1("Correlation Analysis",
     style = "font-family:Times New Roman; font-weight: bold; font-size: 36px; color: #4CAF50;"),

  sidebarLayout(
    sidebarPanel(
      position="right",
      fluid=TRUE,
      actionButton("clipboard_input", "Paste from Clipboard"),
      fileInput("file", "Or choose a file (CSV or Excel)", accept = c(".csv", ".xlsx")),
      selectInput("method", "Correlation Method:",
                  choices = c("pearson", "kendall", "spearman")),
      actionButton("analyze", "Analyze",style = "color: royalblue; font-size: 20px;"),
      hr(),
      selectInput("font_family", "Choose Font Style:",
                  choices = c("serif", "sans", "mono", "Arial", "Helvetica", "Times New Roman")),
      numericInput("font_size", "Font Size:", value = 12, min = 8, max = 20),
      selectInput("color_palette", "Choose Color Palette:",
                  choices = c("RdYlBu", "RdBu", "PuOr", "PRGn", "BrBG", "PiYG")),
      numericInput("plot_width", "Plot Width (inches):", value = 10, min = 1, max = 20),
      numericInput("plot_height", "Plot Height (inches):", value = 8, min = 1, max = 20),
      numericInput("plot_dpi", "Plot DPI:", value = 300, min = 72, max = 600),
      downloadButton("download_plot", "Download Plot",class = "btn-info"),
      downloadButton("download_word", "Download Word Report",class = "btn-info"),
      fluidPage(
        fluidRow(
          column(9, offset = 1,  # This shifts the column 4 spaces to the right
                 card(
                   height = "auto",
                   card_header(
                     div(
                       class = "text-center",
                       "visvaR- VISualize VARiance"
                     ),
                     class = "bg-primary text-white"
                   ),
                   HTML("
        <p style='text-align: center;'>
          <strong>Developed by</strong><br>
          <span style='font-size: 24px;'>Ramesh R</span><br>
          <span>PhD Scholar</span><br>
          <span>Division of Plant Physiology </span><br> ICAR-IARI, New Delhi</span><br>
          <span>ramesh.rahu96@gmail.com</span>
        </p>
        "),
                   card_image(src="https://github.com/rameshram96/visvaR/blob/main/visvaRlogo.png",
                              style = "display: block; margin-left: auto; margin-right: auto;",
                              alt = "",
                              href = NULL,
                              border_radius = c("auto"),
                              mime_type = NULL,
                              class = NULL,
                              height ="30%",
                              fill = FALSE,
                              width = "30%",
                              container = NULL
                   )
                 )
          )
        ),

      ),
    ),

    mainPanel(
      tabsetPanel(
        tabPanel("Data Preview", DTOutput("data_preview")),
        tabPanel("Correlation Plot", plotOutput("correlation_plot", height = "600px")),
        tabPanel("Correlation Matrix", DTOutput("correlation_matrix"))
      )
    )
  )
)
# Define server logic
server <- function(input, output, session) {

  data <- reactiveVal(NULL)
  correlation_matrix <- reactiveVal(NULL)
  original_names <- reactiveVal(NULL)

  observeEvent(input$clipboard_input, {
    tryCatch({
      df <- read.delim("clipboard", header = TRUE, check.names = FALSE)
      data(df)
      original_names(colnames(df))
      showNotification("Data successfully read from clipboard", type = "message")
    }, error = function(e) {
      showNotification("Error reading from clipboard. Please try again.", type = "error")
    })
  })

  observeEvent(input$file, {
    req(input$file)
    ext <- tools::file_ext(input$file$name)

    tryCatch({
      if (ext == "csv") {
        df <- read.csv(input$file$datapath, header = TRUE, check.names = FALSE)
      } else if (ext == "xlsx") {
        df <- read_excel(input$file$datapath, col_names = TRUE)
      } else {
        stop("Unsupported file format")
      }
      data(df)
      original_names(colnames(df))
      showNotification("File successfully uploaded", type = "message")
    }, error = function(e) {
      showNotification(paste("Error reading file:", e$message), type = "error")
    })
  })

  output$data_preview <- renderDT({
    req(data())
    datatable(data(), options = list(scrollX = TRUE, scrollY = "400px", pageLength = 15))
  })

  observeEvent(input$analyze, {
    req(data())
    M <- cor(data(),method = input$method)
    correlation_matrix(M)
  })

  create_correlation_plot <- function() {
    req(correlation_matrix())

    M <- correlation_matrix()
    res1 <- cor.mtest(data(), conf.level = 0.95)

    par(mfrow = c(1, 1), mar = c(5, 4, 4, 2) + 0.1, family = input$font_family)

    corrplot(M, p.mat = res1$p, order = "AOE", type = "upper", tl.pos = "lt", tl.col = "black",
             insig = "label_sig", sig.level = c(0.001, 0.01, 0.05), pch.cex = 2, pch.col = "black",
             tl.cex = input$font_size / 12, cl.cex = input$font_size / 12, tl.srt = 45,
             col = COL2(input$color_palette, 10),
             tl.offset = 0.5)

    corrplot(M, add = TRUE, type = "lower", method = "number", col = "black",
             number.cex = input$font_size / 12, order = "AOE", diag = FALSE,
             tl.pos = "n", cl.pos = "n", tl.cex = input$font_size / 12, tl.srt = 45)
  }

  output$correlation_plot <- renderPlot({
    create_correlation_plot()
  })

  output$correlation_matrix <- renderDT({
    req(correlation_matrix())
    matrix_with_names <- correlation_matrix()
    colnames(matrix_with_names) <- original_names()
    rownames(matrix_with_names) <- original_names()
    datatable(round(matrix_with_names, 2), options = list(scrollX = TRUE, scrollY = "400px"))
  })

  output$download_plot <- downloadHandler(
    filename = function() {
      paste("correlation_plot_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      png(file, width = input$plot_width, height = input$plot_height, units = "in", res = input$plot_dpi)
      create_correlation_plot()
      dev.off()
    }
  )

  output$download_word <- downloadHandler(
    filename = function() {
      paste("correlation_analysis_report_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      doc <- read_docx()
      fp_bold <- fp_text_lite(bold = TRUE)
      fp_refnote <- fp_text_lite(vertical.align = "baseline")
      bl <- block_list(
        fpar(ftext("Analysis done using visvaR R package ", fp_bold)),
        fpar(
          ftext("Developed by Ramesh R PhD scholar, Division of Plant Physiology, ICAR-IARI, New Delhi", fp_bold))
      )
      a_par <- fpar(
        "visvaR- Visulaize Variance                                           ",
        run_footnote(x = bl, prop = fp_refnote),
        ""
      )

      doc <- body_add_fpar(doc, value = a_par, style = "Normal")

      set_flextable_defaults(
        font.size = 12, font.family = "Times New Roman",
        font.color = "#333333",
        table.layout = "autofit",
        border.color = "black")
      # Add title
      doc <- doc %>%
        body_add_par("Correlation Analysis Report", style = "heading 1")
      # Add correlation plot
      temp_plot <- tempfile(fileext = ".png")
      png(temp_plot, width = input$plot_width, height = input$plot_height, units = "in", res = input$plot_dpi)
      create_correlation_plot()
      dev.off()

      doc <- doc %>%
        body_add_par("Correlation Plot", style = "heading 2") %>%
        body_add_img(src = temp_plot, width = 6, height = 6)

      # Add correlation matrix
      doc <- doc %>%
        body_add_par("Correlation Matrix", style = "heading 2")

      matrix_with_names <- correlation_matrix()
      colnames(matrix_with_names) <- original_names()
      rownames(matrix_with_names) <- original_names()
      matrix_table <- flextable(round(as.data.frame(matrix_with_names), 2))
      matrix_table <- autofit(matrix_table)
      matrix_table<-add_footer_lines(matrix_table,paste("Method used:", input$method))
      doc <- doc %>%
        body_add_flextable(matrix_table)

      print(doc, target = file)
    }
  )
}

# Run the Shiny app
shinyApp(ui = ui, server = server)
