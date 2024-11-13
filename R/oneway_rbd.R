#'Run the Shiny App for performing one-way ANOVA on data from a Randomized block design.
#'
#' App allows user to import data through excel/.csv files or through clipboard
#' and user can select the post-hoc test method and download the report which contains anova results and plots
#'
#' @details
#' This Shiny app is part of the `visvaR` package and is designed for
#' analysis of variance on data from randomized block design (one factor) and user can download the report in word format.
#' The analysis of variance was performed using R's aov() function (Chambers & Hastie, 1992; R Core Team, 2024), which implements the classical ANOVA methodology developed by Fisher (1925).To use custom fonts, please install the `extrafont` package and run `extrafont::font_import()` and `extrafont::loadfonts()`.
#' @return
#' This function runs a local instance of the Shiny app in your default web
#' browser. The app interface allows users to upload data, select analysis
#' method, and download outputs.
#'
#' @usage oneway_rbd()
#' @aliases anova1_rbd
#' @aliases oneway_rbd
#' @export oneway_rbd
#' @examples
#'
#' # Example 1: Basic usage
#' if(interactive()) {
#'   oneway_rbd()
#' }
#'
#' # Example 2: Sample workflow with plant growth experiment
#' if(interactive()) {
#'   # Prepare sample data
#'   plant_data <- data.frame(
#'     Treatment = rep(c("Control", "Low", "Medium", "High"), each = 3),
#'     Replication = rep(1:3, times = 4),
#'     Plant_Height = c(25.3, 24.8, 25.1,  # Control
#'                      27.6, 28.1, 27.9,  # Low
#'                      30.2, 29.8, 30.5,  # Medium
#'                      26.8, 27.2, 26.5), # High
#'     Leaf_Count = c(8, 7, 8,    # Control
#'                    10, 11, 10,  # Low
#'                    12, 13, 12,  # Medium
#'                    9, 8, 9)     # High
#'   )
#'
#'   # Save as Excel file
#'   write.xlsx(plant_data, "plant_data.xlsx")
#'
#'   # Launch the app
#'   oneway_rbd()
#'
#'   # Instructions for users:
#'   # 1. Click "Choose .xlsx or .csv" and select plant_data.xlsx
#'   # 2. Select post-hoc test method (e.g., "Tukey HSD")
#'   # 3. Customize plot appearance if desired:
#'   #    - Choose bar color
#'   #    - Select font style
#'   #    - Adjust font size
#'   # 4. Click "Analyze"
#'   # 5. View results in different tabs
#'   # 6. Download Word report
#'
#'   # Clean up
#'   unlink("plant_data.xlsx")
#' }
#'
#' # Example 3: Using clipboard data
#' if(interactive()) {
#'   # Copy this to clipboard:
#'   # Treatment,Replication,Yield
#'   # Control,1,45.2
#'   # Control,2,44.8
#'   # Control,3,45.5
#'   # Treatment1,1,48.6
#'   # Treatment1,2,49.2
#'   # Treatment1,3,48.9
#'   # Treatment2,1,52.3
#'   # Treatment2,2,51.8
#'   # Treatment2,3,52.7
#'
#'   oneway_rbd()
#'   # Click "Use Clipboard Data" after copying data
#' }
#'
#' # Example 4: Multiple response variables
#' if(interactive()) {
#'   # Create data with multiple responses
#'   multi_response_data <- data.frame(
#'     Treatment = rep(c("Control", "Treatment1", "Treatment2"), each = 4),
#'     Replication = rep(1:4, times = 3),
#'     Height = rnorm(12, mean = c(rep(20,4), rep(25,4), rep(30,4)), sd = 2),
#'     Weight = rnorm(12, mean = c(rep(50,4), rep(60,4), rep(70,4)), sd = 5),
#'     Length = rnorm(12, mean = c(rep(10,4), rep(12,4), rep(15,4)), sd = 1)
#'   )
#'
#'   # Save as Excel file
#'   write.xlsx(multi_response_data, "multi_response.xlsx")
#'
#'   # Launch the app
#'   oneway_rbd()
#'   }
#'
#' @references Fisher, R. A. (1925). Statistical Methods for Research Workers. Oliver and Boyd, Edinburgh.
#'             Scheffe, H. (1959). The Analysis of Variance. John Wiley & Sons, New York.
#'             R Core Team (2024). R: A language and environment for statistical computing. R Foundation for Statistical Computing, Vienna, Austria. URL https://www.R-project.org/
#' @name oneway_rbd
#' @author Ramesh Ramasamy
#' @author Mathiyarsai Kulandaivadivel
#' @author Tamilselvan Arumugam
#' @import dplyr
#' @import ggplot2
#' @import agricolae
#' @import bslib
#' @import corrplot
#' @import flextable
#' @import ggcorrplot
#' @import officer
#' @import patchwork
#' @import tibble
#' @import tidyr
#' @importFrom rlang .data
#' @importFrom ggplot2 aes
#' @importFrom shiny h1 div fluidRow fileInput actionButton strong selectInput textInput numericInput downloadButton tabsetPanel tabPanel uiOutput reactiveVal renderUI observeEvent req showNotification renderPrint renderPlot tagList h3 verbatimTextOutput plotOutput hr removeNotification downloadHandler
#' @importFrom stats aov anova fitted resid sd deviance
#' @importFrom grDevices png dev.off
#' @importFrom graphics par
#' @importFrom shiny tags shinyApp
NULL
utils::globalVariables(c("Fitted",'read_excel', "Residuals", "Sample", "Factor_A", "Response", "avg_A", "se",'captionpaste'))
oneway_rbd<-function(){
  ui <- page_fluid(
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

    # Center-aligned title
    h1("One-way Randomized Comlete Block Design (RCBD)",
       style = "font-family:Times New Roman; font-weight: bold; font-size: 36px; color: #890304;"),
    layout_columns(
      col_widths = c(3, 9),
      card( height = "45",
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
      ),
      card( height = "auto",
            card_header(
              div( class="text-center",
                   "Data Input and Analysis Options",
                   height= "120px",
                   class = "bg-primary text-white"
              )),
            card_body(
              fluidRow(
                fileInput("file", "Choose .xlsx or .csv", accept = c(".xlsx", ".csv")),
                actionButton("clipboard_button","Use Clipboard Data", class = "text-black",
                             style = "width: 700px; height:35px; float:center;margin-top: 30px;",
                             title="Click here to paste your clipboard data"),
                strong("Choose acccording to your need", align='center'),
                selectInput("test_method", "Choose Post-Hoc Test",
                            choices = c("LSD" = "LSD.test",
                                        "Duncan" = "duncan.test",
                                        "Tukey HSD" = "HSD.test",
                                        "Student-Newman-Keuls" = "SNK.test")),
                textInput("bar_color", "Bar Color (name or hex code)", value = "coral1"),
                selectInput("font_family", "Font Style for plot",
                            choices = c("Sans" = "sans", "Serif" = "serif", "Mono" = "mono", "Times New Roman" = "Times New Roman")),
                numericInput("font_size", "Font Size for plot", value = 12, min = 8, max = 20),
                textInput("output_filename", "Output Filename", value = "result"),
                actionButton("analyze_button", "Analyze", class = "btn-success"),
                strong("Downlod and save your results in document format by cliclking link below", align='center'),
                downloadButton("download_report", "Download Report", class = "btn-info")
              )
            )
      )
    ),
    layout_columns(
      col_widths = c(12),
      card(
        card_header(
          "Results",
          class = "bg-primary text-white"
        ),
        card_body(
          tabsetPanel(
            tabPanel("Data Preview", DT::DTOutput("preview")),
            tabPanel("",
                     uiOutput("analysis_outputs")
            )
          ),
          uiOutput("analysis_status")  # Add this line to display the analysis status
        )
      )
    )
  )

  server <- function(input, output, session) {
    data_reactive <- reactiveVal(NULL)
    report <- reactiveVal(NULL)
    analysis_complete <- reactiveVal(FALSE)
    session$onSessionEnded(function() {
      shiny::stopApp()
    })
    output$analysis_status <- renderUI({
      if (analysis_complete()) {
        div( class="text-center",
             style = "margin-top: 20px; text-align: center;",
             tags$strong("Analysis complete!", style = "color: green; font-size: 30px;")
        )
      }
    })
    observeEvent(input$file, {
      req(input$file)
      data_reactive(read_excel(input$file$datapath))
    })

    observeEvent(input$clipboard_button, {
      clipboard_data <- read.delim("clipboard", header = TRUE)
      data_reactive(clipboard_data)
    })

    output$preview <- DT::renderDT({
      req(data_reactive())
      data_reactive()
    })

    observeEvent(input$analyze_button, {
      req(data_reactive())

      # Reset the analysis_complete status
      analysis_complete(FALSE)

      # Show a notification that analysis has started
      showNotification("Analysis in progress...", type = "message", duration = NULL, id = "analysis_notification")

      data_full <- data_reactive()
      factor_col <- colnames(data_full)[1]
      rep_col <- colnames(data_full)[2]
      response_cols <- colnames(data_full)[3:ncol(data_full)]

      doc <- read_docx()

      for (i in seq_along(response_cols)) {
        response_col <- response_cols[i]

        data1 <- data_full %>%
          select(all_of(c(factor_col, rep_col, response_col))) %>%
          rename(Factor_A = 1, Replication = 2, Response = 3)

        data1$Factor_A <- as.factor(data1$Factor_A)
        data1$Replication <- as.factor(data1$Replication)

        anova_model <- aov(Response ~ Factor_A + Replication, data = data1)
        aov_result <- anova(anova_model)

        fitted_values <- fitted(anova_model)
        residuals <- resid(anova_model)

        res_fit_plot <- ggplot(data = data.frame(Fitted = fitted_values, Residuals = residuals), aes(x = Fitted, y = Residuals)) +
          geom_point() +
          geom_hline(yintercept = 0, linetype = "dashed", color = "red") +
          labs(title = paste("Residuals vs Fitted -", response_col), x = "Fitted Values", y = "Residuals") +
          theme_minimal() +
          theme(text = element_text(family = input$font_family, size = input$font_size))

        qq_plot <- ggplot(data = data.frame(Sample = residuals), aes(sample = Sample)) +
          stat_qq() +
          stat_qq_line() +
          labs(title = paste("Q-Q Plot -", response_col), x = "Theoretical Quantiles", y = "Sample Quantiles") +
          theme_minimal() +
          theme(text = element_text(family = input$font_family, size = input$font_size))
        combined_plot <- res_fit_plot + qq_plot + plot_layout(nrow = 2)
        perform_test <- function(response, treatment, DFerror, MSerror) {
          test_function <- get(input$test_method)
          test_function(y = response,
                        trt = treatment,
                        DFerror = DFerror,
                        MSerror = MSerror,
                        alpha = 0.05,
                        group = TRUE,
                        console = TRUE)}

        MeanSE_A <- data1 %>%
          group_by(Factor_A) %>%
          summarise(avg_A = mean(Response),
                    se = sd(Response)/sqrt(length(Response)))
        LSD_A <- perform_test(data1$Response, data1$Factor_A, anova_model$df.residual, deviance(anova_model)/anova_model$df.residual)

        ascend_A <- LSD_A$groups %>%
          group_by(rownames(LSD_A$groups)) %>%
          arrange(rownames(LSD_A$groups))

        main_plot <- ggplot(MeanSE_A, aes(x = Factor_A, y = avg_A, fill = Factor_A)) +
          geom_bar(stat = "identity", color = "black",
                   width = 0.5, position = position_dodge(width=0.5)) +
          geom_errorbar(aes(ymax = avg_A + se, ymin = avg_A - se),
                        position = position_dodge(width=0.1), width = 0.1) +
          theme(panel.background = element_rect(fill="white", colour="white", linetype="solid", color="white"),
                legend.position = "none",
                axis.text.x = element_text(angle = 0, face = "bold"),
                panel.border = element_blank(),
                panel.grid.major = element_blank(),
                panel.grid.minor = element_blank(),
                axis.line = element_line(linewidth = 0.5, linetype = "solid", colour = "black"),
                text = element_text(family = input$font_family, size = input$font_size)) +
          labs(title = paste("", response_col), x = factor_col, y = response_col,color = "factor_col") +
          geom_text(aes(x = Factor_A, y = avg_A + se, label = as.matrix(ascend_A$groups)),
                    position = position_dodge(width = 0.9), vjust = -(0.5)) +
          scale_fill_manual(values = rep(input$bar_color, length(unique(MeanSE_A$Factor_A))))+
          scale_y_continuous(expand = c(0,0))+
          expand_limits(y = c(0, max(MeanSE_A$avg_A+ (MeanSE_A$avg_A*0.25))))+
          labs(color = NULL)
      oldpar <- par(no.readonly = TRUE)
      on.exit(par(oldpar))
        # Generate Word document content
        aov_df <- data.frame(Source = rownames(aov_result), aov_result)
        aov_df[, "Significance"] <- ifelse(aov_df$Pr..F. <= 0.001, "***",
                                           ifelse(aov_df$Pr..F. <= 0.01, "**",
                                                  ifelse(aov_df$Pr..F. <= 0.05, "*",
                                                         ifelse(aov_df$Pr..F. <= 0.1, ".", " "))))

        fp_bold <- fp_text_lite(bold = TRUE)
        fp_refnote <- fp_text_lite(vertical.align = "baseline")
        bl <- block_list(
          fpar(ftext("Analysis done using visvaR R package ", fp_bold)),
          fpar(
            ftext("Developed by Ramesh R PhD scholar, Division of Plant Physiology, ICAR-IARI, New Delhi", fp_bold))
        )
        a_par <- fpar(
          "visvaR- Visualize Variance                                           ",
          run_footnote(x = bl, prop = fp_refnote),
          ""
        )

        doc <- body_add_fpar(doc, value = a_par, style = "Normal")

        set_flextable_defaults(
          font.size = 12, font.family = "Times New Roman",
          font.color = "#333333",
          table.layout = "autofit",
          border.color = "black")

        ft <- flextable(aov_df)
        ft <- set_caption(ft, caption = paste("ANOVA Results -", response_col))
        ft <- add_footer_lines(ft, values = "Signif. codes: 0.001='***', 0.01='**',0.05='*',0.1='.'")
        ft <- set_table_properties(ft, width = 0.9, layout = "autofit")

        colnames(ascend_A)[3] <- "Factor_A"
        combined_df <- merge(ascend_A, MeanSE_A, by = "Factor_A")
        ft_m <- flextable(combined_df)
        ft_m <- set_caption(ft_m, caption = paste("Mean Comparison -", response_col))
        ft_m <- set_table_properties(ft_m, width = 0.75, layout = "autofit")
        ft_m<-add_footer_lines(ft_m,paste("Post-hoc test used:", input$test_method))

        ft_cv<-flextable(LSD_A$statistics)
        ft_cv<- set_caption(ft_cv, caption=paste("ANOVA Stats -", response_col))
        ft_cv <- set_table_properties(ft_cv, width = 0.75, layout = "autofit")
        doc <- body_add_flextable(doc, ft)
        doc <- body_add_flextable(doc, ft_m)
        doc <- body_add_flextable(doc, ft_cv)
        doc <- body_add_gg(doc, main_plot, width = 6, height = 4)
        doc <- body_add_gg(doc, combined_plot, width = 4, height = 4)
        doc<-body_add_break(doc,pos = "after")

        # Create dynamic UI outputs
        output[[paste0("analysis_output_", i)]] <- DT::renderDT({
          print(aov_result)
        })
        output[[paste0("analysis_output_", i)]] <- renderPrint({
          cat("Post-hoc test used:", input$test_method, "\n\n")
          print(aov_result)
        })

        output[[paste0("plot_", i)]] <- renderPlot({
          main_plot
        })

        output[[paste0("combined_plot", i)]] <- renderPlot({
          res_fit_plot
        })
      }

      # Store the generated document in the reactive value
      report(doc)

      # Update UI with dynamic outputs
      output$analysis_outputs <- renderUI({
        lapply(seq_along(response_cols), function(i) {
          tagList(
            h3(paste("Analysis for", response_cols[i])),
            verbatimTextOutput(paste0("analysis_output_", i)),
            plotOutput(paste0("plot_", i)),
            plotOutput(paste0("combined_plot", i)),
            hr()
          )
        })
      })

      # Set analysis_complete to TRUE
      analysis_complete(TRUE)

      # Remove the "in progress" notification and show a completion notification
      removeNotification(id = "analysis_notification")
      showNotification("Analysis complete!", type = "message", duration = 20)
    })



    # Download Handler for the report
    output$download_report <- downloadHandler(
      filename = function() {
        paste0(input$output_filename, ".docx")
      },
      content = function(file) {
        req(report())
        print(report(), target = file)
      }
    )
  }

  shinyApp(ui, server)
  }
