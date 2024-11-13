#'Run the Shiny App for performing two-way ANOVA on data from a completely randomized design.
#'
#' App allows user to import data through excel/.csv files or through clipboard
#' and user can select the post-hoc test method and download the report which contains anova results and plots
#'
#' @details
#' This Shiny app is part of the `visvaR` package and is designed for
#' analysis of variance on data from completely randomized design (two factor) and user can download the report in word format.
#' The analysis of variance was performed using R's aov() function (Chambers & Hastie, 1992; R Core Team, 2024), which implements the classical ANOVA methodology developed by Fisher (1925). To use custom fonts, please install the `extrafont` package and run `extrafont::font_import()` and `extrafont::loadfonts()`.
#' @return
#' This function runs a local instance of the Shiny app in your default web
#' browser. The app interface allows users to upload data, select analysis
#' method, and download outputs.
#'
#' @usage twoway_crd()
#' @aliases anova2_crd
#' @aliases twoway_crd
#' @examples
#'
#' # Example 1: Basic usage
#' if(interactive()) {
#'   twoway_crd()
#' }
#'
#' # Example 2: Sample workflow with factorial fertilizer experiment
#' if(interactive()) {
#'   # Prepare sample data with two factors: Nitrogen and Phosphorus levels
#'   fertilizer_data <- data.frame(
#'     Nitrogen = rep(c("N0", "N30", "N60"), each = 12),
#'     Phosphorus = rep(rep(c("P0", "P30", "P60", "P90"), each = 3), times = 3),
#'     Replication = rep(1:3, times = 12),
#'     Grain_Yield = c(
#'       4.2, 4.0, 4.1,  # N0-P0
#'       4.8, 4.6, 4.7,  # N0-P30
#'       5.1, 5.0, 5.2,  # N0-P60
#'       5.0, 4.9, 5.1,  # N0-P90
#'       5.5, 5.3, 5.4,  # N30-P0
#'       6.2, 6.0, 6.1,  # N30-P30
#'       6.8, 6.6, 6.7,  # N30-P60
#'       6.7, 6.5, 6.6,  # N30-P90
#'       6.0, 5.8, 5.9,  # N60-P0
#'       6.8, 6.6, 6.7,  # N60-P30
#'       7.5, 7.3, 7.4,  # N60-P60
#'       7.4, 7.2, 7.3   # N60-P90
#'     ),
#'     Protein_Content = c(
#'       9.0, 8.8, 8.9,   # N0-P0
#'       9.5, 9.3, 9.4,   # N0-P30
#'       9.8, 9.6, 9.7,   # N0-P60
#'       9.7, 9.5, 9.6,   # N0-P90
#'       10.5, 10.3, 10.4, # N30-P0
#'       11.2, 11.0, 11.1, # N30-P30
#'       11.8, 11.6, 11.7, # N30-P60
#'       11.7, 11.5, 11.6, # N30-P90
#'       12.0, 11.8, 11.9, # N60-P0
#'       12.8, 12.6, 12.7, # N60-P30
#'       13.5, 13.3, 13.4, # N60-P60
#'       13.4, 13.2, 13.3  # N60-P90
#'     )
#'   )
#'
#'   # Save as Excel file
#'   write.xlsx(fertilizer_data, "fertilizer_data.xlsx")
#'
#'   # Launch the app
#'   twoway_crd()
#'
#'   # Instructions for users:
#'   # 1. Click "Choose .xlsx or .csv" and select fertilizer_data.xlsx
#'   # 2. Select post-hoc test method (e.g., "LSD" or "Tukey HSD")
#'   # 3. Customize plot appearance:
#'   #    - Select font style
#'   #    - Adjust font size
#'   # 4. Set output filename
#'   # 5. Click "Analyze"
#'   # 6. View results in different tabs
#'   # 7. Download comprehensive Word report
#'
#'   # Clean up
#'   unlink("fertilizer_data.xlsx")
#' }
#'
#' # Example 3: Using clipboard data
#' if(interactive()) {
#'   # Copy this to clipboard:
#'   # Temperature,Light,Replication,Growth_Rate,Chlorophyll
#'   # Low,Dark,1,2.1,15.2
#'   # Low,Dark,2,2.0,15.0
#'   # Low,Dark,3,2.2,15.1
#'   # Low,Medium,1,2.8,16.5
#'   # Low,Medium,2,2.7,16.3
#'   # Low,Medium,3,2.9,16.4
#'   # Low,High,1,3.2,17.8
#'   # Low,High,2,3.1,17.6
#'   # Low,High,3,3.3,17.7
#'   # High,Dark,1,2.5,14.8
#'   # High,Dark,2,2.4,14.6
#'   # High,Dark,3,2.6,14.7
#'   # High,Medium,1,3.5,16.2
#'   # High,Medium,2,3.4,16.0
#'   # High,Medium,3,3.6,16.1
#'   # High,High,1,4.2,17.5
#'   # High,High,2,4.1,17.3
#'   # High,High,3,4.3,17.4
#'
#'   twoway_crd()
#'   # Click "Use Clipboard Data" after copying data
#' }
#'
#' # Example 4: Multiple response variables with variety trial
#' if(interactive()) {
#'   # Create data with multiple responses
#'   variety_trial <- data.frame(
#'     Irrigation = rep(c("Full", "Deficit"), each = 18),
#'     Variety = rep(rep(c("V1", "V2", "V3"), each = 6), times = 2),
#'     Replication = rep(1:6, times = 6),
#'     Yield = rnorm(36, mean = rep(c(5.5, 5.0, 4.8, 4.2, 3.8, 3.5), each = 6), sd = 0.2),
#'     WUE = rnorm(36, mean = rep(c(2.2, 2.4, 2.3, 2.8, 3.0, 2.9), each = 6), sd = 0.1),
#'     Biomass = rnorm(36, mean = rep(c(12.5, 11.8, 11.2, 10.2, 9.8, 9.5), each = 6), sd = 0.5)
#'   )
#'
#'   # Save as Excel file
#'   write.xlsx(variety_trial, "variety_trial.xlsx")
#'
#'   # Launch the app
#'   twoway_crd()
#'
#'   # Clean up
#'   unlink("variety_trial.xlsx")
#' }
#'
#' @references Fisher, R. A. (1925). Statistical Methods for Research Workers. Oliver and Boyd, Edinburgh.
#'             Scheffe, H. (1959). The Analysis of Variance. John Wiley & Sons, New York.
#'             R Core Team (2024). R: A language and environment for statistical computing. R Foundation for Statistical Computing, Vienna, Austria. URL https://www.R-project.org/
#' @name twoway_crd
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
utils::globalVariables(c("Fitted", "Residuals", "Sample", "Factor_A", "Response", "avg_AB", "se","Factor_B","LSD_AB$groups","response",'read_excel','rownames(LSD_AB$groups)'))
#' @export
twoway_crd<-function(){
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
    h1("Two-Way Completely Randomized Design (CRD)",
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
            tabPanel("OUTPUT",
                     uiOutput("analysis_outputs")
            )
          ),
          uiOutput("analysis_status")
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
             tags$strong("Analysis complete!", style = "color: green; font-size: 45px;")
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


      analysis_complete(FALSE)

      showNotification("Analysis in progress...", type = "message", duration = NULL, id = "analysis_notification")

      data_full <- data_reactive()
      factor_a <- colnames(data_full)[1]
      factor_b <- colnames(data_full)[2]
      rep_col <- colnames(data_full)[3]
      response_cols <- colnames(data_full)[4:ncol(data_full)]

      doc <- read_docx()

      for (i in seq_along(response_cols)) {
        response_col <- response_cols[i]

        data1 <- data_full %>%
          select(all_of(c(factor_a,factor_b, rep_col, response_col))) %>%
          rename(Factor_A = 1,Factor_B = 2, Replication = 3, Response = 4)

        data1$Factor_A <- as.factor(data1$Factor_A)
        data1$Factor_B <- as.factor(data1$Factor_B)
        anova_model <- aov(Response ~ Factor_A*Factor_B , data = data1)
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
                        console = TRUE)
        }


        LSD_A <- perform_test(data1$Response, data1$Factor_A, anova_model$df.residual, deviance(anova_model)/anova_model$df.residual)
        LSD_B <- perform_test(data1$Response, data1$Factor_B, anova_model$df.residual, deviance(anova_model)/anova_model$df.residual)
        LSD_AB <- perform_test(data1$Response, data1$Factor_A:data1$Factor_B, anova_model$df.residual, deviance(anova_model)/anova_model$df.residual)


        ascend_AB <- LSD_AB$groups %>%
          group_by(rownames(LSD_AB$groups)) %>%
          arrange(rownames(LSD_AB$groups))
        MeanSE_AB <- data1 %>%
          group_by(Factor_A,Factor_B) %>%
          summarise(avg_AB = mean(Response),
                    se = sd(Response)/sqrt(length(Response)))

        main_plot <- ggplot(MeanSE_AB, aes(x = Factor_A, y = avg_AB, fill = Factor_B)) +
          geom_bar(stat = "identity", color = "black",
                   width = 0.5, position = position_dodge(width=0.5)) +
          geom_errorbar(aes(ymax = avg_AB + se, ymin = avg_AB - se),
                        position = position_dodge(width=0.5), width = 0.1) +
          theme(panel.background = element_rect(fill="white", colour="white", linetype="solid", color="white"),
                legend.position = "top",
                axis.text.x = element_text(angle = 0, face = "bold"),
                panel.border = element_blank(),
                panel.grid.major = element_blank(),
                panel.grid.minor = element_blank(),
                axis.line = element_line(linewidth = 0.5, linetype = "solid", colour = "black"),
                text = element_text(family = input$font_family, size = input$font_size)) +
          labs(title = paste("", response_col), x = factor_a, y = response_col,color = "factor_b") +
          geom_text(aes(x = Factor_A, y = avg_AB + se, label = as.matrix(ascend_AB$groups)),
                    position = position_dodge(width = 0.5), vjust = -(0.5)) +
          scale_y_continuous(expand = c(0,0))+
          expand_limits(y = c(0, max(MeanSE_AB$avg_AB+ (MeanSE_AB$avg_AB*0.25))))+
          guides(fill=guide_legend(title=factor_b))+
          labs(color = NULL)
        oldpar <- par(no.readonly = TRUE)
        on.exit(par(oldpar))

        #Word document content
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
          "visvaR - Visualize Variance                                                                                         ",
          run_footnote(x = bl, prop = fp_refnote),
          "")

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

        ascend_AB_clean <- ascend_AB %>%
          separate(`rownames(LSD_AB$groups)`, into = c("Factor_A", "Factor_B"), sep = ":") %>%
          rename(avg_AB = `response`)

        colnames(ascend_AB)[3] <- "Factor_A"
        combined_df <- ascend_AB_clean %>%
          left_join(MeanSE_AB, by = c("Factor_A", "Factor_B", "avg_AB")) %>%
          select(Factor_A, Factor_B, avg_AB, se, groups)
        ft_m <- flextable(combined_df)
        ft_m <- set_caption(ft_m, caption = paste("Mean Comparison Interaction Effect -", response_col))
        ft_m <- set_table_properties(ft_m, width = 0.75, layout = "autofit")
        ft_m<-add_footer_lines(ft_m,paste("Post-hoc test used:", input$test_method))

        ft_cv<-flextable(LSD_A$statistics)
        ft_cv<- set_caption(ft_cv, caption = paste("Statistics-", response_col))
        ft_cv <- set_table_properties(ft_cv, width = 0.75, layout = "autofit")

        lsd_a <- data.frame(factor = rownames(LSD_A$groups), LSD_A$groups)
        ft_a <- flextable(lsd_a)
        ft_a <- set_caption( ft_a, caption=paste("Factor A Mean Comparision -", response_col))
        ft_a <- set_table_properties( ft_a, width = 0.75, layout = "autofit")
        lsd_b <- data.frame(factor = rownames(LSD_B$groups), LSD_B$groups)
        ft_b <- flextable(lsd_b)
        ft_b <- set_caption( ft_b, caption = paste("Factor B Mean Comparision -", response_col))
        ft_b <- set_table_properties( ft_b, width = 0.75, layout = "autofit")



        doc <- body_add_flextable(doc, ft)
        doc <- body_add_flextable(doc, ft_m)
        doc <- body_add_flextable(doc, ft_cv)
        doc <- body_add_flextable(doc, ft_a)
        doc <- body_add_flextable(doc, ft_b)
        doc <- body_add_gg(doc, main_plot, width = 6, height = 4)
        doc <- body_add_gg(doc, combined_plot, width = 4, height = 4)
        doc<-body_add_break(doc,pos = "after")

        #dynamic UI outputs
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

      # UI with dynamic outputs
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

      analysis_complete(TRUE)

      #Remove the "in progress" notification and show a completion notification
      removeNotification(id = "analysis_notification")
      showNotification("Analysis complete!", type = "message", duration = 30)
    })



    # downloadHandler for the report
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
