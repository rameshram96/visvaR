library(shiny)
library(readxl)
library(ggplot2)
library(dplyr)
library(agricolae)
library(flextable)
library(officer)
library(tibble)
library(tidyr)
library(bslib)
library(patchwork)

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
  h1("One-way Completely Randomized Design (CRD)",
     style = "font-family:Times New Roman; font-weight: bold; font-size: 36px; color: #4CAF50;"),
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
          card_image(src="https://github.com/rameshram96/visvaR/blob/main/visvaRlogo.png",
                     style = "display: block; margin-left: auto; margin-right: auto;",
                     alt = "",
                     href = NULL,
                     border_radius = c("auto"),
                     mime_type = NULL,
                     class = NULL,
                     height ="50%",
                     fill = FALSE,
                     width = "50%",
                     container = NULL
          )
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
  data <- reactiveVal(NULL)
  report <- reactiveVal(NULL)
  analysis_complete <- reactiveVal(FALSE)  # Add this line
  # Add this to create the analysis status UI
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
    data(read_excel(input$file$datapath))
  })

  observeEvent(input$clipboard_button, {
    clipboard_data <- read.delim("clipboard", header = TRUE)
    data(clipboard_data)
  })

  output$preview <- DT::renderDT({
    req(data())
    data()
  })

  observeEvent(input$analyze_button, {
    req(data())

    # Reset the analysis_complete status
    analysis_complete(FALSE)

    # Show a notification that analysis has started
    showNotification("Analysis in progress...", type = "message", duration = NULL, id = "analysis_notification")

    data_full <- data()
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

      anova_model <- aov(Response ~ Factor_A , data = data1)
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

      ft <- flextable(aov_df)
      ft <- set_caption(ft, caption = paste("ANOVA Results -", response_col))
      ft <- add_footer_lines(ft, values = "Signif. codes:  0 '***' 0.001 '**' 0.01 '*' 0.05 '.' 0.1 ' ' 1")
      ft <- set_table_properties(ft, width = 0.9, layout = "autofit")

      colnames(ascend_A)[3] <- "Factor_A"
      combined_df <- merge(ascend_A, MeanSE_A, by = "Factor_A")
      ft_m <- flextable(combined_df)
      ft_m <- set_caption(ft_m, caption = paste("Mean Comparison -", response_col))
      ft_m <- set_table_properties(ft_m, width = 0.75, layout = "autofit")
      ft_m<-add_footer_lines(ft_m,paste("Post-hoc test used:", input$test_method))

      ft_cv<-flextable(LSD_A$statistics)
      ft_cv<- set_caption(ft_cv, captionpaste("ANOVA Stats -", response_col))
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
