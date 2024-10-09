#' Run the Shiny App for performing one-way ANOVA on data from a Randomized block design
#'
#' App allows user to import data through excel/.csv files or through clipboard
#' and user can select the post-hoc test method and download the report which contains anova results and plots
#'
#' @details
#' This Shiny app is part of the `visvaR` package and is designed for
#' analysis of variance randomized block design (two factor) and user can download the report in word format.
#'
#' @return
#' This function runs a local instance of the Shiny app in your default web
#' browser. The app interface allows users to upload data, select analysis
#' method, and download outputs.
#'
#' @usage twoway_rbd()
#' @name twoway_rbd
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
utils::globalVariables(c("Fitted", "Residuals",'read_excel', "Sample", "Factor_A", "Response", "avg_AB", "se","Factor_B","LSD_AB$groups","response",'rownames(LSD_AB$groups)'))
twoway_rbd<-function(){
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
  h1("Two-Way Randomized Comlete Block Design (RCBD)",
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
  data <- reactiveVal(NULL)
  report <- reactiveVal(NULL)
  analysis_complete <- reactiveVal(FALSE)
  #create the analysis status UI
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


    analysis_complete(FALSE)

    showNotification("Analysis in progress...", type = "message", duration = NULL, id = "analysis_notification")

    data_full <- data()
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
      data1$Replication <- as.factor(data1$Replication)

      anova_model <- aov(Response ~ Factor_A*Factor_B + Replication, data = data1)
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
      ft <- add_footer_lines(ft, values = "Signif. codes:  0 '***' 0.001 '**' 0.01 '*' 0.05 '.' 0.1 ' ' 1")
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
