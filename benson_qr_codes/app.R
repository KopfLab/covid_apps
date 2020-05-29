library(shiny)
library(shinyjs)
library(shinyBS)
library(DT)

library(readxl)
library(dplyr)
library(stringr)
library(lubridate)
library(googlesheets4)

source("utils.R")
source("module_selector_table.R")

# google sheets ids
auth_token <- "viewer_client_secret.json"
rooms_sheet_id <- "10m4I_pJuUz_w_ILQEoxLChJAHDUFTA5zC-rm0HHly4w"
logs_sheet_id <- "1jfj2u5hm-vIrq1oXiNEGgvEQKpsbovvCZ2iD6uZ3zMM"

# app
shinyApp(
  ui = fluidPage(
    h1("Generate QR Codes for Benson (BESC)"),
    h4("Select Rooms"),
    h4(style = "color: red;", textOutput("rooms_info")),
    selector_table_ui("rooms"),
    inline(
      actionButton("refresh_rooms", "Refresh Rooms", icon = icon("refresh")),
      spaces(1),
      selector_table_buttons_ui("rooms"),
      spaces(1),
      downloadButton("generate", "Generate QR Codes for Selected Rooms")
    )
    # h1("Logs (TZ = Denver)"),
    # DTOutput("logs"),
    # actionButton("refresh_logs", "Refresh Logs", icon = icon("refresh"))
  ),
  server = function(input, output, session) {
    
    # get rooms ======
    get_rooms <- reactive({
      input$refresh_rooms
      
      data <- withProgress(
        {
          # read only authentication
          googlesheets4::gs4_auth(
            scopes = "https://www.googleapis.com/auth/spreadsheets.readonly",
            path = auth_token
          )
          
          if (googlesheets4::gs4_has_token()) {
            data <- 
              tryCatch(
                googlesheets4::read_sheet(rooms_sheet_id),
                error = function(e) {
                  msg <- paste0("google sheet access failed: ", e$message)
                  warning(msg, immediate. = TRUE, call. = FALSE)
                  return(msg)
                }
              )
            return(data)
          } else {
            msg <- "google sheets authentication failed"
            warning(msg, immediate. = TRUE, call. = FALSE)
            return(msg)
          }
        }, 
        min = 0, max = 1, 0.5,
        message = "Loading building rooms..."
      )
      return(data)
    })
    
    # rooms selector table
    rooms <- callModule(selector_table_server, "rooms", id_column = "Room", initial_page_length = 10)
    output$rooms_info <- renderText({
      if (is.data.frame(get_rooms())) {
        ""
      } else if (is.character(get_rooms())) {
        get_rooms()
      } else {
        "cannot retrieve rooms"
      }
    })
    observe({
      req(is.data.frame(get_rooms()) && nrow(get_rooms()) > 0)
      rooms$set_table(get_rooms())
    })

    # download report
    output$generate <- downloadHandler(
      filename = reactive(paste0(paste(rooms$get_selected(), collapse = "_"), "_qr_codes.docx")),
      content = function(file) {
        temp_dir <- tempdir()
        temp_report <- file.path(temp_dir, "report.Rmd")
        file.copy("report.Rmd", temp_report, overwrite = TRUE)
        file.copy("template1.docx", file.path(temp_dir, "template1.docx"))

        # Set up parameters to pass to Rmd document
        params <- list(rooms = rooms$get_selected())

        # Knit the document, passing in the `params` list, and eval it in a
        # child of the global environment (this isolates the code in the document
        # from the code in this app).
        withProgress(
          rmarkdown::render(temp_report, output_file = file,
                            params = params,
                            envir = new.env(parent = globalenv())
          ), 
          min = 0, max = 1, 0.5,
          message = sprintf("Generating QR codes for %d rooms...", isolate(length(rooms$get_selected()))), 
          detail = isolate(paste(rooms$get_selected(), collapse = "_"))
        )
      }
    )
    
    # logs
    output$logs <- renderDT({
      input$refresh_logs
      isolate({
        req(get_rooms())
        withProgress(
          {
            # read only authentication
            googlesheets4::gs4_auth(
              scopes = "https://www.googleapis.com/auth/spreadsheets.readonly",
              path = auth_token
            )
            
            if (googlesheets4::gs4_has_token()) {
              data <- 
                tryCatch(
                  googlesheets4::read_sheet(logs_sheet_id, sheet = "Form Responses 1"),
                  error = function(e) {
                    stop("google sheet access failed: ", e$message, call. = FALSE)
                  }
                )
            } else {
              stop("google sheets authentication failed")
            }
            data %>% 
              mutate(Room = as.character(Room)) %>% 
              left_join(get_rooms(), by = "Room") %>% 
              select(When = Timestamp, Who = `Email Address`, Direction, Room, Floor, Category) %>% 
              arrange(desc(When)) %>% 
              mutate(When = When %>% lubridate::force_tz("America/Denver") %>% format("%b %d, %I:%M:%S %p"))
          }, 
          min = 0, max = 1, 0.5,
          message = "Refreshing logs..."
        )
      })
    }, 
    options = list()
    )
    
  }
)
