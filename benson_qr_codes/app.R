library(rlang)
library(dplyr)
library(tidyr)
library(purrr)
library(stringr)
library(readxl)
library(officer)
library(googlesheets4)

library(shiny)
library(shinyjs)
library(shinyBS)
library(DT)

source("qr_code_funcs.R")
source("log_funcs.R")
source("shiny_utils.R")
source("shiny_module_selector_table.R")

# google sheets ids
auth_token <- "viewer_client_secret.json"
rooms_sheet_id <- "10m4I_pJuUz_w_ILQEoxLChJAHDUFTA5zC-rm0HHly4w"
logs_sheet_id <- "1jfj2u5hm-vIrq1oXiNEGgvEQKpsbovvCZ2iD6uZ3zMM"
form_url <- "https://docs.google.com/forms/d/e/1FAIpQLSdjnHpnkhpCF6vn8RS_X48pXG4Y3U3DQP_sNo1Yctw56FcmuQ"
room_field_id <- "entry.28605957"
direction_field_id <- "entry.1525768046"

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
      selector_table_buttons_ui("rooms")
    ),
    h4(downloadButton("generate", "Generate QR Codes for Selected Rooms", style = "color: #03a329; font-weight: bold;")),
    h1("Recent Logs (TZ = Denver)"),
    DTOutput("logs"),
    h4(actionButton("refresh_logs", "Refresh Logs", icon = icon("refresh")))
  ),
  server = function(input, output, session) {
    
    # get rooms ======
    get_rooms <- reactive({
      input$refresh_rooms
      
      data <- withProgress(
        read_gs_sheet(sheet_id = rooms_sheet_id, auth_token = auth_token), 
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

    # download QR codes =====
    output$generate <- downloadHandler(
      filename = reactive(paste0(paste(rooms$get_selected(), collapse = "_"), "_qr_codes.docx")),
      content = function(file) {
        withProgress(
          generate_qr_codes_doc(
            template_path = "benson_template.docx",
            rooms = rooms$get_selected(), 
            form_url = form_url, 
            room_field_id = room_field_id, 
            direction_field_id = direction_field_id,
            data = rooms$get_selected_items(),
            save_path = file
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
            logs <- read_gs_sheet(sheet = logs_sheet_id, auth_token = auth_token)
            print(logs)
            logs %>% 
              arrange(desc(Timestamp)) %>% 
              filter(dplyr::row_number() <= 10) %>% 
              mutate(Room = as.character(Room)) %>% 
              left_join(get_rooms(), by = "Room") %>% 
              select(When = Timestamp, Direction, Room, Floor, Category) %>% 
              mutate(When = When %>% lubridate::force_tz("America/Denver") %>% format("%b %d, %I:%M:%S %p"))
          }, 
          min = 0, max = 1, 0.5,
          message = "Refreshing logs..."
        )
      })
    }, 
    options = list(
      dom = "t"
    ))
    
  }
)
